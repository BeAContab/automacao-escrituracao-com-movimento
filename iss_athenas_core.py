"""ISS Fortaleza → Athenas: converte a escrituração exportada do ISS Fortaleza
(abas "Serviços Tomados" e "Serviços Prestados") para o layout de importação do
sistema Athenas ERP.

Migrado do projeto standalone `importar/Import ISS Fortaleza Athenas/iss_para_athenas.py`
(ver `MIGRACAO.md` na mesma pasta). A lógica de negócio abaixo (constantes, helpers,
`processar_aba`, `salvar_xlsx`) foi trazida **byte a byte igual** ao original — só a
interface Tkinter ficou para trás, substituída por `processar_arquivo`/`salvar_log_erros`,
que a GUI desta aplicação chama em loop.

ATENÇÃO: se a Prefeitura mudar a ordem das colunas do arquivo exportado pelo ISS
Fortaleza, os índices COL_* abaixo (posicionais, 0-based) quebram silenciosamente —
ver a tabela de referência no `MIGRACAO.md`, seção 7.
"""

from __future__ import annotations

import calendar
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any, Callable

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

# ─── Índices das colunas no arquivo ISS Fortaleza (ambas as abas) ───────────
COL_NUMERO      = 1
COL_COMPETENCIA = 3
COL_DATA        = 4
COL_ITEM_LISTA  = 13
COL_VALOR       = 19
COL_RET_FED     = 23
COL_PIS         = 25   # Presente quando ISS fornece breakdown individual
COL_COFINS      = 26   # Presente quando ISS fornece breakdown individual
COL_IRRF        = 27
COL_CSLL        = 28   # No ISS: às vezes individual, às vezes PIS+COFINS+CSLL somados
COL_INSS        = 29
COL_ISS_RETIDO  = 33
COL_VALOR_ISS   = 34
COL_STATUS      = 36
COL_CNPJ        = 38   # CPF/CNPJ Prestador (tomados) ou Tomador (prestados)
COL_RAZAO       = 39   # Razão Social Prestador (tomados) ou Tomador (prestados)

# ─── Alíquotas fixas para cálculo dos impostos federais ────────────────────
TAXA_PIS    = 0.0065
TAXA_COFINS = 0.03
TAXA_IRRF   = 0.015
TAXA_CSLL   = 0.01

# ─── Cabeçalhos da planilha Athenas ────────────────────────────────────────
_HDR_BASE = [
    'Número', 'Data', 'Data Registro', 'CFOP', 'Item da Lista',
    'Cod Ref Produto', 'Valor dos Serviços', 'PIS', 'COFINS', 'IRRF',
    'CSLL', 'INSS', 'Valor do ISS Retido',
]
HDR_TOMADOS   = _HDR_BASE + ['CPF/CNPJ Prestador', 'Razão Social/Nome do Prestador',  'Situação', 'Observação', 'Valor Líquido']
# Prestados: 'Valor do ISS Próprio' (ISS não retido, para importação) e
# 'Cod Cliente' (=1 quando CPF/CNPJ do Tomador não vem informado)
HDR_PRESTADOS = _HDR_BASE + ['Valor do ISS Próprio', 'Cod Cliente', 'CPF/CNPJ Tomador', 'Razão Social/Nome do Tomador', 'Situação', 'Observação', 'Valor Líquido']

# Nomes das colunas com formato de moeda e das colunas de impostos (highlight amarelo).
# Localizados por nome no cabeçalho — Tomados e Prestados têm layouts diferentes.
NOMES_MOEDA    = {'Valor dos Serviços', 'PIS', 'COFINS', 'IRRF', 'CSLL', 'INSS', 'Valor do ISS Retido', 'Valor do ISS Próprio', 'Valor Líquido'}
NOMES_IMPOSTOS = {'PIS', 'COFINS', 'IRRF', 'CSLL'}
FMT_MOEDA      = '#,##0.00'


# ─── Funções de transformação ───────────────────────────────────────────────

def to_float(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def nz(raw):
    """Retorna o valor float se > 0, caso contrário 0.0 (trata nan e zero como ausente)."""
    if pd.isna(raw):
        return 0.0
    v = to_float(raw)
    return v if v > 0 else 0.0


def fmt_date(v):
    """Retorna string DD/MM/YYYY a partir de datetime ou string."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    if hasattr(v, 'strftime'):
        return v.strftime('%d/%m/%Y')
    s = str(v).strip()
    # Já no formato DD/MM/YYYY
    if re.match(r'\d{2}/\d{2}/\d{4}', s):
        return s
    return s


def formatar_cnpj_cpf(v):
    """Converte valor numérico para string com zeros à esquerda (11 CPF, 14 CNPJ)."""
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return ''
    s = str(v).strip()
    # Remove casas decimais indesejadas (ex: '50336159000150.0')
    if s.endswith('.0'):
        s = s[:-2]
    # Remove formatação (pontos, traços, barras)
    s = re.sub(r'[.\-/]', '', s)
    try:
        n = int(float(s)) if '.' in str(v) else int(s)
        digits = str(n)
        if len(digits) <= 11:
            return digits.zfill(11)   # CPF
        return digits.zfill(14)        # CNPJ
    except (ValueError, OverflowError):
        return s


def item_lista_int(valor):
    """'31.01' → 3101  |  '1.05' → 105  |  17.1 (float de 17.10) → 1710"""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return ''
    # Quando pandas lê 17.10 do Excel vira float 17.1 (perde o zero).
    # round(float * 100) recupera o código correto independente do zero final.
    if isinstance(valor, float):
        return int(round(valor * 100))
    s = re.sub(r'[.\s]', '', str(valor).split()[0])
    try:
        return int(s)
    except ValueError:
        return str(valor)


def get_competencia(df_vals):
    """Extrai MM/AAAA da coluna Competência para nomear o arquivo."""
    for v in df_vals:
        s = str(v).strip()
        if re.match(r'\d{2}/\d{4}', s):
            return s.replace('/', '-')
    return 'competencia'


def ultimo_dia_competencia(competencia_str):
    """'04/2026' → '30/04/2026'"""
    m = re.match(r'(\d{2})/(\d{4})', str(competencia_str).strip())
    if not m:
        return None
    mes, ano = int(m.group(1)), int(m.group(2))
    ultimo = calendar.monthrange(ano, mes)[1]
    return date(ano, mes, ultimo).strftime('%d/%m/%Y')


def ajustar_data(data_raw, competencia_str):
    """
    Se a Data da NF for de mês diferente da Competência,
    retorna o último dia da Competência.
    """
    data_fmt = fmt_date(data_raw)
    comp_ultimo = ultimo_dia_competencia(competencia_str)
    if not data_fmt or not comp_ultimo:
        return data_fmt
    # Extrai MM/AAAA da data da NF
    m = re.match(r'\d{2}/(\d{2}/\d{4})', data_fmt)
    if not m:
        return data_fmt
    data_mes_ano = m.group(1)          # MM/YYYY da NF
    comp_mes_ano = competencia_str.strip()  # MM/YYYY da competência
    if data_mes_ano != comp_mes_ano:
        return comp_ultimo
    return data_fmt


def encontrar_col_valor_liq(df):
    """Localiza a coluna Valor Líquido pelo cabeçalho (busca por 'l' + 'quido')."""
    for i, col in enumerate(df.columns):
        nome = str(col).lower()
        if 'quido' in nome or 'quida' in nome:
            return i
    return None


def processar_aba(df, tipo, regime=None, col_valor_liq=None):
    """
    Processa Tomados ou Prestados.
    Retorna lista de dicts:
      { 'valores': [...18 campos...], 'cancelada': bool, 'amarelo': bool }
    tipo          : 'tomados' ou 'prestados'
    regime        : 'normal' ou 'simples'  (apenas para prestados)
    col_valor_liq : índice (0-based) da coluna Valor Líquido no ISS, ou None
    """
    linhas = []
    arr = df.values

    for row in arr:
        status_raw = str(row[COL_STATUS]).strip().upper()
        if status_raw in ('NAN', ''):
            status_raw = 'NORMAL'
        cancelada = status_raw not in ('NORMAL',)

        competencia  = str(row[COL_COMPETENCIA]).strip()
        numero       = row[COL_NUMERO]
        data_str     = ajustar_data(row[COL_DATA], competencia)
        valor        = to_float(row[COL_VALOR])
        item         = item_lista_int(row[COL_ITEM_LISTA])
        ret_fed      = to_float(row[COL_RET_FED])
        irrf_raw     = row[COL_IRRF]
        csll_raw     = row[COL_CSLL]
        pis_raw      = row[COL_PIS]
        cofins_raw   = row[COL_COFINS]
        inss_v       = to_float(row[COL_INSS])
        iss_ret      = str(row[COL_ISS_RETIDO]).strip().lower() == 'sim'
        val_iss      = to_float(row[COL_VALOR_ISS])
        cnpj         = formatar_cnpj_cpf(row[COL_CNPJ])
        razao        = str(row[COL_RAZAO]).strip() if pd.notna(row[COL_RAZAO]) else ''

        tem_ret = (ret_fed > 0) or iss_ret or (inss_v > 0)

        # CFOP
        if tipo == 'tomados':
            cfop = 8013 if tem_ret else 8012
        else:
            if regime == 'simples':
                cfop = 8003 if tem_ret else 8001
            else:
                cfop = 8011 if tem_ret else 8010

        # ── Cálculo dos impostos federais ────────────────────────────────────
        # CSLL = 4,65% do valor → é o lump (PIS+COFINS+CSLL agrupados) → desmembrar.
        # CSLL + PIS e COFINS preenchidos → usar os três direto e validar percentuais.
        # CSLL sozinho fora de 1% → também é lump, porém sobre base menor que o
        #   valor dos serviços → derivar a base pelo próprio CSLL e desmembrar.
        # PIS e COFINS sem CSLL → nunca retidos, ignorar com aviso.
        # IRRF → somente da coluna, nunca calculado.
        obs     = ''
        amarelo = False
        pis = cofins = irrf = csll = None

        if ret_fed > 0:
            pis_col    = nz(pis_raw)
            cofins_col = nz(cofins_raw)
            irrf_col   = nz(irrf_raw)
            csll_col   = nz(csll_raw)

            # IRRF: sempre direto da coluna
            irrf = irrf_col or None

            lump_esp     = round(valor * 0.0465, 2)
            csll_esp     = round(valor * TAXA_CSLL, 2)
            eh_lump      = csll_col > 0 and abs(csll_col - lump_esp) < max(0.10, lump_esp * 0.02)
            csll_sozinho = csll_col > 0 and pis_col == 0 and cofins_col == 0
            eh_1pct      = csll_col > 0 and abs(csll_col - csll_esp) < max(0.10, csll_esp * 0.02)
            usou_colunas = False

            if eh_lump:
                # CSLL agrupa PIS+COFINS+CSLL sobre o valor dos serviços
                pis    = round(valor * TAXA_PIS,    5)
                cofins = round(valor * TAXA_COFINS, 5)
                csll   = round(valor * TAXA_CSLL,   5)
            elif csll_sozinho and not eh_1pct:
                # Lump sobre base reduzida (nem todo o serviço sofre retenção federal).
                # Deriva a base pelo próprio CSLL e aplica as alíquotas individuais.
                base   = csll_col / 0.0465
                pis    = round(base * TAXA_PIS,    5)
                cofins = round(base * TAXA_COFINS, 5)
                csll   = round(base * TAXA_CSLL,   5)
            elif csll_col > 0 and pis_col > 0 and cofins_col > 0:
                # Os três vieram individualizados → usar direto e validar percentuais
                pis    = pis_col
                cofins = cofins_col
                csll   = csll_col
                usou_colunas = True
                divergencias = []
                for nome, recebido, taxa in [
                    ('PIS',    pis_col,    TAXA_PIS),
                    ('COFINS', cofins_col, TAXA_COFINS),
                    ('CSLL',   csll_col,   TAXA_CSLL),
                ]:
                    esperado = round(valor * taxa, 2)
                    if esperado > 0 and abs(recebido - esperado) > max(0.10, esperado * 0.02):
                        divergencias.append(
                            f'{nome}: recebido {recebido:.2f}, esperado {esperado:.2f} ({taxa*100:.2f}%)'
                        )
                if divergencias:
                    obs = 'Percentual divergente — ' + ' | '.join(divergencias)
            elif csll_col > 0:
                csll = csll_col

            # PIS e COFINS das colunas quando não foram usados como retenção:
            # sinalizar na observação e, se as células ficaram vazias, marcar amarelo.
            if (pis_col > 0 or cofins_col > 0) and not usou_colunas:
                partes = []
                if pis_col > 0:    partes.append(f'PIS={pis_col:.2f}')
                if cofins_col > 0: partes.append(f'COFINS={cofins_col:.2f}')
                obs = ('No arquivo de origem veio ' + ' e '.join(partes) +
                       ' mas por regra nao sao retidos e foram desconsiderados.')
                if pis is None:
                    # Células PIS/COFINS ficaram em branco → amarelo
                    amarelo = True

        inss_out        = inss_v if inss_v > 0 else None
        iss_ret_out     = val_iss if iss_ret else None
        iss_proprio_out = val_iss if (not iss_ret and val_iss > 0) else None

        val_liq = None
        if col_valor_liq is not None:
            v = row[col_valor_liq]
            val_liq = to_float(v) if pd.notna(v) else None

        campos_base = [
            numero, data_str, data_str, cfop, item, item, valor,
            pis, cofins, irrf, csll, inss_out, iss_ret_out,
        ]
        if tipo == 'prestados':
            campos_base.append(iss_proprio_out)
            # Cod Cliente = 1 (Clientes Diversos no Athenas) somente quando
            # o CPF/CNPJ do Tomador não vier informado no ISS
            cod_cliente = 1 if cnpj == '' else None
            campos_base.append(cod_cliente)
        campos_base += [cnpj, razao, status_raw, obs, val_liq]

        linhas.append({
            'valores': campos_base,
            'cancelada': cancelada,
            'amarelo':   amarelo,
        })

    return linhas


LARGURA_POR_NOME = {
    'Número': 15, 'Data': 12, 'Data Registro': 12, 'CFOP': 8, 'Item da Lista': 13,
    'Cod Ref Produto': 13, 'Valor dos Serviços': 16, 'PIS': 10, 'COFINS': 10,
    'IRRF': 10, 'CSLL': 10, 'INSS': 10, 'Valor do ISS Retido': 16,
    'Valor do ISS Próprio': 18, 'Cod Cliente': 12,
    'CPF/CNPJ Tomador': 22, 'CPF/CNPJ Prestador': 22,
    'Razão Social/Nome do Tomador': 45, 'Razão Social/Nome do Prestador': 45,
    'Situação': 12, 'Observação': 60, 'Valor Líquido': 16,
}


def salvar_xlsx(linhas, cabecalho, caminho):
    wb = Workbook()
    ws = wb.active
    ws.title = 'Importação'

    cor_cab      = PatternFill('solid', start_color='1F4E79')
    cor_vermelho = PatternFill('solid', start_color='FF0000')
    cor_amarelo  = PatternFill('solid', start_color='FFFF00')
    fonte_cab    = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    fonte_dado   = Font(name='Arial', size=10)
    fonte_cancel = Font(name='Arial', size=10, color='FFFFFF')
    alinhar_c    = Alignment(horizontal='center', vertical='center')

    # Localiza colunas por nome — layout difere entre Tomados e Prestados
    col_obs       = cabecalho.index('Observação') + 1
    cols_impostos = [i + 1 for i, h in enumerate(cabecalho) if h in NOMES_IMPOSTOS]
    cols_moeda    = [i + 1 for i, h in enumerate(cabecalho) if h in NOMES_MOEDA]
    cols_amarelo  = cols_impostos + [col_obs]

    for ci, h in enumerate(cabecalho, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill      = cor_cab
        c.font      = fonte_cab
        c.alignment = alinhar_c

    for ri, item in enumerate(linhas, 2):
        if isinstance(item, dict):
            valores   = item['valores']
            cancelada = item.get('cancelada', False)
            amarelo   = item.get('amarelo', False)
        else:
            valores   = item
            cancelada = False
            amarelo   = False

        for ci, val in enumerate(valores, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            if cancelada:
                c.fill = cor_vermelho
                c.font = fonte_cancel
            else:
                c.font = fonte_dado
                if amarelo and ci in cols_amarelo:
                    c.fill = cor_amarelo
            if ci in (2, 3):
                c.alignment = Alignment(horizontal='center')
            if ci in cols_moeda and val is not None:
                c.number_format = FMT_MOEDA

    for ci, h in enumerate(cabecalho, 1):
        ws.column_dimensions[get_column_letter(ci)].width = LARGURA_POR_NOME.get(h, 14)

    ws.freeze_panes = 'A2'
    wb.save(caminho)


# ---------------------------------------------------------------------------
# Orquestração por arquivo — código novo desta migração (não existia no
# standalone, que fazia isto direto em `App._gerar`, misturado com Tkinter).
# Ver MIGRACAO.md, §5.4.
# ---------------------------------------------------------------------------

OPCOES_REGIME = {"normal": "Regime Normal (8010/8011)", "simples": "Simples Nacional (8001/8003)"}


def processar_arquivo(
    caminho_iss: str | Path,
    regime: str,
    pasta_saida: str | Path,
    log: Callable[[str], None] = print,
) -> dict[str, int]:
    """Converte um arquivo do ISS Fortaleza para o(s) arquivo(s) Athenas correspondente(s).

    `regime`: 'normal' ou 'simples' (só é considerado na aba Prestados).
    Devolve {'tomados': N, 'prestados': N} com a quantidade de linhas geradas em
    cada aba (0 se a aba não existir no arquivo de entrada). Levanta `ValueError`
    se o arquivo não tiver nenhuma aba de Tomados nem de Prestados.
    """
    caminho_iss = Path(caminho_iss)
    pasta_saida = Path(pasta_saida)
    base = caminho_iss.stem

    xl = pd.ExcelFile(caminho_iss)
    aba_tomados = next((s for s in xl.sheet_names if 'tomado' in s.lower()), None)
    aba_prestados = next((s for s in xl.sheet_names if 'prestado' in s.lower()), None)
    if not aba_tomados and not aba_prestados:
        raise ValueError(
            f"Nenhuma aba de Tomados ou Prestados encontrada. Abas presentes: {', '.join(xl.sheet_names)}"
        )

    n_tomados = n_prestados = 0

    if aba_tomados:
        df_tom = pd.read_excel(caminho_iss, sheet_name=aba_tomados, header=0)
        linhas_tom = processar_aba(df_tom, 'tomados', col_valor_liq=encontrar_col_valor_liq(df_tom))
        caminho_saida_tom = pasta_saida / f'Athenas_Tomados_{base}.xlsx'
        salvar_xlsx(linhas_tom, HDR_TOMADOS, caminho_saida_tom)
        n_tomados = len(linhas_tom)
        log(f'  {n_tomados} nota(s) de Serviços Tomados exportada(s) para: {caminho_saida_tom.name}')
    else:
        log('  Sem aba Serviços Tomados — ignorado.')

    if aba_prestados:
        df_prest = pd.read_excel(caminho_iss, sheet_name=aba_prestados, header=0)
        linhas_prest = processar_aba(df_prest, 'prestados', regime, col_valor_liq=encontrar_col_valor_liq(df_prest))
        caminho_saida_prest = pasta_saida / f'Athenas_Prestados_{base}.xlsx'
        salvar_xlsx(linhas_prest, HDR_PRESTADOS, caminho_saida_prest)
        n_prestados = len(linhas_prest)
        reg_txt = 'Simples Nacional' if regime == 'simples' else 'Regime Normal'
        log(f'  {n_prestados} nota(s) de Serviços Prestados exportada(s) ({reg_txt}) para: {caminho_saida_prest.name}')
    else:
        log('  Sem aba Serviços Prestados — ignorado.')

    return {'tomados': n_tomados, 'prestados': n_prestados}


def salvar_log_erros(log_erros: list[dict[str, str]], pasta_saida: str | Path) -> Path:
    """Grava `Log_Erros_ISS_Athenas.xlsx` com uma linha por arquivo que falhou.

    Cada item de `log_erros` é um dict com as chaves: Arquivo, Caminho, Regime, Erro, Detalhes.
    """
    pasta_saida = Path(pasta_saida)
    wb = Workbook()
    ws = wb.active
    ws.title = 'Erros'

    cabecalho = ['Data/Hora', 'Arquivo', 'Caminho Completo', 'Regime', 'Erro', 'Detalhes (Traceback)']
    cor_err = PatternFill('solid', start_color='C00000')
    fonte_cab = Font(bold=True, color='FFFFFF', name='Arial', size=10)
    fonte_dado = Font(name='Arial', size=9)
    agora = datetime.now().strftime('%d/%m/%Y %H:%M:%S')

    for ci, h in enumerate(cabecalho, 1):
        c = ws.cell(row=1, column=ci, value=h)
        c.fill = cor_err
        c.font = fonte_cab
        c.alignment = Alignment(horizontal='center', vertical='center')

    for ri, err in enumerate(log_erros, 2):
        dados = [agora, err['Arquivo'], err['Caminho'], err['Regime'], err['Erro'], err['Detalhes']]
        for ci, val in enumerate(dados, 1):
            c = ws.cell(row=ri, column=ci, value=val)
            c.font = fonte_dado
            c.alignment = Alignment(wrap_text=(ci == 6), vertical='top')

    ws.column_dimensions['A'].width = 18
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 55
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 50
    ws.column_dimensions['F'].width = 80
    ws.freeze_panes = 'A2'

    caminho_log = pasta_saida / 'Log_Erros_ISS_Athenas.xlsx'
    wb.save(caminho_log)
    return caminho_log
