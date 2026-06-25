from __future__ import annotations

import argparse
import sys
import re
import unicodedata
from dataclasses import dataclass
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Iterable

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

try:
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - fallback local
    fitz = None

from PyPDF2 import PdfReader

from iss_fortaleza_automacao import (
    executar_automacao_iss,
    interpretar_competencia,
    solicitar_competencia,
)
from cnae_final import enriquecer_registros_cnae_final
from gemini_extracao import extrair_campos_gemini, ErroCotaGemini
from tratamento_erros import registrar_erro, registrar_evento_execucao

@dataclass
class NotaFiscalExtraida:
    arquivo_pdf: str
    cnpj_prestador: str = ""
    numero_nf: str = ""
    data_emissao: str = ""
    id_cnae: str = ""
    desc_cnae: str = ""
    descricao_servico: str = ""
    uf_local_prestacao: str = ""
    cidade_local_prestacao: str = ""
    natureza_operacao: str = ""
    iss_retido: str = ""
    valor_servico: str = ""
    aliquota: str = ""
    id_cnae_final: str = ""
    desc_cnae_final: str = ""
    # Novos campos
    prefeitura: str = ""
    valor_deducoes: str = "0,00"
    descontos_incondicionados: str = "0,00"
    descontos_condicionados: str = "0,00"
    outras_retencoes: str = "0,00"
    ir: str = "0,00"
    pis_nao_retido: str = "0,00"
    cofins_nao_retido: str = "0,00"
    csrf: str = "0,00"
    inss: str = "0,00"
    analisado_pela_ia: str = "NÃO"
    revisao_manual: str = ""

    def como_linha(self) -> list[str]:
        return [
            self.arquivo_pdf,
            self.prefeitura,
            self.cnpj_prestador,
            self.numero_nf,
            self.data_emissao,
            self.id_cnae,
            self.desc_cnae,
            self.descricao_servico,
            self.uf_local_prestacao,
            self.cidade_local_prestacao,
            self.natureza_operacao,
            self.iss_retido,
            self.valor_servico,
            self.valor_deducoes,
            self.descontos_incondicionados,
            self.descontos_condicionados,
            self.outras_retencoes,
            self.ir,
            self.pis_nao_retido,
            self.cofins_nao_retido,
            self.csrf,
            self.inss,
            self.aliquota,
            self.id_cnae_final,
            self.desc_cnae_final,
            self.analisado_pela_ia,
            self.revisao_manual,
        ]

    def campos_vazios(self) -> list[str]:
        """Lista as colunas que não foram preenchidas para esta NF."""
        campos = {
            "PREFEITURA": self.prefeitura,
            "CNPJ_PRESTADOR": self.cnpj_prestador,
            "NUMERO_NF": self.numero_nf,
            "DATA_EMISSAO": self.data_emissao,
            "ID_CNAE": self.id_cnae,
            "DESC_CNAE": self.desc_cnae,
            "DESCRICAO_SERVICO": self.descricao_servico,
            "UF_LOCAL_PRESTACAO": self.uf_local_prestacao,
            "CIDADE_LOCAL_PRESTACAO": self.cidade_local_prestacao,
            "NATUREZA_OPERACAO": self.natureza_operacao,
            "ISS_RETIDO": self.iss_retido,
            "VALOR_SERVICO": self.valor_servico,
            "ALIQUOTA": self.aliquota,
        }
        return [nome for nome, valor in campos.items() if not str(valor).strip()]

    def resumo_campos(self) -> str:
        """Devolve um resumo curto para facilitar o diagnóstico no log."""
        return (
            f"Prefeitura={self.prefeitura or 'vazio'} | "
            f"CNPJ={self.cnpj_prestador or 'vazio'} | "
            f"NF={self.numero_nf or 'vazio'} | "
            f"Data={self.data_emissao or 'vazio'} | "
            f"Valor={self.valor_servico or 'vazio'}"
        )


def normalizar_texto(texto: str) -> str:
    """Consolida espaços para facilitar a extração por expressão regular."""
    return " ".join(texto.split())


def texto_para_busca(texto: str) -> str:
    """Remove acentos mantendo o tamanho do texto para preservar os índices.

    Os PDFs das três prefeituras usam rótulos parecidos, mas com acentos,
    espaços e quebras diferentes. Esta versão facilita a busca sem perder a
    possibilidade de recortar o conteúdo original.
    """
    equivalencias = {"º": "O", "ª": "A", "–": "-", "—": "-", "\u00a0": " "}
    caracteres: list[str] = []

    for caractere in texto:
        caractere = equivalencias.get(caractere, caractere)
        decomposicao = unicodedata.normalize("NFKD", caractere)
        base = "".join(c for c in decomposicao if not unicodedata.combining(c))
        caracteres.append((base[:1] or " ").upper())

    return "".join(caracteres)


def ler_texto_pdf(caminho_pdf: Path) -> str:
    """Lê o texto de todas as páginas do PDF."""
    if fitz is not None:
        documento = fitz.open(str(caminho_pdf))
        partes = [pagina.get_text("text") or "" for pagina in documento]
        return normalizar_texto("\n".join(partes))

    leitor = PdfReader(str(caminho_pdf))
    partes: list[str] = []
    for pagina in leitor.pages:
        partes.append(pagina.extract_text() or "")
    return normalizar_texto(" ".join(partes))
















def limpar_cnpj(cnpj: str) -> str:
    return re.sub(r"\s+", "", cnpj)


def normalizar_numero_nf(numero: str) -> str:
    """Remove zeros à esquerda sem alterar números já normalizados."""
    numero = re.sub(r"\D+", "", numero or "")
    if not numero:
        return ""
    return str(int(numero))



def data_valida_recente(data_str: str) -> bool:
    if not data_str:
        return False
    try:
        from datetime import datetime
        from datetime import timedelta
        data_parsed = datetime.strptime(data_str, "%d/%m/%Y")
        limite_inferior = datetime.now() - timedelta(days=90)
        return data_parsed >= limite_inferior
    except Exception:
        return False

def normalizar_data_br(valor: str) -> str:
    """Retorna apenas a data DD/MM/AAAA, desde que não seja muito antiga."""
    if not valor:
        return ""
    encontrado = re.search(r"\d{2}/\d{2}/\d{4}", valor)
    if encontrado:
        data_str = encontrado.group(0)
        if data_valida_recente(data_str):
            return data_str
    return ""





































def formatar_moeda_br(valor: Decimal) -> str:
    valor = Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")


def formatar_percentual(valor: Decimal) -> str:
    valor = Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
    texto = f"{valor:,.2f}"
    return texto.replace(",", "X").replace(".", ",").replace("X", ".")
















def determinar_natureza_operacao(cidade_local_prestacao: str) -> str:
    return (
        "Tributação no Município"
        if cidade_local_prestacao.strip().lower() == "fortaleza"
        else "Tributação Fora do Município"
    )









def extrair_nota_fiscal(caminho_pdf: Path, api_key: str = None) -> NotaFiscalExtraida:
    texto = ler_texto_pdf(caminho_pdf)
    dados_gemini = extrair_campos_gemini(caminho_pdf, texto, api_key)
    
    prefeitura = dados_gemini.get("prefeitura", "")
    cnpj_prestador = dados_gemini.get("cnpj_prestador", "")
    numero_nf = dados_gemini.get("numero_nf", "")
    data_emissao = dados_gemini.get("data_emissao", "")
    id_cnae = dados_gemini.get("id_cnae", "")
    desc_cnae = dados_gemini.get("desc_cnae", "")
    descricao_servico = dados_gemini.get("descricao_servico", "")
    uf_local_prestacao = dados_gemini.get("uf_local_prestacao", "")
    cidade_local_prestacao = dados_gemini.get("cidade_local_prestacao", "")
    iss_retido = dados_gemini.get("iss_retido", "").strip().upper()
    valor_servico = dados_gemini.get("valor_servico", "")
    aliquota = dados_gemini.get("aliquota", "")
    valor_deducoes = dados_gemini.get("valor_deducoes", "0,00")
    descontos_incondicionados = dados_gemini.get("descontos_incondicionados", "0,00")
    descontos_condicionados = dados_gemini.get("descontos_condicionados", "0,00")
    outras_retencoes = dados_gemini.get("outras_retencoes", "0,00")
    ir = dados_gemini.get("ir", "0,00")
    pis_nao_retido = dados_gemini.get("pis_nao_retido", "0,00")
    cofins_nao_retido = dados_gemini.get("cofins_nao_retido", "0,00")
    csrf = dados_gemini.get("csrf", "0,00")
    inss = dados_gemini.get("inss", "0,00")

    # Anti-alucinação: validar se os dígitos de CNPJ e Número da NF constam no texto bruto
    digitos_texto = "".join(c for c in texto if c.isdigit())
    
    digitos_cnpj = "".join(c for c in cnpj_prestador if c.isdigit())
    if digitos_cnpj and digitos_cnpj not in digitos_texto:
        cnpj_prestador = ""
        
    digitos_nf = "".join(c for c in numero_nf if c.isdigit())
    if digitos_nf and digitos_nf not in digitos_texto:
        numero_nf = ""

    natureza_operacao = determinar_natureza_operacao(cidade_local_prestacao)

    return NotaFiscalExtraida(
        arquivo_pdf=caminho_pdf.name,
        prefeitura=prefeitura,
        cnpj_prestador=limpar_cnpj(cnpj_prestador),
        numero_nf=normalizar_numero_nf(numero_nf),
        data_emissao=normalizar_data_br(data_emissao),
        id_cnae=id_cnae,
        desc_cnae=desc_cnae,
        descricao_servico=descricao_servico,
        uf_local_prestacao=uf_local_prestacao,
        cidade_local_prestacao=cidade_local_prestacao,
        natureza_operacao=natureza_operacao,
        iss_retido=iss_retido,
        valor_servico=valor_servico,
        aliquota=aliquota,
        valor_deducoes=valor_deducoes,
        descontos_incondicionados=descontos_incondicionados,
        descontos_condicionados=descontos_condicionados,
        outras_retencoes=outras_retencoes,
        ir=ir,
        pis_nao_retido=pis_nao_retido,
        cofins_nao_retido=cofins_nao_retido,
        csrf=csrf,
        inss=inss,
        revisao_manual="",
        analisado_pela_ia="SIM" if dados_gemini else "NÃO",
    )



def listar_pdfs(pasta: Path) -> Iterable[Path]:
    return sorted(pasta.glob("*.pdf"))


def ajustar_largura_colunas(planilha) -> None:
    for coluna in planilha.columns:
        maior = 0
        letra_coluna = get_column_letter(coluna[0].column)
        for celula in coluna:
            if celula.value is not None:
                maior = max(maior, len(str(celula.value)))
        planilha.column_dimensions[letra_coluna].width = min(maior + 2, 80)


def gerar_xlsx(registros: list[NotaFiscalExtraida], destino: Path) -> None:
    # Enriquecemos os registros com o CNAE oficial antes de gravar a planilha final.
    enriquecer_registros_cnae_final(registros)

    wb = Workbook()
    ws = wb.active
    ws.title = "NF Extraídas"

    cabecalhos = [
        "ARQUIVO_PDF",
        "PREFEITURA",
        "CNPJ_PRESTADOR",
        "NUMERO_NF",
        "DATA_EMISSAO",
        "ID_CNAE",
        "DESC_CNAE",
        "DESCRICAO_SERVICO",
        "UF_LOCAL_PRESTACAO",
        "CIDADE_LOCAL_PRESTACAO",
        "NATUREZA_OPERACAO",
        "ISS_RETIDO",
        "VALOR_SERVICO",
        "VALOR_DEDUCOES",
        "DESCONTOS_INCONDICIONADOS",
        "DESCONTOS_CONDICIONADOS",
        "OUTRAS_RETENCOES",
        "IR",
        "PIS_NAO_RETIDO",
        "COFINS_NAO_RETIDO",
        "CSRF (CSLL + PIS + Cofins Retidos)",
        "INSS",
        "ALIQUOTA",
        "ID_CNAE_FINAL",
        "DESC_CNAE_FINAL",
        "ANALISADO_PELA_IA",
        "REVISAO_MANUAL",
    ]

    ws.append(cabecalhos)
    for registro in registros:
        ws.append(registro.como_linha())

    for celula in ws[1]:
        celula.font = Font(bold=True, color="FFFFFF")
        celula.fill = PatternFill("solid", fgColor="1F4E78")

    ajustar_largura_colunas(ws)
    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)


def registrar_log_incompletos(
    registros: list[NotaFiscalExtraida],
    destino_xlsx: Path,
) -> Path:
    """Cria um log detalhado quando alguma NF não preencher todas as colunas.

    O objetivo é permitir auditoria rápida dos arquivos que precisam de ajuste
    manual ou de novas regras de extração.
    """
    linhas_log: list[str] = []

    for registro in registros:
        campos_vazios = registro.campos_vazios()
        if not campos_vazios:
            continue

        linhas_log.append(
            f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {registro.arquivo_pdf}"
        )
        linhas_log.append(f"Campos ausentes: {', '.join(campos_vazios)}")
        linhas_log.append(f"Resumo: {registro.resumo_campos()}")
        linhas_log.append("")

    caminho_log = destino_xlsx.with_name(f"{destino_xlsx.stem}_log_extracao.txt")
    if linhas_log:
        caminho_log.write_text("\n".join(linhas_log), encoding="utf-8")
    elif caminho_log.exists():
        caminho_log.unlink()

    return caminho_log


def abrir_site_iss_fortaleza() -> None:
    """Mantido apenas por compatibilidade histórica com versões anteriores."""
    pass



def solicitar_opcao() -> str:
    """Pede ao usuário que escolha entre gerar o XLSX (Apenas IA) ou executar a automação."""
    while True:
        print("Escolha uma opção:")
        print("1 - Gerar o XLSX a partir dos PDFs (Apenas IA)")
        print("2 - Executar a automação visível da ISS de Fortaleza")
        resposta = input("Digite 1 ou 2: ").strip()
        if resposta in {"1", "2"}:
            return resposta
        print("Opção inválida. Por favor, digite 1 ou 2.")
        print()



def solicitar_planilha_automacao() -> Path:
    """Pergunta ao usuário qual planilha será usada na automação da ISS."""

    entrada = input(
        "Informe o caminho da planilha XLSX que será usada na automação da ISS "
        "(Enter para `nf_compilado.xlsx`): "
    ).strip().strip('"')
    if not entrada:
        return Path("nf_compilado.xlsx")
    return Path(entrada)


def oferecer_proximo_passo_apos_funcao1(destino: Path, args: argparse.Namespace) -> int:
    """Pergunta se o usuário quer seguir para a função 2 ou encerrar o programa.

    A planilha recém-gerada vira a entrada padrão da automação da ISS para
    evitar retrabalho e manter o fluxo contínuo entre as duas etapas.
    """

    if not sys.stdin.isatty():
        print(
            "Próximo passo sugerido: escolha a opção 2 para iniciar a automação visível da ISS de Fortaleza."
        )
        return 0

    while True:
        print()
        print("A função 1 foi concluída com sucesso.")
        print(f"Planilha gerada: {destino.resolve()}")
        print("O que você deseja fazer agora?")
        print("1 - Iniciar a função 2 (automação visível da ISS de Fortaleza)")
        print("2 - Encerrar o programa")
        resposta = input("Digite 1 ou 2: ").strip()

        if resposta == "1":
            registrar_evento_execucao(
                "Usuário optou por iniciar a função 2 após a geração do XLSX",
                "extrair_nf_pdfs.py",
            )
            try:
                competencia = (
                    interpretar_competencia(args.competencia)
                    if args.competencia is not None
                    else solicitar_competencia()
                )
            except ValueError as exc:
                raise SystemExit(f"Competência inválida: {exc}") from exc

            # A planilha gerada na função 1 é a base natural para a automação da ISS.
            executar_automacao_iss(destino, competencia, depuracao=args.debug_funcao2)
            return 0

        if resposta == "2":
            registrar_evento_execucao(
                "Usuário optou por encerrar o programa após a função 1",
                "extrair_nf_pdfs.py",
            )
            print("Programa encerrado.")
            return 0

        print("Opção inválida. Digite 1 para iniciar a função 2 ou 2 para encerrar.")
        print()



def construir_argumentos() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Extrai dados de NFS-e em PDF, gera XLSX ou executa a automação da ISS."
    )
    parser.add_argument(
        "-m",
        "--modo",
        choices=("1", "2"),
        default=None,
        help="1 para gerar o XLSX (Apenas IA); 2 para executar a automação da ISS de Fortaleza.",
    )
    parser.add_argument(
        "-c",
        "--competencia",
        default=None,
        help="Competência a trabalhar na ISS, por exemplo '5/2026', '5 2026' ou 'maio 2026'.",
    )
    parser.add_argument(
        "--planilha",
        type=Path,
        default=None,
        help="Planilha XLSX usada pela automação da ISS.",
    )
    parser.add_argument(
        "--debug-funcao2",
        action="store_true",
        help="Abre a função 2 com pausas extras e navegador em modo de teste visível.",
    )
    parser.add_argument(
        "pasta",
        nargs="?",
        type=Path,
        default=None,
        help="Pasta que contém os arquivos PDF. Se omitido, o script solicita no terminal.",
    )
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("nf_compilado.xlsx"),
        help="Caminho do arquivo XLSX de saída.",
    )
    return parser.parse_args()



def main() -> int:
    args = construir_argumentos()
    registrar_evento_execucao("Fluxo principal iniciado no extrair_nf_pdfs.py", "extrair_nf_pdfs.py")

    modo = args.modo
    if modo is None:
        if args.pasta is not None:
            modo = "1"
        else:
            modo = solicitar_opcao()

    if modo == "2":
        try:
            competencia = (
                interpretar_competencia(args.competencia)
                if args.competencia is not None
                else solicitar_competencia()
            )
        except ValueError as exc:
            raise SystemExit(f"Competência inválida: {exc}") from exc
        planilha = args.planilha if args.planilha is not None else solicitar_planilha_automacao()
        if not planilha.exists():
            raise SystemExit(
                f"A planilha informada para a automação não existe: {planilha}"
            )
        registrar_evento_execucao(
            f"Modo 2 selecionado com planilha {planilha.resolve()}",
            "extrair_nf_pdfs.py",
        )
        executar_automacao_iss(planilha, competencia, depuracao=args.debug_funcao2)
        return 0

    pasta = args.pasta
    if pasta is None:
        entrada = input("Informe o caminho da pasta com os PDFs: ").strip().strip('"')
        pasta = Path(entrada)

    destino = args.output

    if not pasta.exists() or not pasta.is_dir():
        raise SystemExit(f"A pasta informada não existe ou não é válida: {pasta}")

    if not destino.is_absolute():
        destino = (pasta / destino.name).resolve()

    pdfs = list(listar_pdfs(pasta))
    if not pdfs:
        raise SystemExit("Nenhum PDF foi encontrado na pasta informada.")

    registrar_evento_execucao(
        f"Modo {modo} selecionado com {len(pdfs)} PDF(s) em {pasta.resolve()}",
        "extrair_nf_pdfs.py",
    )

    registros: list[NotaFiscalExtraida] = []
    notas_com_erro_cota: list[tuple[str, str]] = []
    for indice, pdf in enumerate(pdfs, start=1):
        print(f"[{indice}/{len(pdfs)}] Processando PDF: {pdf.name}", flush=True)
        registrar_evento_execucao(
            f"Iniciando processamento do PDF {indice}/{len(pdfs)}: {pdf.name}",
            "extrair_nf_pdfs.py",
        )
        try:
            registros.append(extrair_nota_fiscal(pdf))
        except ErroCotaGemini as exc:
            print(f"[{indice}/{len(pdfs)}] Limite de cota da IA atingido ao processar {pdf.name}", flush=True)
            registrar_evento_execucao(
                f"Limite de cota atingido para {pdf.name}: {exc}",
                "extrair_nf_pdfs.py",
            )
            notas_com_erro_cota.append((pdf.name, str(exc)))
            registros.append(
                NotaFiscalExtraida(
                    arquivo_pdf=pdf.name,
                    descricao_servico="NÃO ANALISADO: Limite de cota da IA excedido.",
                )
            )
        except Exception as exc:
            print(f"[{indice}/{len(pdfs)}] Falha ao processar {pdf.name}: {exc}", flush=True)
            registrar_evento_execucao(
                f"Falha ao processar {pdf.name}: {exc}",
                "extrair_nf_pdfs.py",
            )
            registros.append(
                NotaFiscalExtraida(
                    arquivo_pdf=pdf.name,
                    descricao_servico=f"ERRO NA EXTRAÇÃO: {exc}",
                )
            )
        else:
            registrar_evento_execucao(
                f"Processamento concluído com sucesso para {pdf.name}",
                "extrair_nf_pdfs.py",
            )

    registros_completos = [registro for registro in registros if not registro.campos_vazios()]
    gerar_xlsx(registros_completos, destino)
    caminho_log = registrar_log_incompletos(registros, destino)
    print(f"Arquivo XLSX gerado com sucesso: {destino.resolve()}")
    print(f"Total de PDFs processados: {len(pdfs)}")
    print(f"Linhas completas exportadas: {len(registros_completos)}")
    if caminho_log.exists():
        print(f"Log de extração gerado em: {caminho_log.resolve()}")

    if notas_com_erro_cota:
        caminho_log_cota = destino.parent / f"{destino.stem}_log_cota_excedida.txt"
        try:
            with open(caminho_log_cota, "w", encoding="utf-8") as f:
                f.write(f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] NOTAS NÃO ANALISADAS POR LIMITE DE COTA DA IA:\n")
                for nome_pdf, erro in notas_com_erro_cota:
                    f.write(f"- {nome_pdf} (Motivo: {erro})\n")
            print(f"Aviso: Algumas notas não foram analisadas por limite de cota da API. Detalhes gravados em: {caminho_log_cota.resolve()}", flush=True)
        except Exception as exc_log:
            registrar_evento_execucao(
                f"Erro ao salvar log de cota excedida {caminho_log_cota.name}: {exc_log}",
                "extrair_nf_pdfs.py",
            )

    registrar_evento_execucao(
        f"XLSX gerado em {destino.resolve()} com {len(registros_completos)} linha(s) exportada(s)",
        "extrair_nf_pdfs.py",
    )
    return oferecer_proximo_passo_apos_funcao1(destino, args)


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except SystemExit as exc:
        if exc.code not in (0, None):
            registrar_erro(exc, "extrair_nf_pdfs.py")
        raise
    except Exception as exc:
        registrar_erro(exc, "extrair_nf_pdfs.py")
        raise

