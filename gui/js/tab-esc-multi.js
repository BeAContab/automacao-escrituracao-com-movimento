// ---- Escrituração — Multi-CNPJ ----
// Três modos: validar acessos (login, seleção/conferência do CNPJ e logout), simular (preenche os campos
// e NÃO grava) e escriturar de verdade (grava no portal — sempre com confirmação do operador).

let escMultiPausado = false;

async function selecionarPlanilhaEscMulti() {
    if (window.pywebview) {
        const arquivo = await window.pywebview.api.selecionar_arquivo();
        if (arquivo) {
            window._planilhaEscMulti = arquivo;
            document.getElementById('lbl-planilha-esc-multi').innerText = arquivo;
        }
    }
}

// Modo escolhido na tela: 'validar' | 'simular' | 'escriturar'
function modoEscMulti() {
    if (document.getElementById('chk-apenas-validar-esc-multi').checked) return 'validar';
    const escolhido = document.querySelector('input[name="modo-esc-multi"]:checked');
    return escolhido ? escolhido.value : 'simular';
}

const TEXTOS_MODO_ESC_MULTI = {
    validar: { icone: 'verified_user', rotulo: 'Validar Acessos', perigo: false },
    simular: { icone: 'science', rotulo: 'Simular Escrituração (não grava)', perigo: false },
    escriturar: { icone: 'edit_note', rotulo: 'Escriturar de Verdade', perigo: true },
};

// Ajusta a tela ao modo: mostra/oculta competência e opções, e troca o texto/cor do botão principal.
function atualizarModoEscMulti() {
    const modo = modoEscMulti();
    document.getElementById('bloco-modo-esc-multi').classList.toggle('hidden', modo === 'validar');
    const btn = document.getElementById('btn-iniciar-esc-multi');
    if (btn.disabled) return;  // em execução: o botão mostra "Executando..."
    const t = TEXTOS_MODO_ESC_MULTI[modo];
    btn.innerHTML = `<span class="material-symbols-outlined">${t.icone}</span> ${t.rotulo}`;
    btn.classList.toggle('bg-primary-container', !t.perigo);
    btn.classList.toggle('bg-error', t.perigo);
}

// Trava/destrava as opções enquanto a execução está em andamento.
function bloquearOpcoesEscMulti(bloquear) {
    ['chk-apenas-validar-esc-multi', 'select-mes-esc-multi', 'select-ano-esc-multi'].forEach(id => {
        document.getElementById(id).disabled = bloquear;
    });
    document.querySelectorAll('input[name="modo-esc-multi"]').forEach(r => { r.disabled = bloquear; });
}

async function iniciarEscrituracaoMulti() {
    const planilha = (window._planilhaEscMulti || '').trim();
    if (!planilha) {
        addLogEscMulti('Selecione a planilha do Multi-CNPJ (com CNPJ_TOMADOR, LOGIN e SENHA preenchidos).', true);
        return;
    }

    const modo = modoEscMulti();
    let competencia = '';
    if (modo !== 'validar') {
        const mes = document.getElementById('select-mes-esc-multi').value;
        const ano = document.getElementById('select-ano-esc-multi').value;
        if (!mes || !ano) {
            addLogEscMulti('Selecione o mês e o ano da competência.', true);
            return;
        }
        competencia = `${mes}/${ano}`;
    }

    // Escriturar de verdade GRAVA no portal: nunca começa sem a confirmação explícita do operador.
    let confirmado = false;
    if (modo === 'escriturar') {
        confirmado = confirm(
            'ESCRITURAR DE VERDADE\n\n' +
            'O robô vai GRAVAR as notas da planilha no portal ISS, empresa por empresa.\n\n' +
            `Planilha: ${planilha}\n` +
            `Competência: ${competencia} (para todas as empresas)\n\n` +
            'Você já fez a simulação e conferiu o resultado?\n\n' +
            'Deseja realmente começar a escriturar?'
        );
        if (!confirmado) {
            addLogEscMulti('Escrituração cancelada antes de começar. Nada foi feito.');
            return;
        }
    }

    document.getElementById('log-esc-multi').innerHTML = '';
    document.getElementById('resumo-esc-multi').classList.add('hidden');
    document.getElementById('progresso-container-esc-multi').classList.remove('hidden');
    updateProgressoEscMulti(0, 'Iniciando...');

    const btn = document.getElementById('btn-iniciar-esc-multi');
    btn.disabled = true;
    btn.innerHTML = '<span class="material-symbols-outlined" style="animation: spin 1s linear infinite; display:inline-block;">progress_activity</span> Executando...';
    bloquearOpcoesEscMulti(true);

    escMultiPausado = false;
    document.getElementById('lbl-pausar-esc-multi').innerText = 'Pausar';
    document.getElementById('icone-pausar-esc-multi').innerText = 'pause';
    document.getElementById('btn-parar-esc-multi').disabled = false;
    document.getElementById('controles-esc-multi').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_escrituracao_multi_gui(planilha, modo, competencia, confirmado);
        } catch (e) {
            addLogEscMulti('Erro ao iniciar a Escrituração Multi-CNPJ: ' + e, true);
            escMultiFinalizado();
        }
    }
    // O botão volta ao normal quando o Python avisa o fim (escMultiFinalizado).
}

function alternarPausaEscMulti() {
    if (!window.pywebview || !window.pywebview.api) return;
    escMultiPausado = !escMultiPausado;
    if (escMultiPausado) {
        window.pywebview.api.pausar_processamento_xml('esc_multi');
        document.getElementById('lbl-pausar-esc-multi').innerText = 'Continuar';
        document.getElementById('icone-pausar-esc-multi').innerText = 'play_arrow';
    } else {
        window.pywebview.api.retomar_processamento_xml('esc_multi');
        document.getElementById('lbl-pausar-esc-multi').innerText = 'Pausar';
        document.getElementById('icone-pausar-esc-multi').innerText = 'pause';
    }
}

function pararEscrituracaoMulti() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        'Parar a execução?\n\nO robô interrompe a empresa em andamento, faz logout do portal e fecha o navegador. ' +
        'Notas já escrituradas ficam salvas (não serão repetidas ao executar de novo). ' +
        'As empresas que ainda não foram acessadas ficam como "Cancelada".'
    );
    if (!confirmado) return;
    window.pywebview.api.cancelar_processamento_xml('esc_multi');
    document.getElementById('btn-parar-esc-multi').disabled = true;
}

function escMultiFinalizado() {
    escMultiPausado = false;
    document.getElementById('controles-esc-multi').classList.add('hidden');
    const btn = document.getElementById('btn-iniciar-esc-multi');
    btn.disabled = false;
    bloquearOpcoesEscMulti(false);
    atualizarModoEscMulti();
}
window.esc_multi_finalizado = escMultiFinalizado;

function addLogEscMulti(msg, isError = false) {
    const el = document.getElementById('log-esc-multi');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_esc_multi = addLogEscMulti;

function limparLogEscMulti() {
    const el = document.getElementById('log-esc-multi');
    if (el) el.innerHTML = '';
}

function updateProgressoEscMulti(porcentagem, status) {
    const container = document.getElementById('progresso-container-esc-multi');
    const bar = document.getElementById('progresso-barra-esc-multi');
    const pct = document.getElementById('progresso-porcentagem-esc-multi');
    const st = document.getElementById('progresso-status-esc-multi');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_esc_multi = updateProgressoEscMulti;

// Resumo final: uma linha por empresa (CNPJ + situação). Os textos vêm do Python e são
// inseridos com innerText, nunca como HTML.
function mostrarResumoEscMulti(resumo) {
    const card = document.getElementById('resumo-esc-multi');
    const titulo = document.getElementById('resumo-esc-multi-titulo');
    const lista = document.getElementById('resumo-esc-multi-lista');
    if (!card || !lista) return;

    const partes = Object.entries(resumo.contagem || {}).map(([rotulo, qtd]) => `${qtd} × ${rotulo}`);
    titulo.innerText = (resumo.cancelado ? 'Execução cancelada — ' : 'Resumo — ') + (partes.join(' · ') || 'nenhuma empresa');

    lista.innerHTML = '';
    (resumo.empresas || []).forEach(emp => {
        const linha = document.createElement('div');
        linha.className = 'flex items-start gap-sm py-xs border-b border-outline-variant last:border-b-0';

        const icone = document.createElement('span');
        icone.className = 'material-symbols-outlined text-[18px] ' + (emp.ok ? 'text-primary-container' : 'text-error');
        icone.innerText = emp.ok ? 'check_circle' : 'error';

        const texto = document.createElement('div');
        texto.className = 'flex-1 min-w-0';
        const cabecalho = document.createElement('p');
        cabecalho.className = 'font-label-md text-label-md text-on-surface';
        cabecalho.innerText = `${emp.cnpj} — ${emp.status}`;
        texto.appendChild(cabecalho);
        if (emp.mensagem) {
            const detalhe = document.createElement('p');
            detalhe.className = 'font-body-sm text-body-sm text-on-surface-variant break-words';
            detalhe.innerText = emp.mensagem;
            texto.appendChild(detalhe);
        }

        linha.appendChild(icone);
        linha.appendChild(texto);
        lista.appendChild(linha);
    });
    card.classList.remove('hidden');
}
window.mostrar_resumo_esc_multi = mostrarResumoEscMulti;
