// ---- Escrituração — Multi-CNPJ ----
// Etapa 1: apenas valida o acesso de cada empresa (login, seleção/conferência do CNPJ, logout).

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

async function iniciarEscrituracaoMulti() {
    const planilha = (window._planilhaEscMulti || '').trim();
    if (!planilha) {
        addLogEscMulti('Selecione a planilha do Multi-CNPJ (com CNPJ_TOMADOR, LOGIN e SENHA preenchidos).', true);
        return;
    }

    document.getElementById('log-esc-multi').innerHTML = '';
    document.getElementById('resumo-esc-multi').classList.add('hidden');
    document.getElementById('progresso-container-esc-multi').classList.remove('hidden');
    updateProgressoEscMulti(0, 'Iniciando...');

    const btn = document.getElementById('btn-iniciar-esc-multi');
    btn.disabled = true;
    btn.innerHTML = '<span class="material-symbols-outlined" style="animation: spin 1s linear infinite; display:inline-block;">progress_activity</span> Executando...';

    escMultiPausado = false;
    document.getElementById('lbl-pausar-esc-multi').innerText = 'Pausar';
    document.getElementById('icone-pausar-esc-multi').innerText = 'pause';
    document.getElementById('btn-parar-esc-multi').disabled = false;
    document.getElementById('controles-esc-multi').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_escrituracao_multi_gui(planilha);
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
        'Parar a execução?\n\nO robô conclui a etapa em andamento, faz logout do portal e fecha o navegador. ' +
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
    btn.innerHTML = '<span class="material-symbols-outlined">verified_user</span> Validar Acessos';
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
