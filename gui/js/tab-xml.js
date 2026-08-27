// ---- Processamento de XMLs (Tess AI) ----

async function selecionarPastaXML() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            window._pastaXmlSelecionada = pasta;
            document.getElementById('lbl-pasta-xml').innerText = pasta;
        }
    }
}

async function iniciarProcessamentoXML() {
    const pasta = (window._pastaXmlSelecionada || '').trim();
    if (!pasta) {
        addLogXml('Selecione a pasta de origem dos XMLs.', true);
        return;
    }

    const tessKey = document.getElementById('api-key-tess').value.trim();
    const tessAgent = document.getElementById('agent-id-tess').value.trim();

    document.getElementById('log-xml').innerHTML = '';
    // Oculta botão de abrir planilha de execução anterior (MELHORIA-05)
    document.getElementById('container-abrir-planilha').classList.add('hidden');
    window._ultimaPlanilhaGerada = null;
    // Oculta botão de abrir pasta com planilhas (caso de subpastas) de execução anterior
    document.getElementById('container-abrir-pasta-resultados').classList.add('hidden');
    window._ultimaPastaResultados = null;

    addLogXml('Iniciando varredura e importação de XMLs...');
    document.getElementById('progresso-container-xml').classList.remove('hidden');

    // Muda o botão para estado "processando" com ícone animado
    const btnProcessar = document.getElementById('btn-processar-xml');
    btnProcessar.disabled = true;
    btnProcessar.innerHTML = '<span class="material-symbols-outlined" style="animation: spin 1s linear infinite; display:inline-block;">progress_activity</span> Processando...';

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.processar_xmls_gui(pasta, tessKey, tessAgent);
        } catch (e) {
            addLogXml('Erro ao chamar processamento de XMLs: ' + e, true);
        }
    }

    // Restaura o botão após o processamento
    btnProcessar.disabled = false;
    btnProcessar.innerHTML = '<span class="material-symbols-outlined">play_arrow</span> Iniciar Processamento';
}

function addLogXml(msg, isError = false) {
    const el = document.getElementById('log-xml');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_xml = addLogXml;

function limparLogXml() {
    const el = document.getElementById('log-xml');
    if (el) el.innerHTML = '';
}

// MELHORIA-05: Exibe o botão de abrir planilha e armazena o caminho gerado pelo Python
function mostrarBotaoPlanilha(caminho) {
    window._ultimaPlanilhaGerada = caminho;
    const container = document.getElementById('container-abrir-planilha');
    if (container) container.classList.remove('hidden');
}
window.mostrar_botao_planilha = mostrarBotaoPlanilha;

function abrirPlanilhaGerada() {
    if (window._ultimaPlanilhaGerada && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimaPlanilhaGerada.replace(/\\/g, '/'));
    }
}

// Exibe o botão de abrir a pasta-pai quando o processamento gerou uma planilha
// por subpasta (várias planilhas), em vez de um único arquivo consolidado.
function mostrarBotaoPastaResultados(pasta, quantidade) {
    window._ultimaPastaResultados = pasta;
    const lbl = document.getElementById('lbl-abrir-pasta-resultados');
    if (lbl) lbl.innerText = `Abrir Pasta com as ${quantidade} Planilhas Geradas`;
    const container = document.getElementById('container-abrir-pasta-resultados');
    if (container) container.classList.remove('hidden');
}
window.mostrar_botao_pasta_resultados = mostrarBotaoPastaResultados;

function abrirPastaResultados() {
    if (window._ultimaPastaResultados && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimaPastaResultados.replace(/\\/g, '/'));
    }
}

function updateProgressoXml(porcentagem, status) {
    const container = document.getElementById('progresso-container-xml');
    const bar = document.getElementById('progresso-barra-xml');
    const pct = document.getElementById('progresso-porcentagem-xml');
    const st = document.getElementById('progresso-status-xml');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_xml = updateProgressoXml;
