// ---- Processar XMLs — Multi-CNPJ ----

let xmlMultiPausado = false;

async function selecionarPastaXmlMulti() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            window._pastaXmlMultiSelecionada = pasta;
            document.getElementById('lbl-pasta-xml-multi').innerText = pasta;
        }
    }
}

async function iniciarProcessamentoXmlMulti() {
    const pasta = (window._pastaXmlMultiSelecionada || '').trim();
    if (!pasta) {
        addLogXmlMulti('Selecione a pasta que contém as pastas de CNPJ.', true);
        return;
    }

    const tessKey = document.getElementById('api-key-tess').value.trim();
    const tessAgent = document.getElementById('agent-id-tess').value.trim();

    document.getElementById('log-xml-multi').innerHTML = '';
    document.getElementById('container-abrir-planilha-multi').classList.add('hidden');
    window._ultimaPlanilhaMulti = null;

    addLogXmlMulti('Iniciando processamento Multi-CNPJ...');
    document.getElementById('progresso-container-xml-multi').classList.remove('hidden');

    const btn = document.getElementById('btn-processar-xml-multi');
    btn.disabled = true;
    btn.innerHTML = '<span class="material-symbols-outlined" style="animation: spin 1s linear infinite; display:inline-block;">progress_activity</span> Processando...';

    xmlMultiPausado = false;
    document.getElementById('lbl-pausar-xml-multi').innerText = 'Pausar';
    document.getElementById('icone-pausar-xml-multi').innerText = 'pause';
    document.getElementById('btn-parar-xml-multi').disabled = false;
    document.getElementById('controles-xml-multi').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.processar_xmls_multi_cnpj_gui(pasta, tessKey, tessAgent);
        } catch (e) {
            addLogXmlMulti('Erro ao chamar o processamento Multi-CNPJ: ' + e, true);
            xmlMultiFinalizado();
        }
    }
    // O botão volta ao normal quando o Python avisa o fim (xmlMultiFinalizado).
}

function alternarPausaXmlMulti() {
    if (!window.pywebview || !window.pywebview.api) return;
    xmlMultiPausado = !xmlMultiPausado;
    if (xmlMultiPausado) {
        window.pywebview.api.pausar_processamento_xml('multi');
        document.getElementById('lbl-pausar-xml-multi').innerText = 'Continuar';
        document.getElementById('icone-pausar-xml-multi').innerText = 'play_arrow';
    } else {
        window.pywebview.api.retomar_processamento_xml('multi');
        document.getElementById('lbl-pausar-xml-multi').innerText = 'Pausar';
        document.getElementById('icone-pausar-xml-multi').innerText = 'pause';
    }
}

function pararProcessamentoXmlMulti() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        'Parar o processamento?\n\nO que já foi processado fica salvo na planilha. ' +
        'Ao processar a mesma pasta de novo, a planilha será continuada de onde parou.'
    );
    if (!confirmado) return;
    window.pywebview.api.cancelar_processamento_xml('multi');
    document.getElementById('btn-parar-xml-multi').disabled = true;
}

function xmlMultiFinalizado() {
    xmlMultiPausado = false;
    document.getElementById('controles-xml-multi').classList.add('hidden');
    const btn = document.getElementById('btn-processar-xml-multi');
    btn.disabled = false;
    btn.innerHTML = '<span class="material-symbols-outlined">play_arrow</span> Iniciar Processamento';
}
window.xml_multi_finalizado = xmlMultiFinalizado;

function addLogXmlMulti(msg, isError = false) {
    const el = document.getElementById('log-xml-multi');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_xml_multi = addLogXmlMulti;

function limparLogXmlMulti() {
    const el = document.getElementById('log-xml-multi');
    if (el) el.innerHTML = '';
}

function updateProgressoXmlMulti(porcentagem, status) {
    const container = document.getElementById('progresso-container-xml-multi');
    const bar = document.getElementById('progresso-barra-xml-multi');
    const pct = document.getElementById('progresso-porcentagem-xml-multi');
    const st = document.getElementById('progresso-status-xml-multi');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_xml_multi = updateProgressoXmlMulti;

function mostrarBotaoPlanilhaMulti(caminho) {
    window._ultimaPlanilhaMulti = caminho;
    const container = document.getElementById('container-abrir-planilha-multi');
    if (container) container.classList.remove('hidden');
}
window.mostrar_botao_planilha_multi = mostrarBotaoPlanilhaMulti;

function abrirPlanilhaMulti() {
    if (window._ultimaPlanilhaMulti && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimaPlanilhaMulti.replace(/\\/g, '/'));
    }
}
