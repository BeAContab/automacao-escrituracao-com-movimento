// ---- Exportar XML de Prestados ----

async function selecionarPastaExportar() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            document.getElementById('input-pasta-exportar').value = pasta;
        }
    }
}

async function iniciarExportacaoXML() {
    const pasta = document.getElementById('input-pasta-exportar').value.trim();
    const mes = document.getElementById('select-mes-exportar').value;
    const ano = document.getElementById('select-ano-exportar').value;
    const competencia = `${mes}/${ano}`;

    if (!pasta) {
        addLogExportar('Selecione a pasta de destino dos downloads.', true);
        return;
    }

    document.getElementById('log-exportar').innerHTML = '';
    addLogExportar(`Iniciando exportação de XMLs para competência ${competencia}...`);
    document.getElementById('progresso-container-exportar').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_exportacao_xml_gui(pasta, competencia);
        } catch (e) {
            addLogExportar('Erro ao iniciar exportação: ' + e, true);
        }
    }
}

async function confirmarLoginExportacao() {
    if (window.pywebview && window.pywebview.api) {
        const ok = await window.pywebview.api.confirmar_login_exportacao();
        if (ok) {
            addLogExportar('Login verificado pelo operador. Retomando robô exportador...');
            document.getElementById('card-confirmar-login-exportar').classList.add('hidden');
        } else {
            addLogExportar('Falha ao enviar sinal de confirmação do login.', true);
        }
    }
}

function addLogExportar(msg, isError = false) {
    const el = document.getElementById('log-exportar');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_exportar = addLogExportar;

function limparLogExportar() {
    const el = document.getElementById('log-exportar');
    if (el) el.innerHTML = '';
}

function updateProgressoExportar(porcentagem, status) {
    const container = document.getElementById('progresso-container-exportar');
    const bar = document.getElementById('progresso-barra-exportar');
    const pct = document.getElementById('progresso-porcentagem-exportar');
    const st = document.getElementById('progresso-status-exportar');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_exportar = updateProgressoExportar;
