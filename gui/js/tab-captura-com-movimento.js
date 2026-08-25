// ---- Captura Escrituração (Com Movimento) ----

async function selecionarPlanilhaCaptura() {
    if (window.pywebview) {
        const file = await window.pywebview.api.selecionar_arquivo();
        if (file) {
            document.getElementById('input-planilha-captura').value = file;
        }
    }
}

async function selecionarSaidaCaptura() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            document.getElementById('input-saida-captura').value = pasta;
        }
    }
}

async function iniciarCapturaComMovimento() {
    const planilha = document.getElementById('input-planilha-captura').value.trim();
    const saida = document.getElementById('input-saida-captura').value.trim();
    const mes = document.getElementById('select-mes-captura').value;
    const ano = document.getElementById('select-ano-captura').value;
    const competencia = `${mes}/${ano}`;
    const dataEncerramento = document.getElementById('input-data-encerramento-captura').value;
    const cpf = document.getElementById('input-cpf-captura').value.trim();
    const senha = document.getElementById('input-senha-captura').value;
    const chromePath = document.getElementById('input-chrome-captura').value.trim();
    const urlPortal = document.getElementById('input-url-captura').value.trim();

    if (!planilha) { addLogCaptura('Selecione a planilha fiscal.', true); return; }
    if (!saida) { addLogCaptura('Selecione a pasta de saída.', true); return; }
    if (!dataEncerramento) { addLogCaptura('Informe a data de encerramento (filtro).', true); return; }
    if (!cpf || !senha) { addLogCaptura('Informe o CPF e a senha do portal ISS Fortaleza.', true); return; }

    document.getElementById('log-captura').innerHTML = '';
    addLogCaptura(`Iniciando captura de escrituração — competência ${competencia}...`);

    document.getElementById('btn-iniciar-captura').disabled = true;
    document.getElementById('btn-cancelar-captura').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_captura_com_movimento_gui(
                planilha, saida, competencia, dataEncerramento, chromePath, urlPortal, cpf, senha
            );
        } catch (e) {
            addLogCaptura('Erro ao iniciar a captura: ' + e, true);
        }
    }

    document.getElementById('btn-iniciar-captura').disabled = false;
    document.getElementById('btn-cancelar-captura').classList.add('hidden');
}

function cancelarCapturaComMovimento() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        "Tem certeza que deseja cancelar a captura?\n\nO navegador será fechado e o processo precisará ser reiniciado do zero."
    );
    if (!confirmado) return;

    window.pywebview.api.cancelar_captura_com_movimento();
    addLogCaptura('Cancelamento solicitado pelo operador. Encerrando...', true);
    document.getElementById('btn-cancelar-captura').disabled = true;
}

function addLogCaptura(msg, isError = false) {
    const el = document.getElementById('log-captura');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_captura = addLogCaptura;

function limparLogCaptura() {
    const el = document.getElementById('log-captura');
    if (el) el.innerHTML = '';
}
