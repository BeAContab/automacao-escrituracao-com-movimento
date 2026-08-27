// ---- Automação: Escrituração ----

async function selecionarExcelAutomacao() {
    if (window.pywebview) {
        const file = await window.pywebview.api.selecionar_arquivo();
        if (file) {
            document.getElementById('input-excel-automacao').value = file;
        }
    }
}

let automacaoPausada = false;

async function iniciarAutomacao() {
    const inputExcel = document.getElementById('input-excel-automacao').value;
    const mes = document.getElementById('select-mes-automacao').value;
    const ano = document.getElementById('select-ano-automacao').value;
    const competencia = `${mes}/${ano}`;

    if (!inputExcel) {
        addLogAutomacao("Selecione o arquivo Excel de origem.", true);
        return;
    }

    // MELHORIA-06: Aviso visual se o arquivo não foi selecionado nesta sessão
    if (!window._arquivoExcelSelecionadoNestaSessao) {
        addLogAutomacao("Atenção: verifique se o arquivo Excel selecionado ainda existe no caminho informado antes de continuar.", false);
    }

    // Configura e exibe a barra de progresso da automação
    const container = document.getElementById('progresso-container-automacao');
    const bar = document.getElementById('progresso-barra-automacao');
    const pct = document.getElementById('progresso-porcentagem-automacao');
    const st = document.getElementById('progresso-status-automacao');

    container.classList.remove('hidden');
    bar.style.width = '0%';
    pct.innerText = '0%';
    st.innerText = 'Iniciando automação...';

    // Garante que o botão comece ocultado e no estado não-pausado
    document.getElementById('btn-pausar-retomar-automacao').classList.add('hidden');
    automacaoPausada = false;
    const btnPausaAut = document.getElementById('btn-pausar-retomar-automacao');
    const lblPausaAut = document.getElementById('lbl-pausar-retomar-automacao');
    const iconPausaAut = btnPausaAut.querySelector('.material-symbols-outlined');
    lblPausaAut.innerText = "PARAR";
    iconPausaAut.innerText = "pause";
    btnPausaAut.className = "hidden px-sm py-1 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded text-[10px] font-bold tracking-wider transition-all flex items-center gap-1";

    // Diferente do botão de pausa (só aparece após confirmar login), o de Cancelar
    // já aparece aqui — o caso mais importante para cancelar é durante a própria
    // espera da confirmação de login.
    const btnCancelar = document.getElementById('btn-cancelar-automacao');
    btnCancelar.classList.remove('hidden');
    btnCancelar.disabled = false;

    // FALHA-08: O card de confirmação de login é exibido APÓS a chamada da API Python,
    // não antes — evita que o card fique preso visível em caso de erro precoce.
    addLogAutomacao(`Iniciando portal da ISS Fortaleza (Competência: ${competencia})...`);
    document.getElementById('btn-iniciar-automacao').disabled = true;

    if (window.pywebview) {
        await window.pywebview.api.iniciar_automacao(inputExcel, competencia);
    }
    document.getElementById('btn-iniciar-automacao').disabled = false;
}

async function confirmarLoginFeito() {
    if (window.pywebview && window.pywebview.api) {
        const ok = await window.pywebview.api.confirmar_login_feito();
        if (ok) {
            addLogAutomacao("Login verificado pelo operador. Retomando robô...");
            // FALHA-08: O card só é ocultado AQUI (após confirmação OK) e é exibido pela API Python
            document.getElementById('card-confirmar-login').classList.add('hidden');
            document.getElementById('btn-pausar-retomar-automacao').classList.remove('hidden');
        } else {
            addLogAutomacao("Falha ao enviar sinal de confirmação do login.", true);
        }
    }
}

function alternarPausaAutomacao() {
    if (!window.pywebview || !window.pywebview.api) return;
    automacaoPausada = !automacaoPausada;
    const lbl = document.getElementById('lbl-pausar-retomar-automacao');
    const btn = document.getElementById('btn-pausar-retomar-automacao');
    const icon = btn.querySelector('.material-symbols-outlined');

    if (automacaoPausada) {
        window.pywebview.api.pausar_automacao();
        lbl.innerText = "CONTINUAR";
        icon.innerText = "play_arrow";
        btn.className = "px-sm py-1 bg-green-500/20 hover:bg-green-500/30 border border-green-500/40 text-white rounded text-[10px] font-bold tracking-wider transition-all flex items-center gap-1";
        addLogAutomacao("Automação PAUSADA pelo operador.");
    } else {
        window.pywebview.api.retomar_automacao();
        lbl.innerText = "PARAR";
        icon.innerText = "pause";
        btn.className = "px-sm py-1 bg-white/10 hover:bg-white/20 border border-white/20 text-white rounded text-[10px] font-bold tracking-wider transition-all flex items-center gap-1";
        addLogAutomacao("Automação RETOMADA pelo operador.");
    }
}

function cancelarAutomacao() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        "Tem certeza que deseja cancelar a automação?\n\n" +
        "O navegador será fechado e você precisará começar do zero: selecionar a " +
        "planilha e a competência novamente, e refazer o login manual."
    );
    if (!confirmado) return;

    window.pywebview.api.cancelar_automacao();
    addLogAutomacao("Cancelamento solicitado pelo operador. Encerrando a automação...", true);
    document.getElementById('btn-cancelar-automacao').disabled = true;
}

function addLogAutomacao(msg, isError = false) {
    const logContainer = document.getElementById('log-automacao');
    if (!logContainer) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? "text-error font-bold" : "";
    logContainer.appendChild(p);
    logContainer.scrollTop = logContainer.scrollHeight;
}
window.log_automacao = addLogAutomacao;

function limparLogAutomacao() {
    const el = document.getElementById('log-automacao');
    if (el) el.innerHTML = '';
}

function updateProgressoAutomacao(porcentagem, status) {
    const container = document.getElementById('progresso-container-automacao');
    const bar = document.getElementById('progresso-barra-automacao');
    const pct = document.getElementById('progresso-porcentagem-automacao');
    const st = document.getElementById('progresso-status-automacao');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_automacao = updateProgressoAutomacao;
