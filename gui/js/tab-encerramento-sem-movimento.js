// ---- Encerramento ISS (Sem Movimento) ----

async function selecionarPlanilhaEncerramento() {
    if (window.pywebview) {
        const file = await window.pywebview.api.selecionar_arquivo();
        if (file) {
            document.getElementById('input-planilha-encerramento').value = file;
        }
    }
}

async function selecionarSaidaEncerramento() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            document.getElementById('input-saida-encerramento').value = pasta;
        }
    }
}

async function iniciarEncerramentoSemMovimento() {
    const planilha = document.getElementById('input-planilha-encerramento').value.trim();
    const saida = document.getElementById('input-saida-encerramento').value.trim();
    const mes = document.getElementById('select-mes-encerramento').value;
    const ano = document.getElementById('select-ano-encerramento').value;
    const competencia = `${mes}/${ano}`;
    const cpf = document.getElementById('input-cpf-encerramento').value.trim();
    const senha = document.getElementById('input-senha-encerramento').value;
    const chromePath = document.getElementById('input-chrome-encerramento').value.trim();
    const urlPortal = document.getElementById('input-url-encerramento').value.trim();

    if (!planilha) { addLogEncerramento('Selecione a planilha fiscal.', true); return; }
    if (!saida) { addLogEncerramento('Selecione a pasta de saída.', true); return; }
    if (!cpf || !senha) { addLogEncerramento('Informe o CPF e a senha do portal ISS Fortaleza.', true); return; }

    document.getElementById('log-encerramento').innerHTML = '';
    document.getElementById('resumo-encerramento').classList.add('hidden');
    document.getElementById('btn-abrir-relatorio-encerramento').classList.add('hidden');
    addLogEncerramento(`Iniciando encerramento ISS — competência ${competencia}...`);

    document.getElementById('btn-iniciar-encerramento').disabled = true;
    document.getElementById('btn-cancelar-encerramento').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_encerramento_sem_movimento_gui(
                planilha, saida, competencia, chromePath, urlPortal, cpf, senha
            );
        } catch (e) {
            addLogEncerramento('Erro ao iniciar o encerramento: ' + e, true);
        }
    }

    document.getElementById('btn-iniciar-encerramento').disabled = false;
    document.getElementById('btn-cancelar-encerramento').classList.add('hidden');
}

function cancelarEncerramentoSemMovimento() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        "Tem certeza que deseja cancelar o encerramento?\n\nO navegador será fechado e o processo precisará ser reiniciado do zero."
    );
    if (!confirmado) return;

    window.pywebview.api.cancelar_encerramento_sem_movimento();
    addLogEncerramento('Cancelamento solicitado pelo operador. Encerrando...', true);
    document.getElementById('btn-cancelar-encerramento').disabled = true;
}

function mostrarResumoEncerramento(resumo) {
    document.getElementById('resumo-processadas').innerText = resumo.processadas;
    document.getElementById('resumo-encerradas').innerText = resumo.encerradas;
    document.getElementById('resumo-problemas').innerText = resumo.problemas;
    document.getElementById('resumo-encerramento').classList.remove('hidden');

    window._ultimoRelatorioEncerramento = resumo.report_path;
    const btnRelatorio = document.getElementById('btn-abrir-relatorio-encerramento');
    if (resumo.report_path) {
        btnRelatorio.classList.remove('hidden');
    } else {
        btnRelatorio.classList.add('hidden');
    }
}
window.mostrar_resumo_encerramento = mostrarResumoEncerramento;

function abrirRelatorioEncerramento() {
    if (window._ultimoRelatorioEncerramento && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimoRelatorioEncerramento.replace(/\\/g, '/'));
    }
}

function addLogEncerramento(msg, isError = false) {
    const el = document.getElementById('log-encerramento');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_encerramento = addLogEncerramento;

function limparLogEncerramento() {
    const el = document.getElementById('log-encerramento');
    if (el) el.innerHTML = '';
}
