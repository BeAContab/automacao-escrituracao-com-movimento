// ---- Encerramento ISS — Múltiplos Meses ----
// Mesma automação do Encerramento ISS, mas para uma lista de competências: o operador
// adiciona um ou mais meses/anos antes de iniciar.

let competenciasEncerramentoMulti = []; // lista de strings "MM/AAAA", na ordem em que foram adicionadas

async function selecionarPlanilhaEncerramentoMulti() {
    if (window.pywebview) {
        const file = await window.pywebview.api.selecionar_arquivo();
        if (file) {
            document.getElementById('input-planilha-encerramento-multi').value = file;
        }
    }
}

async function selecionarSaidaEncerramentoMulti() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            document.getElementById('input-saida-encerramento-multi').value = pasta;
        }
    }
}

function nomeDoMes(mes) {
    const nomes = {
        '01': 'Janeiro', '02': 'Fevereiro', '03': 'Março', '04': 'Abril', '05': 'Maio', '06': 'Junho',
        '07': 'Julho', '08': 'Agosto', '09': 'Setembro', '10': 'Outubro', '11': 'Novembro', '12': 'Dezembro',
    };
    return nomes[mes] || mes;
}

function adicionarCompetenciaEncerramentoMulti() {
    const mes = document.getElementById('select-mes-encerramento-multi').value;
    const ano = document.getElementById('select-ano-encerramento-multi').value;
    const competencia = `${mes}/${ano}`;

    if (competenciasEncerramentoMulti.includes(competencia)) {
        addLogEncerramentoMulti(`A competência ${nomeDoMes(mes)}/${ano} já está na lista.`, true);
        return;
    }

    competenciasEncerramentoMulti.push(competencia);
    // Ordem cronológica na lista visual, independente da ordem em que foram clicadas
    competenciasEncerramentoMulti.sort((a, b) => {
        const [ma, aa] = a.split('/');
        const [mb, ab] = b.split('/');
        return aa === ab ? ma.localeCompare(mb) : aa.localeCompare(ab);
    });
    renderizarCompetenciasEncerramentoMulti();
}

function removerCompetenciaEncerramentoMulti(competencia) {
    competenciasEncerramentoMulti = competenciasEncerramentoMulti.filter(c => c !== competencia);
    renderizarCompetenciasEncerramentoMulti();
}

function renderizarCompetenciasEncerramentoMulti() {
    const container = document.getElementById('lista-competencias-encerramento-multi');
    if (!container) return;
    container.innerHTML = '';

    if (competenciasEncerramentoMulti.length === 0) {
        const vazio = document.createElement('p');
        vazio.className = 'font-body-sm text-body-sm text-on-surface-variant/70 italic';
        vazio.innerText = 'Nenhuma competência adicionada ainda.';
        container.appendChild(vazio);
        return;
    }

    competenciasEncerramentoMulti.forEach(competencia => {
        const [mes, ano] = competencia.split('/');
        const chip = document.createElement('span');
        chip.className = 'inline-flex items-center gap-1 pl-sm pr-1 py-1 bg-primary-container/15 text-primary-container rounded-full text-xs font-bold';

        const texto = document.createElement('span');
        texto.innerText = `${nomeDoMes(mes)}/${ano}`;
        chip.appendChild(texto);

        const remover = document.createElement('button');
        remover.type = 'button';
        remover.title = 'Remover';
        remover.className = 'material-symbols-outlined text-[16px] leading-none hover:text-error transition-colors';
        remover.innerText = 'close';
        remover.onclick = () => removerCompetenciaEncerramentoMulti(competencia);
        chip.appendChild(remover);

        container.appendChild(chip);
    });
}

async function iniciarEncerramentoMultiPeriodo() {
    const planilha = document.getElementById('input-planilha-encerramento-multi').value.trim();
    const saida = document.getElementById('input-saida-encerramento-multi').value.trim();
    const cpf = document.getElementById('input-cpf-encerramento-multi').value.trim();
    const senha = document.getElementById('input-senha-encerramento-multi').value;
    const chromePath = document.getElementById('input-chrome-encerramento-multi').value.trim();
    const urlPortal = document.getElementById('input-url-encerramento-multi').value.trim();

    if (!planilha) { addLogEncerramentoMulti('Selecione a planilha fiscal.', true); return; }
    if (!saida) { addLogEncerramentoMulti('Selecione a pasta de saída.', true); return; }
    if (competenciasEncerramentoMulti.length === 0) { addLogEncerramentoMulti('Adicione ao menos uma competência.', true); return; }
    if (!cpf || !senha) { addLogEncerramentoMulti('Informe o CPF e a senha do portal ISS Fortaleza.', true); return; }

    document.getElementById('log-encerramento-multi').innerHTML = '';
    document.getElementById('resumo-encerramento-multi').classList.add('hidden');
    document.getElementById('btn-abrir-relatorios-encerramento-multi').classList.add('hidden');
    addLogEncerramentoMulti(`Iniciando encerramento ISS — ${competenciasEncerramentoMulti.length} competência(s): ${competenciasEncerramentoMulti.join(', ')}...`);

    document.getElementById('btn-iniciar-encerramento-multi').disabled = true;
    document.getElementById('btn-cancelar-encerramento-multi').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_encerramento_multi_periodo_gui(
                planilha, saida, competenciasEncerramentoMulti, chromePath, urlPortal, cpf, senha
            );
        } catch (e) {
            addLogEncerramentoMulti('Erro ao iniciar o encerramento: ' + e, true);
        }
    }

    document.getElementById('btn-iniciar-encerramento-multi').disabled = false;
    document.getElementById('btn-cancelar-encerramento-multi').classList.add('hidden');
}

function cancelarEncerramentoMultiPeriodo() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        "Tem certeza que deseja cancelar o encerramento?\n\nO navegador será fechado e o processo precisará ser reiniciado do zero."
    );
    if (!confirmado) return;

    window.pywebview.api.cancelar_encerramento_multi_periodo();
    addLogEncerramentoMulti('Cancelamento solicitado pelo operador. Encerrando...', true);
    document.getElementById('btn-cancelar-encerramento-multi').disabled = true;
}

function mostrarResumoEncerramentoMulti(resumo) {
    document.getElementById('resumo-multi-processadas').innerText = resumo.processadas;
    document.getElementById('resumo-multi-encerradas').innerText = resumo.encerradas;
    document.getElementById('resumo-multi-problemas').innerText = resumo.problemas;
    document.getElementById('resumo-encerramento-multi').classList.remove('hidden');

    window._ultimaPastaRelatoriosEncerramentoMulti = resumo.pasta_saida;
    const btnRelatorios = document.getElementById('btn-abrir-relatorios-encerramento-multi');
    if (resumo.report_paths && resumo.report_paths.length > 0) {
        btnRelatorios.innerText = `Abrir Pasta com ${resumo.report_paths.length} Relatório(s) Gerado(s)`;
        btnRelatorios.classList.remove('hidden');
    } else {
        btnRelatorios.classList.add('hidden');
    }
}
window.mostrar_resumo_encerramento_multi = mostrarResumoEncerramentoMulti;

function abrirPastaRelatoriosEncerramentoMulti() {
    if (window._ultimaPastaRelatoriosEncerramentoMulti && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimaPastaRelatoriosEncerramentoMulti.replace(/\\/g, '/'));
    }
}

function addLogEncerramentoMulti(msg, isError = false) {
    const el = document.getElementById('log-encerramento-multi');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_encerramento_multi = addLogEncerramentoMulti;

function limparLogEncerramentoMulti() {
    const el = document.getElementById('log-encerramento-multi');
    if (el) el.innerHTML = '';
}
