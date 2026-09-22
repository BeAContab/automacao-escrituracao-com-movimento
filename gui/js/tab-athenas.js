// ---- Importar para Athenas ----
// Converte a escrituração exportada do ISS Fortaleza para o layout de importação do
// Athenas ERP. Não usa portal nem navegador — só leitura/escrita de planilha, sem
// pausar/continuar/parar (cada arquivo leva segundos).

let arquivosAthenas = []; // [{ caminho, nome, regime: 'normal' | 'simples' }]

const OPCOES_REGIME_ATHENAS = [
    { valor: 'normal', rotulo: 'Regime Normal (8010/8011)' },
    { valor: 'simples', rotulo: 'Simples Nacional (8001/8003)' },
];

async function adicionarArquivosAthenas() {
    if (!window.pywebview || !window.pywebview.api) return;
    const caminhos = await window.pywebview.api.selecionar_arquivos_athenas();
    if (!caminhos || caminhos.length === 0) return;

    let adicionados = 0;
    caminhos.forEach(caminho => {
        if (arquivosAthenas.some(a => a.caminho === caminho)) return; // ignora duplicata
        const nome = caminho.split(/[\\/]/).pop();
        arquivosAthenas.push({ caminho, nome, regime: 'normal' });
        adicionados++;
    });

    // Ao adicionar o primeiro arquivo, preenche a pasta de saída com a pasta dele
    const pastaSaida = document.getElementById('input-pasta-athenas');
    if (adicionados > 0 && !pastaSaida.value.trim() && caminhos[0]) {
        const partes = caminhos[0].split(/[\\/]/);
        partes.pop();
        pastaSaida.value = partes.join('\\');
    }

    renderizarListaAthenas();
}

function removerArquivoAthenas(caminho) {
    arquivosAthenas = arquivosAthenas.filter(a => a.caminho !== caminho);
    renderizarListaAthenas();
}

function limparListaAthenas() {
    arquivosAthenas = [];
    renderizarListaAthenas();
}

function definirRegimeAthenas(caminho, regime) {
    const item = arquivosAthenas.find(a => a.caminho === caminho);
    if (item) item.regime = regime;
}

function renderizarListaAthenas() {
    const container = document.getElementById('lista-arquivos-athenas');
    if (!container) return;
    container.innerHTML = '';

    if (arquivosAthenas.length === 0) {
        const vazio = document.createElement('p');
        vazio.id = 'lista-arquivos-athenas-vazia';
        vazio.className = 'font-body-sm text-body-sm text-on-surface-variant/70 italic p-sm';
        vazio.innerText = 'Nenhum arquivo adicionado. Clique em "Adicionar Arquivos".';
        container.appendChild(vazio);
        return;
    }

    arquivosAthenas.forEach(item => {
        const linha = document.createElement('div');
        linha.className = 'flex items-center gap-sm p-sm rounded-lg bg-surface-container-low';

        const nome = document.createElement('span');
        nome.className = 'flex-1 font-body-sm text-body-sm text-on-surface truncate';
        nome.title = item.caminho;
        nome.innerText = item.nome;
        linha.appendChild(nome);

        const select = document.createElement('select');
        select.className = 'input-border rounded-lg py-1 px-2 text-xs outline-none bg-white';
        OPCOES_REGIME_ATHENAS.forEach(op => {
            const opt = document.createElement('option');
            opt.value = op.valor;
            opt.innerText = op.rotulo;
            if (op.valor === item.regime) opt.selected = true;
            select.appendChild(opt);
        });
        select.onchange = () => definirRegimeAthenas(item.caminho, select.value);
        linha.appendChild(select);

        const remover = document.createElement('button');
        remover.type = 'button';
        remover.title = 'Remover';
        remover.className = 'material-symbols-outlined text-[18px] text-error hover:opacity-70 transition-opacity';
        remover.innerText = 'close';
        remover.onclick = () => removerArquivoAthenas(item.caminho);
        linha.appendChild(remover);

        container.appendChild(linha);
    });
}

async function selecionarPastaAthenas() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            document.getElementById('input-pasta-athenas').value = pasta;
        }
    }
}

async function gerarPlanilhasAthenas() {
    const pasta = document.getElementById('input-pasta-athenas').value.trim();

    if (arquivosAthenas.length === 0) {
        addLogAthenas('Adicione ao menos um arquivo do ISS Fortaleza.', true);
        return;
    }
    if (!pasta) {
        addLogAthenas('Selecione a pasta de saída.', true);
        return;
    }

    document.getElementById('log-athenas').innerHTML = '';
    document.getElementById('resumo-athenas').classList.add('hidden');
    document.getElementById('btn-abrir-log-erros-athenas').classList.add('hidden');
    document.getElementById('progresso-container-athenas').classList.remove('hidden');
    addLogAthenas(`Iniciando importação de ${arquivosAthenas.length} arquivo(s) para o Athenas...`);

    const btn = document.getElementById('btn-gerar-athenas');
    btn.disabled = true;
    btn.innerHTML = '<span class="material-symbols-outlined" style="animation: spin 1s linear infinite; display:inline-block;">progress_activity</span> Processando...';

    if (window.pywebview && window.pywebview.api) {
        try {
            const caminhos = arquivosAthenas.map(a => a.caminho);
            const regimes = arquivosAthenas.map(a => a.regime);
            await window.pywebview.api.iniciar_athenas_gui(caminhos, regimes, pasta);
        } catch (e) {
            addLogAthenas('Erro ao iniciar a importação: ' + e, true);
            athenasFinalizado();
        }
    }
    // O botão volta ao normal quando o Python avisa o fim (athenasFinalizado).
}

function athenasFinalizado() {
    const btn = document.getElementById('btn-gerar-athenas');
    btn.disabled = false;
    btn.innerHTML = '<span class="material-symbols-outlined">file_download</span> Gerar Planilhas';
}
window.athenas_finalizado = athenasFinalizado;

function mostrarResumoAthenas(resumo) {
    document.getElementById('resumo-athenas-tomados').innerText = resumo.tomados;
    document.getElementById('resumo-athenas-prestados').innerText = resumo.prestados;
    document.getElementById('resumo-athenas-erros').innerText = resumo.erros;
    document.getElementById('resumo-athenas').classList.remove('hidden');

    window._ultimaPastaSaidaAthenas = resumo.pasta_saida;
    window._ultimoLogErrosAthenas = resumo.caminho_log_erros;
    const btnLogErros = document.getElementById('btn-abrir-log-erros-athenas');
    btnLogErros.classList.toggle('hidden', !resumo.caminho_log_erros);
}
window.mostrar_resumo_athenas = mostrarResumoAthenas;

function abrirPastaSaidaAthenas() {
    if (window._ultimaPastaSaidaAthenas && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimaPastaSaidaAthenas.replace(/\\/g, '/'));
    }
}

function abrirLogErrosAthenas() {
    if (window._ultimoLogErrosAthenas && window.pywebview && window.pywebview.api) {
        window.pywebview.api.abrir_link('file:///' + window._ultimoLogErrosAthenas.replace(/\\/g, '/'));
    }
}

function addLogAthenas(msg, isError = false) {
    const el = document.getElementById('log-athenas');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_athenas = addLogAthenas;

function limparLogAthenas() {
    const el = document.getElementById('log-athenas');
    if (el) el.innerHTML = '';
}

function updateProgressoAthenas(porcentagem, status) {
    const container = document.getElementById('progresso-container-athenas');
    const bar = document.getElementById('progresso-barra-athenas');
    const pct = document.getElementById('progresso-porcentagem-athenas');
    const st = document.getElementById('progresso-status-athenas');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_athenas = updateProgressoAthenas;
