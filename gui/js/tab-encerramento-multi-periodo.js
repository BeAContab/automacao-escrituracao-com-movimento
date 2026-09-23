// ---- Encerramento ISS — Múltiplos Meses ----
// Mesma automação do Encerramento ISS, mas para uma lista de competências: o operador
// adiciona um ou mais meses/anos antes de iniciar.
//
// Dois modos: "Apenas verificar" (padrão; confere tudo sem encerrar nem baixar nada e mostra uma
// lista para o operador aprovar) e "Encerrar" (encerra de verdade). Pausar/Continuar/Parar agem
// nos limites seguros do robô (entre competências/empresas — nunca no meio de um encerramento).

let competenciasEncerramentoMulti = []; // lista de strings "MM/AAAA", na ordem em que foram adicionadas
let encerramentoMultiPausado = false;
let verificacaoEncerramentoMulti = []; // itens da última verificação (vindos do Python)

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

function modoEncerramentoMulti() {
    const marcado = document.querySelector('input[name="modo-encerramento-multi"]:checked');
    return marcado ? marcado.value : 'verificar';
}

function atualizarBotaoIniciarEncerramentoMulti() {
    const verificar = modoEncerramentoMulti() === 'verificar';
    document.getElementById('lbl-btn-iniciar-encerramento-multi').innerText = verificar ? 'Verificar (não encerra)' : 'Realizar Encerramento';
    document.getElementById('icone-btn-iniciar-encerramento-multi').innerText = verificar ? 'fact_check' : 'send';
}

// Estado da tela enquanto o robô roda: botão de iniciar desligado, controles de pausa/parada visíveis.
function definirExecutandoEncerramentoMulti(executando) {
    document.getElementById('btn-iniciar-encerramento-multi').disabled = executando;
    const executar = document.getElementById('btn-executar-verificadas-encerramento-multi');
    if (executar) executar.disabled = executando;
    document.getElementById('controles-encerramento-multi').classList.toggle('hidden', !executando);
    encerramentoMultiPausado = false;
    document.getElementById('lbl-pausar-encerramento-multi').innerText = 'Pausar';
    document.getElementById('icone-pausar-encerramento-multi').innerText = 'pause';
    document.getElementById('btn-parar-encerramento-multi').disabled = false;
}

// Chamado pelo Python quando a execução termina (com sucesso, erro ou parada).
function encerramentoMultiFinalizado() {
    definirExecutandoEncerramentoMulti(false);
}
window.encerramento_multi_finalizado = encerramentoMultiFinalizado;

function lerCamposEncerramentoMulti() {
    return {
        planilha: document.getElementById('input-planilha-encerramento-multi').value.trim(),
        saida: document.getElementById('input-saida-encerramento-multi').value.trim(),
        cpf: document.getElementById('input-cpf-encerramento-multi').value.trim(),
        senha: document.getElementById('input-senha-encerramento-multi').value,
        chromePath: document.getElementById('input-chrome-encerramento-multi').value.trim(),
        urlPortal: document.getElementById('input-url-encerramento-multi').value.trim(),
    };
}

async function iniciarEncerramentoMultiPeriodo() {
    const { planilha, saida, cpf, senha, chromePath, urlPortal } = lerCamposEncerramentoMulti();
    const modo = modoEncerramentoMulti();

    if (!planilha) { addLogEncerramentoMulti('Selecione a planilha fiscal.', true); return; }
    if (!saida) { addLogEncerramentoMulti('Selecione a pasta de saída.', true); return; }
    if (competenciasEncerramentoMulti.length === 0) { addLogEncerramentoMulti('Adicione ao menos uma competência.', true); return; }
    if (!cpf || !senha) { addLogEncerramentoMulti('Informe o CPF e a senha do portal ISS Fortaleza.', true); return; }

    if (modo === 'encerrar') {
        const confirmado = confirm(
            'ENCERRAR de verdade?\n\n' +
            `Competências: ${competenciasEncerramentoMulti.join(', ')}\n\n` +
            'O robô vai ENCERRAR no portal as escriturações aptas de todas as empresas sem movimento da planilha ' +
            'e baixar os certificados. O encerramento é irreversível.\n\n' +
            'Dica: use "Apenas verificar" antes para conferir o que seria encerrado.'
        );
        if (!confirmado) return;
    }

    document.getElementById('log-encerramento-multi').innerHTML = '';
    document.getElementById('resumo-encerramento-multi').classList.add('hidden');
    document.getElementById('verificacao-encerramento-multi').classList.add('hidden');
    document.getElementById('btn-abrir-relatorios-encerramento-multi').classList.add('hidden');
    verificacaoEncerramentoMulti = [];
    addLogEncerramentoMulti(
        `${modo === 'verificar' ? 'Iniciando VERIFICAÇÃO (nada será encerrado)' : 'Iniciando encerramento ISS'} — ` +
        `${competenciasEncerramentoMulti.length} competência(s): ${competenciasEncerramentoMulti.join(', ')}...`
    );

    definirExecutandoEncerramentoMulti(true);

    if (window.pywebview && window.pywebview.api) {
        try {
            const iniciou = await window.pywebview.api.iniciar_encerramento_multi_periodo_gui(
                planilha, saida, competenciasEncerramentoMulti, chromePath, urlPortal, cpf, senha, modo
            );
            if (iniciou === false) encerramentoMultiFinalizado(); // já havia uma execução em andamento
        } catch (e) {
            addLogEncerramentoMulti('Erro ao iniciar o encerramento: ' + e, true);
            encerramentoMultiFinalizado();
        }
    }
    // Os controles voltam ao normal quando o Python avisa o fim (encerramento_multi_finalizado).
}

function alternarPausaEncerramentoMulti() {
    if (!window.pywebview || !window.pywebview.api) return;
    encerramentoMultiPausado = !encerramentoMultiPausado;
    if (encerramentoMultiPausado) {
        window.pywebview.api.pausar_processamento_xml('encerramento_multi');
        document.getElementById('lbl-pausar-encerramento-multi').innerText = 'Continuar';
        document.getElementById('icone-pausar-encerramento-multi').innerText = 'play_arrow';
    } else {
        window.pywebview.api.retomar_processamento_xml('encerramento_multi');
        document.getElementById('lbl-pausar-encerramento-multi').innerText = 'Pausar';
        document.getElementById('icone-pausar-encerramento-multi').innerText = 'pause';
    }
}

function pararEncerramentoMulti() {
    if (!window.pywebview || !window.pywebview.api) return;
    const confirmado = confirm(
        'Parar a execução?\n\nO robô termina a competência em andamento (nunca para no meio de um encerramento), ' +
        'gera os relatórios do que já foi feito e fecha o navegador. Para continuar depois será preciso iniciar de novo; ' +
        'o que já foi encerrado é reconhecido no portal.'
    );
    if (!confirmado) return;
    window.pywebview.api.cancelar_processamento_xml('encerramento_multi');
    document.getElementById('btn-parar-encerramento-multi').disabled = true;
}

// Compatibilidade com quem ainda chamar o nome antigo do botão "Cancelar".
function cancelarEncerramentoMultiPeriodo() {
    pararEncerramentoMulti();
}

// ---- Revisão da verificação ---------------------------------------------------------------------------

const GRUPOS_VERIFICACAO_ENCERRAMENTO_MULTI = [
    { categoria: 'apta', titulo: 'Aptas a encerrar', classe: 'text-secondary' },
    { categoria: 'ja_encerrada_sem_certificado', titulo: 'Já encerradas — falta baixar o certificado', classe: 'text-secondary' },
    { categoria: 'ja_encerrada_certificado_existente', titulo: 'Já encerradas — certificado já existe na pasta', classe: 'text-on-surface-variant' },
    { categoria: 'problema', titulo: 'Com problema (não serão encerradas)', classe: 'text-error' },
];

function itemComPendenciaNoPortal(item) {
    return /pend[eê]ncia/i.test(item.situacao || '');
}

function renderizarVerificacaoEncerramentoMulti() {
    const painel = document.getElementById('verificacao-encerramento-multi');
    const lista = document.getElementById('lista-verificacao-encerramento-multi');
    lista.innerHTML = '';

    if (verificacaoEncerramentoMulti.length === 0) {
        painel.classList.add('hidden');
        return;
    }

    GRUPOS_VERIFICACAO_ENCERRAMENTO_MULTI.forEach(grupo => {
        const itens = verificacaoEncerramentoMulti
            .map((item, indice) => ({ item, indice }))
            .filter(({ item }) => item.categoria === grupo.categoria);
        if (itens.length === 0) return;

        const titulo = document.createElement('p');
        titulo.className = `font-label-md text-label-md font-bold ${grupo.classe} mt-xs`;
        titulo.innerText = `${grupo.titulo} (${itens.length})`;
        lista.appendChild(titulo);

        itens.forEach(({ item, indice }) => {
            const linha = document.createElement('label');
            linha.className = 'flex items-start gap-sm p-sm rounded-lg bg-surface-container-low';

            if (item.executavel) {
                const caixa = document.createElement('input');
                caixa.type = 'checkbox';
                caixa.className = 'mt-1 accent-primary-container caixa-verificacao-encerramento-multi';
                caixa.dataset.indice = String(indice);
                // "Aberta - Com Pendência": vem desmarcada, para o operador decidir conscientemente.
                caixa.checked = !itemComPendenciaNoPortal(item);
                caixa.onchange = atualizarContadorExecutarVerificadas;
                linha.appendChild(caixa);
            } else {
                const espaco = document.createElement('span');
                espaco.className = 'w-4';
                linha.appendChild(espaco);
            }

            const texto = document.createElement('span');
            texto.className = 'flex-1 font-body-sm text-body-sm text-on-surface';
            const [mes, ano] = item.competencia.split('/');
            let detalhe = `${item.nome} — ${item.cnpj_m} · ${nomeDoMes(mes)}/${ano}`;
            if (item.categoria === 'problema' && item.motivo) detalhe += ` — ${item.motivo}`;
            if (item.situacao) detalhe += ` (portal: ${item.situacao})`;
            texto.innerText = detalhe;
            linha.appendChild(texto);

            if (item.executavel && itemComPendenciaNoPortal(item)) {
                const alerta = document.createElement('span');
                alerta.className = 'px-sm py-0.5 bg-error/10 text-error border border-error/30 rounded-full text-[10px] font-bold uppercase';
                alerta.innerText = 'Com pendência';
                linha.appendChild(alerta);
            }

            lista.appendChild(linha);
        });
    });

    painel.classList.remove('hidden');
    atualizarContadorExecutarVerificadas();
}

function itensSelecionadosVerificacaoEncerramentoMulti() {
    return Array.from(document.querySelectorAll('.caixa-verificacao-encerramento-multi:checked'))
        .map(caixa => verificacaoEncerramentoMulti[Number(caixa.dataset.indice)]);
}

function atualizarContadorExecutarVerificadas() {
    const total = itensSelecionadosVerificacaoEncerramentoMulti().length;
    document.getElementById('lbl-btn-executar-verificadas-encerramento-multi').innerText = `Executar o que foi verificado (${total})`;
    document.getElementById('btn-executar-verificadas-encerramento-multi').classList.toggle('opacity-50', total === 0);
}

async function executarVerificadasEncerramentoMulti() {
    const selecionados = itensSelecionadosVerificacaoEncerramentoMulti();
    if (selecionados.length === 0) { addLogEncerramentoMulti('Marque ao menos um item da lista para executar.', true); return; }

    const { planilha, saida, cpf, senha, chromePath, urlPortal } = lerCamposEncerramentoMulti();
    if (!planilha || !saida) { addLogEncerramentoMulti('A planilha fiscal e a pasta de saída precisam continuar selecionadas.', true); return; }
    if (!cpf || !senha) { addLogEncerramentoMulti('Informe o CPF e a senha do portal ISS Fortaleza.', true); return; }

    const aEncerrar = selecionados.filter(i => i.categoria === 'apta').length;
    const certificados = selecionados.filter(i => i.categoria === 'ja_encerrada_sem_certificado').length;
    const confirmado = confirm(
        'Executar o que foi verificado?\n\n' +
        `• ${aEncerrar} competência(s) serão ENCERRADAS no portal (irreversível)\n` +
        `• ${certificados} certificado(s) de competências já encerradas serão baixados\n\n` +
        'Será aberto um novo Chrome com novo login, e cada item é conferido de novo no portal antes de ser encerrado ' +
        '(se algo mudou desde a verificação, o item é registrado como problema e NÃO é encerrado).'
    );
    if (!confirmado) return;

    const itens = selecionados.map(i => ({ cnpj: i.cnpj, competencia: i.competencia }));
    document.getElementById('log-encerramento-multi').innerHTML = '';
    addLogEncerramentoMulti(`Executando ${itens.length} item(ns) aprovado(s) na verificação...`);
    definirExecutandoEncerramentoMulti(true);

    if (window.pywebview && window.pywebview.api) {
        try {
            const iniciou = await window.pywebview.api.executar_verificadas_encerramento_multi_gui(
                planilha, saida, itens, chromePath, urlPortal, cpf, senha
            );
            if (iniciou === false) encerramentoMultiFinalizado();
        } catch (e) {
            addLogEncerramentoMulti('Erro ao iniciar a execução: ' + e, true);
            encerramentoMultiFinalizado();
        }
    }
}

// ---- Resumo -------------------------------------------------------------------------------------------------

function mostrarResumoEncerramentoMulti(resumo) {
    const verificacao = resumo.modo === 'verificar';
    document.getElementById('resumo-multi-processadas').innerText = resumo.processadas;
    // Na verificação nada foi encerrado: a coluna do meio mostra as aptas.
    document.getElementById('resumo-multi-encerradas').innerText = verificacao ? resumo.aptas : (resumo.encerradas_agora ?? resumo.encerradas);
    document.getElementById('rotulo-resumo-multi-encerradas').innerText = verificacao ? 'Aptas' : 'Encerradas agora';
    document.getElementById('resumo-multi-ja-encerradas').innerText = resumo.ja_encerradas ?? 0;
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

    if (verificacao) {
        verificacaoEncerramentoMulti = resumo.verificacao || [];
        renderizarVerificacaoEncerramentoMulti();
    } else {
        document.getElementById('verificacao-encerramento-multi').classList.add('hidden');
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
