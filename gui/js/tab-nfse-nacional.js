// ---- Baixar NFS-e — Portal Nacional ----

let certificadoWindowsSelecionado = null; // { thumbprint, nome, documento, cnpj, tipo_certificado, validade_ate, exibicao }

function exibirCertificadoSelecionado(cert) {
    const detalhes = [];
    if (cert.documento) detalhes.push(cert.documento);
    if (cert.tipo_certificado) detalhes.push(cert.tipo_certificado);
    if (cert.validade_ate) detalhes.push(`válido até ${cert.validade_ate}`);

    document.getElementById('lbl-certificado-nome-nfse-nacional').innerText = cert.nome || cert.exibicao || '';
    document.getElementById('lbl-certificado-detalhe-nfse-nacional').innerText = detalhes.join(' · ');
    document.getElementById('card-certificado-selecionado-nfse-nacional').classList.remove('hidden');
    document.getElementById('btn-selecionar-certificado-windows').classList.add('hidden');

    // Certificados e-CNPJ trazem o CNPJ do titular embutido — preenche o campo
    // "CNPJ da Filial" automaticamente, já que normalmente é o mesmo CNPJ do certificado.
    if (cert.cnpj) {
        document.getElementById('input-cnpj-filial-nfse-nacional').value = cert.cnpj;
        // Se a pasta de destino já tinha sido escolhida antes do certificado, refaz a
        // sugestão de NSU agora que o CNPJ está disponível (só preenche se ainda vazio).
        const pasta = document.getElementById('input-pasta-nfse-nacional').value.trim();
        if (pasta) {
            sugerirNsuPelaPastaNfseNacional(pasta);
        }
    }
}

async function selecionarCertificadoWindows() {
    if (!(window.pywebview && window.pywebview.api)) return;
    addLogNfseNacional('Abrindo seletor de certificados do Windows...');
    try {
        const resultado = await window.pywebview.api.selecionar_certificado_windows_gui();
        if (resultado && resultado.thumbprint) {
            certificadoWindowsSelecionado = resultado;
            exibirCertificadoSelecionado(resultado);
            addLogNfseNacional(`Certificado selecionado: ${resultado.exibicao || resultado.nome}`);
        } else {
            addLogNfseNacional('Nenhum certificado selecionado.');
        }
    } catch (e) {
        addLogNfseNacional('Erro ao abrir o seletor de certificados: ' + e, true);
    }
}

async function carregarUltimoCertificadoNfseNacional() {
    if (!(window.pywebview && window.pywebview.api)) return;
    const ultimo = await window.pywebview.api.obter_ultimo_certificado_nfse_nacional();
    if (ultimo && ultimo.thumbprint) {
        certificadoWindowsSelecionado = ultimo;
        exibirCertificadoSelecionado(ultimo);
    }
}
document.addEventListener('DOMContentLoaded', carregarUltimoCertificadoNfseNacional);

async function selecionarCertificadoNfseNacional() {
    if (window.pywebview) {
        const arquivo = await window.pywebview.api.selecionar_arquivo_certificado();
        if (arquivo) {
            document.getElementById('input-pfx-nfse-nacional').value = arquivo;
        }
    }
}

async function selecionarPastaNfseNacional() {
    if (window.pywebview) {
        const pasta = await window.pywebview.api.selecionar_pasta();
        if (pasta) {
            document.getElementById('input-pasta-nfse-nacional').value = pasta;
            await sugerirNsuPelaPastaNfseNacional(pasta);
        }
    }
}

async function sugerirNsuPelaPastaNfseNacional(pasta) {
    // Só sugere (pré-preenche) se o campo estiver vazio — nunca sobrescreve um valor
    // que o operador já digitou, inclusive de propósito para reprocessar um
    // intervalo de NSU que ele tenha apagado da pasta.
    const campoNsu = document.getElementById('input-nsu-inicial-nfse-nacional');
    if (campoNsu.value.trim() !== '') return;
    if (!(window.pywebview && window.pywebview.api)) return;

    const cnpj = document.getElementById('input-cnpj-filial-nfse-nacional').value.trim();
    const sugestao = await window.pywebview.api.detectar_nsu_pasta_nfse_nacional(pasta, cnpj);
    if (sugestao !== null && sugestao !== undefined) {
        campoNsu.value = sugestao;
        addLogNfseNacional(
            `Detectados XMLs já baixados nesta pasta — sugerindo NSU inicial ${sugestao} ` +
            `(considerando eventuais lacunas de arquivos apagados). Se quiser baixar de novo ` +
            `algum intervalo apagado, é só trocar o valor do campo.`
        );
    }
}

async function usarUltimoNsuSalvoNfseNacional() {
    const cnpj = document.getElementById('input-cnpj-filial-nfse-nacional').value.trim();
    if (window.pywebview && window.pywebview.api) {
        const ultimo = await window.pywebview.api.obter_ultimo_nsu_nfse_nacional(cnpj);
        if (ultimo !== null && ultimo !== undefined) {
            document.getElementById('input-nsu-inicial-nfse-nacional').value = ultimo;
            addLogNfseNacional(`Último NSU salvo para este CNPJ: ${ultimo}.`);
        } else {
            addLogNfseNacional('Nenhum NSU salvo encontrado para este CNPJ ainda.');
        }
    }
}

async function iniciarNfseNacional() {
    const pfx = document.getElementById('input-pfx-nfse-nacional').value.trim();
    const senha = document.getElementById('input-senha-cert-nfse-nacional').value;
    const pasta = document.getElementById('input-pasta-nfse-nacional').value.trim();
    const nsuInicial = document.getElementById('input-nsu-inicial-nfse-nacional').value || '0';
    const cnpjFilial = document.getElementById('input-cnpj-filial-nfse-nacional').value.trim();
    const apenasNotasTomadas = document.getElementById('checkbox-apenas-notas-tomadas-nfse-nacional').checked;
    const anoFiltro = document.getElementById('input-ano-filtro-nfse-nacional').value.trim();
    const mesFiltro = document.getElementById('input-mes-filtro-nfse-nacional').value.trim();
    const thumbprint = certificadoWindowsSelecionado ? certificadoWindowsSelecionado.thumbprint : '';

    if (!thumbprint && !(pfx && senha)) {
        addLogNfseNacional('Selecione um certificado do Windows ou informe um arquivo .pfx com senha.', true);
        return;
    }
    if (!pasta) {
        addLogNfseNacional('Selecione a pasta de destino dos XMLs.', true);
        return;
    }

    document.getElementById('log-nfse-nacional').innerHTML = '';
    addLogNfseNacional('Iniciando download de NFS-e do Portal Nacional...');
    document.getElementById('progresso-container-nfse-nacional').classList.remove('hidden');

    if (window.pywebview && window.pywebview.api) {
        try {
            await window.pywebview.api.iniciar_nfse_nacional_gui(
                pasta, nsuInicial, cnpjFilial, apenasNotasTomadas, anoFiltro, mesFiltro, thumbprint, pfx, senha
            );
        } catch (e) {
            addLogNfseNacional('Erro ao iniciar download: ' + e, true);
        }
    }

    // Limpa o campo de senha da tela logo após enviar — não fica exposta na UI depois de disparado.
    document.getElementById('input-senha-cert-nfse-nacional').value = '';
}

async function cancelarNfseNacional() {
    if (window.pywebview && window.pywebview.api) {
        await window.pywebview.api.cancelar_nfse_nacional();
        addLogNfseNacional('Cancelamento solicitado — aguardando o lote atual terminar...');
    }
}

function addLogNfseNacional(msg, isError = false) {
    const el = document.getElementById('log-nfse-nacional');
    if (!el) return;
    const p = document.createElement('p');
    const agora = new Date().toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit', second: '2-digit' });
    p.innerText = `[${agora}] ${msg}`;
    p.className = isError ? 'text-error font-bold' : '';
    el.appendChild(p);
    el.scrollTop = el.scrollHeight;
}
window.log_nfse_nacional = addLogNfseNacional;

function limparLogNfseNacional() {
    const el = document.getElementById('log-nfse-nacional');
    if (el) el.innerHTML = '';
}

function updateProgressoNfseNacional(porcentagem, status) {
    const container = document.getElementById('progresso-container-nfse-nacional');
    const bar = document.getElementById('progresso-barra-nfse-nacional');
    const pct = document.getElementById('progresso-porcentagem-nfse-nacional');
    const st = document.getElementById('progresso-status-nfse-nacional');

    if (container) container.classList.remove('hidden');
    if (bar) bar.style.width = `${porcentagem}%`;
    if (pct) pct.innerText = `${porcentagem}%`;
    if (st) st.innerText = status;
}
window.update_progresso_nfse_nacional = updateProgressoNfseNacional;
