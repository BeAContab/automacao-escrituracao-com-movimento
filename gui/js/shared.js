// Helpers compartilhados entre abas: credenciais Tess AI e inicialização geral

async function carregarChavesAPI() {
    if (window.pywebview && window.pywebview.api) {
        try {
            const chaves = await window.pywebview.api.obter_chaves_salvas();
            if (chaves) {
                if (chaves.tess_key) {
                    document.getElementById('api-key-tess').value = chaves.tess_key;
                }
                if (chaves.tess_agent_id) {
                    document.getElementById('agent-id-tess').value = chaves.tess_agent_id;
                }
            }
        } catch (e) {
            console.error("Erro ao carregar chaves:", e);
        }
    }
}

async function salvarChavesAPI() {
    const tessKey = document.getElementById('api-key-tess').value.trim();
    const tessAgent = document.getElementById('agent-id-tess').value.trim();

    if (window.pywebview && window.pywebview.api) {
        const ok = await window.pywebview.api.salvar_chaves(tessKey, tessAgent);
        if (ok) {
            addLogXml("Configurações do Tess AI salvas com sucesso!");
        } else {
            addLogXml("Falha ao salvar configurações do Tess AI.", true);
        }
    }
}

window.addEventListener('DOMContentLoaded', () => {
    const hoje = new Date();
    const mesAtual = String(hoje.getMonth() + 1).padStart(2, '0');
    const anoAtual = String(hoje.getFullYear());

    // MELHORIA-06: Flag para detectar se o arquivo foi selecionado nesta sessão
    window._arquivoExcelSelecionadoNestaSessao = false;

    // Função reutilizável para inicializar seletores de ano em todas as abas
    function inicializarSeletorAno(idSelect) {
        const select = document.getElementById(idSelect);
        if (!select) return;
        let anoExiste = Array.from(select.options).some(o => o.value === anoAtual);
        if (!anoExiste) {
            const opt = document.createElement('option');
            opt.value = anoAtual;
            opt.text = anoAtual;
            select.add(opt, select.options[0]);
        }
        select.value = anoAtual;
    }

    ['select-ano-automacao', 'select-ano-exportar', 'select-ano-captura', 'select-ano-encerramento'].forEach(inicializarSeletorAno);

    ['select-mes-automacao', 'select-mes-exportar', 'select-mes-captura', 'select-mes-encerramento'].forEach(id => {
        const select = document.getElementById(id);
        if (select) select.value = mesAtual;
    });

    if (window.pywebview && window.pywebview.api) {
        carregarChavesAPI();
    } else {
        window.addEventListener('pywebviewready', carregarChavesAPI);
    }
    setActiveTab('xml');
});
