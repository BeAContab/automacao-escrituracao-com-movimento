// Menu lateral: grupo recolhível "Outras Funções" e modal de Configuração Tess AI

const CHAVE_OUTRAS_FUNCOES_ABERTAS = 'beacontab_outras_funcoes_abertas';

function alternarOutrasFuncoes(forcarAberto) {
    const grupo = document.getElementById('grupo-outras-funcoes');
    const icone = document.getElementById('icone-outras-funcoes');
    if (!grupo || !icone) return;

    const abrir = forcarAberto !== undefined ? forcarAberto : grupo.classList.contains('hidden');
    grupo.classList.toggle('hidden', !abrir);
    icone.style.transform = abrir ? 'rotate(180deg)' : '';
    try {
        localStorage.setItem(CHAVE_OUTRAS_FUNCOES_ABERTAS, abrir ? '1' : '0');
    } catch (e) {
        // localStorage indisponível (ex.: modo privado): apenas não lembra a preferência
    }
}

function abrirConfigTess() {
    const modal = document.getElementById('modal-config-tess');
    if (modal) modal.classList.remove('hidden');
}

function fecharConfigTess() {
    const modal = document.getElementById('modal-config-tess');
    if (modal) modal.classList.add('hidden');
}

window.addEventListener('DOMContentLoaded', () => {
    // Lembra se "Outras Funções" estava aberto na última vez (começa fechado por padrão)
    let aberto = false;
    try {
        aberto = localStorage.getItem(CHAVE_OUTRAS_FUNCOES_ABERTAS) === '1';
    } catch (e) {
        // segue fechado por padrão
    }
    alternarOutrasFuncoes(aberto);

    document.addEventListener('keydown', (evento) => {
        if (evento.key === 'Escape') fecharConfigTess();
    });
});
