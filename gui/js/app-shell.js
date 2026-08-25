// Navegação entre abas do menu lateral

const TABS = [
    { id: 'xml', btnId: 'btn-aba-xml', mainId: 'aba-xml', title: 'Processamento de XMLs' },
    { id: 'automacao', btnId: 'btn-aba-automacao', mainId: 'aba-automacao', title: 'Automação: Escrituração' },
    { id: 'exportar', btnId: 'btn-aba-exportar', mainId: 'aba-exportar', title: 'Exportar XML de Prestados' },
    { id: 'captura-movimento', btnId: 'btn-aba-captura-movimento', mainId: 'aba-captura-movimento', title: 'Captura Escrituração (Com Movimento)' },
    { id: 'encerramento-sem-movimento', btnId: 'btn-aba-encerramento-sem-movimento', mainId: 'aba-encerramento-sem-movimento', title: 'Encerramento ISS (Sem Movimento)' },
];

const NAV_ACTIVE_CLASSES = ['border-l-4', 'border-primary-container', 'bg-on-secondary-fixed-variant', 'text-white'];
const NAV_INACTIVE_CLASSES = ['text-secondary-fixed-dim'];

const headerTitle = document.getElementById('header-title');

function setActiveTab(tabId) {
    TABS.forEach(tab => {
        const btn = document.getElementById(tab.btnId);
        const main = document.getElementById(tab.mainId);
        if (!btn || !main) return;

        if (tab.id === tabId) {
            btn.classList.add(...NAV_ACTIVE_CLASSES);
            btn.classList.remove(...NAV_INACTIVE_CLASSES);
            main.classList.remove('hidden');
            if (headerTitle) headerTitle.innerText = tab.title;
        } else {
            btn.classList.remove(...NAV_ACTIVE_CLASSES);
            btn.classList.add(...NAV_INACTIVE_CLASSES);
            main.classList.add('hidden');
        }
    });
}

TABS.forEach(tab => {
    const btn = document.getElementById(tab.btnId);
    if (btn) btn.addEventListener('click', () => setActiveTab(tab.id));
});
