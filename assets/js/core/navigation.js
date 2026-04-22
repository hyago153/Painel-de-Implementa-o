/*  NAVEGAO  */
const panelMeta = {
  overview:   { title: 'Visão geral',           sub: 'Painel de Implementação',        action: 'Testar conexão' },
  pipelines:  { title: 'Pipelines',             sub: 'Gerenciar funis e estágios',     action: '+ Novo pipeline' },
  campos:     { title: 'Consulta de campos',    sub: 'Nativos e personalizados',       action: 'Exportar CSV' },
  criar:      { title: 'Criar campo',           sub: 'Campo individual',               action: 'Documentao API' },
  massa:      { title: 'Criar em massa',        sub: 'Formulrio e CSV',               action: 'Baixar template XLS' },
  card:       { title: 'Config. do card',       sub: 'Visibilidade e seções',          action: 'Carregar config.' },
  envmanager: { title: 'Ambientes',             sub: 'Gerenciar webhooks Bitrix24',    action: '+ Novo ambiente' },
};

const appModeMeta = {
  crm: 'CRM',
  spa: 'SPA',
};
const LS_SIDEBAR_STATE = 'b24panel_sidebar_state';

function setSidebarPinned(pinned) {
  const app = document.getElementById('app');
  const btn = document.getElementById('sidebar-pin-btn');
  if (!app) return;
  app.classList.toggle('sidebar-pinned', pinned);
  app.classList.toggle('sidebar-minimized', !pinned);
  if (btn) {
    btn.setAttribute('aria-pressed', String(pinned));
    btn.setAttribute('title', pinned ? 'Minimizar menu' : 'Fixar menu expandido');
    btn.setAttribute('aria-label', pinned ? 'Minimizar menu' : 'Fixar menu expandido');
  }
  try { localStorage.setItem(LS_SIDEBAR_STATE, pinned ? 'expanded' : 'minimized'); } catch (_) {}
}

function toggleSidebarPinned() {
  const app = document.getElementById('app');
  setSidebarPinned(!(app && app.classList.contains('sidebar-pinned')));
}

function restoreSidebarState() {
  let saved = 'minimized';
  try { saved = localStorage.getItem(LS_SIDEBAR_STATE) || saved; } catch (_) {}
  setSidebarPinned(saved === 'expanded');
}

function getAppModeLabel() {
  return appModeMeta[appMode] || 'CRM';
}

function setAppMode(mode) {
  if (!appModeMeta[mode]) return;
  if (window.EntityContext && window.EntityContext.setCurrentMode) {
    window.EntityContext.setCurrentMode(mode);
  } else {
    appMode = mode;
  }
  try { localStorage.setItem(LS_APP_MODE, mode); } catch (_) {}
  updateModeSwitcher();
  if (currentPanel === 'envmanager') {
    navigate('overview');
    return;
  }
  navigate(currentPanel || 'overview');
}

function updateModeSwitcher() {
  document.querySelectorAll('.mode-btn').forEach(btn => btn.classList.remove('active'));
  const btn = document.getElementById('mode-' + appMode);
  if (btn) btn.classList.add('active');
  updateMobileModeSwitcher();
}

function updateMobileModeSwitcher() {
  document.querySelectorAll('.mobile-mode-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.mobileMode === appMode);
  });
}

function updateTopbar(id, meta) {
  const isGlobal = id === 'envmanager';
  document.getElementById('tb-title').textContent = isGlobal ? meta.title : getAppModeLabel();
  document.getElementById('tb-sub').textContent = isGlobal ? meta.sub : (meta.title || id);
  document.getElementById('tb-action').textContent = (!isGlobal && appMode === 'spa' && id === 'overview')
    ? 'Atualizar lista'
    : (meta.action || '');
  document.getElementById('tb-action').dataset.panel = id;
}

function globalSearchKey(event) {
  if (event.key !== 'Enter') return;
  const input = event.currentTarget;
  const query = String(input.value || '').trim().toLowerCase();
  if (!query) return;

  const matches = [
    ['overview', ['visao', 'viso', 'geral', 'inicio', 'dashboard', 'painel']],
    ['pipelines', ['pipeline', 'pipelines', 'funil', 'funis', 'estágio', 'estágio']],
    ['campos', ['campo', 'campos', 'consulta', 'customizado', 'personalizado']],
    ['criar', ['criar', 'novo', 'field', 'individual']],
    ['massa', ['massa', 'csv', 'xls', 'xlsx', 'planilha', 'importar']],
    ['card', ['card', 'config', 'configuracao', 'configuração', 'visibilidade', 'secao', 'seção']],
    ['envmanager', ['ambiente', 'ambientes', 'webhook', 'conexao', 'conexão']],
  ];

  const hit = matches.find(([, terms]) => terms.some(term => term.includes(query) || query.includes(term)));
  if (hit) {
    navigate(hit[0]);
    input.value = '';
  }
}

document.addEventListener('keydown', event => {
  if ((event.ctrlKey || event.metaKey) && event.key.toLowerCase() === 'k') {
    event.preventDefault();
    const input = document.getElementById('global-search');
    if (input) input.focus();
  }
});

function navigate(id) {
  document.querySelectorAll('.panel').forEach(p => p.classList.remove('active'));
  document.querySelectorAll('.nav-item').forEach(n => n.classList.remove('active'));
  const panel = document.getElementById('panel-' + id);
  if (panel) panel.classList.add('active');
  const navEl = document.getElementById('nav-' + id);
  if (navEl) navEl.classList.add('active');
  updateMobileNavigation(id);
  toggleMobileMoreMenu(false);
  const meta = panelMeta[id] || {};
  updateModeSwitcher();
  updateTopbar(id, meta);
  currentPanel = id;
  if (id !== 'envmanager') {
    try { localStorage.setItem(LS_PANEL, id); } catch (_) {}
  }
  if (id === 'overview' && window.spaOverviewEnsureLoaded) spaOverviewEnsureLoaded();
  if (window.spaOverviewModeSync) spaOverviewModeSync();
  if (id === 'pipelines' && window.pipSyncContextUI) pipSyncContextUI();
  if (id === 'campos' && window.camposSyncContextUI) camposSyncContextUI();
  if (id === 'criar' && window.criarSyncContextUI) criarSyncContextUI();
  if (id === 'massa' && window.massaSyncContextUI) massaSyncContextUI();
  if (id === 'card' && window.cardSyncContextUI) cardSyncContextUI();
  if (id === 'envmanager') renderEnvList();
}

function updateMobileNavigation(id) {
  document.querySelectorAll('.mobile-nav-item').forEach(item => {
    const target = item.dataset.mobilePanel;
    item.classList.toggle('active', target === id);
  });
}

function toggleMobileMoreMenu(open) {
  const overlay = document.getElementById('mobile-more-overlay');
  if (!overlay) return;
  const shouldOpen = typeof open === 'boolean' ? open : !overlay.classList.contains('open');
  overlay.classList.toggle('open', shouldOpen);
  document.querySelector('.mobile-bottom-nav')?.classList.toggle('module-open', shouldOpen);
  document.querySelectorAll('.mobile-module-trigger').forEach(item => {
    item.classList.toggle('module-open', shouldOpen);
  });
}

document.addEventListener('keydown', event => {
  if (event.key === 'Escape') toggleMobileMoreMenu(false);
});

function restoreNavigationState() {
  restoreSidebarState();

  try {
    const savedMode = localStorage.getItem(LS_APP_MODE);
    if (appModeMeta[savedMode]) {
      if (window.EntityContext && window.EntityContext.setCurrentMode) {
        window.EntityContext.setCurrentMode(savedMode);
      } else {
        appMode = savedMode;
      }
    }
  } catch (_) {}

  try {
    const savedPanel = localStorage.getItem(LS_PANEL);
    if (savedPanel && panelMeta[savedPanel] && savedPanel !== 'envmanager') return savedPanel;
  } catch (_) {}
  return 'overview';
}

window.toggleSidebarPinned = toggleSidebarPinned;
window.restoreSidebarState = restoreSidebarState;
window.toggleMobileMoreMenu = toggleMobileMoreMenu;

function tbAction() {
  const panel = document.getElementById('tb-action').dataset.panel || currentPanel;
  switch(panel) {
    case 'overview':   appMode === 'spa' && window.spaOverviewLoad ? spaOverviewLoad(true) : testConnection(false); break;
    case 'pipelines':
      document.getElementById('pip-new-name')?.scrollIntoView({ behavior: 'smooth', block: 'center' });
      document.getElementById('pip-new-name')?.focus();
      break;
    case 'campos':     camposAllFields.length > 0 ? camposExportCSV() : camposLoad(); break;
    case 'criar':      window.open('https://apidocs.bitrix24.com/api-reference/crm/userfields/'); break;
    case 'massa':      massaDownloadTemplate(); break;
    case 'card':       cardLoad(); break;
    case 'envmanager': document.getElementById('new-env-name').focus(); break;
  }
}
