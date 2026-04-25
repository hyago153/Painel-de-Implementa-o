/*  AMBIENTES  */
function loadEnvs() {
  try { envs = JSON.parse(localStorage.getItem(LS_ENVS)) || []; } catch { envs = []; }
  try { activeIdx = parseInt(localStorage.getItem(LS_ACTIVE)) || 0; } catch { activeIdx = 0; }
  if (activeIdx >= envs.length) activeIdx = 0;
}
function saveEnvs() {
  localStorage.setItem(LS_ENVS, JSON.stringify(envs));
  localStorage.setItem(LS_ACTIVE, String(activeIdx));
}

function rebuildSelect() {
  const sel = document.getElementById('env-select');
  sel.innerHTML = '';
  envs.forEach((e, i) => {
    const opt = document.createElement('option');
    opt.value = i;
    opt.textContent = e.name;
    if (i === activeIdx) opt.selected = true;
    sel.appendChild(opt);
  });
}

function switchEnv(idx) {
  activeIdx = parseInt(idx, 10);
  saveEnvs();
  if (window.spaOverviewOnEnvironmentChanged) spaOverviewOnEnvironmentChanged();
  if (window.projOnEnvironmentChanged) projOnEnvironmentChanged();
  testConnection(false);
}

function addEnv() {
  const name = document.getElementById('new-env-name').value.trim();
  const url  = document.getElementById('new-env-url').value.trim();
  if (!name) { toast('Informe um nome para o ambiente.', 'wn'); return; }
  if (!url || !url.startsWith('http')) { toast('Informe uma URL vlida.', 'wn'); return; }
  envs.push({ name, url: url.endsWith('/') ? url : url + '/' });
  saveEnvs();
  clearNewEnvForm();
  renderEnvList();
  rebuildSelect();
  toast(`Ambiente "${name}" adicionado.`, 'ok');
}

function clearNewEnvForm() {
  document.getElementById('new-env-name').value = '';
  document.getElementById('new-env-url').value  = '';
}

function activateEnv(idx) {
  activeIdx = idx;
  saveEnvs();
  if (window.spaOverviewOnEnvironmentChanged) spaOverviewOnEnvironmentChanged();
  if (window.projOnEnvironmentChanged) projOnEnvironmentChanged();
  rebuildSelect();
  renderEnvList();
  testConnection(false);
  navigate('overview');
}

function deleteEnv(idx) {
  if (envs.length === 1) { toast('Voc precisa ter ao menos um ambiente.', 'wn'); return; }
  const name = envs[idx].name;
  const wasActive = idx === activeIdx;
  envs.splice(idx, 1);
  if (activeIdx >= envs.length) activeIdx = envs.length - 1;
  saveEnvs();
  if (window.spaOverviewOnEnvironmentChanged) spaOverviewOnEnvironmentChanged();
  if (window.projOnEnvironmentChanged) projOnEnvironmentChanged();
  rebuildSelect();
  renderEnvList();
  if (wasActive) testConnection(true);
  toast(`Ambiente "${name}" removido.`, 'wn');
}

function renderEnvList() {
  const container = document.getElementById('env-list-ui');
  container.innerHTML = '';
  if (envs.length === 0) {
    container.innerHTML = '<div style="font-size:12px;color:#aaa;padding:8px 0;">Nenhum ambiente cadastrado.</div>';
    return;
  }
  envs.forEach((env, i) => {
    const isActive = i === activeIdx;
    const initials = escHtml(env.name.substring(0, 2).toUpperCase());
    const urlShort = env.url.replace('https://', '').replace('http://', '');
    const row = document.createElement('div');
    row.className = 'env-row' + (isActive ? ' active-env' : '');
    row.innerHTML = `
      <div class="env-row-avatar">${initials}</div>
      <div style="flex:1;min-width:0;">
        <div style="display:flex;align-items:center;gap:7px;">
          <span class="env-row-name">${escHtml(env.name)}</span>
          ${isActive ? '<span class="env-row-badge">Ativo</span>' : ''}
        </div>
        <div class="env-row-url" title="${escHtml(env.url)}">${escHtml(urlShort)}</div>
      </div>
      <div class="env-row-actions">
        ${!isActive ? `<button class="icon-btn use" title="Usar este ambiente" onclick="activateEnv(${i})"></button>` : ''}
        <button class="icon-btn del" title="Remover" onclick="deleteEnv(${i})"></button>
      </div>
    `;
    container.appendChild(row);
  });
}

/*  SETUP  */
function setupSave() {
  const name = document.getElementById('setup-name').value.trim();
  const url  = document.getElementById('setup-webhook').value.trim();
  if (!name) { toast('Informe um nome para o ambiente.', 'wn'); return; }
  if (!url || !url.startsWith('http')) { toast('Informe uma URL de webhook vlida.', 'wn'); return; }
  envs.push({ name, url: url.endsWith('/') ? url : url + '/' });
  activeIdx = 0;
  saveEnvs();
  hideSetup();
  rebuildSelect();
  testConnection(false);
}

function setupSkip() {
  hideSetup();
  rebuildSelect();
  testConnection(true);
}

function hideSetup() {
  document.getElementById('screen-setup').style.display = 'none';
}

/*  BOOT  */
function boot() {
  loadEnvs();
  const initialPanel = window.restoreNavigationState ? restoreNavigationState() : 'overview';
  if (envs.length === 0) {
    document.getElementById('screen-setup').style.display = 'flex';
  } else {
    document.getElementById('setup-skip').style.display = 'block';
    document.getElementById('screen-setup').style.display = 'none';
    rebuildSelect();
    testConnection(true);
  }
  if (window.spaOverviewOnEnvironmentChanged) spaOverviewOnEnvironmentChanged();
  if (window.projOnEnvironmentChanged) projOnEnvironmentChanged();
  navigate(initialPanel);
  criarInitTipos();
  massaInit();
}

/* boot() movido para o final do arquivo  chamado aps todas as const/let serem inicializadas */


