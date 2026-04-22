/*  CONEXO  */
async function testConnection(silent = false) {
  setConnStatus('checking', 'Verificando...');
  try {
    const data = await call('profile');
    if (data && data.result) {
      setConnStatus('ok', `Conectado  ${extractDomain()}`);
      updateOverview(data.result);
      hideError();
      if (!silent) toast('Conexão estabelecida com sucesso.', 'ok');
      return true;
    } else {
      throw new Error('Resposta invlida');
    }
  } catch (e) {
    setConnStatus('err', 'Falha na conexão');
    showError();
    if (!silent) toast('Não foi possível conectar ao webhook.', 'er');
    return false;
  }
}

function extractDomain() {
  try {
    const env = envs[activeIdx];
    if (!env) return '';
    return new URL(env.url).hostname;
  } catch { return ''; }
}

function setConnStatus(state, text) {
  const dot = document.getElementById('conn-dot');
  const txt = document.getElementById('env-status-txt');
  dot.className = `env-dot ${state}`;
  txt.textContent = text;
  txt.style.color = state === 'ok' ? 'rgba(11,191,126,0.7)' : state === 'err' ? 'rgba(248,113,113,0.8)' : 'rgba(255,255,255,0.38)';

  const topDot = document.getElementById('tb-conn-dot');
  const topTxt = document.getElementById('tb-conn-text');
  if (topDot) topDot.className = `tb-conn-dot ${state}`;
  if (topTxt) topTxt.textContent = state === 'ok' ? 'Conectado' : state === 'err' ? 'Offline' : 'Verificando...';
}

function showError() {
  const env = envs[activeIdx];
  document.getElementById('error-env-label').textContent = env ? env.url : '';
  document.getElementById('screen-error').classList.add('visible');
}
function hideError() {
  document.getElementById('screen-error').classList.remove('visible');
}
function retryConnection() { testConnection(false); }
function showEnvManager() { hideError(); navigate('envmanager'); }

function updateOverview(profile) {
  const env = envs[activeIdx];
  document.getElementById('ov-env-name').textContent   = env ? env.name : '';
  document.getElementById('ov-env-domain').textContent = extractDomain();
  document.getElementById('ov-status').textContent     = 'Ativo ';
}

