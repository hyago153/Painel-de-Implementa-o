/*  CONSTANTES  */
const LS_ENVS   = 'b24panel_envs';
const LS_ACTIVE = 'b24panel_active';
const LS_APP_MODE = 'b24panel_app_mode';
const LS_PANEL = 'b24panel_panel';

/*  ESTADO  */
let envs       = [];
let activeIdx  = 0;
let currentPanel = 'overview';
let appMode = 'crm';

/*  API CALL  */
function getWebhook() {
  const env = envs[activeIdx];
  return env ? env.url : null;
}

async function call(method, params = {}) {
  const webhook = getWebhook();
  if (!webhook) throw new Error('Nenhum webhook configurado');
  const url  = webhook.replace(/\/$/, '') + '/' + method;
  const form = new FormData();
  function appendFlat(prefix, val) {
    if (val === null || val === undefined) return;
    if (typeof val === 'object' && !Array.isArray(val)) {
      for (const [k, v] of Object.entries(val)) appendFlat(`${prefix}[${k}]`, v);
    } else if (Array.isArray(val)) {
      val.forEach((v, i) => appendFlat(`${prefix}[${i}]`, v));
    } else {
      form.append(prefix, val);
    }
  }
  for (const [k, v] of Object.entries(params)) appendFlat(k, v);
  const res = await fetch(url, { method: 'POST', body: form });
  if (!res.ok) {
    let detail = '';
    try { const j = await res.json(); detail = j.error_description || j.error || ''; } catch(_) {}
    throw new Error(`HTTP ${res.status}${detail ? ': ' + detail : ''}`);
  }
  return res.json();
}

