/* ── LABEL FROM BITRIX ───────────────────────────────── */
// Bitrix24 pode retornar EDIT_FORM_LABEL em vários formatos:
//   1. string pura:            "Nome do Campo"
//   2. objeto localizado:      {"pt_BR":"Nome","en":"Name"}
//   3. array {LANGUAGE_ID/VALUE}: [{LANGUAGE_ID:"pt_BR",VALUE:"Nome"}]
//   4. array de strings:       ["Nome do Campo"]
function labelFromBitrix(val) {
  if (!val) return '';
  if (typeof val === 'string') return val;
  if (typeof val === 'object') {
    // Tenta formato [{LANGUAGE_ID, VALUE}] — se encontrar, retorna o VALUE
    if (Array.isArray(val)) {
      const pref = ['pt_BR', 'pt', 'en', 'ru'];
      for (const lang of pref) {
        const entry = val.find(v => v && v.LANGUAGE_ID === lang && v.VALUE);
        if (entry) return entry.VALUE;
      }
      const any = val.find(v => v && typeof v.VALUE === 'string' && v.VALUE);
      if (any) return any.VALUE;
      // Não era formato {LANGUAGE_ID,VALUE} — continua para Object.values abaixo
      // (lida com arrays simples como ["Nome do Campo"])
    }
    // Trata como mapa de idiomas {"pt_BR":"..","en":".."}
    // Object.values também funciona para arrays simples ["Nome"] → retorna "Nome"
    return val['pt_BR'] || val['br'] || val['pt'] || val['en'] || val['ru'] ||
           Object.values(val).find(v => typeof v === 'string' && v) || '';
  }
  return '';
}

/* ── ESC HTML ────────────────────────────────────────── */
function escHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;').replace(/'/g,'&#39;');
}

/* ── TOAST ───────────────────────────────────────────── */
function toast(msg, type = 'ok', duration = 4000) {
  const area = document.getElementById('toast-area');
  const el   = document.createElement('div');
  el.className = `toast ${type}`;
  el.innerHTML = `<div class="toast-dot"></div><div class="toast-msg">${msg}</div><span class="toast-close" onclick="dismissToast(this.parentElement)">×</span>`;
  area.appendChild(el);
  const timer = setTimeout(() => dismissToast(el), duration);
  el._timer = timer;
}
function dismissToast(el) {
  if (!el || !el.parentElement) return;
  clearTimeout(el._timer);
  el.classList.add('out');
  setTimeout(() => el.remove(), 250);
}

