/* -- MODULO: PIPELINES ----------------------------------------------------- */

let pipEntityTypeId = 2;
let pipelines       = [];
let pipSelectedId   = null;
let pipSelectedName = '';
let pipLastContextKey = '';
const stageCollapsedIds = new Set();
let stageColorPickerState = null;

function pipMode() {
  return window.EntityContext ? window.EntityContext.getCurrentMode() : 'crm';
}

function pipGetContext(requireSelection = true) {
  const mode = pipMode();
  if (mode === 'spa') {
    const spa = requireSelection
      ? window.EntityContext.requireActiveSpa()
      : window.EntityContext.getActiveSpa();

    return {
      mode,
      adapter: window.EntityAdapter.spa,
      spa,
      entityTypeId: spa ? spa.entityTypeId : null,
      label: spa ? spa.title : 'SPA',
    };
  }

  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(pipEntityTypeId);
  const meta = window.CRM_ENTITY_META && window.CRM_ENTITY_META[pipEntityTypeId];
  return {
    mode,
    adapter: window.EntityAdapter.crm,
    entityTypeId: pipEntityTypeId,
    label: meta ? meta.label : 'CRM',
  };
}

function pipContextKey(ctx = pipGetContext(false)) {
  if (ctx.mode === 'spa') return ctx.spa ? `spa:${ctx.spa.id}:${ctx.entityTypeId}` : 'spa:none';
  return `crm:${ctx.entityTypeId}`;
}

function pipResetSelection() {
  pipSelectedId = null;
  pipSelectedName = '';
  const stagesCard = document.getElementById('pip-stages-card');
  if (stagesCard) stagesCard.style.display = 'none';
}

function pipRenderEmptyState(message) {
  const list = document.getElementById('pip-list');
  if (list) list.innerHTML = `<div style="font-size:12px;color:#aaa;">${escHtml(message)}</div>`;
  pipResetSelection();
}

function pipJsArg(value) {
  return escHtml(JSON.stringify(String(value)));
}

function pipActionIcon(name) {
  const common = 'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"';
  const icons = {
    edit: `<svg ${common}><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>`,
    delete: `<svg ${common}><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v5"/><path d="M14 11v5"/></svg>`,
  };
  return icons[name] || '';
}

function pipSyncContextUI() {
  const ctx = pipGetContext(false);
  const key = pipContextKey(ctx);
  const tabs = document.getElementById('pip-entity-tabs');
  const banner = document.getElementById('pip-banner');
  const defaultCheck = document.getElementById('pip-new-default');
  const defaultLabel = defaultCheck ? defaultCheck.closest('label') : null;

  if (tabs) tabs.style.display = ctx.mode === 'spa' ? 'none' : '';

  if (ctx.mode === 'spa') {
    if (banner) {
      banner.innerHTML = ctx.spa
        ? `SPA ativo: <strong>${escHtml(ctx.spa.title)}</strong> (entityTypeId ${escHtml(String(ctx.entityTypeId))}). As etapas usam ENTITY_ID <strong>DYNAMIC_${escHtml(String(ctx.entityTypeId))}_STAGE_{pipelineId}</strong>.`
        : 'Selecione um processo inteligente na Visão geral do módulo SPA antes de gerenciar pipelines.';
    }
    if (defaultLabel) defaultLabel.style.display = 'none';
  } else {
    document.querySelectorAll('#pip-entity-tabs .etab').forEach(t => {
      t.classList.toggle('active', String(t.dataset.entity) === String(pipEntityTypeId));
    });
    if (banner) {
      banner.innerHTML = pipEntityTypeId === 2
        ? 'Os pipelines de Negócio têm restrições: <strong>isDefault não pode ser alterado via REST API para Negócios.</strong>'
        : 'Lead: <strong>O pipeline padrao pode ser definido normalmente via API.</strong>';
    }
    if (defaultLabel) defaultLabel.style.display = '';
  }

  if (key !== pipLastContextKey) {
    pipelines = [];
    pipLastContextKey = key;
    pipResetSelection();
    pipRenderEmptyState(ctx.mode === 'spa' && !ctx.spa
      ? 'Selecione um processo inteligente na Visão geral do módulo SPA.'
      : 'Clique em "Recarregar" para carregar os pipelines.');
  }
}

function pipSelectEntity(eid, btn) {
  pipEntityTypeId = eid;
  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(eid);
  if (btn) {
    document.querySelectorAll('#pip-entity-tabs .etab').forEach(t => t.classList.remove('active'));
    btn.classList.add('active');
  }
  pipLastContextKey = '';
  pipSyncContextUI();
  pipLoad();
}

function pipExtractCategories(data) {
  const result = data && data.result;
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.categories)) return result.categories;
  if (result && Array.isArray(result.items)) return result.items;
  return [];
}

async function pipLoad() {
  pipSyncContextUI();
  const list = document.getElementById('pip-list');
  if (list) list.innerHTML = '<div style="font-size:12px;color:#aaa;">Carregando...</div>';

  try {
    const ctx = pipGetContext(true);
    const data = ctx.mode === 'spa'
      ? await ctx.adapter.listPipelines()
      : await ctx.adapter.listPipelines(ctx.entityTypeId);

    if (data.error) throw new Error(data.error_description || data.error);
    pipelines = pipExtractCategories(data);
    pipRenderList();
  } catch(e) {
    if (list) list.innerHTML = `<div style="font-size:12px;color:#ef4444;">Erro: ${escHtml(e.message)}</div>`;
  }
}

function pipRenderList() {
  const list = document.getElementById('pip-list');
  if (!list) return;

  if (!pipelines.length) {
    list.innerHTML = '<div style="font-size:12px;color:#aaa;">Nenhum pipeline encontrado.</div>';
    return;
  }

  list.innerHTML = pipelines.map(pip => {
    const isDefault = pip.isDefault === 'Y' || pip.IS_DEFAULT === 'Y' || String(pip.id) === '0';
    const isSelected = String(pip.id) === String(pipSelectedId);
    const rowClass = isSelected ? 'pip-row active' : 'pip-row';
    return `
      <div class="${rowClass}" id="pip-row-${escHtml(pip.id)}">
        <div class="pip-arrow" onclick="pipOpenStages(${Number(pip.id)},${pipJsArg(pip.name)})" title="Ver estágios">&#9656;</div>
        <div class="pip-body">
          <div class="pip-name">${escHtml(pip.name)}</div>
          <div class="pip-meta">Ordem: ${escHtml(String(pip.sort ?? pip.SORT ?? '-'))}</div>
          <div class="pip-badges">
            ${isDefault ? '<span class="pip-badge-default">Padrão</span>' : ''}
            <span class="pip-badge-id">ID: ${escHtml(String(pip.id))}</span>
          </div>
        </div>
        <div class="pip-actions">
          ${isDefault
            ? '<span style="font-size:10px;color:#aaa;">somente leitura</span>'
            : `<button class="icon-btn" title="Renomear" aria-label="Renomear" onclick="pipToggleRename(${Number(pip.id)})">${pipActionIcon('edit')}</button>
               <button class="icon-btn del" title="Excluir" aria-label="Excluir" onclick="pipDelete(${Number(pip.id)},${pipJsArg(pip.name)})">${pipActionIcon('delete')}</button>`
          }
        </div>
      </div>
      <div id="pip-rename-${escHtml(pip.id)}" style="display:none;" class="pip-rename-wrap">
        <div class="pip-rename-inner">
          <input class="form-input" id="pip-rename-input-${escHtml(pip.id)}" value="${escHtml(pip.name)}" placeholder="Novo nome" style="max-width:240px;" />
          <button class="btn sm primary" onclick="pipRename(${Number(pip.id)})">Salvar</button>
          <button class="btn sm" onclick="pipToggleRename(${Number(pip.id)})">Cancelar</button>
        </div>
      </div>
    `;
  }).join('');
}

function pipToggleRename(id) {
  const el = document.getElementById(`pip-rename-${id}`);
  if (!el) return;
  el.style.display = el.style.display === 'none' ? 'block' : 'none';
  if (el.style.display === 'block') document.getElementById(`pip-rename-input-${id}`).focus();
}

async function pipRename(id) {
  const input = document.getElementById(`pip-rename-input-${id}`);
  const name = input.value.trim();
  if (!name) { toast('Informe um nome.', 'wn'); return; }

  try {
    const ctx = pipGetContext(true);
    const data = ctx.mode === 'spa'
      ? await ctx.adapter.updatePipeline(id, { name })
      : await ctx.adapter.updatePipeline(ctx.entityTypeId, id, { name });

    if (data.error) throw new Error(data.error_description || data.error);
    toast('Pipeline renomeado.', 'ok');
    pipLoad();
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function pipCreate() {
  const name = document.getElementById('pip-new-name').value.trim();
  const sort = parseInt(document.getElementById('pip-new-sort').value, 10) || 100;
  const ctx = pipGetContext(false);
  const isDefault = ctx.mode === 'crm' && document.getElementById('pip-new-default').checked ? 'Y' : 'N';

  if (!name) { toast('Informe o nome do pipeline.', 'wn'); return; }

  try {
    const fullCtx = pipGetContext(true);
    const fields = fullCtx.mode === 'crm' ? { name, sort, isDefault } : { name, sort };
    const data = fullCtx.mode === 'spa'
      ? await fullCtx.adapter.addPipeline(fields)
      : await fullCtx.adapter.addPipeline(fullCtx.entityTypeId, fields);

    if (data.error) throw new Error(data.error_description || data.error);
    toast(`Pipeline "${escHtml(name)}" criado!`, 'ok');
    document.getElementById('pip-new-name').value = '';
    document.getElementById('pip-new-sort').value = '100';
    document.getElementById('pip-new-default').checked = false;
    pipLoad();
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function pipDelete(id, name) {
  if (!confirm(`Excluir pipeline "${name}"? Esta ação não pode ser desfeita.`)) return;

  try {
    const ctx = pipGetContext(true);
    const data = ctx.mode === 'spa'
      ? await ctx.adapter.deletePipeline(id)
      : await ctx.adapter.deletePipeline(ctx.entityTypeId, id);

    if (data.error) throw new Error(data.error_description || data.error);
    toast(`Pipeline "${escHtml(name)}" excluido.`, 'wn');
    if (String(pipSelectedId) === String(id)) pipResetSelection();
    pipLoad();
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function pipOpenStages(id, name) {
  pipSelectedId   = id;
  pipSelectedName = name;
  stageCollapsedIds.clear();
  document.getElementById('pip-stages-card').style.display = 'block';
  document.getElementById('pip-stages-title').textContent  = name;
  document.getElementById('pip-stages-card').scrollIntoView({ behavior: 'smooth', block: 'nearest' });
  pipRenderList();
  pipLoadStages();
}

function pipGetStageEntityId(ctx = pipGetContext(true)) {
  return window.EntityContext.getStageEntityId(ctx.entityTypeId, pipSelectedId);
}

async function pipLoadStages() {
  const tbody = document.getElementById('pip-stages-tbody');
  tbody.innerHTML = '<tr><td colspan="5" style="color:#aaa;font-size:12px;padding:12px;">Carregando...</td></tr>';

  try {
    const ctx = pipGetContext(true);
    if (pipSelectedId === null || pipSelectedId === undefined) throw new Error('Selecione um pipeline antes de carregar os estágios.');

    const data = ctx.mode === 'spa'
      ? await ctx.adapter.listStages(pipSelectedId)
      : await ctx.adapter.listStages(ctx.entityTypeId, pipSelectedId);

    if (data.error) throw new Error(data.error_description || data.error);
    const stages = Array.isArray(data.result) ? data.result : [];
    if (!stages.length) {
      tbody.innerHTML = '<tr><td colspan="5" style="color:#aaa;font-size:12px;padding:12px;">Nenhum estágio encontrado.</td></tr>';
      stageUpdateMobileBulkToggle();
      return;
    }

    tbody.innerHTML = stages.map(st => {
      const color  = st.COLOR || st.color || '#cccccc';
      const semRaw = st.SEMANTICS || st.semantics || '';
      const sType  = semRaw === 'S' ? 'SUCCESS' : semRaw === 'F' ? 'FAIL' : 'WORK';
      const statusId = st.STATUS_ID || st.statusId || st.id || '-';
      const numId    = String(st.ID || st.id || statusId);
      const sort   = st.SORT || st.sort || '-';
      const name   = st.NAME || st.name || '-';
      const collapsed = stageCollapsedIds.has(numId);
      return `
        <tr id="stage-row-${escHtml(numId)}" data-stage-id="${escHtml(numId)}" class="${collapsed ? 'stage-mobile-collapsed' : ''}">
          <td>
            <div class="stage-mobile-row-head">
              <div class="stage-mobile-row-summary">
                <span class="stage-mobile-row-title"><span class="stage-dot" style="background:${escHtml(color)};"></span>${escHtml(name)}</span>
                <span class="stage-mobile-row-id">${escHtml(String(statusId))}</span>
              </div>
              <button type="button" class="icon-btn section-collapse-btn stage-mobile-row-toggle" onclick="stageToggleMobileRow(this.closest('tr').dataset.stageId)" title="${collapsed ? 'Expandir estÃ¡gio' : 'Minimizar estÃ¡gio'}" aria-label="${collapsed ? 'Expandir estÃ¡gio' : 'Minimizar estÃ¡gio'}" aria-expanded="${collapsed ? 'false' : 'true'}">${collapsed ? '+' : '-'}</button>
            </div>
            <span class="stage-name-desktop"><span class="stage-dot" style="background:${escHtml(color)};"></span>${escHtml(name)}</span>
          </td>
          <td><span class="stage-badge ${sType}">${sType}</span></td>
          <td style="font-family:monospace;font-size:10px;color:#aaa;">${escHtml(String(statusId))}</td>
          <td>${escHtml(String(sort))}</td>
          <td>
            <div style="display:flex;gap:4px;">
              <button class="icon-btn" title="Renomear" aria-label="Renomear" onclick="stageToggleRename(${pipJsArg(numId)})">${pipActionIcon('edit')}</button>
              <button class="icon-btn del" title="Excluir" aria-label="Excluir" onclick="stageDelete(${pipJsArg(numId)},${pipJsArg(name)})">${pipActionIcon('delete')}</button>
            </div>
          </td>
        </tr>
        <tr id="stage-rename-row-${escHtml(numId)}" class="stage-rename-row" style="display:none;">
          <td colspan="5">
            <div class="stage-rename-inner">
              <input class="form-input" id="stage-rename-name-${escHtml(numId)}" value="${escHtml(name)}" placeholder="Nome" style="max-width:180px;" />
              <div class="color-input-wrap">
                <input class="form-input" id="stage-rename-hex-${escHtml(numId)}" value="${escHtml(color)}" placeholder="#RRGGBB" style="max-width:100px;"
                  oninput="stageSyncColorButton('stage-rename-swatch-${escHtml(numId)}', this.value)" />
                <button id="stage-rename-swatch-${escHtml(numId)}" class="stage-color-trigger" type="button"
                  style="--stage-color:${escHtml(color)};" title="Selecionar cor" aria-label="Selecionar cor"
                  onclick="stageOpenColorPicker('stage-rename-hex-${escHtml(numId)}','stage-rename-swatch-${escHtml(numId)}')"></button>
              </div>
              <button class="btn sm primary" onclick="stageRename(${pipJsArg(numId)})">Salvar</button>
              <button class="btn sm" onclick="stageToggleRename(${pipJsArg(numId)})">Cancelar</button>
            </div>
          </td>
        </tr>
      `;
    }).join('');
    stageUpdateMobileBulkToggle();
  } catch(e) {
    tbody.innerHTML = `<tr><td colspan="5" style="color:#ef4444;font-size:12px;padding:12px;">Erro: ${escHtml(e.message)}</td></tr>`;
    stageUpdateMobileBulkToggle();
  }
}

function stageUpdateMobileBulkToggle() {
  const btn = document.getElementById('stage-mobile-bulk-toggle');
  if (!btn) return;
  const rows = [...document.querySelectorAll('#pip-stages-tbody tr[data-stage-id]')];
  const collapsedCount = rows.filter(row => row.classList.contains('stage-mobile-collapsed')).length;
  const shouldExpand = rows.length > 0 && collapsedCount === rows.length;
  btn.textContent = shouldExpand ? 'Expandir todos' : 'Minimizar todos';
  btn.setAttribute('aria-expanded', shouldExpand ? 'false' : 'true');
}

function stageToggleMobileRow(stageId) {
  const id = String(stageId || '');
  if (!id) return;
  if (stageCollapsedIds.has(id)) stageCollapsedIds.delete(id);
  else stageCollapsedIds.add(id);

  const safeId = window.CSS && CSS.escape ? CSS.escape(id) : id.replace(/"/g, '\\"');
  const row = document.querySelector(`#pip-stages-tbody tr[data-stage-id="${safeId}"]`);
  if (row) {
    const collapsed = stageCollapsedIds.has(id);
    row.classList.toggle('stage-mobile-collapsed', collapsed);
    const btn = row.querySelector('.stage-mobile-row-toggle');
    if (btn) {
      btn.textContent = collapsed ? '+' : '-';
      btn.title = collapsed ? 'Expandir estÃ¡gio' : 'Minimizar estÃ¡gio';
      btn.setAttribute('aria-label', collapsed ? 'Expandir estÃ¡gio' : 'Minimizar estÃ¡gio');
      btn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    }
    const renameRow = document.getElementById(`stage-rename-row-${id}`);
    if (collapsed && renameRow) renameRow.style.display = 'none';
  }
  stageUpdateMobileBulkToggle();
}

function stageToggleAllMobileRows() {
  const rows = [...document.querySelectorAll('#pip-stages-tbody tr[data-stage-id]')];
  if (!rows.length) return;
  const shouldCollapse = rows.some(row => !row.classList.contains('stage-mobile-collapsed'));
  rows.forEach(row => {
    const id = row.dataset.stageId;
    if (!id) return;
    if (shouldCollapse) stageCollapsedIds.add(id);
    else stageCollapsedIds.delete(id);
    row.classList.toggle('stage-mobile-collapsed', shouldCollapse);
    const btn = row.querySelector('.stage-mobile-row-toggle');
    if (btn) {
      btn.textContent = shouldCollapse ? '+' : '-';
      btn.title = shouldCollapse ? 'Expandir estÃ¡gio' : 'Minimizar estÃ¡gio';
      btn.setAttribute('aria-label', shouldCollapse ? 'Expandir estÃ¡gio' : 'Minimizar estÃ¡gio');
      btn.setAttribute('aria-expanded', shouldCollapse ? 'false' : 'true');
    }
    const renameRow = document.getElementById(`stage-rename-row-${id}`);
    if (shouldCollapse && renameRow) renameRow.style.display = 'none';
  });
  stageUpdateMobileBulkToggle();
}

function stageToggleRename(stId) {
  const row = document.getElementById(`stage-rename-row-${stId}`);
  if (!row) return;
  row.style.display = row.style.display === 'none' ? '' : 'none';
}

function stageNormalizeHex(value) {
  const raw = String(value || '').trim();
  const withHash = raw.startsWith('#') ? raw : `#${raw}`;
  if (/^#[0-9a-fA-F]{6}$/.test(withHash)) return withHash.toUpperCase();
  return '';
}

function stageSyncColorButton(buttonId, value) {
  const btn = document.getElementById(buttonId);
  const hex = stageNormalizeHex(value);
  if (btn && hex) btn.style.setProperty('--stage-color', hex);
}

/* ── Conversão HSV ↔ RGB ↔ HEX ──────────────────────────────── */
function scpHsvToRgb(h, s, v) {
  const c = v * s, x = c * (1 - Math.abs((h / 60) % 2 - 1)), m = v - c;
  let r0, g0, b0;
  if      (h <  60) { r0 = c; g0 = x; b0 = 0; }
  else if (h < 120) { r0 = x; g0 = c; b0 = 0; }
  else if (h < 180) { r0 = 0; g0 = c; b0 = x; }
  else if (h < 240) { r0 = 0; g0 = x; b0 = c; }
  else if (h < 300) { r0 = x; g0 = 0; b0 = c; }
  else              { r0 = c; g0 = 0; b0 = x; }
  return { r: Math.round((r0+m)*255), g: Math.round((g0+m)*255), b: Math.round((b0+m)*255) };
}

function scpRgbToHsv(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b), d = max - min;
  const s = max === 0 ? 0 : d / max;
  const v = max;
  let h;
  if (d === 0)       { h = 0; }
  else if (max === r) { h = ((g - b) / d % 6 + 6) % 6 * 60; }
  else if (max === g) { h = ((b - r) / d + 2) * 60; }
  else                { h = ((r - g) / d + 4) * 60; }
  return { h, s, v };
}

function scpHexToRgb(hex) {
  const m = /^#([0-9a-f]{6})$/i.exec(hex);
  if (!m) return null;
  return { r: parseInt(m[1].slice(0,2),16), g: parseInt(m[1].slice(2,4),16), b: parseInt(m[1].slice(4,6),16) };
}

function scpRgbToHex(r, g, b) {
  return '#' + [r,g,b].map(n => Math.max(0,Math.min(255,Math.round(n))).toString(16).padStart(2,'0')).join('').toUpperCase();
}

/* ── Renderização dos canvas ─────────────────────────────────── */
function scpPaintSvCanvas(h) {
  const canvas = document.getElementById('scp-sv-canvas');
  if (!canvas || !canvas.width) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width, ht = canvas.height;
  const { r, g, b } = scpHsvToRgb(h, 1, 1);
  ctx.fillStyle = `rgb(${r},${g},${b})`;
  ctx.fillRect(0, 0, w, ht);
  const wg = ctx.createLinearGradient(0, 0, w, 0);
  wg.addColorStop(0, 'rgba(255,255,255,1)');
  wg.addColorStop(1, 'rgba(255,255,255,0)');
  ctx.fillStyle = wg; ctx.fillRect(0, 0, w, ht);
  const bg = ctx.createLinearGradient(0, 0, 0, ht);
  bg.addColorStop(0, 'rgba(0,0,0,0)');
  bg.addColorStop(1, 'rgba(0,0,0,1)');
  ctx.fillStyle = bg; ctx.fillRect(0, 0, w, ht);
}

function scpPaintHueCanvas() {
  const canvas = document.getElementById('scp-hue-canvas');
  if (!canvas || !canvas.width) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width, ht = canvas.height;
  const grad = ctx.createLinearGradient(0, 0, w, 0);
  for (let i = 0; i <= 6; i++) {
    const rgb = scpHsvToRgb(i * 60, 1, 1);
    grad.addColorStop(i / 6, `rgb(${rgb.r},${rgb.g},${rgb.b})`);
  }
  ctx.fillStyle = grad;
  ctx.fillRect(0, 0, w, ht);
}

function scpResizeCanvases() {
  ['scp-sv-canvas', 'scp-hue-canvas'].forEach(id => {
    const c = document.getElementById(id);
    if (!c) return;
    const rect = c.getBoundingClientRect();
    if (rect.width > 0) { c.width = Math.round(rect.width); c.height = Math.round(rect.height); }
  });
}

/* ── Cursores ────────────────────────────────────────────────── */
function scpMoveSvCursor(s, v) {
  const canvas = document.getElementById('scp-sv-canvas');
  const cursor = document.getElementById('scp-sv-cursor');
  if (!canvas || !cursor) return;
  cursor.style.left = (s * canvas.offsetWidth)  + 'px';
  cursor.style.top  = ((1 - v) * canvas.offsetHeight) + 'px';
}

function scpMoveHueCursor(h) {
  const canvas = document.getElementById('scp-hue-canvas');
  const cursor = document.getElementById('scp-hue-cursor');
  if (!canvas || !cursor) return;
  cursor.style.left = ((h / 360) * canvas.offsetWidth) + 'px';
}

/* ── Atualização central ─────────────────────────────────────── */
function scpUpdate(hsv) {
  if (!stageColorPickerState) return;
  stageColorPickerState.hsv = hsv;
  const { h, s, v } = hsv;
  const rgb = scpHsvToRgb(h, s, v);
  const hex = scpRgbToHex(rgb.r, rgb.g, rgb.b);
  stageColorPickerState.value = hex;
  scpPaintSvCanvas(h);
  scpMoveSvCursor(s, v);
  scpMoveHueCursor(h);
  const preview = document.getElementById('scp-preview');
  if (preview) preview.style.background = hex;
  const hexInput = document.getElementById('stage-color-custom-input');
  if (hexInput && document.activeElement !== hexInput) hexInput.value = hex;
}

/* ── Handlers de ponteiro ────────────────────────────────────── */
function scpSvFromEvent(e) {
  const canvas = document.getElementById('scp-sv-canvas');
  if (!canvas || !stageColorPickerState) return;
  const rect = canvas.getBoundingClientRect();
  const s = Math.max(0, Math.min(1, (e.clientX - rect.left) / rect.width));
  const v = Math.max(0, Math.min(1, 1 - (e.clientY - rect.top) / rect.height));
  scpUpdate({ h: stageColorPickerState.hsv.h, s, v });
}

function scpHueFromEvent(e) {
  const canvas = document.getElementById('scp-hue-canvas');
  if (!canvas || !stageColorPickerState) return;
  const rect = canvas.getBoundingClientRect();
  const h = Math.max(0, Math.min(359.99, ((e.clientX - rect.left) / rect.width) * 360));
  scpUpdate({ h, s: stageColorPickerState.hsv.s, v: stageColorPickerState.hsv.v });
}

function scpOnHexInput(value) {
  if (!stageColorPickerState) return;
  const hex = stageNormalizeHex(value);
  if (!hex) return;
  const rgb = scpHexToRgb(hex);
  if (rgb) scpUpdate(scpRgbToHsv(rgb.r, rgb.g, rgb.b));
}

async function scpPickFromScreen() {
  if (!window.EyeDropper) { toast('EyeDropper não é suportado neste navegador.', 'wn'); return; }
  try {
    const result = await new EyeDropper().open();
    const hex = stageNormalizeHex(result.sRGBHex);
    if (hex) { const rgb = scpHexToRgb(hex); if (rgb) scpUpdate(scpRgbToHsv(rgb.r, rgb.g, rgb.b)); }
  } catch (_) { /* usuário cancelou */ }
}

function scpInitEvents() {
  if (scpInitEvents._done) return;
  scpInitEvents._done = true;
  const svWrap  = document.getElementById('scp-sv-wrap');
  const hueWrap = document.getElementById('scp-hue-wrap');
  if (svWrap) {
    svWrap.addEventListener('pointerdown', e => { svWrap.setPointerCapture(e.pointerId); scpSvFromEvent(e); });
    svWrap.addEventListener('pointermove', e => { if (e.buttons) scpSvFromEvent(e); });
  }
  if (hueWrap) {
    hueWrap.addEventListener('pointerdown', e => { hueWrap.setPointerCapture(e.pointerId); scpHueFromEvent(e); });
    hueWrap.addEventListener('pointermove', e => { if (e.buttons) scpHueFromEvent(e); });
  }
}

/* ── API pública ─────────────────────────────────────────────── */
function stageSetPendingColor(value) {
  if (!stageColorPickerState) return;
  const hex = stageNormalizeHex(value);
  if (!hex) return;
  const rgb = scpHexToRgb(hex);
  if (rgb) scpUpdate(scpRgbToHsv(rgb.r, rgb.g, rgb.b));
}

function stageOpenColorPicker(inputId, buttonId) {
  const input   = document.getElementById(inputId);
  const overlay = document.getElementById('stage-color-overlay');
  if (!input || !overlay) return;
  scpInitEvents();
  const hex = stageNormalizeHex(input.value) || '#9BC2E6';
  const rgb = scpHexToRgb(hex) || { r: 155, g: 194, b: 230 };
  const hsv = scpRgbToHsv(rgb.r, rgb.g, rgb.b);
  stageColorPickerState = { inputId, buttonId, value: hex, hsv };
  overlay.classList.add('open');
  document.body.classList.add('stage-color-open');
  requestAnimationFrame(() => {
    scpResizeCanvases();
    scpPaintHueCanvas();
    scpUpdate(hsv);
  });
}

function stageApplyColorPicker() {
  if (!stageColorPickerState) return;
  const input = document.getElementById(stageColorPickerState.inputId);
  if (input) {
    input.value = stageColorPickerState.value;
    input.dispatchEvent(new Event('input', { bubbles: true }));
  }
  stageSyncColorButton(stageColorPickerState.buttonId, stageColorPickerState.value);
  stageCloseColorPicker();
}

function stageCloseColorPicker() {
  const overlay = document.getElementById('stage-color-overlay');
  if (overlay) overlay.classList.remove('open');
  document.body.classList.remove('stage-color-open');
  stageColorPickerState = null;
}

async function stageRename(stId) {
  const name  = document.getElementById(`stage-rename-name-${stId}`).value.trim();
  const color = document.getElementById(`stage-rename-hex-${stId}`).value.trim();
  if (!name) { toast('Informe o nome.', 'wn'); return; }

  try {
    const ctx = pipGetContext(true);
    const data = await ctx.adapter.updateStage(stId, { NAME: name, COLOR: color });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Estagio atualizado.', 'ok');
    pipLoadStages();
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function stageDelete(stId, name) {
  if (!confirm(`Excluir estágio "${name}"? Se houver itens ativos nesta etapa, o Bitrix24 pode bloquear a exclusão.`)) return;

  try {
    const ctx = pipGetContext(true);
    const data = await ctx.adapter.deleteStage(stId);
    if (data.error) throw new Error(data.error_description || data.error);
    toast(`Estagio "${escHtml(name)}" excluido.`, 'wn');
    pipLoadStages();
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function stageBuildStatusId(name) {
  let statusId = name.normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toUpperCase().replace(/[^A-Z0-9]/g,'_').replace(/_+/g,'_')
    .replace(/^_|_$/g,'').substring(0,47) || '';

  if (!statusId || /^\d/.test(statusId)) statusId = 'ST_' + statusId;
  if (!statusId || statusId === 'ST_') statusId = 'STAGE_' + Date.now();
  return statusId;
}

async function stageCreate() {
  const name  = document.getElementById('stage-new-name').value.trim();
  const sort  = parseInt(document.getElementById('stage-new-sort').value, 10) || 100;
  const color = document.getElementById('stage-new-color-hex').value.trim() || '#9BC2E6';
  const stypeRaw = document.getElementById('stage-new-type').value;
  const semantics = stypeRaw === 'SUCCESS' ? 'S' : stypeRaw === 'FAIL' ? 'F' : null;

  if (!name) { toast('Informe o nome do estágio.', 'wn'); return; }

  try {
    const ctx = pipGetContext(true);
    if (pipSelectedId === null || pipSelectedId === undefined) throw new Error('Selecione um pipeline antes de criar o estágio.');

    const entityId = pipGetStageEntityId(ctx);
    const statusId = stageBuildStatusId(name);
    const data = await ctx.adapter.addStage({
      ENTITY_ID: entityId,
      STATUS_ID: statusId,
      NAME: name,
      SORT: sort,
      COLOR: color,
      SEMANTICS: semantics,
    });

    if (data.error) throw new Error(data.error_description || data.error);
    toast(`Estagio "${escHtml(name)}" criado!`, 'ok');
    document.getElementById('stage-new-name').value = '';
    document.getElementById('stage-new-sort').value = '100';
    pipLoadStages();
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

window.pipSyncContextUI = pipSyncContextUI;
window.stageToggleMobileRow = stageToggleMobileRow;
window.stageToggleAllMobileRows = stageToggleAllMobileRows;
window.stageOpenColorPicker = stageOpenColorPicker;
window.stageCloseColorPicker = stageCloseColorPicker;
window.stageSetPendingColor = stageSetPendingColor;
window.stageApplyColorPicker = stageApplyColorPicker;
window.stageSyncColorButton = stageSyncColorButton;
