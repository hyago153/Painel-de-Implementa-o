/* -- MODULO: PIPELINES ----------------------------------------------------- */

let pipEntityTypeId = 2;
let pipelines       = [];
let pipSelectedId   = null;
let pipSelectedName = '';
let pipLastContextKey = '';
const pipCheckedIds = new Set();
let pipImportPlan = null;
const stageCollapsedIds = new Set();
let stageColorPickerState = null;
let _scpH = 0, _scpS = 1, _scpV = 1;
let _scpSvDrag = false, _scpHueDrag = false;

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
  pipCheckedIds.clear();
  pipUpdateSelectionBar();
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

function pipGetPipelineId(pip) {
  return pip && (pip.id ?? pip.ID ?? pip.categoryId ?? pip.CATEGORY_ID);
}

function pipGetPipelineName(pip) {
  return pip && (pip.name ?? pip.NAME ?? pip.title ?? pip.TITLE ?? '');
}

function pipGetPipelineSort(pip) {
  return pip && (pip.sort ?? pip.SORT ?? '');
}

function pipGetPipelineById(id) {
  return pipelines.find(p => String(pipGetPipelineId(p)) === String(id)) || null;
}

function pipToggleChecked(id, checked) {
  const key = String(id);
  if (checked) pipCheckedIds.add(key);
  else pipCheckedIds.delete(key);
  const row = document.getElementById(`pip-row-${key}`);
  if (row) row.classList.toggle('selected', checked);
  pipUpdateSelectionBar();
}

function pipUpdateSelectionBar() {
  const bar = document.getElementById('pip-sel-bar');
  const count = document.getElementById('pip-sel-count');
  const n = pipCheckedIds.size;
  if (count) count.textContent = String(n);
  if (bar) bar.classList.toggle('show', n > 0);
}

function pipClearSelection() {
  pipCheckedIds.clear();
  document.querySelectorAll('.pip-cb').forEach(cb => { cb.checked = false; });
  document.querySelectorAll('.pip-row.selected').forEach(row => row.classList.remove('selected'));
  pipUpdateSelectionBar();
}

function pipSelectAll() {
  if (!pipelines.length) {
    toast('Carregue os pipelines antes de selecionar todos.', 'wn');
    return;
  }
  pipelines.forEach(pip => {
    const id = pipGetPipelineId(pip);
    if (id !== null && id !== undefined) pipCheckedIds.add(String(id));
  });
  document.querySelectorAll('.pip-cb').forEach(cb => { cb.checked = true; });
  document.querySelectorAll('.pip-row').forEach(row => row.classList.add('selected'));
  pipUpdateSelectionBar();
}

function pipOpenModal(id, mode) {
  const el = document.getElementById(id);
  if (!el) return;
  if (id === 'pip-modal-export' && mode) {
    const radio = document.querySelector(`input[name="pip-export-mode"][value="${mode}"]`);
    if (radio) radio.checked = true;
    pipSyncExportOptions();
  }
  el.classList.add('open');
}

function pipCloseModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('open');
}

function pipSyncExportOptions() {
  document.querySelectorAll('.export-option').forEach(option => {
    const input = option.querySelector('input[type="radio"]');
    option.classList.toggle('selected', !!input && input.checked);
  });
}

function pipEnhanceListRows() {
  pipelines.forEach(pip => {
    const id = pipGetPipelineId(pip);
    const row = document.getElementById(`pip-row-${id}`);
    if (!row || row.querySelector('.pip-check')) return;
    const checkWrap = document.createElement('div');
    checkWrap.className = 'pip-check';
    const check = document.createElement('input');
    check.type = 'checkbox';
    check.className = 'pip-cb';
    check.checked = pipCheckedIds.has(String(id));
    check.setAttribute('aria-label', 'Selecionar pipeline');
    check.addEventListener('change', () => pipToggleChecked(id, check.checked));
    checkWrap.appendChild(check);
    row.insertBefore(checkWrap, row.firstChild);
    row.classList.toggle('selected', check.checked);
  });
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
  pipEnhanceListRows();
  pipUpdateSelectionBar();
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

/* ---- HSV Color Picker -------------------------------------------------- */

function _scpHsvToRgb(h, s, v) {
  const i = Math.floor(h / 60) % 6;
  const f = h / 60 - Math.floor(h / 60);
  const p = v * (1 - s), q = v * (1 - f * s), t = v * (1 - (1 - f) * s);
  return [[v,q,p,p,t,v],[t,v,v,q,p,p],[p,p,t,v,v,q]].map(a => Math.round(a[i] * 255));
}

function _scpRgbToHsv(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r,g,b), min = Math.min(r,g,b), d = max - min;
  let h = 0;
  const s = max === 0 ? 0 : d / max, v = max;
  if (d > 0) {
    if (max === r)      h = ((g - b) / d + (g < b ? 6 : 0)) * 60;
    else if (max === g) h = ((b - r) / d + 2) * 60;
    else                h = ((r - g) / d + 4) * 60;
  }
  return [h, s, v];
}

function _scpToHex(r, g, b) {
  return '#' + [r,g,b].map(n => n.toString(16).padStart(2,'0')).join('').toUpperCase();
}

function _scpHexToRgb(hex) {
  const h = hex.replace('#','');
  if (!/^[0-9a-fA-F]{6}$/.test(h)) return null;
  const n = parseInt(h, 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

function _scpPaintSv() {
  const canvas = document.getElementById('scp-canvas');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const w = canvas.width, h = canvas.height;
  const [hr, hg, hb] = _scpHsvToRgb(_scpH, 1, 1);
  const gradS = ctx.createLinearGradient(0, 0, w, 0);
  gradS.addColorStop(0, '#fff');
  gradS.addColorStop(1, `rgb(${hr},${hg},${hb})`);
  ctx.fillStyle = gradS;
  ctx.fillRect(0, 0, w, h);
  const gradV = ctx.createLinearGradient(0, 0, 0, h);
  gradV.addColorStop(0, 'rgba(0,0,0,0)');
  gradV.addColorStop(1, '#000');
  ctx.fillStyle = gradV;
  ctx.fillRect(0, 0, w, h);
}

function _scpPaintHue() {
  const canvas = document.getElementById('scp-hue');
  if (!canvas) return;
  const ctx = canvas.getContext('2d');
  const g = ctx.createLinearGradient(0, 0, canvas.width, 0);
  for (let i = 0; i <= 6; i++) g.addColorStop(i / 6, `hsl(${i * 60},100%,50%)`);
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, canvas.width, canvas.height);
}

function _scpRelPos(canvas, e) {
  const rect = canvas.getBoundingClientRect();
  const t = e.touches ? e.touches[0] : e;
  return {
    x: Math.max(0, Math.min(canvas.width,  (t.clientX - rect.left) / rect.width  * canvas.width)),
    y: Math.max(0, Math.min(canvas.height, (t.clientY - rect.top)  / rect.height * canvas.height)),
  };
}

function _scpSync() {
  const hex = _scpToHex(..._scpHsvToRgb(_scpH, _scpS, _scpV));
  const svCanvas  = document.getElementById('scp-canvas');
  const hCanvas   = document.getElementById('scp-hue');
  const thumb     = document.getElementById('scp-thumb');
  const hThumb    = document.getElementById('scp-hue-thumb');
  const preview   = document.getElementById('scp-preview');
  const hexInp    = document.getElementById('scp-hex-input');
  if (svCanvas && thumb) {
    thumb.style.left = `${_scpS * 100}%`;
    thumb.style.top  = `${(1 - _scpV) * 100}%`;
  }
  if (hCanvas && hThumb) {
    hThumb.style.left = `${(_scpH / 360) * 100}%`;
  }
  if (preview) preview.style.background = hex;
  if (hexInp && document.activeElement !== hexInp) hexInp.value = hex.slice(1);
  if (stageColorPickerState) stageColorPickerState.value = hex;
}

function _scpBindEvents() {
  const sv  = document.getElementById('scp-canvas');
  const hue = document.getElementById('scp-hue');
  if (!sv || sv.dataset.bound === '1') return;

  function moveSv(e) {
    e.preventDefault();
    const p = _scpRelPos(sv, e);
    _scpS = p.x / sv.width;
    _scpV = 1 - p.y / sv.height;
    _scpSync();
  }
  function moveHue(e) {
    e.preventDefault();
    _scpH = (_scpRelPos(hue, e).x / hue.width) * 360;
    _scpPaintSv();
    _scpSync();
  }

  sv.addEventListener('mousedown',  e => { _scpSvDrag  = true; moveSv(e); });
  sv.addEventListener('touchstart', e => { _scpSvDrag  = true; moveSv(e); }, {passive: false});
  hue.addEventListener('mousedown',  e => { _scpHueDrag = true; moveHue(e); });
  hue.addEventListener('touchstart', e => { _scpHueDrag = true; moveHue(e); }, {passive: false});

  document.addEventListener('mousemove', e => {
    if (_scpSvDrag)  moveSv(e);
    if (_scpHueDrag) moveHue(e);
  });
  document.addEventListener('touchmove', e => {
    if (_scpSvDrag)  moveSv(e);
    if (_scpHueDrag) moveHue(e);
  }, {passive: false});
  document.addEventListener('mouseup',  () => { _scpSvDrag = false; _scpHueDrag = false; });
  document.addEventListener('touchend', () => { _scpSvDrag = false; _scpHueDrag = false; });

  sv.dataset.bound = '1';
}

function scpHexInput(val) {
  const hex = stageNormalizeHex(val.length === 6 ? '#' + val : val);
  if (!hex) return;
  const rgb = _scpHexToRgb(hex);
  if (!rgb) return;
  [_scpH, _scpS, _scpV] = _scpRgbToHsv(...rgb);
  _scpPaintSv();
  _scpSync();
}

async function scpEyedropper() {
  if (!window.EyeDropper) { toast('EyeDropper não suportado neste navegador.', 'wn'); return; }
  try {
    const {sRGBHex} = await new EyeDropper().open();
    const rgb = _scpHexToRgb(sRGBHex);
    if (!rgb) return;
    [_scpH, _scpS, _scpV] = _scpRgbToHsv(...rgb);
    _scpPaintSv();
    _scpSync();
  } catch (_) {}
}

function stageOpenColorPicker(inputId, buttonId) {
  const input   = document.getElementById(inputId);
  const overlay = document.getElementById('stage-color-overlay');
  if (!input || !overlay) return;
  const hex = stageNormalizeHex(input.value) || '#9BC2E6';
  const rgb = _scpHexToRgb(hex);
  [_scpH, _scpS, _scpV] = rgb ? _scpRgbToHsv(...rgb) : [210, 0.36, 0.9];
  stageColorPickerState = { inputId, buttonId, value: hex };
  overlay.classList.add('open');
  document.body.classList.add('stage-color-open');
  requestAnimationFrame(() => {
    _scpPaintHue();
    _scpPaintSv();
    _scpBindEvents();
    _scpSync();
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

    if (semantics === 'S') {
      const existingStages = await pipFetchStagesForPipeline(pipSelectedId);
      if (pipFindExistingStageBySemantics(existingStages, 'S')) {
        throw new Error('O Bitrix permite apenas uma etapa de sucesso por pipeline. Renomeie/ajuste a etapa de sucesso existente ou crie uma etapa de falha.');
      }
    }

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

function pipRequireXlsx() {
  if (!window.XLSX) throw new Error('Biblioteca XLSX nao carregada.');
}

function pipSemanticToSheet(value) {
  const raw = String(value || '').toUpperCase();
  if (raw === 'S' || raw === 'SUCCESS') return 'SUCESSO';
  if (raw === 'F' || raw === 'FAIL') return 'FALHA';
  return 'ANDAMENTO';
}

function pipSemanticToApi(value) {
  const raw = String(value || '').trim().toUpperCase();
  if (raw === 'SUCESSO' || raw === 'SUCCESS') return 'S';
  if (raw === 'FALHA' || raw === 'FAIL') return 'F';
  return null;
}

function pipPadCode(prefix, index) {
  return `${prefix}_${String(index).padStart(4, '0')}`;
}

function pipStageCode(pipelineNumber, stageNumber) {
  const p = Math.max(0, parseInt(pipelineNumber, 10) || 0);
  const s = Math.max(0, parseInt(stageNumber, 10) || 0);
  return `STG_${String(p).padStart(2, '0')}_${String(s).padStart(2, '0')}`;
}

function pipStageImportRank(row) {
  const sort = parseInt(row && row.sort, 10);
  if (Number.isFinite(sort) && sort > 0) return sort;
  return (row && row.__sourceIndex !== undefined ? row.__sourceIndex + 1 : 1) * 10;
}

function pipStageSemanticOrder(row) {
  const semantics = pipSemanticToApi(row && row.semantica);
  if (semantics === 'S') return 1;
  if (semantics === 'F') return 2;
  return 0;
}

function pipStageSortForOrderedIndex(index) {
  return (index + 1) * 10;
}

function pipSortStagesForImport(stageRows) {
  return [...stageRows].sort((a, b) => {
    const semDiff = pipStageSemanticOrder(a) - pipStageSemanticOrder(b);
    if (semDiff) return semDiff;
    const sortDiff = pipStageImportRank(a) - pipStageImportRank(b);
    if (sortDiff) return sortDiff;
    return (a.__sourceIndex || 0) - (b.__sourceIndex || 0);
  });
}

function pipExistingStageSemantics(stage) {
  return String(stage && (stage.SEMANTICS ?? stage.semantics ?? '') || '').toUpperCase();
}

function pipExistingStageSort(stage) {
  const sort = parseInt(stage && (stage.SORT ?? stage.sort), 10);
  return Number.isFinite(sort) ? sort : 0;
}

function pipExistingStageId(stage) {
  if (!stage) return null;
  return stage.ID ?? stage.id ?? stage.STATUS_ID ?? stage.statusId ?? null;
}

function pipFindExistingStageBySemantics(stages, semantics) {
  const sem = String(semantics || '').toUpperCase();
  return pipFindExistingStagesBySemantics(stages, sem)[0] || null;
}

function pipFindExistingStagesBySemantics(stages, semantics) {
  const sem = String(semantics || '').toUpperCase();
  return (stages || [])
    .filter(stage => pipExistingStageSemantics(stage) === sem)
    .sort((a, b) => pipExistingStageSort(a) - pipExistingStageSort(b));
}

function pipFindReusableWorkStage(stages, keepStageIds) {
  return (stages || [])
    .filter(stage => {
      const stageId = pipExistingStageId(stage);
      return stageId && !keepStageIds.has(String(stageId)) && !pipExistingStageSemantics(stage);
    })
    .sort((a, b) => pipExistingStageSort(a) - pipExistingStageSort(b))[0] || null;
}

async function pipUpdateExistingSemanticStage(ctx, stage, row, results) {
  const semantics = pipSemanticToApi(row.semantica);
  if (!semantics || !stage) return null;

  const stageId = pipExistingStageId(stage);
  if (!stageId) return null;

  const targetSort = parseInt(row.sort, 10) || pipStageImportRank(row);
  const sortData = await ctx.adapter.updateStage(stageId, { SORT: targetSort });
  if (sortData.error) throw new Error(sortData.error_description || sortData.error);

  const fields = {
    NAME: row.nome_etapa,
    COLOR: row.cor_hex || (semantics === 'S' ? '#22C55E' : '#EF4444'),
  };
  const data = await ctx.adapter.updateStage(stageId, fields);
  if (data.error) throw new Error(data.error_description || data.error);

  Object.assign(stage, fields);
  stage.SORT = targetSort;
  stage.sort = targetSort;
  stage.SEMANTICS = semantics;
  stage.name = row.nome_etapa;
  stage.color = fields.COLOR;
  stage.semantics = semantics;
  if (results) {
    const label = semantics === 'S' ? 'ganho' : 'falha';
    results.push(`Etapa de ${label} ajustada: ${row.nome_etapa} (SORT ${targetSort}).`);
  }
  return stage;
}

async function pipUpdateExistingWorkStage(ctx, stage, row, results) {
  if (!stage || !row) return null;

  const stageId = pipExistingStageId(stage);
  if (!stageId) return null;

  const targetSort = parseInt(row.sort, 10) || pipStageImportRank(row);
  const sortData = await ctx.adapter.updateStage(stageId, { SORT: targetSort });
  if (sortData.error) throw new Error(sortData.error_description || sortData.error);

  const fields = {
    NAME: row.nome_etapa,
    COLOR: row.cor_hex || '#9BC2E6',
  };
  const data = await ctx.adapter.updateStage(stageId, fields);
  if (data.error) throw new Error(data.error_description || data.error);

  Object.assign(stage, fields);
  stage.SORT = targetSort;
  stage.sort = targetSort;
  stage.name = row.nome_etapa;
  stage.color = fields.COLOR;
  stage.semantics = null;
  if (results) results.push(`Primeira etapa ajustada: ${row.nome_etapa} (SORT ${targetSort}).`);
  return stage;
}

async function pipDeleteStagesNotInImport(ctx, existingStages, keepStageIds, results) {
  for (const stage of existingStages || []) {
    const stageId = pipExistingStageId(stage);
    if (!stageId || keepStageIds.has(String(stageId))) continue;

    try {
      const data = await ctx.adapter.deleteStage(stageId);
      if (data.error) throw new Error(data.error_description || data.error);
      if (results) results.push(`Etapa automatica removida: ${stage.NAME || stage.name || stage.STATUS_ID || stageId}`);
    } catch(e) {
      if (results) results.push(`Aviso: nao foi possivel remover etapa automatica ${stage.NAME || stage.name || stageId}: ${e.message}`);
    }
  }
}

function pipImportGroupStages(stageRows) {
  return stageRows.reduce((acc, row) => {
    const key = row.chave_pipeline || '';
    acc[key] = acc[key] || [];
    acc[key].push(row);
    return acc;
  }, {});
}

function pipDownloadWorkbook(workbook, filename) {
  pipRequireXlsx();
  XLSX.writeFile(workbook, filename);
}

function pipDownloadModel() {
  try {
    pipRequireXlsx();
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Como usar'],
      ['1. Preencha a aba Pipelines com uma linha por pipeline.'],
      ['2. Preencha a aba Estagios com uma ou mais etapas por chave_pipeline.'],
      ['3. Campos sort e codigos sugeridos podem ficar vazios; o painel preenche automaticamente.'],
      ['4. Semantica aceita: ANDAMENTO, SUCESSO, FALHA.'],
    ]), 'Instrucoes');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
      { chave_pipeline: 'pipeline_1', nome_pipeline: 'Pipeline Exemplo', sort: '', codigo_pipeline_sugerido: '', is_default: 'N' },
    ]), 'Pipelines');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet([
      { chave_pipeline: 'pipeline_1', nome_etapa: 'Primeiro contato', semantica: 'ANDAMENTO', sort: '', cor_hex: '#9BC2E6', codigo_etapa_sugerido: 'STG_01_01' },
      { chave_pipeline: 'pipeline_1', nome_etapa: 'Proposta', semantica: 'ANDAMENTO', sort: '', cor_hex: '#6FA8DC', codigo_etapa_sugerido: 'STG_01_02' },
      { chave_pipeline: 'pipeline_1', nome_etapa: 'Ganho', semantica: 'SUCESSO', sort: '', cor_hex: '#22C55E', codigo_etapa_sugerido: 'STG_01_03' },
      { chave_pipeline: 'pipeline_1', nome_etapa: 'Perdido', semantica: 'FALHA', sort: '', cor_hex: '#EF4444', codigo_etapa_sugerido: 'STG_01_04' },
      { chave_pipeline: 'pipeline_1', nome_etapa: 'Sem resposta', semantica: 'FALHA', sort: '', cor_hex: '#F97316', codigo_etapa_sugerido: 'STG_01_05' },
    ]), 'Estagios');
    pipDownloadWorkbook(wb, 'modelo-importacao-pipelines.xlsx');
    toast('Modelo baixado.', 'ok');
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function pipFetchStagesForPipeline(pipelineId) {
  const ctx = pipGetContext(true);
  const data = ctx.mode === 'spa'
    ? await ctx.adapter.listStages(pipelineId)
    : await ctx.adapter.listStages(ctx.entityTypeId, pipelineId);
  if (data.error) throw new Error(data.error_description || data.error);
  return Array.isArray(data.result) ? data.result : [];
}

async function pipExportXlsx() {
  try {
    pipRequireXlsx();
    const modeEl = document.querySelector('input[name="pip-export-mode"]:checked');
    const mode = modeEl ? modeEl.value : 'active';
    const selectedIds = mode === 'selected' ? [...pipCheckedIds] : [pipSelectedId].filter(v => v !== null && v !== undefined);
    if (!selectedIds.length) throw new Error(mode === 'selected' ? 'Selecione ao menos uma pipeline.' : 'Abra uma pipeline antes de exportar.');

    const pipeRows = [];
    const stageRows = [];
    for (let i = 0; i < selectedIds.length; i++) {
      const id = selectedIds[i];
      const pip = pipGetPipelineById(id);
      if (!pip) continue;
      const chave = `pipeline_${i + 1}`;
      pipeRows.push({
        chave_pipeline: chave,
        nome_pipeline: pipGetPipelineName(pip),
        sort: pipGetPipelineSort(pip),
        codigo_pipeline_sugerido: pipPadCode('PIP', i + 1),
        is_default: (pip.isDefault === 'Y' || pip.IS_DEFAULT === 'Y' || String(id) === '0') ? 'Y' : 'N',
        id_bitrix: id,
      });
      const stages = await pipFetchStagesForPipeline(id);
      stages.forEach((st, stageIndex) => {
        stageRows.push({
          chave_pipeline: chave,
          nome_etapa: st.NAME || st.name || '',
          semantica: pipSemanticToSheet(st.SEMANTICS || st.semantics || ''),
          sort: st.SORT || st.sort || '',
          cor_hex: st.COLOR || st.color || '',
          codigo_etapa_sugerido: pipStageCode(id, stageIndex + 1),
          id_bitrix: st.STATUS_ID || st.statusId || st.ID || st.id || '',
        });
      });
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(pipeRows), 'Pipelines');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stageRows), 'Estagios');
    pipDownloadWorkbook(wb, `pipelines-export-${new Date().toISOString().slice(0,10)}.xlsx`);
    pipCloseModal('pip-modal-export');
    toast('Exportacao gerada.', 'ok');
  } catch(e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function pipSheetRows(wb, names) {
  const sheetName = names.find(name => wb.SheetNames.includes(name));
  if (!sheetName) return null;
  return XLSX.utils.sheet_to_json(wb.Sheets[sheetName], { defval: '' });
}

function pipCleanRow(row) {
  const out = {};
  Object.entries(row || {}).forEach(([key, value]) => {
    out[String(key).trim()] = typeof value === 'string' ? value.trim() : value;
  });
  return out;
}

function pipIsEmptyRow(row) {
  return Object.values(row || {}).every(v => String(v ?? '').trim() === '');
}

function pipValidateImport(pipelineRowsRaw, stageRowsRaw) {
  const errors = [];
  const warnings = [];
  const autofill = [];
  const pipelineRows = pipelineRowsRaw.map(pipCleanRow).filter(row => !pipIsEmptyRow(row)).map((row, i) => ({ ...row, __sourceIndex: i }));
  const stageRows = stageRowsRaw.map(pipCleanRow).filter(row => !pipIsEmptyRow(row)).map((row, i) => ({ ...row, __sourceIndex: i }));
  const requiredP = ['chave_pipeline', 'nome_pipeline'];
  const requiredS = ['chave_pipeline', 'nome_etapa', 'semantica'];
  if (!pipelineRows.length) errors.push('Aba Pipelines sem linhas validas.');
  if (!stageRows.length) errors.push('Aba Estagios sem linhas validas.');
  requiredP.forEach(col => { if (!pipelineRowsRaw.length || !(col in pipCleanRow(pipelineRowsRaw[0]))) errors.push(`Aba Pipelines sem coluna ${col}.`); });
  requiredS.forEach(col => { if (!stageRowsRaw.length || !(col in pipCleanRow(stageRowsRaw[0]))) errors.push(`Aba Estagios sem coluna ${col}.`); });

  const pipeKeys = new Set();
  const pipeCodes = new Set();
  const stageCodes = new Set();
  const validSem = new Set(['ANDAMENTO', 'SUCESSO', 'FALHA']);
  pipelineRows.forEach((row, i) => {
    const line = i + 2;
    if (!row.chave_pipeline) errors.push(`Pipelines linha ${line}: chave_pipeline obrigatoria.`);
    if (!row.nome_pipeline) errors.push(`Pipelines linha ${line}: nome_pipeline obrigatorio.`);
    if (pipeKeys.has(row.chave_pipeline)) errors.push(`Pipelines linha ${line}: chave_pipeline duplicada.`);
    pipeKeys.add(row.chave_pipeline);
    if (!row.sort) { row.sort = (i + 1) * 100; autofill.push(`Pipeline ${row.chave_pipeline}: sort ${row.sort}`); }
    row.sort = parseInt(row.sort, 10) || ((i + 1) * 100);
    if (!row.codigo_pipeline_sugerido) { row.codigo_pipeline_sugerido = pipPadCode('PIP', i + 1); autofill.push(`Pipeline ${row.chave_pipeline}: codigo ${row.codigo_pipeline_sugerido}`); }
    if (pipeCodes.has(row.codigo_pipeline_sugerido)) errors.push(`Pipelines linha ${line}: codigo_pipeline_sugerido duplicado.`);
    pipeCodes.add(row.codigo_pipeline_sugerido);
  });
  const pipelineNumberByKey = {};
  pipelineRows.forEach((row, i) => {
    pipelineNumberByKey[row.chave_pipeline] = i + 1;
  });

  stageRows.forEach((row, i) => {
    const line = i + 2;
    if (!row.chave_pipeline) errors.push(`Estagios linha ${line}: chave_pipeline obrigatoria.`);
    if (!row.nome_etapa) errors.push(`Estagios linha ${line}: nome_etapa obrigatorio.`);
    const sem = String(row.semantica || '').trim().toUpperCase();
    if (!validSem.has(sem)) errors.push(`Estagios linha ${line}: semantica invalida.`);
    row.semantica = sem;
    if (row.chave_pipeline && !pipeKeys.has(row.chave_pipeline)) errors.push(`Estagios linha ${line}: chave_pipeline nao existe em Pipelines.`);
    if (row.cor_hex && !stageNormalizeHex(row.cor_hex)) errors.push(`Estagios linha ${line}: cor_hex invalida.`);
    row.cor_hex = stageNormalizeHex(row.cor_hex) || '#9BC2E6';
  });

  const stagesByKey = pipImportGroupStages(stageRows);
  pipelineRows.forEach(pipe => {
    const rows = stagesByKey[pipe.chave_pipeline] || [];
    const orderedRows = pipSortStagesForImport(rows);
    const successRows = rows.filter(row => row.semantica === 'SUCESSO');
    if (successRows.length > 1) errors.push(`Pipeline ${pipe.chave_pipeline}: use apenas uma etapa com semantica SUCESSO.`);

    orderedRows.forEach((row, index) => {
      const nextSort = pipStageSortForOrderedIndex(index);
      const currentSort = parseInt(row.sort, 10);
      if (currentSort !== nextSort) {
        const action = currentSort ? `sort ajustado de ${currentSort} para ${nextSort}` : `sort ${nextSort}`;
        row.sort = nextSort;
        autofill.push(`Etapa ${row.nome_etapa}: ${action}`);
      }

      const nextCode = pipStageCode(pipelineNumberByKey[row.chave_pipeline], index + 1);
      if (!row.codigo_etapa_sugerido || row.codigo_etapa_autopreenchido) {
        row.codigo_etapa_sugerido = nextCode;
        row.codigo_etapa_autopreenchido = true;
        autofill.push(`Etapa ${row.nome_etapa}: codigo ${row.codigo_etapa_sugerido}`);
      }
    });
  });

  stageRows.forEach((row, i) => {
    const line = i + 2;
    if (row.codigo_etapa_sugerido) {
      if (stageCodes.has(row.codigo_etapa_sugerido)) errors.push(`Estagios linha ${line}: codigo_etapa_sugerido duplicado.`);
      stageCodes.add(row.codigo_etapa_sugerido);
    }
  });

  pipelineRows.forEach(row => {
    if (!stageRows.some(st => st.chave_pipeline === row.chave_pipeline)) warnings.push(`Pipeline ${row.chave_pipeline} nao possui etapas.`);
  });
  pipelineRows.forEach((row, i) => { row.__planIndex = i; });
  stageRows.forEach((row, i) => { row.__planIndex = i; });
  return { pipelineRows, stageRows, errors, warnings, autofill };
}

function pipImportPublicRows(rows) {
  return rows.map(row => {
    const out = {};
    Object.entries(row).forEach(([key, value]) => {
      if (!key.startsWith('__')) out[key] = value;
    });
    return out;
  });
}

function pipImportCell(kind, index, col, value, type = 'text') {
  if (!pipImportPlan) return;
  const rows = kind === 'pipeline' ? pipImportPlan.pipelineRows : pipImportPlan.stageRows;
  if (!rows || !rows[index]) return;
  rows[index][col] = type === 'number' ? (value === '' ? '' : parseInt(value, 10) || '') : value;
  if (kind === 'stage' && col === 'codigo_etapa_sugerido') rows[index].codigo_etapa_autopreenchido = false;
  const pipelineRows = pipImportPublicRows(pipImportPlan.pipelineRows);
  const stageRows = pipImportPublicRows(pipImportPlan.stageRows);
  pipImportPlan = pipValidateImport(pipelineRows, stageRows);
  pipRenderImportPreview(pipImportPlan);
}

function pipImportInput(kind, index, col, value, type = 'text') {
  const safeValue = escHtml(value ?? '');
  return `<input class="pip-import-input" value="${safeValue}" onchange="pipImportCell('${kind}', ${index}, '${col}', this.value, '${type}')" />`;
}

function pipImportSelect(kind, index, col, value) {
  const options = ['ANDAMENTO', 'SUCESSO', 'FALHA'];
  return `<select class="pip-import-input" onchange="pipImportCell('${kind}', ${index}, '${col}', this.value)">
    ${options.map(opt => `<option value="${opt}" ${String(value).toUpperCase() === opt ? 'selected' : ''}>${opt}</option>`).join('')}
  </select>`;
}

function pipRenderPipelineImportRows(rows) {
  return rows.map((row, index) => `
    <tr>
      <td>${pipImportInput('pipeline', row.__planIndex ?? index, 'chave_pipeline', row.chave_pipeline)}</td>
      <td>${pipImportInput('pipeline', row.__planIndex ?? index, 'nome_pipeline', row.nome_pipeline)}</td>
      <td>${pipImportInput('pipeline', row.__planIndex ?? index, 'sort', row.sort, 'number')}</td>
      <td>${pipImportInput('pipeline', row.__planIndex ?? index, 'codigo_pipeline_sugerido', row.codigo_pipeline_sugerido)}</td>
      <td>${pipImportInput('pipeline', row.__planIndex ?? index, 'is_default', row.is_default || 'N')}</td>
    </tr>
  `).join('');
}

function pipRenderStageImportRows(rows) {
  return rows.map((row, index) => `
    <tr>
      <td>${pipImportInput('stage', row.__planIndex ?? index, 'chave_pipeline', row.chave_pipeline)}</td>
      <td>${pipImportInput('stage', row.__planIndex ?? index, 'nome_etapa', row.nome_etapa)}</td>
      <td>${pipImportSelect('stage', row.__planIndex ?? index, 'semantica', row.semantica)}</td>
      <td>${pipImportInput('stage', row.__planIndex ?? index, 'sort', row.sort, 'number')}</td>
      <td>${pipImportInput('stage', row.__planIndex ?? index, 'cor_hex', row.cor_hex)}</td>
      <td>${pipImportInput('stage', row.__planIndex ?? index, 'codigo_etapa_sugerido', row.codigo_etapa_sugerido)}</td>
    </tr>
  `).join('');
}

function pipRenderImportPreview(plan) {
  const box = document.getElementById('pip-import-preview');
  const btn = document.getElementById('pip-import-run-btn');
  if (!box || !btn) return;
  const stageCountByKey = {};
  plan.stageRows.forEach(row => { stageCountByKey[row.chave_pipeline] = (stageCountByKey[row.chave_pipeline] || 0) + 1; });
  const details = plan.pipelineRows.map(row => `${escHtml(row.nome_pipeline)}: ${stageCountByKey[row.chave_pipeline] || 0} etapa(s)`).join('<br>');
  const orderedStageRows = [];
  const stagesByKey = pipImportGroupStages(plan.stageRows);
  plan.pipelineRows.forEach(row => {
    orderedStageRows.push(...pipSortStagesForImport(stagesByKey[row.chave_pipeline] || []));
  });
  const remainingStages = plan.stageRows.filter(row => !plan.pipelineRows.some(pipe => pipe.chave_pipeline === row.chave_pipeline));
  orderedStageRows.push(...pipSortStagesForImport(remainingStages));
  box.innerHTML = `
    <div class="${plan.errors.length ? 'err' : 'ok'}">${plan.errors.length ? 'Erros encontrados. Corrija a planilha antes de executar.' : 'Validacao concluida. Pronto para executar.'}</div>
    <div>${plan.pipelineRows.length} pipeline(s), ${plan.stageRows.length} etapa(s).</div>
    <div style="margin-top:6px;">${details || 'Nenhuma linha valida encontrada.'}</div>
    ${plan.autofill.length ? `<div class="warn" style="margin-top:8px;">Autopreenchimentos/ajustes: ${plan.autofill.length}</div><ul>${plan.autofill.slice(0,8).map(x => `<li>${escHtml(x)}</li>`).join('')}</ul>` : ''}
    <div class="pip-import-editor-note">Revise e edite os valores abaixo antes de executar. A ordem das etapas segue o sort da planilha.</div>
    <div class="pip-import-editor">
      <div class="pip-import-table-title">Aba Pipelines</div>
      <div class="pip-import-table-wrap">
        <table class="pip-import-table">
          <thead><tr><th>chave_pipeline</th><th>nome_pipeline</th><th>sort</th><th>codigo_pipeline_sugerido</th><th>is_default</th></tr></thead>
          <tbody>${pipRenderPipelineImportRows(plan.pipelineRows)}</tbody>
        </table>
      </div>
      <div class="pip-import-table-title">Aba Estagios</div>
      <div class="pip-import-table-wrap">
        <table class="pip-import-table">
          <thead><tr><th>chave_pipeline</th><th>nome_etapa</th><th>semantica</th><th>sort</th><th>cor_hex</th><th>codigo_etapa_sugerido</th></tr></thead>
          <tbody>${pipRenderStageImportRows(orderedStageRows)}</tbody>
        </table>
      </div>
    </div>
    ${plan.warnings.length ? `<div class="warn" style="margin-top:8px;">Avisos</div><ul>${plan.warnings.map(x => `<li>${escHtml(x)}</li>`).join('')}</ul>` : ''}
    ${plan.errors.length ? `<div class="err" style="margin-top:8px;">Erros</div><ul>${plan.errors.map(x => `<li>${escHtml(x)}</li>`).join('')}</ul>` : ''}
  `;
  btn.disabled = plan.errors.length > 0 || plan.pipelineRows.length === 0;
}

async function pipHandleImportFile(file) {
  if (!file) return;
  try {
    pipRequireXlsx();
    const buffer = await file.arrayBuffer();
    const wb = XLSX.read(buffer, { type: 'array' });
    const pRows = pipSheetRows(wb, ['Pipelines']);
    const sRows = pipSheetRows(wb, ['Estagios', 'Estágios']);
    if (!pRows || !sRows) throw new Error('A planilha precisa conter as abas Pipelines e Estagios.');
    pipImportPlan = pipValidateImport(pRows, sRows);
    pipRenderImportPreview(pipImportPlan);
  } catch(e) {
    pipImportPlan = null;
    const box = document.getElementById('pip-import-preview');
    const btn = document.getElementById('pip-import-run-btn');
    if (box) box.innerHTML = `<span class="err">Erro: ${escHtml(e.message)}</span>`;
    if (btn) btn.disabled = true;
  }
}

function pipExtractCreatedPipelineId(data) {
  const r = data && data.result;
  if (!r) return null;
  if (r.id !== undefined) return r.id;
  if (r.category && r.category.id !== undefined) return r.category.id;
  if (r.item && r.item.id !== undefined) return r.item.id;
  return r;
}

async function pipRunImport() {
  if (!pipImportPlan || pipImportPlan.errors.length) return;
  const btn = document.getElementById('pip-import-run-btn');
  if (btn) btn.disabled = true;
  const results = [];
  try {
    const ctx = pipGetContext(true);
    const idByKey = {};
    const existingStagesByKey = {};
    for (const row of pipImportPlan.pipelineRows) {
      const fields = ctx.mode === 'crm'
        ? { name: row.nome_pipeline, sort: parseInt(row.sort, 10) || 100, isDefault: String(row.is_default || 'N').toUpperCase() === 'Y' ? 'Y' : 'N' }
        : { name: row.nome_pipeline, sort: parseInt(row.sort, 10) || 100 };
      const data = ctx.mode === 'spa'
        ? await ctx.adapter.addPipeline(fields)
        : await ctx.adapter.addPipeline(ctx.entityTypeId, fields);
      if (data.error) throw new Error(data.error_description || data.error);
      idByKey[row.chave_pipeline] = pipExtractCreatedPipelineId(data);
      if (idByKey[row.chave_pipeline] === null || idByKey[row.chave_pipeline] === undefined) {
        throw new Error(`Nao foi possivel identificar o ID da pipeline criada: ${row.nome_pipeline}`);
      }
      results.push(`Pipeline criada: ${row.nome_pipeline}`);
      existingStagesByKey[row.chave_pipeline] = await pipFetchStagesForPipeline(idByKey[row.chave_pipeline]);
    }
    const stagesByKey = pipImportGroupStages(pipImportPlan.stageRows);

    for (const pipeRow of pipImportPlan.pipelineRows) {
      const pipelineId = idByKey[pipeRow.chave_pipeline];
      const entityId = window.EntityContext.getStageEntityId(ctx.entityTypeId, pipelineId);
      const orderedRows = pipSortStagesForImport(stagesByKey[pipeRow.chave_pipeline] || []);
      const workRows = orderedRows.filter(row => !pipSemanticToApi(row.semantica));
      const semanticRows = orderedRows.filter(row => pipSemanticToApi(row.semantica));
      const existingStages = existingStagesByKey[pipeRow.chave_pipeline] || [];
      const keepStageIds = new Set();
      const semanticPools = {
        S: pipFindExistingStagesBySemantics(existingStages, 'S'),
        F: pipFindExistingStagesBySemantics(existingStages, 'F'),
      };
      const rowsToCreateAfterWork = [];
      const workRowsToCreate = [...workRows];

      for (const row of semanticRows) {
        const semantics = pipSemanticToApi(row.semantica);
        const reusableSemanticStage = semanticPools[semantics].shift() || null;
        if (!reusableSemanticStage) {
          rowsToCreateAfterWork.push(row);
          continue;
        }
        const existingSemanticStage = await pipUpdateExistingSemanticStage(ctx, reusableSemanticStage, row, results);
        const existingSemanticStageId = pipExistingStageId(existingSemanticStage);
        if (existingSemanticStageId) keepStageIds.add(String(existingSemanticStageId));
      }

      if (workRowsToCreate.length) {
        const reusableWorkStage = pipFindReusableWorkStage(existingStages, keepStageIds);
        if (reusableWorkStage) {
          const existingWorkStage = await pipUpdateExistingWorkStage(ctx, reusableWorkStage, workRowsToCreate[0], results);
          const existingWorkStageId = pipExistingStageId(existingWorkStage);
          if (existingWorkStageId) {
            keepStageIds.add(String(existingWorkStageId));
            workRowsToCreate.shift();
          }
        }
      }

      await pipDeleteStagesNotInImport(ctx, existingStages, keepStageIds, results);

      for (let index = 0; index < workRowsToCreate.length; index++) {
        const row = workRowsToCreate[index];
        const stageNumber = orderedRows.indexOf(row) + 1;
        const statusId = row.codigo_etapa_autopreenchido
          ? pipStageCode(pipelineId, stageNumber)
          : (row.codigo_etapa_sugerido || pipStageCode(pipelineId, stageNumber));
        const data = await ctx.adapter.addStage({
          ENTITY_ID: entityId,
          STATUS_ID: statusId,
          NAME: row.nome_etapa,
          SORT: parseInt(row.sort, 10) || pipStageImportRank(row),
          COLOR: row.cor_hex || '#9BC2E6',
          SEMANTICS: null,
        });
        if (data.error) throw new Error(data.error_description || data.error);
        results.push(`Etapa criada: ${row.nome_etapa}`);
      }

      for (const row of rowsToCreateAfterWork) {
        const semantics = pipSemanticToApi(row.semantica);

        const stageNumber = orderedRows.indexOf(row) + 1;
        const statusId = row.codigo_etapa_autopreenchido
          ? pipStageCode(pipelineId, stageNumber)
          : (row.codigo_etapa_sugerido || pipStageCode(pipelineId, stageNumber));
        const data = await ctx.adapter.addStage({
          ENTITY_ID: entityId,
          STATUS_ID: statusId,
          NAME: row.nome_etapa,
          SORT: parseInt(row.sort, 10) || pipStageImportRank(row),
          COLOR: row.cor_hex || (semantics === 'S' ? '#22C55E' : '#EF4444'),
          SEMANTICS: semantics,
        });
        if (data.error) throw new Error(data.error_description || data.error);
        results.push(`Etapa semantica criada: ${row.nome_etapa}`);
      }
    }
    const box = document.getElementById('pip-import-preview');
    if (box) box.innerHTML = `<div class="ok">Importacao concluida.</div><ul>${results.map(x => `<li>${escHtml(x)}</li>`).join('')}</ul>`;
    toast('Importacao concluida.', 'ok');
    pipLoad();
  } catch(e) {
    const box = document.getElementById('pip-import-preview');
    if (box) box.innerHTML += `<div class="err" style="margin-top:8px;">Falha: ${escHtml(e.message)}</div>`;
    toast('Erro: ' + e.message, 'er');
  } finally {
    if (btn) btn.disabled = false;
  }
}

function pipBindImportDrop() {
  const zone = document.getElementById('pip-import-drop');
  if (!zone || zone.dataset.bound === '1') return;
  zone.dataset.bound = '1';
  ['dragenter', 'dragover'].forEach(evt => {
    zone.addEventListener(evt, e => {
      e.preventDefault();
      zone.classList.add('dragover');
    });
  });
  ['dragleave', 'drop'].forEach(evt => {
    zone.addEventListener(evt, e => {
      e.preventDefault();
      zone.classList.remove('dragover');
    });
  });
  zone.addEventListener('drop', e => {
    const file = e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0];
    pipHandleImportFile(file);
  });
}

function stageSetPendingColor() {}

pipBindImportDrop();

window.pipSyncContextUI = pipSyncContextUI;
window.pipOpenModal = pipOpenModal;
window.pipCloseModal = pipCloseModal;
window.pipToggleChecked = pipToggleChecked;
window.pipClearSelection = pipClearSelection;
window.pipSelectAll = pipSelectAll;
window.pipDownloadModel = pipDownloadModel;
window.pipExportXlsx = pipExportXlsx;
window.pipHandleImportFile = pipHandleImportFile;
window.pipImportCell = pipImportCell;
window.pipRunImport = pipRunImport;
window.pipSyncExportOptions = pipSyncExportOptions;
window.stageToggleMobileRow = stageToggleMobileRow;
window.stageToggleAllMobileRows = stageToggleAllMobileRows;
window.stageOpenColorPicker = stageOpenColorPicker;
window.stageCloseColorPicker = stageCloseColorPicker;
window.stageSetPendingColor = stageSetPendingColor;
window.stageApplyColorPicker = stageApplyColorPicker;
window.stageSyncColorButton = stageSyncColorButton;
