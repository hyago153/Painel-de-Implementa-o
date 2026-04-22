/* -- SPA: VISAO GERAL E PROCESSO ATIVO ------------------------------------ */
const LS_ACTIVE_SPA_PREFIX = 'b24panel_active_spa';

let spaTypes = [];
let spaOverviewLoaded = false;
let spaOverviewLoading = false;
let spaOverviewError = '';
let spaTypeEditingId = null;

const SPA_TYPE_FLAGS = [
  { key: 'isAutomationEnabled', input: 'spa-type-automation', aliases: ['automationEnabled'] },
  { key: 'isStagesEnabled', input: 'spa-type-stages', aliases: ['stagesEnabled'] },
  { key: 'isCategoriesEnabled', input: 'spa-type-categories', aliases: ['categoriesEnabled'] },
  { key: 'isClientEnabled', input: 'spa-type-client', aliases: ['clientEnabled'] },
  { key: 'isLinkWithProductsEnabled', input: 'spa-type-products', aliases: ['linkWithProductsEnabled'] },
  { key: 'isObserversEnabled', input: 'spa-type-observers', aliases: ['observersEnabled'] },
  { key: 'isDocumentsEnabled', input: 'spa-type-documents', aliases: ['documentsEnabled'] },
  { key: 'isRecyclebinEnabled', input: 'spa-type-recyclebin', aliases: ['recyclebinEnabled'] },
  { key: 'isBizProcEnabled', input: 'spa-type-bizproc', aliases: ['bizProcEnabled'] },
];

function spaOverviewEnvKey() {
  const env = envs[activeIdx];
  const raw = env ? (env.url || env.name || String(activeIdx)) : String(activeIdx);
  return `${LS_ACTIVE_SPA_PREFIX}:${encodeURIComponent(raw)}`;
}

function spaBool(value) {
  return value === true || value === 'Y' || value === '1' || value === 1;
}

function spaFlag(type, keys) {
  for (const key of keys) {
    if (Object.prototype.hasOwnProperty.call(type, key)) return spaBool(type[key]);
  }
  return false;
}

function spaFlagBadge(label, enabled) {
  return `<span class="spa-flag ${enabled ? 'on' : 'off'}">${label}: ${enabled ? 'Sim' : 'Não'}</span>`;
}

function spaAssertOk(data) {
  if (data && data.error) throw new Error(data.error_description || data.error);
  return data;
}

function spaNormalizeType(type) {
  const id = parseInt(type.id, 10);
  const entityTypeId = parseInt(type.entityTypeId, 10);
  return {
    ...type,
    id,
    entityTypeId,
    title: type.title || type.name || `Processo ${id || ''}`.trim(),
  };
}

function spaExtractTypes(data) {
  const result = data && data.result;
  const list = Array.isArray(result) ? result :
    Array.isArray(result && result.types) ? result.types :
    Array.isArray(result && result.items) ? result.items : [];
  return list.map(spaNormalizeType).filter(type => type.id && type.entityTypeId);
}

function spaSaveActive(spa) {
  if (!spa) {
    localStorage.removeItem(spaOverviewEnvKey());
    return;
  }
  localStorage.setItem(spaOverviewEnvKey(), JSON.stringify({
    id: spa.id,
    entityTypeId: spa.entityTypeId,
    title: spa.title,
  }));
}

function spaRestoreActive() {
  let saved = null;
  try { saved = JSON.parse(localStorage.getItem(spaOverviewEnvKey())); } catch (_) { saved = null; }
  if (!saved || !saved.id) return null;
  const fromList = spaTypes.find(type => String(type.id) === String(saved.id));
  return window.EntityContext.setActiveSpa(fromList || saved);
}

function spaOverviewModeSync() {
  const isSpa = appMode === 'spa';
  document.querySelectorAll('.overview-crm').forEach(el => { el.style.display = isSpa ? 'none' : ''; });
  document.querySelectorAll('.overview-spa').forEach(el => { el.style.display = isSpa ? '' : 'none'; });

  const state = document.getElementById('spa-overview-state');
  if (isSpa && state && !spaOverviewLoaded && !spaOverviewLoading) {
    state.textContent = 'Clique em "Atualizar lista" para carregar os processos inteligentes.';
  }
}

function spaOverviewRenderSelect() {
  const select = document.getElementById('spa-active-select');
  if (!select) return;

  const active = window.EntityContext.getActiveSpa();
  select.innerHTML = '';

  if (spaTypes.length === 0) {
    const opt = document.createElement('option');
    opt.value = '';
    opt.textContent = 'Nenhum processo carregado';
    select.appendChild(opt);
  } else {
    const empty = document.createElement('option');
    empty.value = '';
    empty.textContent = 'Selecione um processo';
    select.appendChild(empty);

    spaTypes.forEach(type => {
      const opt = document.createElement('option');
      opt.value = String(type.id);
      opt.textContent = `${type.title} (ID ${type.id} / entityTypeId ${type.entityTypeId})`;
      if (active && String(active.id) === String(type.id)) opt.selected = true;
      select.appendChild(opt);
    });
  }

  spaOverviewRenderActiveSummary();
}

function spaOverviewRenderActiveSummary() {
  const box = document.getElementById('spa-active-summary');
  const hint = document.getElementById('spa-active-hint');
  if (!box) return;

  const active = window.EntityContext.getActiveSpa();
  if (!active) {
    box.innerHTML = '<div class="entity-meta">Nenhum processo selecionado.</div>';
    if (hint) hint.textContent = 'O processo selecionado fica salvo por ambiente.';
    return;
  }

  box.innerHTML = `
    <div class="spa-active-title">${escHtml(active.title)}</div>
    <div class="entity-meta">id: ${active.id}  entityTypeId: ${active.entityTypeId}</div>
  `;
  if (hint) hint.textContent = 'Este processo sera usado pelas abas SPA.';
}

function spaOverviewRenderTypes() {
  const state = document.getElementById('spa-overview-state');
  const grid = document.getElementById('spa-type-grid');
  if (!state || !grid) return;

  grid.innerHTML = '';
  if (spaOverviewLoading) {
    state.style.display = '';
    state.textContent = 'Carregando processos inteligentes...';
    return;
  }

  if (spaTypes.length === 0) {
    state.style.display = '';
    state.textContent = spaOverviewError || (spaOverviewLoaded
      ? 'Nenhum processo inteligente foi encontrado neste portal.'
      : 'Clique em "Atualizar lista" para carregar os processos inteligentes.');
    return;
  }

  state.style.display = 'none';
  const active = window.EntityContext.getActiveSpa();
  spaTypes.forEach(type => {
    const isActive = active && String(active.id) === String(type.id);
    const card = document.createElement('div');
    card.className = 'spa-type-card' + (isActive ? ' active' : '');
    card.innerHTML = `
      <div class="spa-type-head">
        <div>
          <div class="spa-type-title">${escHtml(type.title)}</div>
          <div class="entity-meta">id: ${type.id}  entityTypeId: ${type.entityTypeId}</div>
        </div>
        ${isActive ? '<span class="spa-active-badge">Ativo</span>' : ''}
      </div>
      <div class="spa-flag-row">
        ${spaFlagBadge('Automacao', spaFlag(type, ['isAutomationEnabled', 'automationEnabled']))}
        ${spaFlagBadge('Funis', spaFlag(type, ['isCategoriesEnabled', 'categoriesEnabled']))}
        ${spaFlagBadge('Etapas', spaFlag(type, ['isStagesEnabled', 'stagesEnabled']))}
        ${spaFlagBadge('Campos UF', spaFlag(type, ['isUseInUserfieldEnabled', 'isUserfieldEnabled', 'userfieldEnabled']))}
      </div>
      <div class="spa-type-actions">
        <button class="btn sm" type="button" onclick="spaOverviewSelect(${type.id})">${isActive ? 'Selecionado' : 'Selecionar'}</button>
        <button class="btn sm" type="button" onclick="spaTypeOpenEdit(${type.id})">Editar</button>
        <button class="btn sm danger" type="button" onclick="spaTypeDelete(${type.id})">Excluir</button>
      </div>
    `;
    grid.appendChild(card);
  });
}

function spaOverviewRender() {
  spaOverviewModeSync();
  spaOverviewRenderSelect();
  spaOverviewRenderTypes();
}

function spaOverviewSelect(id) {
  const selected = spaTypes.find(type => String(type.id) === String(id));
  if (!selected) {
    window.EntityContext.setActiveSpa(null);
    spaSaveActive(null);
    spaOverviewRender();
    return;
  }

  const active = window.EntityContext.setActiveSpa(selected);
  spaSaveActive(active);
  spaOverviewRender();
  toast(`Processo ativo: ${escHtml(active.title)}.`, 'ok');
}

async function spaOverviewLoad(force = false) {
  if (spaOverviewLoading) return;
  if (spaOverviewLoaded && !force) {
    spaOverviewRender();
    return;
  }

  spaOverviewLoading = true;
  spaOverviewError = '';
  spaOverviewRender();

  try {
    const data = spaAssertOk(await window.EntityAdapter.spa.listTypes());
    spaTypes = spaExtractTypes(data);
    spaOverviewLoaded = true;

    const restored = spaRestoreActive();
    if (!restored && spaTypes.length === 1) {
      const only = window.EntityContext.setActiveSpa(spaTypes[0]);
      spaSaveActive(only);
    }

    spaOverviewRender();
    toast('Lista de processos inteligentes atualizada.', 'ok');
  } catch (e) {
    spaOverviewError = `Não foi possível listar os processos inteligentes: ${e.message}`;
    toast('Não foi possível listar os processos inteligentes.', 'er');
  } finally {
    spaOverviewLoading = false;
    spaOverviewRender();
  }
}

function spaTypeSetModalOpen(open) {
  const overlay = document.getElementById('spa-type-overlay');
  if (overlay) overlay.classList.toggle('open', open);
}

function spaTypeSetBusy(busy) {
  const saveBtn = document.getElementById('spa-type-save-btn');
  if (!saveBtn) return;
  saveBtn.disabled = busy;
  saveBtn.textContent = busy ? 'Salvando...' : 'Salvar';
}

function spaTypeFlagValue(type, flag) {
  return spaFlag(type, [flag.key, ...(flag.aliases || [])]);
}

function spaTypeFillForm(type = null) {
  const title = document.getElementById('spa-type-title');
  const info = document.getElementById('spa-type-id-info');
  if (title) title.value = type ? (type.title || '') : '';
  if (info) {
    info.textContent = type
      ? `id: ${type.id} · entityTypeId: ${type.entityTypeId}`
      : 'Novo processo inteligente';
  }

  SPA_TYPE_FLAGS.forEach(flag => {
    const input = document.getElementById(flag.input);
    if (input) input.checked = type ? spaTypeFlagValue(type, flag) : true;
  });
}

function spaTypeFieldsFromForm() {
  const titleEl = document.getElementById('spa-type-title');
  const title = titleEl ? titleEl.value.trim() : '';
  if (!title) throw new Error('Informe o titulo do processo.');

  const fields = { title };
  SPA_TYPE_FLAGS.forEach(flag => {
    const input = document.getElementById(flag.input);
    fields[flag.key] = input && input.checked ? 'Y' : 'N';
  });
  return fields;
}

function spaTypeOpenNew() {
  spaTypeEditingId = null;
  const heading = document.getElementById('spa-type-modal-title');
  if (heading) heading.textContent = 'Novo processo';
  spaTypeFillForm(null);
  spaTypeSetBusy(false);
  spaTypeSetModalOpen(true);
  setTimeout(() => document.getElementById('spa-type-title')?.focus(), 0);
}

function spaTypeOpenEdit(id) {
  const type = spaTypes.find(item => String(item.id) === String(id));
  if (!type) {
    toast('Processo não encontrado na lista carregada.', 'er');
    return;
  }

  spaTypeEditingId = type.id;
  const heading = document.getElementById('spa-type-modal-title');
  if (heading) heading.textContent = 'Editar processo';
  spaTypeFillForm(type);
  spaTypeSetBusy(false);
  spaTypeSetModalOpen(true);
  setTimeout(() => document.getElementById('spa-type-title')?.focus(), 0);
}

function spaTypeClose() {
  spaTypeSetModalOpen(false);
  spaTypeEditingId = null;
}

async function spaTypeSave() {
  try {
    const fields = spaTypeFieldsFromForm();
    spaTypeSetBusy(true);
    if (spaTypeEditingId) {
      spaAssertOk(await window.EntityAdapter.spa.updateType(spaTypeEditingId, fields));
      toast('Processo inteligente atualizado.', 'ok');
    } else {
      spaAssertOk(await window.EntityAdapter.spa.addType(fields));
      toast('Processo inteligente criado.', 'ok');
    }
    spaTypeClose();
    await spaOverviewLoad(true);
  } catch (e) {
    toast(e.message || 'Não foi possível salvar o processo inteligente.', 'er');
  } finally {
    spaTypeSetBusy(false);
  }
}

async function spaTypeDelete(id) {
  const type = spaTypes.find(item => String(item.id) === String(id));
  if (!type) {
    toast('Processo não encontrado na lista carregada.', 'er');
    return;
  }

  if (!confirm(`Excluir o processo "${type.title}"? Esta ação não pode ser desfeita.`)) return;
  const typed = prompt(`Digite EXCLUIR para confirmar a exclusão de "${type.title}".`);
  if (typed !== 'EXCLUIR') {
    toast('Exclusao cancelada.', 'er');
    return;
  }

  try {
    spaAssertOk(await window.EntityAdapter.spa.deleteType(type.id));
    const active = window.EntityContext.getActiveSpa();
    if (active && String(active.id) === String(type.id)) {
      window.EntityContext.setActiveSpa(null);
      spaSaveActive(null);
    }
    toast('Processo inteligente excluido.', 'ok');
    await spaOverviewLoad(true);
  } catch (e) {
    toast(e.message || 'Não foi possível excluir o processo inteligente.', 'er');
  }
}

function spaOverviewEnsureLoaded() {
  spaOverviewModeSync();
  spaRestoreActive();
  spaOverviewRender();
  if (appMode === 'spa' && !spaOverviewLoaded && !spaOverviewLoading && getWebhook()) {
    spaOverviewLoad(false);
  }
}

function spaOverviewOnEnvironmentChanged() {
  spaTypes = [];
  spaOverviewLoaded = false;
  spaOverviewLoading = false;
  spaOverviewError = '';
  window.EntityContext.setActiveSpa(null);
  spaRestoreActive();
  spaOverviewRender();
}

window.spaOverviewLoad = spaOverviewLoad;
window.spaOverviewSelect = spaOverviewSelect;
window.spaOverviewEnsureLoaded = spaOverviewEnsureLoaded;
window.spaOverviewModeSync = spaOverviewModeSync;
window.spaOverviewOnEnvironmentChanged = spaOverviewOnEnvironmentChanged;
window.spaTypeOpenNew = spaTypeOpenNew;
window.spaTypeOpenEdit = spaTypeOpenEdit;
window.spaTypeClose = spaTypeClose;
window.spaTypeSave = spaTypeSave;
window.spaTypeDelete = spaTypeDelete;
