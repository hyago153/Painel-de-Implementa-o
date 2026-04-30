/* 
   MDULO: CONFIG DO CARD (Tpico 6)
    */

let cardEntityTypeId = 4;
let cardLeadType     = 1;
let cardPipelineId   = '0';
let cardPipelines    = [];
let cardConfig       = [];
let cardAllFields    = {};
let cardLoaded       = false;
let cardSelectedHiddenFields = new Set();
let cardExpandedHiddenFields = new Set();
let cardSelectedSectionFields = new Set();
let cardHiddenBulkCollapsed = false;
let cardHiddenPanelCollapsed = false;
let cardEditFieldRef = null;
let cardLastContextKey = '';

function cardMode() {
  return window.EntityContext ? window.EntityContext.getCurrentMode() : 'crm';
}

function cardGetContext(requireSelection = true) {
  if (cardMode() === 'spa') {
    const spa = requireSelection
      ? window.EntityContext.requireActiveSpa()
      : window.EntityContext.getActiveSpa();
    return {
      mode: 'spa',
      spa,
      entityTypeId: spa ? spa.entityTypeId : null,
      entityId: spa ? 'CRM_' + spa.id : '',
    };
  }
  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(cardEntityTypeId);
  return {
    mode: 'crm',
    entityTypeId: cardEntityTypeId,
    entityId: CAMPOS_API[cardEntityTypeId] ? CAMPOS_API[cardEntityTypeId].entityId : '',
  };
}

function cardContextKey(ctx = cardGetContext(false)) {
  if (ctx.mode === 'spa') return ctx.spa ? `spa:${ctx.spa.id}:${ctx.entityTypeId}:${cardPipelineId}` : 'spa:none';
  return `crm:${cardEntityTypeId}:${cardPipelineId}:${cardLeadType}`;
}

function cardEnsureContextBanner() {
  let banner = document.getElementById('card-context-banner');
  if (banner) return banner;
  const tabs = document.getElementById('card-entity-tabs');
  if (!tabs || !tabs.parentElement) return null;
  banner = document.createElement('div');
  banner.id = 'card-context-banner';
  banner.className = 'form-hint';
  banner.style.marginTop = '8px';
  tabs.parentElement.insertBefore(banner, tabs.nextSibling);
  return banner;
}

function cardSyncContextUI() {
  const ctx = cardGetContext(false);
  const tabs = document.getElementById('card-entity-tabs');
  const leadRow = document.getElementById('card-lead-type-row');
  const pipRow = document.getElementById('card-pipeline-row');
  const pipLabel = pipRow ? pipRow.querySelector('.form-label') : null;
  const banner = cardEnsureContextBanner();

  if (ctx.mode === 'spa') {
    if (tabs) tabs.style.display = 'none';
    if (leadRow) leadRow.style.display = 'none';
    if (pipRow) pipRow.style.display = ctx.spa ? 'block' : 'none';
    if (pipLabel) pipLabel.textContent = 'Pipeline/Funil da SPA';
    if (ctx.spa) cardEntityTypeId = ctx.entityTypeId;
    if (banner) {
      banner.innerHTML = ctx.spa
        ? `SPA ativo: <strong>${escHtml(ctx.spa.title)}</strong> (id ${escHtml(String(ctx.spa.id))} / entityTypeId ${escHtml(String(ctx.entityTypeId))}). Selecione o pipeline antes de carregar o layout.`
        : 'Selecione um processo inteligente na Visão geral do módulo SPA antes de editar o card.';
      banner.style.display = '';
    }
    if (ctx.spa) cardLoadPipelines();
  } else {
    if (tabs) tabs.style.display = '';
    document.querySelectorAll('#card-entity-tabs .etab').forEach(t => {
      t.classList.toggle('active', String(t.dataset.entity) === String(cardEntityTypeId));
    });
    if (leadRow) leadRow.style.display = cardEntityTypeId === 1 ? 'flex' : 'none';
    if (pipRow) pipRow.style.display = cardEntityTypeId === 2 ? 'block' : 'none';
    if (pipLabel) pipLabel.textContent = 'Pipeline';
    if (banner) banner.style.display = 'none';
    if (cardEntityTypeId === 2) cardLoadPipelines();
  }

  const key = cardContextKey(ctx);
  if (key !== cardLastContextKey) {
    cardShowEditor(false);
    cardConfig = [];
    cardAllFields = {};
    cardSelectedHiddenFields.clear();
    cardExpandedHiddenFields.clear();
    cardSelectedSectionFields.clear();
    cardLoaded = false;
    cardLastContextKey = key;
  }
}

function cardSelectEntity(eid, btn) {
  cardEntityTypeId = eid;
  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(eid);
  document.querySelectorAll('#card-entity-tabs .etab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('card-lead-type-row').style.display    = eid === 1 ? 'flex' : 'none';
  document.getElementById('card-pipeline-row').style.display     = eid === 2 ? 'block' : 'none';
  if (eid === 2) cardLoadPipelines();
  cardShowEditor(false);
}

function cardSelectLeadType(type) {
  cardLeadType = type;
  document.getElementById('card-lt-simples').classList.toggle('active',  type === 1);
  document.getElementById('card-lt-repetido').classList.toggle('active', type === 2);
}

function cardSelectPipeline(value) {
  cardPipelineId = String(value || '0');
  cardShowEditor(false);
  cardConfig = [];
  cardAllFields = {};
  cardSelectedHiddenFields.clear();
  cardExpandedHiddenFields.clear();
  cardSelectedSectionFields.clear();
  cardLoaded = false;
  cardLastContextKey = cardContextKey(cardGetContext(false));
}

async function cardLoadPipelines() {
  try {
    const ctx = cardGetContext(false);
    const entityTypeId = ctx.mode === 'spa' ? ctx.entityTypeId : 2;
    if (!entityTypeId) return;

    const data = await call('crm.category.list', { entityTypeId });
    cardPipelines = (data.result && data.result.categories) ? data.result.categories : (Array.isArray(data.result) ? data.result : []);
    const sel = document.getElementById('card-pipeline-select');
    const previous = cardPipelineId || '0';
    sel.innerHTML = '<option value="0">Pipeline padrão (ID 0)</option>';
    cardPipelines.forEach(c => {
      if (String(c.id) !== '0') {
        const opt = document.createElement('option');
        opt.value = String(c.id);
        opt.textContent = escHtml(c.name) + ' (#' + c.id + ')';
        sel.appendChild(opt);
      }
    });
    sel.value = [...sel.options].some(opt => opt.value === previous) ? previous : '0';
    cardPipelineId = sel.value;
  } catch(e) { /* silencioso */ }
}

function buildCardConfigParams(entityTypeId, scope = 'P', extrasValues = {}, mode = 'crm') {
  const params = { entityTypeId, scope };
  const extras = {};
  if (mode === 'spa') extras.categoryId = parseInt(extrasValues.categoryId) || 0;
  else if (entityTypeId === 2) extras.dealCategoryId = parseInt(extrasValues.dealCategoryId) || 0;
  else if (entityTypeId === 1) extras.leadCustomerType = parseInt(extrasValues.leadCustomerType) || 1;
  if (Object.keys(extras).length) params.extras = extras;
  return params;
}

function cardBuildCallParams(scope = 'P') {
  const ctx = cardGetContext(true);
  const base = buildCardConfigParams(ctx.entityTypeId, scope, {
    dealCategoryId: cardPipelineId,
    categoryId: cardPipelineId,
    leadCustomerType: cardLeadType,
  }, ctx.mode);
  return { methodPrefix: 'crm.item.details.configuration', base };
}

function cardExtractConfigSections(data) {
  const result = data && data.result;
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.configuration)) return result.configuration;
  if (result && Array.isArray(result.data)) return result.data;
  if (result && Array.isArray(result.items)) return result.items;
  return [];
}

async function cardFetchConfig() {
  const method = 'crm.item.details.configuration.get';
  // Tenta carregar a config pessoal (scope P) primeiro; se vazia, carrega a comum (scope C)
  const dataP = await call(method, cardBuildCallParams('P').base);
  if (dataP.error) throw new Error(dataP.error_description || dataP.error);
  const personalConfig = cardExtractConfigSections(dataP);
  if (personalConfig.length > 0) return personalConfig;
  const dataC = await call(method, cardBuildCallParams('C').base);
  if (dataC.error) throw new Error(dataC.error_description || dataC.error);
  return cardExtractConfigSections(dataC);
}

async function cardFetchAllFields() {
  if (cardMode() === 'spa') return cardFetchSpaAllFields();

  const eid = cardEntityTypeId;
  const api  = CAMPOS_API[eid];

  // Busca campos nativos e primeira pgina de campos personalizados em paralelo
  const ufFilter = { ENTITY_ID: api.entityId };
  const [nRes, firstUfRes] = await Promise.all([
    call(api.fields),
    call(api.uf, { filter: ufFilter, start: 0 }),
  ]);

  // Paginão completa: continua buscando enquanto houver mais pginas (Bitrix retorna 50 por pgina)
  const allCustoms = Array.isArray(firstUfRes.result) ? [...firstUfRes.result] : [];
  const PAGE_SIZE = 50;
  let start = PAGE_SIZE;
  while (allCustoms.length === start) {
    const pageRes = await call(api.uf, { filter: ufFilter, start });
    const page = Array.isArray(pageRes.result) ? pageRes.result : [];
    if (!page.length) break;
    allCustoms.push(...page);
    start += PAGE_SIZE;
  }

  const customOptionsMap = new Map();
  const customRawMetaMap = new Map();
  allCustoms.forEach(f => {
    const id = f.FIELD_NAME || f.fieldName;
    if (!id) return;
    if (Array.isArray(f.LIST) && f.LIST.length) {
      customOptionsMap.set(id, f.LIST.map(o => o.VALUE || o.value || '').filter(Boolean));
    }
    customRawMetaMap.set(id, {
      rawId:     f.ID    || f.id    || null,
      mandatory: (f.MANDATORY || f.mandatory || 'N') === 'Y',
      multiple:  (f.MULTIPLE  || f.multiple  || 'N') === 'Y',
      settings:  f.SETTINGS || f.settings || {},
      helpMessage: labelFromBitrix(f.HELP_MESSAGE || f.helpMessage) || '',
    });
  });

  // Tenta userfieldconfig.list para obter labels reais (camelCase conforme o módulo exige)
  const ucfgMap = new Map();
  try {
    let ucfgStart = 0;
    while (true) {
      const ucfgRes = await call('userfieldconfig.list', {
        moduleId: 'crm',
        select: { 0: '*', language: 'br' },
        filter: { entityId: api.entityId },
        start: ucfgStart,
      });
      const ucfgPage = Array.isArray(ucfgRes.result)
        ? ucfgRes.result
        : (ucfgRes.result && Array.isArray(ucfgRes.result.fields)
          ? ucfgRes.result.fields
          : (Array.isArray(ucfgRes.fields) ? ucfgRes.fields : []));
      if (!ucfgPage.length) break;
      ucfgPage.forEach(f => {
        const id = f.fieldName || f.FIELD_NAME;
        if (!id) return;
        const label = labelFromBitrix(f.editFormLabel)
          || labelFromBitrix(f.listColumnLabel)
          || labelFromBitrix(f.EDIT_FORM_LABEL)
          || '';
        const type = f.userTypeId || f.USER_TYPE_ID || '';
        if (label) ucfgMap.set(id, { label, type });
      });
      ucfgStart += PAGE_SIZE;
      if (ucfgPage.length < PAGE_SIZE) break;
    }
  } catch(e) {
    console.warn('[Bitrix24 Panel] userfieldconfig.list indisponível (card):', e.message);
  }

  const result = {};
  const native = nRes.result || {};
  Object.entries(native).forEach(([id, info]) => {
    result[id] = { label: (info && info.title) ? info.title : id, type: (info && info.type) ? info.type : '' };
  });

  // Aplica labels: prioridade ucfgMap > EDIT_FORM_LABEL > title crm.fields > FIELD_NAME
  allCustoms.forEach(f => {
    const id = f.FIELD_NAME || f.fieldName;
    if (!id) return;
    const existing = result[id];
    if (ucfgMap.has(id)) {
      const u = ucfgMap.get(id);
      result[id] = { label: u.label, type: u.type };
      return;
    }
    const label = labelFromBitrix(f.EDIT_FORM_LABEL)
      || labelFromBitrix(f.LIST_COLUMN_LABEL)
      || (existing && existing.label && existing.label !== id ? existing.label : id);
    const type = f.USER_TYPE_ID || (existing && existing.type) || '';
    result[id] = { label, type };
  });
  Object.entries(result).forEach(([id, info]) => {
    const meta = customRawMetaMap.get(id) || {};
    const isCustom = id.startsWith('UF_') || !!meta.rawId;
    result[id] = {
      ...info,
      origin: isCustom ? 'custom' : 'native',
      rawId: meta.rawId ?? null,
      mandatory: meta.mandatory ?? false,
      multiple: meta.multiple ?? false,
      options: customOptionsMap.get(id) || [],
      settings: meta.settings || {},
      helpMessage: meta.helpMessage || '',
    };
  });
  return result;
}

async function cardFetchSpaAllFields() {
  const ctx = cardGetContext(true);
  const [nRes, firstUfRes] = await Promise.all([
    call('crm.item.fields', { entityTypeId: ctx.entityTypeId, useOriginalUfNames: 'Y' }),
    call('userfieldconfig.list', {
      moduleId: 'crm',
      select: { 0: '*', language: 'br' },
      filter: { entityId: ctx.entityId },
      start: 0,
    }),
  ]);

  const PAGE_SIZE = 50;
  const allCustoms = camposExtractUserFieldPage(firstUfRes);
  let start = PAGE_SIZE;
  while (allCustoms.length === start) {
    const pageRes = await call('userfieldconfig.list', {
      moduleId: 'crm',
      select: { 0: '*', language: 'br' },
      filter: { entityId: ctx.entityId },
      start,
    });
    const page = camposExtractUserFieldPage(pageRes);
    if (!page.length) break;
    allCustoms.push(...page);
    start += PAGE_SIZE;
  }

  const result = {};
  const keyToFieldId = new Map();
  Object.entries(camposExtractNativeFields(nRes)).forEach(([id, info]) => {
    result[id] = {
      label: (info && (info.title || info.formLabel)) ? (info.title || info.formLabel) : id,
      type: (info && info.type) ? info.type : '',
      origin: id.startsWith('UF_') ? 'custom' : 'native',
    };
    keyToFieldId.set(camposFieldKey(id), id);
  });

  allCustoms.forEach(f => {
    const rawId = camposFieldName(f);
    if (!rawId) return;
    const id = result[rawId] ? rawId : (keyToFieldId.get(camposFieldKey(rawId)) || rawId);
    const options = camposFieldOptions(f);
    result[id] = {
      ...(result[id] || {}),
      label: labelFromBitrix(f.editFormLabel)
        || labelFromBitrix(f.listColumnLabel)
        || labelFromBitrix(f.EDIT_FORM_LABEL)
        || labelFromBitrix(f.LIST_COLUMN_LABEL)
        || id,
      type: f.userTypeId || f.USER_TYPE_ID || (result[id] && result[id].type) || '',
      origin: 'custom',
      rawId: f.ID || f.id || null,
      mandatory: camposBool(f.MANDATORY ?? f.mandatory ?? f.isMandatory),
      multiple: camposBool(f.MULTIPLE ?? f.multiple ?? f.isMultiple),
      options,
      settings: f.SETTINGS || f.settings || {},
      helpMessage: labelFromBitrix(f.HELP_MESSAGE || f.helpMessage) || '',
    };
    keyToFieldId.set(camposFieldKey(rawId), id);
  });

  Object.entries(result).forEach(([id, info]) => {
    result[id] = {
      ...info,
      rawId: info.rawId ?? null,
      mandatory: info.mandatory ?? false,
      multiple: info.multiple ?? false,
      options: info.options || [],
      settings: info.settings || {},
      helpMessage: info.helpMessage || '',
    };
  });
  return result;
}

async function cardLoad() {
  cardSyncContextUI();
  const wrap = document.getElementById('card-editor-wrap');
  wrap.style.display = 'none';
  try {
    toast('Carregando configuração', 'wn', 2000);
    const [cfg, fields] = await Promise.all([cardFetchConfig(), cardFetchAllFields()]);
    cardConfig    = cfg.map(sec => ({
      name:     sec.name     || sec.title || '',
      title:    sec.title    || '',
      type:     sec.type     || 'section',
      elements: (sec.elements || []).map(el => ({
        name: el.name,
        optionFlags: el.optionFlags ?? 0,
        ...(el.options ? { options: el.options } : {}),
      })),
    }));
    cardAllFields = fields;
    cardLoaded    = true;
    cardShowEditor(true);
    cardRender();
    toast('Configuração carregada.', 'ok');
  } catch(e) {
    toast('Erro ao carregar: ' + e.message, 'er');
  }
}

function cardShowEditor(visible) {
  document.getElementById('card-editor-wrap').style.display = visible ? 'block' : 'none';
  if (visible) document.getElementById('card-unsaved-badge').style.display = 'none';
}

function cardRender() {
  cardRenderSections();
  cardRenderHiddenFields();
}

function cardIcon(name) {
  const common = 'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"';
  const icons = {
    up: `<svg ${common}><path d="m18 15-6-6-6 6"/></svg>`,
    down: `<svg ${common}><path d="m6 9 6 6 6-6"/></svg>`,
    expand: `<svg ${common}><path d="m9 18 6-6-6-6"/></svg>`,
    collapse: `<svg ${common}><path d="m6 9 6 6 6-6"/></svg>`,
    edit: `<svg ${common}><path d="M12 20h9"/><path d="M16.5 3.5a2.1 2.1 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>`,
    trash: `<svg ${common}><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v5"/><path d="M14 11v5"/></svg>`
  };
  return icons[name] || '';
}

function cardRenderSections() {
  const container = document.getElementById('card-sections-container');
  if (!cardConfig.length) {
    cardSelectedSectionFields.clear();
    container.innerHTML = '<div class="section-empty">Nenhuma seção. Adicione uma abaixo.</div>';
    return;
  }
  cardPruneSelectedSectionFields();
  container.innerHTML = cardConfig.map((sec, si) => {
    const count = (sec.elements || []).length;
    const collapsed = !!sec._collapsed;
    const selectedCount = cardGetSectionSelectedCount(si);
    const allSelected = count > 0 && selectedCount === count;
    const fieldsHtml = count === 0
      ? '<div class="section-empty">Nenhum campo nesta seção. Arraste campos ocultos para cá.</div>'
      : `<table class="card-fields-table">
          <thead><tr><th>Campo</th><th>Tipo</th><th>Opções</th><th>MOSTRAR SEMPRE</th><th>Mover</th><th>Ações</th></tr></thead>
          <tbody>
            ${sec.elements.map((el, fi) => {
              const finfo  = cardAllFields[el.name] || { label: el.name, type: '' };
              const fixed  = el.optionFlags === 1;
              const typeL  = TIPO_LABEL[finfo.type] || finfo.type;
              const isCustomEditable = finfo.origin === 'custom' && finfo.rawId;
              const fieldIdArg = JSON.stringify(el.name);
              const selected = cardSelectedSectionFields.has(el.name);
              const optBadges = [];
              if (finfo.mandatory) optBadges.push('Obrigatrio');
              if (finfo.multiple) optBadges.push('Mltiplo');
              if (finfo.type === 'string' && finfo.settings && finfo.settings.ROWS) {
                optBadges.push(`${finfo.settings.ROWS} linha${Number(finfo.settings.ROWS) === 1 ? '' : 's'}`);
              }
              if (finfo.options && finfo.options.length) optBadges.push(`${finfo.options.length} opções`);
              const optsHtml = optBadges.length
                ? `<div class="card-field-options">${optBadges.map(o => `<span class="card-field-opt">${escHtml(o)}</span>`).join('')}</div>`
                : '<span style="color:#ccc;font-size:11px;"></span>';
              const editBtn = isCustomEditable
                ? `<button class="card-action-btn edit" onclick="cardOpenEdit(${si},${fi})">Editar</button>
                   <button class="card-action-btn del" onclick="cardDeleteField(${si},${fi})">Excluir</button>`
                : '<span style="color:#ccc;font-size:11px;">Nativo</span>';
              return `
                <tr class="${selected ? 'selected' : ''}">
                  <td>
                    <div class="field-selectable-cell">
                      <label class="section-field-check" title="Selecionar campo">
                        <input type="checkbox" ${selected ? 'checked' : ''} onchange='cardToggleSectionField(${fieldIdArg}, this.checked)' />
                      </label>
                      <div class="field-name-cell">
                      <span class="field-label-txt">${escHtml(finfo.label)}</span>
                      <span class="field-id-mono">${escHtml(el.name)}</span>
                      </div>
                    </div>
                  </td>
                  <td><span class="field-type-badge">${escHtml(typeL)}</span></td>
                  <td>${optsHtml}</td>
                  <td>
                    <label class="fixed-toggle ${fixed ? 'active' : ''}">
                      <input type="checkbox" ${fixed ? 'checked' : ''} onchange="cardToggleFixed(${si},${fi})" />
                      ${fixed ? 'Visível no Card' : 'Oculto na Seção'}
                    </label>
                  </td>
                  <td>
                    <div style="display:flex;gap:3px;">
                      <button class="move-btn" onclick="cardMoveFieldUp(${si},${fi})" ${fi===0?'disabled':''} title="Subir" aria-label="Subir campo">${cardIcon('up')}</button>
                      <button class="move-btn" onclick="cardMoveFieldDown(${si},${fi})" ${fi===count-1?'disabled':''} title="Descer" aria-label="Descer campo">${cardIcon('down')}</button>
                    </div>
                  </td>
                  <td>
                    <div class="card-field-actions">
                      ${editBtn}
                      <button class="card-action-btn" onclick="cardHideField(${si},${fi})">Ocultar</button>
                    </div>
                  </td>
                </tr>
              `;
            }).join('')}
          </tbody>
        </table>`;

    return `
      <div class="card-section-block ${collapsed ? 'collapsed' : ''}">
        <div class="card-section-header">
          <button class="icon-btn section-collapse-btn" onclick="cardToggleSectionCollapse(${si})" title="${collapsed ? 'Expandir secao' : 'Minimizar secao'}" aria-label="${collapsed ? 'Expandir secao' : 'Minimizar secao'}" aria-expanded="${collapsed ? 'false' : 'true'}">${cardIcon(collapsed ? 'expand' : 'collapse')}</button>
          <div class="card-section-title" id="card-sec-title-${si}" onclick="cardShowRenameSection(${si})" title="Clique para renomear">${escHtml(sec.title || sec.name || 'Seção')}</div>
          <span class="card-section-count">${count} campo${count !== 1 ? 's' : ''}</span>
          <label class="section-bulk-check ${count === 0 ? 'disabled' : ''}">
            <input type="checkbox" ${allSelected ? 'checked' : ''} ${count === 0 ? 'disabled' : ''} onchange="cardToggleAllSectionFields(${si}, this.checked)" />
            Selecionar se&ccedil;&atilde;o
          </label>
          <button class="card-action-btn section-hide-selected" onclick="cardHideSelectedSectionFields(${si})" ${selectedCount === 0 ? 'disabled' : ''}>Ocultar selecionados${selectedCount ? ` (${selectedCount})` : ''}</button>
          <div style="display:flex;gap:4px;">
            <button class="move-btn" onclick="cardMoveSectionUp(${si})" ${si===0?'disabled':''} title="Subir seção" aria-label="Subir seção">${cardIcon('up')}</button>
            <button class="move-btn" onclick="cardMoveSectionDown(${si})" ${si===cardConfig.length-1?'disabled':''} title="Descer seção" aria-label="Descer seção">${cardIcon('down')}</button>
            <button class="icon-btn" onclick="cardShowRenameSection(${si})" title="Renomear seção" aria-label="Renomear seção">${cardIcon('edit')}</button>
            <button class="icon-btn del" onclick="cardRemoveSection(${si})" title="Remover seção" aria-label="Remover seção">${cardIcon('trash')}</button>
          </div>
        </div>
        <div id="card-rename-form-${si}" style="display:none;padding:8px 14px;background:#f9f9fb;border-bottom:1px solid #eee;">
          <div class="card-section-rename">
            <input class="card-section-rename-input" id="card-rename-input-${si}" value="${escHtml(sec.title||sec.name||'')}" placeholder="Novo nome da seção" />
            <button class="btn sm primary" onclick="cardRenameSection(${si})">OK</button>
            <button class="btn sm" onclick="cardShowRenameSection(${si})">Cancelar</button>
          </div>
        </div>
        <div class="card-section-body" ${collapsed ? 'hidden' : ''}>
          ${fieldsHtml}
        </div>
      </div>
    `;
  }).join('');
}

function cardRenderHiddenFields() {
  const filter = (document.getElementById('card-hidden-filter').value || '').toLowerCase();
  const originFilter = document.getElementById('card-hidden-origin-filter')?.value || 'all';
  const visibleIds = new Set();
  cardConfig.forEach(sec => (sec.elements || []).forEach(el => visibleIds.add(el.name)));
  cardSelectedHiddenFields = new Set([...cardSelectedHiddenFields].filter(id => !visibleIds.has(id)));
  cardExpandedHiddenFields = new Set([...cardExpandedHiddenFields].filter(id => !visibleIds.has(id)));
  const hiddenAll = Object.entries(cardAllFields).filter(([id]) => !visibleIds.has(id));
  const hidden = hiddenAll.filter(([id, info]) => cardHiddenFieldMatchesFilters(id, info, filter, originFilter));

  const list = document.getElementById('card-hidden-list');
  if (!hiddenAll.length) {
    cardSelectedHiddenFields.clear();
    cardExpandedHiddenFields.clear();
    list.innerHTML = '<div style="font-size:12px;color:#aaa;text-align:center;padding:8px;">Nenhum campo oculto.</div>';
    return;
  }

  const sectionOptions = cardConfig.map((sec, si) =>
    `<option value="${si}">${escHtml(sec.title || sec.name || `Seção ${si+1}`)}</option>`
  ).join('');
  const selectedCount = [...cardSelectedHiddenFields].filter(id => cardAllFields[id] && !visibleIds.has(id)).length;
  const allFilteredSelected = hidden.length > 0 && hidden.every(([id]) => cardSelectedHiddenFields.has(id));
  const bulkDisabled = selectedCount === 0 || cardConfig.length === 0 ? 'disabled' : '';
  const bulkCollapsed = !!cardHiddenBulkCollapsed;
  const filteredExpandedCount = hidden.filter(([id]) => cardExpandedHiddenFields.has(id)).length;

  const actionsHtml = `
    <div class="hidden-fields-actions ${bulkCollapsed ? 'collapsed' : ''}">
      <div class="hidden-fields-actions-head">
        <button class="icon-btn section-collapse-btn hidden-fields-bulk-toggle" onclick="cardToggleHiddenBulkCollapse()" title="${bulkCollapsed ? 'Expandir opcoes em massa' : 'Minimizar opcoes em massa'}" aria-label="${bulkCollapsed ? 'Expandir opcoes em massa' : 'Minimizar opcoes em massa'}" aria-expanded="${bulkCollapsed ? 'false' : 'true'}">${cardIcon(bulkCollapsed ? 'expand' : 'collapse')}<span class="hidden-fields-bulk-toggle-label">${bulkCollapsed ? 'Mostrar opcoes' : 'Ocultar opcoes'}</span></button>
        <label class="hidden-field-check all">
          <input type="checkbox" ${allFilteredSelected ? 'checked' : ''} onchange="cardToggleAllHiddenFields(this.checked)" />
          Selecionar filtrados
        </label>
        <span class="hidden-selected-count">${selectedCount} selecionado${selectedCount !== 1 ? 's' : ''}</span>
        <span class="hidden-selected-count">${hidden.length} no filtro</span>
      </div>
      <div class="hidden-fields-actions-body" ${bulkCollapsed ? 'hidden' : ''}>
        <div class="hidden-fields-bulk-target">
          <span>Adicionar em</span>
          <select class="hidden-field-select" id="card-bulk-add-sec" ${cardConfig.length === 0 ? 'disabled' : ''}>
            ${sectionOptions}
          </select>
        </div>
        <div class="hidden-fields-bulk-buttons">
          <button class="btn sm primary" ${bulkDisabled} onclick="cardAddSelectedHiddenFields(document.getElementById('card-bulk-add-sec').value)">+ Adicionar selecionados</button>
          <button class="btn sm" ${hidden.length === 0 || filteredExpandedCount === hidden.length ? 'disabled' : ''} onclick="cardSetFilteredHiddenFieldsExpanded(true)">Expandir itens</button>
          <button class="btn sm" ${filteredExpandedCount === 0 ? 'disabled' : ''} onclick="cardSetFilteredHiddenFieldsExpanded(false)">Minimizar itens</button>
        </div>
      </div>
    </div>
  `;

  const rowsHtml = hidden.length ? hidden.map(([id, info]) => {
    const typeL = TIPO_LABEL[info.type] || info.type;
    const originL = info.origin === 'custom' ? 'Personalizado' : 'Nativo';
    const safeId = id.replace(/[^a-z0-9]/gi,'_');
    const idArg = JSON.stringify(id);
    const selectIdArg = JSON.stringify(`card-add-sec-${safeId}`);
    const checked = cardSelectedHiddenFields.has(id) ? 'checked' : '';
    const expanded = cardExpandedHiddenFields.has(id);
    return `
      <div class="hidden-field-row ${expanded ? 'expanded' : 'collapsed'}">
        <div class="hidden-field-summary">
          <button class="icon-btn section-collapse-btn hidden-field-expand-btn" onclick='cardToggleHiddenFieldDetails(${idArg})' title="${expanded ? 'Minimizar opcoes do campo' : 'Expandir opcoes do campo'}" aria-label="${expanded ? 'Minimizar opcoes do campo' : 'Expandir opcoes do campo'}" aria-expanded="${expanded ? 'true' : 'false'}">${cardIcon(expanded ? 'collapse' : 'expand')}</button>
          <label class="hidden-field-check" title="Selecionar campo">
            <input type="checkbox" id="card-hidden-check-${safeId}" ${checked} onchange='cardToggleHiddenField(${idArg}, this.checked)' />
          </label>
          <div class="hidden-field-info">
            <div class="hidden-field-label">${escHtml(info.label)}</div>
            <div class="hidden-field-id">${escHtml(id)}</div>
          </div>
          <span class="hidden-origin-badge ${info.origin === 'custom' ? 'custom' : 'native'}">${originL}</span>
          <span class="field-type-badge">${escHtml(typeL)}</span>
        </div>
        <div class="hidden-field-details" ${expanded ? '' : 'hidden'}>
          <select class="hidden-field-select" id="card-add-sec-${safeId}">
            ${sectionOptions}
          </select>
          <button class="btn sm primary" onclick='cardAddHiddenField(${idArg}, document.getElementById(${selectIdArg}).value)'>+ Adicionar</button>
        </div>
      </div>
    `;
  }).join('') : '<div style="font-size:12px;color:#aaa;text-align:center;padding:8px;">Nenhum campo encontrado no filtro.</div>';

  list.innerHTML = actionsHtml + rowsHtml;
}

function cardPruneSelectedSectionFields() {
  const visibleIds = new Set();
  cardConfig.forEach(sec => (sec.elements || []).forEach(el => visibleIds.add(el.name)));
  cardSelectedSectionFields = new Set([...cardSelectedSectionFields].filter(id => visibleIds.has(id)));
}

function cardGetSectionSelectedCount(sectionIdx) {
  const sec = cardConfig[sectionIdx];
  if (!sec || !Array.isArray(sec.elements)) return 0;
  return sec.elements.filter(el => cardSelectedSectionFields.has(el.name)).length;
}

function cardToggleSectionField(fieldId, checked) {
  if (checked) cardSelectedSectionFields.add(fieldId);
  else cardSelectedSectionFields.delete(fieldId);
  cardRenderSections();
}

function cardToggleAllSectionFields(sectionIdx, checked) {
  const sec = cardConfig[sectionIdx];
  if (!sec || !Array.isArray(sec.elements)) return;
  sec.elements.forEach(el => {
    if (checked) cardSelectedSectionFields.add(el.name);
    else cardSelectedSectionFields.delete(el.name);
  });
  cardRenderSections();
}

function cardHideSelectedSectionFields(sectionIdx) {
  const sec = cardConfig[sectionIdx];
  if (!sec || !Array.isArray(sec.elements)) return;
  const before = sec.elements.length;
  sec.elements = sec.elements.filter(el => !cardSelectedSectionFields.has(el.name));
  const hiddenCount = before - sec.elements.length;
  if (!hiddenCount) {
    toast('Selecione pelo menos um campo da se\u00e7\u00e3o.', 'wn');
    return;
  }
  cardSelectedSectionFields.clear();
  cardMarkUnsaved();
  cardRender();
  toast(`${hiddenCount} campo${hiddenCount !== 1 ? 's' : ''} ocultado${hiddenCount !== 1 ? 's' : ''}.`, 'ok');
}

function cardHiddenFieldMatchesFilters(id, info, filter, originFilter) {
  const origin = info?.origin === 'custom' ? 'custom' : 'native';
  if (originFilter !== 'all' && origin !== originFilter) return false;
  if (!filter) return true;
  return id.toLowerCase().includes(filter) || (info.label || '').toLowerCase().includes(filter);
}

function cardToggleHiddenField(fieldId, checked) {
  if (checked) cardSelectedHiddenFields.add(fieldId);
  else cardSelectedHiddenFields.delete(fieldId);
  cardRenderHiddenFields();
}

function cardToggleAllHiddenFields(checked) {
  const filter = (document.getElementById('card-hidden-filter').value || '').toLowerCase();
  const originFilter = document.getElementById('card-hidden-origin-filter')?.value || 'all';
  const visibleIds = new Set();
  cardConfig.forEach(sec => (sec.elements || []).forEach(el => visibleIds.add(el.name)));
  Object.entries(cardAllFields)
    .filter(([id]) => !visibleIds.has(id))
    .filter(([id, info]) => cardHiddenFieldMatchesFilters(id, info, filter, originFilter))
    .forEach(([id]) => {
      if (checked) cardSelectedHiddenFields.add(id);
      else cardSelectedHiddenFields.delete(id);
    });
  cardRenderHiddenFields();
}

function cardToggleHiddenFieldDetails(fieldId) {
  if (cardExpandedHiddenFields.has(fieldId)) cardExpandedHiddenFields.delete(fieldId);
  else cardExpandedHiddenFields.add(fieldId);
  cardRenderHiddenFields();
}

function cardSetFilteredHiddenFieldsExpanded(expanded) {
  const filter = (document.getElementById('card-hidden-filter').value || '').toLowerCase();
  const originFilter = document.getElementById('card-hidden-origin-filter')?.value || 'all';
  const visibleIds = new Set();
  cardConfig.forEach(sec => (sec.elements || []).forEach(el => visibleIds.add(el.name)));
  Object.entries(cardAllFields)
    .filter(([id]) => !visibleIds.has(id))
    .filter(([id, info]) => cardHiddenFieldMatchesFilters(id, info, filter, originFilter))
    .forEach(([id]) => {
      if (expanded) cardExpandedHiddenFields.add(id);
      else cardExpandedHiddenFields.delete(id);
    });
  cardRenderHiddenFields();
}

function cardToggleHiddenBulkCollapse() {
  cardHiddenBulkCollapsed = !cardHiddenBulkCollapsed;
  cardRenderHiddenFields();
}

function cardToggleHiddenPanelCollapse() {
  cardHiddenPanelCollapsed = !cardHiddenPanelCollapsed;
  const body = document.getElementById('card-hidden-panel-body');
  const btn = document.getElementById('card-hidden-panel-toggle');
  if (body) body.hidden = cardHiddenPanelCollapsed;
  if (btn) {
    btn.innerHTML = cardIcon(cardHiddenPanelCollapsed ? 'expand' : 'collapse');
    btn.title = cardHiddenPanelCollapsed ? 'Expandir campos ocultos' : 'Minimizar campos ocultos';
    btn.setAttribute('aria-label', btn.title);
    btn.setAttribute('aria-expanded', cardHiddenPanelCollapsed ? 'false' : 'true');
  }
}

function cardMarkUnsaved() {
  document.getElementById('card-unsaved-badge').style.display = 'inline-flex';
}

function cardAddSection() {
  const input = document.getElementById('card-new-section-input');
  const title = input.value.trim();
  if (!title) { toast('Informe o nome da seção.', 'wn'); return; }
  cardConfig.push({ name: title.toLowerCase().replace(/\s+/g,'_'), title, type: 'section', elements: [] });
  input.value = '';
  cardMarkUnsaved();
  cardRender();
}

function cardRemoveSection(idx) {
  const sec = cardConfig[idx];
  if (sec && sec.elements && sec.elements.length > 0) {
    toast('Remova todos os campos da seção antes de excluí-la.', 'wn'); return;
  }
  cardConfig.splice(idx, 1);
  cardMarkUnsaved();
  cardRender();
}

function cardShowRenameSection(idx) {
  const form = document.getElementById(`card-rename-form-${idx}`);
  if (cardConfig[idx] && cardConfig[idx]._collapsed) {
    cardConfig[idx]._collapsed = false;
    cardRenderSections();
    document.getElementById(`card-rename-form-${idx}`).style.display = 'block';
    document.getElementById(`card-rename-input-${idx}`).focus();
    return;
  }
  form.style.display = form.style.display === 'none' ? 'block' : 'none';
  if (form.style.display === 'block') document.getElementById(`card-rename-input-${idx}`).focus();
}

function cardToggleSectionCollapse(idx) {
  if (!cardConfig[idx]) return;
  cardConfig[idx]._collapsed = !cardConfig[idx]._collapsed;
  cardRenderSections();
}

function cardRenameSection(idx) {
  const input = document.getElementById(`card-rename-input-${idx}`);
  const title = input.value.trim();
  if (!title) { toast('Informe um nome.', 'wn'); return; }
  cardConfig[idx].title = title;
  cardConfig[idx].name  = title.toLowerCase().replace(/\s+/g,'_');
  cardMarkUnsaved();
  cardRender();
}

function cardMoveSectionUp(idx) {
  if (idx === 0) return;
  [cardConfig[idx-1], cardConfig[idx]] = [cardConfig[idx], cardConfig[idx-1]];
  cardMarkUnsaved();
  cardRender();
}

function cardMoveSectionDown(idx) {
  if (idx >= cardConfig.length - 1) return;
  [cardConfig[idx], cardConfig[idx+1]] = [cardConfig[idx+1], cardConfig[idx]];
  cardMarkUnsaved();
  cardRender();
}

function cardMoveFieldUp(si, fi) {
  const els = cardConfig[si].elements;
  if (fi === 0) return;
  [els[fi-1], els[fi]] = [els[fi], els[fi-1]];
  cardMarkUnsaved();
  cardRender();
}

function cardMoveFieldDown(si, fi) {
  const els = cardConfig[si].elements;
  if (fi >= els.length - 1) return;
  [els[fi], els[fi+1]] = [els[fi+1], els[fi]];
  cardMarkUnsaved();
  cardRender();
}

function cardToggleFixed(si, fi) {
  const el = cardConfig[si].elements[fi];
  el.optionFlags = el.optionFlags === 1 ? 0 : 1;
  cardMarkUnsaved();
  cardRender();
}

function cardHideField(si, fi) {
  const removed = cardConfig[si].elements.splice(fi, 1)[0];
  if (removed) cardSelectedSectionFields.delete(removed.name);
  cardMarkUnsaved();
  cardRender();
}

function cardAddHiddenField(fieldId, targetSectionIdx) {
  const idx = parseInt(targetSectionIdx);
  if (isNaN(idx) || !cardConfig[idx]) { toast('Seção inválida.', 'wn'); return; }
  const already = cardConfig.some(s => s.elements.some(e => e.name === fieldId));
  if (already) { toast('Campo já está em uma seção.', 'wn'); return; }
  cardConfig[idx].elements.push({ name: fieldId, optionFlags: 0 });
  cardSelectedHiddenFields.delete(fieldId);
  cardMarkUnsaved();
  cardRender();
}

function cardAddSelectedHiddenFields(targetSectionIdx) {
  const idx = parseInt(targetSectionIdx);
  if (isNaN(idx) || !cardConfig[idx]) { toast('Seção inválida.', 'wn'); return; }
  const selected = [...cardSelectedHiddenFields].filter(fieldId =>
    cardAllFields[fieldId] && !cardConfig.some(s => s.elements.some(e => e.name === fieldId))
  );
  if (!selected.length) { toast('Selecione pelo menos um campo oculto.', 'wn'); return; }

  selected.forEach(fieldId => {
    cardConfig[idx].elements.push({ name: fieldId, optionFlags: 0 });
    cardSelectedHiddenFields.delete(fieldId);
  });
  cardMarkUnsaved();
  cardRender();
  toast(`${selected.length} campo${selected.length !== 1 ? 's' : ''} adicionado${selected.length !== 1 ? 's' : ''}.`, 'ok');
}

function cardFieldLabel(fieldId) {
  return (cardAllFields[fieldId] && cardAllFields[fieldId].label) ? cardAllFields[fieldId].label : fieldId;
}

function cardOpenEdit(si, fi) {
  const el = cardConfig[si] && cardConfig[si].elements && cardConfig[si].elements[fi];
  if (!el) return;
  const f = cardAllFields[el.name];
  if (!f || f.origin !== 'custom' || !f.rawId) {
    toast('Apenas campos personalizados podem ser editados.', 'wn');
    return;
  }

  cardEditFieldRef = { si, fi, fieldId: el.name };
  document.getElementById('card-edit-id-info').textContent = `ID do campo: ${el.name}  |  Tipo: ${TIPO_LABEL[f.type] || f.type}`;
  document.getElementById('card-edit-label').value = f.label || el.name;
  document.getElementById('card-edit-mandatory').checked = !!f.mandatory;
  document.getElementById('card-edit-multiple').checked = !!f.multiple;

  const rowsSection = document.getElementById('card-edit-string-section');
  if (f.type === 'string') {
    rowsSection.style.display = 'block';
    document.getElementById('card-edit-string-rows').value = parseInt(camposGetSettingValue(f.settings, 'ROWS', 1), 10) || 1;
  } else {
    rowsSection.style.display = 'none';
  }

  const doubleSection = document.getElementById('card-edit-double-section');
  if (doubleSection) {
    doubleSection.style.display = f.type === 'double' ? 'block' : 'none';
    document.getElementById('card-edit-double-precision').value = String(parseInt(camposGetSettingValue(f.settings, 'PRECISION', 2), 10) || 0);
  }

  const booleanSection = document.getElementById('card-edit-boolean-section');
  if (booleanSection) {
    const def = camposGetSettingValue(f.settings, 'DEFAULT_VALUE', '');
    booleanSection.style.display = f.type === 'boolean' ? 'block' : 'none';
    document.getElementById('card-edit-boolean-default').value = def === 'Y' || def === true ? '1' : def === 'N' || def === false ? '2' : '0';
  }

  const enumSection = document.getElementById('card-edit-enum-section');
  if (f.type === 'enumeration') {
    enumSection.style.display = 'block';
    document.getElementById('card-edit-enum-options').value = (f.options || []).join('\n');
  } else {
    enumSection.style.display = 'none';
  }
  const helpInput = document.getElementById('card-edit-help');
  if (helpInput) helpInput.value = f.helpMessage || '';

  document.getElementById('card-edit-overlay').classList.add('open');
  document.getElementById('card-edit-label').focus();
}

function cardCloseEdit() {
  document.getElementById('card-edit-overlay').classList.remove('open');
  cardEditFieldRef = null;
}

async function cardSaveFieldEditLegacy() {
  if (!cardEditFieldRef) return;
  const f = cardAllFields[cardEditFieldRef.fieldId];
  if (!f) return;

  const label = document.getElementById('card-edit-label').value.trim();
  if (!label) { toast('Informe o nome do campo.', 'wn'); return; }

  const mandatory = document.getElementById('card-edit-mandatory').checked ? 'Y' : 'N';
  const multiple  = document.getElementById('card-edit-multiple').checked  ? 'Y' : 'N';
  const fields = {
    EDIT_FORM_LABEL: label,
    LIST_COLUMN_LABEL: label,
    MANDATORY: mandatory,
    MULTIPLE: multiple,
  };

  if (f.type === 'enumeration') {
    const opts = document.getElementById('card-edit-enum-options').value
      .split('\n').map(v => v.trim()).filter(Boolean);
    fields.LIST = opts.map((v, i) => ({ SORT: (i + 1) * 10, VALUE: v, DEF: 'N' }));
  }

  if (f.type === 'string') {
    const rows = Math.max(1, parseInt(document.getElementById('card-edit-string-rows').value, 10) || 1);
    fields.SETTINGS = { ...(f.settings || {}), ROWS: rows };
  }

  const methods = ENTITY_METHODS[cardEntityTypeId];
  if (!methods || !methods.update) { toast('Método de atualização não disponível.', 'wn'); return; }

  const saveBtn = document.getElementById('card-edit-save-btn');
  saveBtn.textContent = 'Salvando';
  saveBtn.disabled = true;

  try {
    const data = await call(methods.update, { id: f.rawId, fields });
    if (data.error) throw new Error(data.error_description || data.error);

    cardAllFields[cardEditFieldRef.fieldId] = {
      ...f,
      label,
      mandatory: mandatory === 'Y',
      multiple: multiple === 'Y',
      options: f.type === 'enumeration'
        ? document.getElementById('card-edit-enum-options').value.split('\n').map(v => v.trim()).filter(Boolean)
        : (f.options || []),
      settings: f.type === 'string' ? fields.SETTINGS : (f.settings || {}),
    };

    cardCloseEdit();
    cardRender();
    toast('Campo atualizado com sucesso!', 'ok');
  } catch (e) {
    toast('Erro ao salvar: ' + e.message, 'er');
  } finally {
    saveBtn.textContent = 'Salvar';
    saveBtn.disabled = false;
  }
}

async function cardDeleteFieldLegacy(si, fi) {
  const el = cardConfig[si] && cardConfig[si].elements && cardConfig[si].elements[fi];
  if (!el) return;
  const f = cardAllFields[el.name];
  if (!f || f.origin !== 'custom' || !f.rawId) {
    toast('Apenas campos personalizados podem ser excludos.', 'wn');
    return;
  }

  if (!confirm(`Excluir o campo "${f.label}" (${el.name})?\n\nEsta ação não pode ser desfeita.`)) return;

  const methods = ENTITY_METHODS[cardEntityTypeId];
  if (!methods || !methods.delete) { toast('Método de exclusão não disponível.', 'wn'); return; }

  try {
    const data = await call(methods.delete, { id: f.rawId });
    if (data.error) throw new Error(data.error_description || data.error);

    cardConfig.forEach(sec => {
      sec.elements = (sec.elements || []).filter(item => item.name !== el.name);
    });
    delete cardAllFields[el.name];
    cardSelectedHiddenFields.delete(el.name);
    cardSelectedSectionFields.delete(el.name);
    cardMarkUnsaved();
    cardRender();
    toast('Campo excludo com sucesso.', 'ok');
  } catch (e) {
    toast('Erro ao excluir: ' + e.message, 'er');
  }
}

function cardBuildLayoutData() {
  return cardConfig.map(sec => ({
    name: String(sec.name || sec.title || ''),
    title: String(sec.title || sec.name || ''),
    type: sec.type || 'section',
    elements: (sec.elements || []).map(el => ({
      name: String(el.name || ''),
      optionFlags: Number(el.optionFlags || 0),
      ...(el.options ? { options: el.options } : {}),
    })).filter(el => el.name),
  }));
}

function cardSlug(value) {
  return String(value || 'layout')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-z0-9]+/gi, '-')
    .replace(/^-+|-+$/g, '')
    .toLowerCase() || 'layout';
}

function cardWorkbookAvailable() {
  if (window.XLSX) return true;
  toast('Biblioteca XLSX não carregada. Recarregue a página e tente novamente.', 'er');
  return false;
}

function cardFieldInfo(fieldId) {
  return cardAllFields[fieldId] || {};
}

function cardBoolLabel(value) {
  return value ? 'Sim' : 'Não';
}

function cardTypeLabel(type) {
  return TIPO_LABEL[type] || type || '';
}

function cardOriginLabel(origin) {
  if (origin === 'native') return 'Nativo';
  if (origin === 'custom') return 'Personalizado';
  return origin || '';
}

function cardSectionTypeLabel(type) {
  if (!type || type === 'section') return 'Seção';
  return type;
}

function cardOptionFlagsLabel(optionFlags) {
  return Number(optionFlags || 0) === 1 ? 'Mostrar sempre no card' : 'Oculto na seção';
}

function cardReadBool(value) {
  const normalized = String(value || '').trim().toLowerCase();
  return ['1', 'sim', 's', 'yes', 'y', 'true', 'verdadeiro'].includes(normalized);
}

function cardDownloadImportModel() {
  if (!cardWorkbookAvailable()) return;
  const wb = XLSX.utils.book_new();
  const contextRows = [
    ['chave', 'valor'],
    ['observacao', 'Preencha as abas Secoes e Campos. A importacao usa id_secao, ordem_secao, id_campo, ordem_campo e exibicao_no_card/mostrar_sempre.'],
  ];
  const sectionRows = [
    ['ordem_secao', 'id_secao', 'titulo_secao', 'tipo_secao', 'quantidade_campos'],
    [1, 'dados_principais', 'Dados principais', 'Seção', 2],
  ];
  const fieldRows = [
    ['ordem_secao', 'id_secao', 'titulo_secao', 'ordem_campo', 'id_campo', 'label_campo', 'tipo_campo', 'origem_campo', 'obrigatorio', 'multiplo', 'mostrar_sempre', 'exibicao_no_card', 'opcoes', 'dica'],
    [1, 'dados_principais', 'Dados principais', 1, 'TITLE', 'Nome', 'Texto', 'Nativo', 'Nao', 'Nao', 'Sim', 'Mostrar sempre no card', '', ''],
    [1, 'dados_principais', 'Dados principais', 2, 'UF_CRM_EXEMPLO', 'Campo personalizado exemplo', 'Texto', 'Personalizado', 'Nao', 'Nao', 'Nao', 'Oculto na seção', '', 'Substitua pelo ID real do campo'],
  ];
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(contextRows), 'Contexto');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sectionRows), 'Secoes');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fieldRows), 'Campos');
  XLSX.writeFile(wb, 'modelo-importacao-layout-card.xlsx');
  toast('Modelo de importação baixado.', 'ok');
}

function cardExportLayout() {
  if (!cardLoaded) { toast('Carregue uma configuração antes de exportar.', 'wn'); return; }
  if (!cardWorkbookAvailable()) return;
  const ctx = cardGetContext(false);
  const layout = cardBuildLayoutData();
  const wb = XLSX.utils.book_new();
  const contextRows = [
    ['chave', 'valor'],
    ['versao', 1],
    ['exportado_em', new Date().toISOString()],
    ['origem', 'LibeSales Painel Bitrix24'],
    ['modo', ctx.mode],
    ['entityTypeId', ctx.entityTypeId || ''],
    ['entityId', ctx.entityId || ''],
    ['pipelineId', cardPipelineId || '0'],
    ['leadType', ctx.mode === 'crm' && cardEntityTypeId === 1 ? cardLeadType : ''],
    ['spaId', ctx.mode === 'spa' && ctx.spa ? ctx.spa.id : ''],
    ['spaTitle', ctx.mode === 'spa' && ctx.spa ? ctx.spa.title : ''],
  ];
  const sectionRows = [[
    'ordem_secao',
    'id_secao',
    'titulo_secao',
    'tipo_secao',
    'quantidade_campos',
  ]];
  const fieldRows = [[
    'ordem_secao',
    'id_secao',
    'titulo_secao',
    'ordem_campo',
    'id_campo',
    'label_campo',
    'tipo_campo',
    'origem_campo',
    'obrigatorio',
    'multiplo',
    'mostrar_sempre',
    'exibicao_no_card',
    'opcoes',
    'dica',
  ]];

  layout.forEach((sec, si) => {
    const elements = sec.elements || [];
    sectionRows.push([si + 1, sec.name, sec.title, cardSectionTypeLabel(sec.type), elements.length]);
    elements.forEach((el, fi) => {
      const info = cardFieldInfo(el.name);
      fieldRows.push([
        si + 1,
        sec.name,
        sec.title,
        fi + 1,
        el.name,
        info.label || el.name,
        cardTypeLabel(info.type),
        cardOriginLabel(info.origin),
        cardBoolLabel(!!info.mandatory),
        cardBoolLabel(!!info.multiple),
        cardBoolLabel(Number(el.optionFlags || 0) === 1),
        cardOptionFlagsLabel(el.optionFlags),
        (info.options || []).map(opt => opt.VALUE || opt.value || opt.label || opt).join('\n'),
        info.helpMessage || '',
      ]);
    });
  });

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(contextRows), 'Contexto');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(sectionRows), 'Secoes');
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(fieldRows), 'Campos');

  const contextName = ctx.mode === 'spa' && ctx.spa ? `spa-${ctx.spa.id}` : `crm-${cardEntityTypeId}`;
  XLSX.writeFile(wb, `layout-card-${cardSlug(contextName)}-${new Date().toISOString().slice(0, 10)}.xlsx`);
  toast('Layout exportado.', 'ok');
}

function cardOpenImportLayout() {
  if (!cardLoaded) { toast('Carregue uma configuração antes de importar.', 'wn'); return; }
  const input = document.getElementById('card-layout-import-file');
  if (!input) return;
  input.value = '';
  input.click();
}

function cardGetWorkbookRows(wb, sheetName) {
  const sheet = wb && wb.Sheets && wb.Sheets[sheetName];
  return sheet ? XLSX.utils.sheet_to_json(sheet, { defval: '' }) : [];
}

function cardPick(row, names) {
  for (const name of names) {
    if (row[name] !== undefined && row[name] !== null && String(row[name]).trim() !== '') return row[name];
  }
  return '';
}

function cardSectionsFromWorkbook(wb) {
  const sectionRows = cardGetWorkbookRows(wb, 'Secoes');
  const fieldRows = cardGetWorkbookRows(wb, 'Campos');
  if (!fieldRows.length && !sectionRows.length) throw new Error('A planilha precisa ter as abas Secoes e/ou Campos.');

  const sectionsByKey = new Map();
  sectionRows.forEach((row, idx) => {
    const order = parseInt(cardPick(row, ['ordem_secao', 'ordem']), 10) || idx + 1;
    const name = String(cardPick(row, ['id_secao', 'name', 'secao']) || `secao_${order}`).trim();
    const title = String(cardPick(row, ['titulo_secao', 'title', 'titulo']) || name).trim();
    const type = String(cardPick(row, ['tipo_secao', 'type']) || 'section').trim().toLowerCase() === 'seção' ? 'section' : String(cardPick(row, ['tipo_secao', 'type']) || 'section').trim();
    sectionsByKey.set(name, { order, name, title, type, elements: [] });
  });

  fieldRows.forEach((row, idx) => {
    const sectionOrder = parseInt(cardPick(row, ['ordem_secao', 'ordem']), 10) || 9999;
    const sectionName = String(cardPick(row, ['id_secao', 'name', 'secao']) || `secao_${sectionOrder}`).trim();
    const sectionTitle = String(cardPick(row, ['titulo_secao', 'title', 'titulo']) || sectionName).trim();
    const fieldId = String(cardPick(row, ['id_campo', 'field_id', 'campo', 'name'])).trim();
    if (!fieldId) return;
    if (!sectionsByKey.has(sectionName)) {
      sectionsByKey.set(sectionName, {
        order: sectionOrder,
        name: sectionName,
        title: sectionTitle,
        type: 'section',
        elements: [],
      });
    }
    const section = sectionsByKey.get(sectionName);
    const optionFlagsValue = cardPick(row, ['optionFlags', 'option_flags']);
    const displayLabel = String(cardPick(row, ['exibicao_no_card', 'exibição_no_card'])).trim().toLowerCase();
    const showAlways = cardPick(row, ['mostrar_sempre', 'visivel_no_card']);
    section.elements.push({
      order: parseInt(cardPick(row, ['ordem_campo', 'ordem_field']), 10) || idx + 1,
      name: fieldId,
      optionFlags: optionFlagsValue !== ''
        ? Number(optionFlagsValue || 0)
        : (displayLabel.includes('mostrar') || cardReadBool(showAlways) ? 1 : 0),
    });
  });

  return Array.from(sectionsByKey.values())
    .sort((a, b) => a.order - b.order)
    .map(sec => ({
      name: sec.name,
      title: sec.title,
      type: sec.type || 'section',
      elements: sec.elements
        .sort((a, b) => a.order - b.order)
        .map(el => ({ name: el.name, optionFlags: Number(el.optionFlags || 0) })),
    }));
}

function cardNormalizeImportedLayout(sections) {
  if (!Array.isArray(sections)) throw new Error('Arquivo sem lista de seções.');
  return sections.map((sec, idx) => {
    const title = String(sec.title || sec.name || `Seção ${idx + 1}`).trim();
    const name = String(sec.name || title.toLowerCase().replace(/\s+/g, '_')).trim();
    const elements = Array.isArray(sec.elements) ? sec.elements : [];
    return {
      name,
      title,
      type: sec.type || 'section',
      elements: elements.map(el => ({
        name: String(el.name || '').trim(),
        optionFlags: Number(el.optionFlags || 0),
        ...(el.options ? { options: el.options } : {}),
      })).filter(el => el.name),
    };
  });
}

function cardApplyImportedLayout(sections) {
  const imported = cardNormalizeImportedLayout(sections);
  const knownIds = new Set(Object.keys(cardAllFields || {}));
  const missing = imported.reduce((total, sec) => (
    total + (sec.elements || []).filter(el => !knownIds.has(el.name)).length
  ), 0);
  const msg = missing
    ? `Importar layout? ${missing} campo(s) não existem no contexto atual e podem aparecer apenas pelo ID até serem criados.`
    : 'Importar layout? Isso substitui as seções carregadas na tela.';
  if (!confirm(msg)) return;
  cardConfig = imported;
  cardSelectedHiddenFields.clear();
  cardExpandedHiddenFields.clear();
  cardSelectedSectionFields.clear();
  cardMarkUnsaved();
  cardRender();
  toast('Layout importado. Revise e salve para aplicar no Bitrix24.', 'ok', 4000);
}

function cardImportLayoutFile(input) {
  const file = input && input.files && input.files[0];
  if (!file) return;
  if (!cardWorkbookAvailable()) {
    input.value = '';
    return;
  }
  file.arrayBuffer()
    .then(buffer => {
      const wb = XLSX.read(buffer, { type: 'array' });
      cardApplyImportedLayout(cardSectionsFromWorkbook(wb));
    })
    .catch(e => {
      toast('Erro ao importar layout: ' + e.message, 'er');
    })
    .finally(() => {
      input.value = '';
    });
}

async function cardSave() {
  const data = cardBuildLayoutData();
  // scope 'P' (pessoal) funciona para qualquer usuário sem precisar de admin
  try {
    const { methodPrefix, base } = cardBuildCallParams('P');
    const method = methodPrefix + '.set';
    const r = await call(method, Object.assign({}, base, { data }));
    if (r.error) throw new Error(r.error_description || r.error);
    toast('Configuração salva com sucesso!', 'ok');
    document.getElementById('card-unsaved-badge').style.display = 'none';
  } catch(e) {
    toast('Erro ao salvar: ' + e.message, 'er');
  }
}

async function cardForceAll() {
  if (!confirm('Isso vai sobrescrever a configuração de card de TODOS os usuários. Confirma?')) return;
  const data = cardBuildLayoutData();
  try {
    // Passo 1: salva a config atual no escopção comum (C) - requer admin
    const { methodPrefix, base } = cardBuildCallParams('C');
    const setMethod   = methodPrefix + '.set';
    const forceMethod = methodPrefix + '.forceCommonScopeForAll';
    const rSet = await call(setMethod, Object.assign({}, base, { data }));
    if (rSet.error) throw new Error(rSet.error_description || rSet.error);
    // Passo 2: fora o escopção comum para todos os usuários
    const forceParams = { entityTypeId: base.entityTypeId };
    if (base.extras) forceParams.extras = base.extras;
    const rForce = await call(forceMethod, forceParams);
    if (rForce.error) throw new Error(rForce.error_description || rForce.error);
    toast('Configuração aplicada para todos os usuários!', 'ok');
    document.getElementById('card-unsaved-badge').style.display = 'none';
  } catch(e) {
    toast('Erro ao forar para todos: ' + e.message, 'er');
  }
}

function camposReadEditSettingsFromCard(type, currentSettings = {}) {
  const settings = { ...(currentSettings || {}) };
  if (type === 'string') {
    settings.ROWS = Math.max(1, parseInt(document.getElementById('card-edit-string-rows').value, 10) || 1);
    if (settings.DEFAULT_VALUE === undefined) settings.DEFAULT_VALUE = '';
  }
  if (type === 'double') {
    settings.PRECISION = Math.max(0, parseInt(document.getElementById('card-edit-double-precision').value, 10) || 0);
  }
  if (type === 'boolean') {
    const def = document.getElementById('card-edit-boolean-default').value;
    settings.DEFAULT_VALUE = def === '0' ? '' : def === '1' ? 'Y' : 'N';
  }
  return settings;
}

async function cardSaveFieldEdit() {
  if (!cardEditFieldRef) return;
  const f = cardAllFields[cardEditFieldRef.fieldId];
  if (!f) return;

  const label = document.getElementById('card-edit-label').value.trim();
  if (!label) { toast('Informe o nome do campo.', 'wn'); return; }

  const mandatory = document.getElementById('card-edit-mandatory').checked ? 'Y' : 'N';
  const multiple = document.getElementById('card-edit-multiple').checked ? 'Y' : 'N';
  const options = f.type === 'enumeration'
    ? document.getElementById('card-edit-enum-options').value.split('\n').map(v => v.trim()).filter(Boolean)
    : (f.options || []);
  if (f.type === 'enumeration' && !options.length) {
    toast('Adicione ao menos uma opção à lista.', 'wn');
    return;
  }

  const settings = camposReadEditSettingsFromCard(f.type, f.settings);
  const helpMessage = (document.getElementById('card-edit-help') || { value: '' }).value.trim();
  const adapter = window.EntityContext ? window.EntityContext.getCurrentAdapter() : null;
  if (!adapter || !adapter.updateField || !window.EntityContext.buildUserFieldUpdatePayload) {
    toast('Método de atualização não disponível.', 'wn');
    return;
  }

  const saveBtn = document.getElementById('card-edit-save-btn');
  saveBtn.textContent = 'Salvando...';
  saveBtn.disabled = true;

  try {
    const payload = window.EntityContext.buildUserFieldUpdatePayload({
      label,
      mandatory,
      multiple,
      type: f.type,
      settings,
      options,
      helpMessage,
    }, cardMode());
    const data = cardMode() === 'spa'
      ? await adapter.updateField(f.rawId, payload)
      : await adapter.updateField(f.rawId, payload, cardEntityTypeId);
    if (data.error) throw new Error(data.error_description || data.error);

    cardAllFields[cardEditFieldRef.fieldId] = {
      ...f,
      label,
      mandatory: mandatory === 'Y',
      multiple: multiple === 'Y',
      options,
      settings,
      helpMessage,
    };

    cardCloseEdit();
    cardRender();
    toast('Campo atualizado com sucesso!', 'ok');
  } catch (e) {
    toast('Erro ao salvar: ' + e.message, 'er');
  } finally {
    saveBtn.textContent = 'Salvar';
    saveBtn.disabled = false;
  }
}

async function cardDeleteField(si, fi) {
  const el = cardConfig[si] && cardConfig[si].elements && cardConfig[si].elements[fi];
  if (!el) return;
  const f = cardAllFields[el.name];
  if (!f || f.origin !== 'custom' || !f.rawId) {
    toast('Apenas campos personalizados podem ser excludos.', 'wn');
    return;
  }

  if (!confirm(`Excluir o campo "${f.label}" (${el.name})?\n\nEsta ação não pode ser desfeita.`)) return;
  const adapter = window.EntityContext ? window.EntityContext.getCurrentAdapter() : null;
  if (!adapter || !adapter.deleteField) { toast('Método de exclusão não disponível.', 'wn'); return; }

  try {
    const data = cardMode() === 'spa'
      ? await adapter.deleteField(f.rawId)
      : await adapter.deleteField(f.rawId, cardEntityTypeId);
    if (data.error) throw new Error(data.error_description || data.error);

    cardConfig.forEach(sec => {
      sec.elements = (sec.elements || []).filter(item => item.name !== el.name);
    });
    delete cardAllFields[el.name];
    cardSelectedHiddenFields.delete(el.name);
    cardSelectedSectionFields.delete(el.name);
    cardMarkUnsaved();
    cardRender();
    toast('Campo excludo com sucesso.', 'ok');
  } catch (e) {
    toast('Erro ao excluir: ' + e.message, 'er');
  }
}

window.cardSyncContextUI = cardSyncContextUI;
