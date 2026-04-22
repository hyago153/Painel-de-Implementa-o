/* 
   MDULO: CONSULTA DE CAMPOS (Tpico 3)
    */

let camposEntityTypeId = 4;
let camposPipId        = '0';
let camposAllFields    = [];
let camposTab          = 'all';
let camposSortKey      = 'label';
let camposSortAsc      = true;
let camposEditFieldIndex = -1;
let camposLastContextKey = '';
let camposCollapsedFieldIds = new Set();

const CAMPOS_API = {
  4: { fields: 'crm.company.fields', uf: 'crm.company.userfield.list', entityId: 'CRM_COMPANY', configMethod: 'crm.item.details.configuration.get', configParams: { entityTypeId: 4, scope: 'C' } },
  3: { fields: 'crm.contact.fields', uf: 'crm.contact.userfield.list', entityId: 'CRM_CONTACT', configMethod: 'crm.item.details.configuration.get', configParams: { entityTypeId: 3, scope: 'C' } },
  2: { fields: 'crm.deal.fields',    uf: 'crm.deal.userfield.list',    entityId: 'CRM_DEAL',    configMethod: 'crm.item.details.configuration.get', configParams: { entityTypeId: 2, scope: 'C' } },
  1: { fields: 'crm.lead.fields',    uf: 'crm.lead.userfield.list',    entityId: 'CRM_LEAD',    configMethod: 'crm.item.details.configuration.get', configParams: { entityTypeId: 1, scope: 'C' } },
};

const TIPO_LABEL = {
  string:'Texto', integer:'Inteiro', double:'Decimal', boolean:'Sim/No',
  date:'Data', datetime:'Data e hora', money:'Moeda', url:'URL',
  address:'Endereço', enumeration:'Lista', file:'Arquivo', employee:'Usuário',
  crm_status:'Status CRM', crm:'Vínculo CRM', iblock_section:'Seção IBlock',
  iblock_element:'Elem. IBlock',
};

const CAMPOS_TABLE_COLUMNS = [
  { key: 'label',   min: 260, max: 520, priority: 5, get: f => f.label },
  { key: 'id',      min: 118, max: 230, priority: 2, get: f => `${f.id} ${f.rawId || ''}` },
  { key: 'type',    min: 72,  max: 104, priority: 1, get: f => TIPO_LABEL[f.type] || f.type },
  { key: 'origin',  min: 82,  max: 116, priority: 1, get: f => f.origin === 'native' ? 'Nativo' : 'Personalizado' },
  { key: 'vis',     min: 70,  max: 94,  priority: 1, get: f => f.vis === 'vis' ? 'Visivel' : f.vis === 'fixed' ? 'Fixo' : 'Oculto' },
  { key: 'section', min: 82,  max: 150, priority: 1, get: f => f.section || '' },
  { key: 'options', min: 170, max: 380, priority: 4, get: f => (f.options || []).join(' ') },
  { key: 'acoes',   min: 84,  max: 98,  priority: 1, get: f => f.origin === 'custom' && f.rawId ? 'Editar Excluir' : '' },
];
const camposManualColumnWidths = {};
let camposColumnResizeBound = false;
let camposColumnResizeTimer = null;

function camposMode() {
  return window.EntityContext ? window.EntityContext.getCurrentMode() : 'crm';
}

function camposGetContext(requireSelection = true) {
  const mode = camposMode();
  if (mode === 'spa') {
    const spa = requireSelection
      ? window.EntityContext.requireActiveSpa()
      : window.EntityContext.getActiveSpa();
    return {
      mode,
      spa,
      entityTypeId: spa ? spa.entityTypeId : null,
      entityId: spa ? 'CRM_' + spa.id : '',
      label: spa ? spa.title : 'SPA',
      fieldsMethod: 'crm.item.fields',
      ufMethod: 'userfieldconfig.list',
      configMethod: 'crm.item.details.configuration.get',
    };
  }

  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(camposEntityTypeId);
  const api = CAMPOS_API[camposEntityTypeId];
  const meta = window.CRM_ENTITY_META && window.CRM_ENTITY_META[camposEntityTypeId];
  return {
    mode,
    entityTypeId: camposEntityTypeId,
    entityId: api ? api.entityId : '',
    label: meta ? meta.label : 'CRM',
    fieldsMethod: api ? api.fields : '',
    ufMethod: api ? api.uf : '',
    configMethod: api ? api.configMethod : 'crm.item.details.configuration.get',
  };
}

function camposContextKey(ctx = camposGetContext(false)) {
  if (ctx.mode === 'spa') return ctx.spa ? `spa:${ctx.spa.id}:${ctx.entityTypeId}:${camposPipId}` : 'spa:none';
  return `crm:${ctx.entityTypeId}:${camposPipId}`;
}

function camposEnsureContextBanner() {
  let banner = document.getElementById('campos-context-banner');
  if (banner) return banner;
  const tabs = document.getElementById('campos-entity-tabs');
  if (!tabs || !tabs.parentElement) return null;
  banner = document.createElement('div');
  banner.id = 'campos-context-banner';
  banner.className = 'form-hint';
  banner.style.marginTop = '8px';
  tabs.parentElement.insertBefore(banner, tabs.nextSibling);
  return banner;
}

function camposSyncContextUI() {
  const ctx = camposGetContext(false);
  const tabs = document.getElementById('campos-entity-tabs');
  const pipRow = document.getElementById('campos-pipeline-row');
  const pipLabel = pipRow ? pipRow.querySelector('.form-label') : null;
  const banner = camposEnsureContextBanner();

  if (tabs) tabs.style.display = ctx.mode === 'spa' ? 'none' : '';

  if (ctx.mode === 'spa') {
    if (pipRow) pipRow.style.display = ctx.spa ? 'block' : 'none';
    if (pipLabel) pipLabel.textContent = 'Pipeline/Funil da SPA';
    if (banner) {
      banner.innerHTML = ctx.spa
        ? `SPA ativo: <strong>${escHtml(ctx.spa.title)}</strong> (id ${escHtml(String(ctx.spa.id))} / entityTypeId ${escHtml(String(ctx.entityTypeId))}). Selecione o pipeline para conferir seções e visibilidade.`
        : 'Selecione um processo inteligente na Visão geral do módulo SPA antes de consultar campos.';
      banner.style.display = '';
    }
    if (ctx.spa) camposLoadPipelines();
  } else {
    document.querySelectorAll('#campos-entity-tabs .etab').forEach(t => {
      t.classList.toggle('active', String(t.dataset.entity) === String(camposEntityTypeId));
    });
    if (pipRow) pipRow.style.display = camposEntityTypeId === 2 ? 'block' : 'none';
    if (pipLabel) pipLabel.textContent = 'Pipeline (para config do card)';
    if (banner) banner.style.display = 'none';
    if (camposEntityTypeId === 2) camposLoadPipelines();
  }

  const key = camposContextKey(ctx);
  if (key !== camposLastContextKey) {
    camposAllFields = [];
    const resultCard = document.getElementById('campos-result-card');
    if (resultCard) resultCard.style.display = 'none';
    camposLastContextKey = key;
  }
}

function camposExtractNativeFields(data) {
  const result = data && data.result;
  if (result && result.fields && typeof result.fields === 'object') return result.fields;
  return result && typeof result === 'object' && !Array.isArray(result) ? result : {};
}

function camposExtractUserFieldPage(data) {
  const result = data && data.result;
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.fields)) return result.fields;
  if (result && Array.isArray(result.items)) return result.items;
  if (Array.isArray(data && data.fields)) return data.fields;
  return [];
}

function camposBool(value) {
  return value === true || value === 'Y' || value === '1' || value === 1;
}

function camposFieldName(field) {
  return field.FIELD_NAME || field.fieldName || field.name || field.NAME || '';
}

function camposFieldOptions(field) {
  const list = field.LIST || field.list || field.enum || field.ENUM || field.items || [];
  if (!Array.isArray(list)) return [];
  return list.map(o => o.VALUE || o.value || o.label || o.name || '').filter(Boolean);
}

function camposFieldKey(id) {
  return String(id || '').replace(/_/g, '').toUpperCase();
}

function camposGetSettingValue(settings, key, fallback = '') {
  if (!settings) return fallback;
  return settings[key] ?? settings[key.toLowerCase()] ?? fallback;
}

function camposReadEditSettings(type, currentSettings = {}) {
  const settings = { ...(currentSettings || {}) };
  if (type === 'string') {
    settings.ROWS = Math.max(1, parseInt(document.getElementById('campos-edit-string-rows').value, 10) || 1);
    if (settings.DEFAULT_VALUE === undefined) settings.DEFAULT_VALUE = '';
  }
  if (type === 'double') {
    settings.PRECISION = Math.max(0, parseInt(document.getElementById('campos-edit-double-precision').value, 10) || 0);
  }
  if (type === 'boolean') {
    const def = document.getElementById('campos-edit-boolean-default').value;
    settings.DEFAULT_VALUE = def === '0' ? '' : def === '1' ? 'Y' : 'N';
  }
  return settings;
}

function camposSetEditConditional(type, settings = {}, helpMessage = '') {
  const stringSection = document.getElementById('campos-edit-string-section');
  const doubleSection = document.getElementById('campos-edit-double-section');
  const booleanSection = document.getElementById('campos-edit-boolean-section');
  const enumSection = document.getElementById('campos-edit-enum-section');

  if (stringSection) {
    stringSection.style.display = type === 'string' ? 'block' : 'none';
    document.getElementById('campos-edit-string-rows').value = parseInt(camposGetSettingValue(settings, 'ROWS', 1), 10) || 1;
  }
  if (doubleSection) {
    doubleSection.style.display = type === 'double' ? 'block' : 'none';
    document.getElementById('campos-edit-double-precision').value = String(parseInt(camposGetSettingValue(settings, 'PRECISION', 2), 10) || 0);
  }
  if (booleanSection) {
    const def = camposGetSettingValue(settings, 'DEFAULT_VALUE', '');
    booleanSection.style.display = type === 'boolean' ? 'block' : 'none';
    document.getElementById('campos-edit-boolean-default').value = def === 'Y' || def === true ? '1' : def === 'N' || def === false ? '2' : '0';
  }
  if (enumSection) enumSection.style.display = type === 'enumeration' ? 'block' : 'none';

  const helpInput = document.getElementById('campos-edit-help');
  if (helpInput) helpInput.value = helpMessage || '';
}

function camposRemoveFromLoadedCard(fieldId) {
  if (typeof cardConfig === 'undefined' || !Array.isArray(cardConfig)) return;
  let removed = false;
  cardConfig.forEach(sec => {
    const before = (sec.elements || []).length;
    sec.elements = (sec.elements || []).filter(item => item.name !== fieldId);
    if (sec.elements.length !== before) removed = true;
  });
  if (typeof cardAllFields !== 'undefined' && cardAllFields) delete cardAllFields[fieldId];
  if (typeof cardSelectedHiddenFields !== 'undefined' && cardSelectedHiddenFields) cardSelectedHiddenFields.delete(fieldId);
  if (removed && typeof cardMarkUnsaved === 'function') cardMarkUnsaved();
  if (removed && typeof cardRender === 'function') cardRender();
}

function camposSelectEntity(eid, btn) {
  camposEntityTypeId = eid;
  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(eid);
  document.querySelectorAll('#campos-entity-tabs .etab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  const pipRow = document.getElementById('campos-pipeline-row');
  if (eid === 2) {
    pipRow.style.display = 'block';
    camposLoadPipelines();
  } else {
    pipRow.style.display = 'none';
    camposPipId = '0';
  }
  camposAllFields = [];
  document.getElementById('campos-result-card').style.display = 'none';
}

function camposSelectPipeline(value) {
  camposPipId = String(value || '0');
  camposAllFields = [];
  const resultCard = document.getElementById('campos-result-card');
  if (resultCard) resultCard.style.display = 'none';
  camposLastContextKey = camposContextKey(camposGetContext(false));
}

async function camposLoadPipelines() {
  try {
    const ctx = camposGetContext(false);
    const entityTypeId = ctx.mode === 'spa' ? ctx.entityTypeId : 2;
    if (!entityTypeId) return;

    const data = await call('crm.category.list', { entityTypeId });
    const cats = (data.result && data.result.categories) ? data.result.categories : (Array.isArray(data.result) ? data.result : []);
    const sel = document.getElementById('campos-pip-select');
    const previous = camposPipId || '0';
    sel.innerHTML = '<option value="0">Pipeline padrão (ID 0)</option>';
    cats.forEach(c => {
      if (String(c.id) !== '0') {
        const opt = document.createElement('option');
        opt.value = c.id;
        opt.textContent = escHtml(c.name) + ' (#' + c.id + ')';
        sel.appendChild(opt);
      }
    });
    sel.value = [...sel.options].some(opt => opt.value === previous) ? previous : '0';
    camposPipId = sel.value;
  } catch(e) { /* silencioso */ }
}

async function camposLoad() {
  camposSyncContextUI();
  if (camposMode() === 'spa') return camposLoadSpa();

  const api = CAMPOS_API[camposEntityTypeId];
  if (!api) return;
  const resultCard = document.getElementById('campos-result-card');
  resultCard.style.display = 'none';
  document.getElementById('campos-tbody').innerHTML = '<tr><td colspan="7" style="color:#aaa;font-size:12px;padding:12px;">Carregando</td></tr>';
  resultCard.style.display = 'block';

  try {
    const cfgExtras = {
      dealCategoryId: parseInt(camposPipId) || 0,
      leadCustomerType: 1,
    };
    const cfgPromise = call(api.configMethod, buildCardConfigParams(camposEntityTypeId, 'P', cfgExtras))
      .then(res => (Array.isArray(res.result) && res.result.length)
        ? res
        : call(api.configMethod, buildCardConfigParams(camposEntityTypeId, 'C', cfgExtras)));

    const ufFilter = { ENTITY_ID: api.entityId };
    const [nativeRes, firstUfRes, cfgRes] = await Promise.all([
      call(api.fields),
      call(api.uf, { filter: ufFilter, start: 0 }),
      cfgPromise.catch(() => ({ result: [] })),
    ]);

    // Paginão completa dos campos personalizados (Bitrix retorna 50 por pgina)
    const allCustoms = Array.isArray(firstUfRes.result) ? [...firstUfRes.result] : [];
    const PAGE_SIZE = 50;
    let ufStart = PAGE_SIZE;
    while (allCustoms.length === ufStart) {
      const pageRes = await call(api.uf, { filter: ufFilter, start: ufStart });
      const page = Array.isArray(pageRes.result) ? pageRes.result : [];
      if (!page.length) break;
      allCustoms.push(...page);
      ufStart += PAGE_SIZE;
    }

    // Mapa de opções (LIST) para campos do tipo lista/enumeração
    // e mapa de metadados brutos (ID numrico, MANDATORY, MULTIPLE)
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

    // Tenta userfieldconfig.list (fonte mais confivel para labels de UF_)
    // Usa camelCase nos parmetros e language:'br' conforme portal Bitrix24
    const ucfgMap = new Map();
    try {
      const ucfgFilter = { entityId: api.entityId };
      let ucfgStart = 0;
      while (true) {
        const ucfgRes = await call('userfieldconfig.list', {
          moduleId: 'crm',
          select: { 0: '*', language: 'br' },
          filter: ucfgFilter,
          start: ucfgStart,
        });
        // userfieldconfig retorna { result: { fields: [...] } }
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
            || '';
          const type = f.userTypeId || f.USER_TYPE_ID || '';
          if (label) ucfgMap.set(id, { label, type });
        });
        ucfgStart += PAGE_SIZE;
        if (ucfgPage.length < PAGE_SIZE) break;
      }
    } catch(e) {
      console.warn('[Bitrix24 Panel] userfieldconfig.list indisponível:', e.message);
    }

    // Mapa de labels/tipos: usa userfield.list como fallback quando userfieldconfig não retornou label
    const customLabelMap = new Map();
    allCustoms.forEach(f => {
      const id = f.FIELD_NAME || f.fieldName;
      if (!id) return;
      // Prioridade: ucfgMap (userfieldconfig.list) > EDIT_FORM_LABEL (userfield.list) > fallback
      if (ucfgMap.has(id)) {
        customLabelMap.set(id, ucfgMap.get(id));
        return;
      }
      const label = labelFromBitrix(f.EDIT_FORM_LABEL)
        || labelFromBitrix(f.LIST_COLUMN_LABEL)
        || id;
      const type = f.USER_TYPE_ID || f.userTypeId || '';
      customLabelMap.set(id, { label, type });
    });

    // Build configMap
    const sections = Array.isArray(cfgRes.result) ? cfgRes.result : [];
    const configMap = new Map();
    for (const sec of sections) {
      for (const el of (sec.elements || [])) {
        configMap.set(el.name, { section: sec.title || '', optionFlags: el.optionFlags ?? 0 });
      }
    }

    // Mapa unificado: campos nativos como base, UF_ sobrescritos com labels reais
    const fieldMap = new Map();
    const nativeObj = nativeRes.result || {};
    Object.entries(nativeObj).forEach(([id, info]) => {
      fieldMap.set(id, {
        id,
        label: (typeof info === 'object' && info.title) ? info.title : id,
        type:  (typeof info === 'object' && info.type)  ? info.type  : '',
        origin: id.startsWith('UF_') ? 'custom' : 'native',
      });
    });

    // Sobrescreve UF_ com labels reais (ucfgMap > userfield.list > crm.fields title > FIELD_NAME)
    customLabelMap.forEach(({ label, type }, id) => {
      const existing = fieldMap.get(id);
      const bestLabel = (label && label !== id)
        ? label
        : (existing && existing.label && existing.label !== id ? existing.label : id);
      fieldMap.set(id, { id, label: bestLabel, type, origin: 'custom' });
    });

    camposAllFields = Array.from(fieldMap.values()).map(f => {
      const cfg  = configMap.get(f.id);
      const meta = customRawMetaMap.get(f.id) || {};
      let vis = 'oculto';
      if (cfg) vis = cfg.optionFlags === 1 ? 'fixed' : 'vis';
      const options = customOptionsMap.get(f.id) || [];
      return {
        ...f,
        vis,
        section:   cfg ? cfg.section : '',
        options,
        rawId:     meta.rawId    ?? null,
        mandatory: meta.mandatory ?? false,
        multiple:  meta.multiple  ?? false,
        settings:  meta.settings  || {},
        helpMessage: meta.helpMessage || '',
      };
    });

    camposRenderTable();
  } catch(e) {
    document.getElementById('campos-tbody').innerHTML = `<tr><td colspan="8" style="color:#ef4444;font-size:12px;padding:12px;">Erro: ${escHtml(e.message)}</td></tr>`;
  }
}

async function camposLoadSpa() {
  const resultCard = document.getElementById('campos-result-card');
  resultCard.style.display = 'none';
  document.getElementById('campos-tbody').innerHTML = '<tr><td colspan="8" style="color:#aaa;font-size:12px;padding:12px;">Carregando...</td></tr>';
  resultCard.style.display = 'block';

  try {
    const ctx = camposGetContext(true);
    const cfgParamsP = buildCardConfigParams(ctx.entityTypeId, 'P', { categoryId: camposPipId }, 'spa');
    const cfgParamsC = buildCardConfigParams(ctx.entityTypeId, 'C', { categoryId: camposPipId }, 'spa');
    const cfgPromise = call('crm.item.details.configuration.get', cfgParamsP)
      .then(res => (Array.isArray(res.result) && res.result.length)
        ? res
        : call('crm.item.details.configuration.get', cfgParamsC));

    const [nativeRes, firstUfRes, cfgRes] = await Promise.all([
      call('crm.item.fields', { entityTypeId: ctx.entityTypeId, useOriginalUfNames: 'Y' }),
      call('userfieldconfig.list', {
        moduleId: 'crm',
        select: { 0: '*', language: 'br' },
        filter: { entityId: ctx.entityId },
        start: 0,
      }),
      cfgPromise.catch(() => ({ result: [] })),
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

    const customOptionsMap = new Map();
    const customRawMetaMap = new Map();
    const customLabelMap = new Map();

    allCustoms.forEach(f => {
      const id = camposFieldName(f);
      if (!id) return;
      const options = camposFieldOptions(f);
      if (options.length) customOptionsMap.set(id, options);
      customRawMetaMap.set(id, {
        rawId: f.ID || f.id || null,
        mandatory: camposBool(f.MANDATORY ?? f.mandatory ?? f.isMandatory),
        multiple: camposBool(f.MULTIPLE ?? f.multiple ?? f.isMultiple),
        settings: f.SETTINGS || f.settings || {},
        helpMessage: labelFromBitrix(f.HELP_MESSAGE || f.helpMessage) || '',
      });
      customLabelMap.set(id, {
        label: labelFromBitrix(f.editFormLabel)
          || labelFromBitrix(f.listColumnLabel)
          || labelFromBitrix(f.EDIT_FORM_LABEL)
          || labelFromBitrix(f.LIST_COLUMN_LABEL)
          || id,
        type: f.userTypeId || f.USER_TYPE_ID || f.type || '—',
      });
    });

    const sections = Array.isArray(cfgRes.result) ? cfgRes.result : [];
    const configMap = new Map();
    for (const sec of sections) {
      for (const el of (sec.elements || [])) {
        if (el && el.name) {
          const cfg = { section: sec.title || '—', optionFlags: el.optionFlags ?? 0 };
          configMap.set(el.name, cfg);
          configMap.set(camposFieldKey(el.name), cfg);
        }
      }
    }

    const fieldMap = new Map();
    const fieldKeyMap = new Map();
    const nativeObj = camposExtractNativeFields(nativeRes);
    Object.entries(nativeObj).forEach(([id, info]) => {
      fieldMap.set(id, {
        id,
        label: (typeof info === 'object' && (info.title || info.formLabel)) ? (info.title || info.formLabel) : id,
        type: (typeof info === 'object' && info.type) ? info.type : '—',
        origin: id.startsWith('UF_') ? 'custom' : 'native',
      });
      fieldKeyMap.set(camposFieldKey(id), id);
    });

    customLabelMap.forEach(({ label, type }, id) => {
      const targetId = fieldMap.has(id) ? id : (fieldKeyMap.get(camposFieldKey(id)) || id);
      const existing = fieldMap.get(targetId);
      const bestLabel = label && label !== id
        ? label
        : (existing && existing.label && existing.label !== id ? existing.label : id);
      fieldMap.set(targetId, { id: targetId, label: bestLabel, type, origin: 'custom' });
      fieldKeyMap.set(camposFieldKey(id), targetId);
    });

    camposAllFields = Array.from(fieldMap.values()).map(f => {
      const cfg = configMap.get(f.id) || configMap.get(camposFieldKey(f.id));
      const rawMetaKey = customRawMetaMap.has(f.id)
        ? f.id
        : Array.from(customRawMetaMap.keys()).find(key => camposFieldKey(key) === camposFieldKey(f.id));
      const meta = customRawMetaMap.get(rawMetaKey) || {};
      let vis = 'oculto';
      if (cfg) vis = cfg.optionFlags === 1 ? 'fixed' : 'vis';
      return {
        ...f,
        vis,
        section: cfg ? cfg.section : '—',
        options: customOptionsMap.get(f.id) || customOptionsMap.get(rawMetaKey) || [],
        rawId: meta.rawId ?? null,
        mandatory: meta.mandatory ?? false,
        multiple: meta.multiple ?? false,
        settings: meta.settings || {},
        helpMessage: meta.helpMessage || '',
      };
    });

    camposRenderTable();
  } catch(e) {
    document.getElementById('campos-tbody').innerHTML = `<tr><td colspan="8" style="color:#ef4444;font-size:12px;padding:12px;">Erro: ${escHtml(e.message)}</td></tr>`;
  }
}

function camposSetTab(tab, btn) {
  camposTab = tab;
  document.querySelectorAll('#campos-tabs-row .campos-tab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  camposRenderTable();
}

function camposSort(key) {
  if (camposSortKey === key) {
    camposSortAsc = !camposSortAsc;
  } else {
    camposSortKey = key;
    camposSortAsc = true;
  }
  camposRenderTable();
}

function camposTextWidth(text) {
  const clean = String(text || '').replace(/\s+/g, ' ').trim();
  if (!clean) return 0;
  const longestWord = clean.split(' ').reduce((max, word) => Math.max(max, word.length), 0);
  return Math.max(clean.length * 6.2, longestWord * 8.5);
}

function camposColumnHeader(col) {
  const labels = {
    label: 'NOME',
    id: 'ID',
    type: 'TIPO',
    origin: 'ORIGEM',
    vis: 'VISIBILIDADE',
    section: 'SECAO',
    options: 'OPCOES',
    acoes: 'ACOES',
  };
  return labels[col.key] || col.key;
}

function camposAutoWidthForColumn(col, fields) {
  const sample = fields.slice(0, 250);
  const headerWidth = camposColumnHeader(col).length * 7 + 22;
  const contentWidth = sample.reduce((max, field) => Math.max(max, camposTextWidth(col.get(field))), headerWidth);
  return Math.round(Math.min(col.max, Math.max(col.min, contentWidth + 18)));
}

function camposApplyColumnWidths(fields) {
  const table = document.querySelector('.campos-table');
  const wrap = document.querySelector('.campos-table-wrap');
  if (!table || !wrap) return;

  const widths = {};
  CAMPOS_TABLE_COLUMNS.forEach(col => {
    widths[col.key] = camposManualColumnWidths[col.key] || camposAutoWidthForColumn(col, fields);
  });

  const availableWidth = Math.max(wrap.clientWidth || 0, 760);
  let totalWidth = Object.values(widths).reduce((sum, value) => sum + value, 0);
  if (totalWidth < availableWidth) {
    const spare = availableWidth - totalWidth;
    const spareTargets = [
      ['label', .62],
      ['options', .30],
      ['section', .08],
    ].filter(([key]) => !camposManualColumnWidths[key]);
    const weightTotal = spareTargets.reduce((sum, [, weight]) => sum + weight, 0);
    spareTargets.forEach(([key, weight]) => {
      widths[key] += Math.round(spare * (weight / weightTotal));
    });
    if (spareTargets.length) {
      widths[spareTargets[spareTargets.length - 1][0]] += availableWidth - Object.values(widths).reduce((sum, value) => sum + value, 0);
    }
  } else if (totalWidth > availableWidth) {
    let excess = totalWidth - availableWidth;
    const shrinkables = CAMPOS_TABLE_COLUMNS
      .filter(col => !camposManualColumnWidths[col.key] && widths[col.key] > col.min)
      .sort((a, b) => a.priority - b.priority);
    shrinkables.forEach(col => {
      if (excess <= 0) return;
      const canShrink = widths[col.key] - col.min;
      const shrink = Math.min(canShrink, excess);
      widths[col.key] -= shrink;
      excess -= shrink;
    });
  }

  const finalWidth = Object.values(widths).reduce((sum, value) => sum + value, 0);
  table.style.minWidth = `${Math.min(finalWidth, availableWidth)}px`;
  table.style.width = finalWidth <= (wrap.clientWidth || finalWidth) ? '100%' : `${finalWidth}px`;

  CAMPOS_TABLE_COLUMNS.forEach(col => {
    const colEl = document.querySelector(`#campos-colgroup col[data-col="${col.key}"]`);
    if (colEl) colEl.style.width = `${Math.max(col.min, widths[col.key])}px`;
  });

  camposBindColumnResize();
}

function camposBindColumnResize() {
  if (camposColumnResizeBound) return;
  camposColumnResizeBound = true;

  window.addEventListener('resize', () => {
    clearTimeout(camposColumnResizeTimer);
    camposColumnResizeTimer = setTimeout(() => {
      if (document.getElementById('panel-campos')?.classList.contains('active')) camposRenderTable();
    }, 120);
  });

  document.querySelectorAll('.campos-table .col-resizer').forEach(handle => {
    handle.addEventListener('click', event => event.stopPropagation());
    handle.addEventListener('dblclick', event => {
      event.preventDefault();
      event.stopPropagation();
      const key = handle.closest('th')?.dataset.col;
      if (key) {
        delete camposManualColumnWidths[key];
        camposRenderTable();
      }
    });
    handle.addEventListener('pointerdown', event => {
      const th = handle.closest('th');
      const key = th?.dataset.col;
      const col = CAMPOS_TABLE_COLUMNS.find(item => item.key === key);
      if (!th || !key || !col) return;

      event.preventDefault();
      event.stopPropagation();
      th.classList.add('resizing');

      const startX = event.clientX;
      const startWidth = th.getBoundingClientRect().width;
      const move = moveEvent => {
        const nextWidth = Math.max(col.min, Math.round(startWidth + moveEvent.clientX - startX));
        camposManualColumnWidths[key] = nextWidth;
        camposApplyColumnWidths(camposAllFields);
      };
      const up = () => {
        th.classList.remove('resizing');
        window.removeEventListener('pointermove', move);
        window.removeEventListener('pointerup', up);
      };

      window.addEventListener('pointermove', move);
      window.addEventListener('pointerup', up, { once: true });
    });
  });
}

function camposUpdateMobileBulkToggle() {
  const btn = document.getElementById('campos-mobile-bulk-toggle');
  if (!btn) return;
  const rows = [...document.querySelectorAll('#campos-tbody tr[data-field-id]')];
  const collapsedCount = rows.filter(row => row.classList.contains('campos-mobile-collapsed')).length;
  const shouldExpand = rows.length > 0 && collapsedCount === rows.length;
  btn.textContent = shouldExpand ? 'Expandir todos' : 'Minimizar todos';
  btn.setAttribute('aria-expanded', shouldExpand ? 'false' : 'true');
}

function camposToggleMobileRow(fieldId) {
  const id = String(fieldId || '');
  if (!id) return;
  if (camposCollapsedFieldIds.has(id)) {
    camposCollapsedFieldIds.delete(id);
  } else {
    camposCollapsedFieldIds.add(id);
  }
  const safeId = window.CSS && CSS.escape ? CSS.escape(id) : id.replace(/"/g, '\\"');
  const row = document.querySelector(`#campos-tbody tr[data-field-id="${safeId}"]`);
  if (row) {
    const collapsed = camposCollapsedFieldIds.has(id);
    row.classList.toggle('campos-mobile-collapsed', collapsed);
    const btn = row.querySelector('.campos-mobile-row-toggle');
    if (btn) {
      btn.textContent = collapsed ? '+' : '-';
      btn.title = collapsed ? 'Expandir campo' : 'Minimizar campo';
      btn.setAttribute('aria-label', collapsed ? 'Expandir campo' : 'Minimizar campo');
      btn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
    }
  }
  camposUpdateMobileBulkToggle();
}

function camposToggleAllMobileRows() {
  const rows = [...document.querySelectorAll('#campos-tbody tr[data-field-id]')];
  if (!rows.length) return;
  const shouldCollapse = rows.some(row => !row.classList.contains('campos-mobile-collapsed'));
  rows.forEach(row => {
    const id = row.dataset.fieldId;
    if (!id) return;
    if (shouldCollapse) camposCollapsedFieldIds.add(id);
    else camposCollapsedFieldIds.delete(id);
    row.classList.toggle('campos-mobile-collapsed', shouldCollapse);
    const btn = row.querySelector('.campos-mobile-row-toggle');
    if (btn) {
      btn.textContent = shouldCollapse ? '+' : '-';
      btn.title = shouldCollapse ? 'Expandir campo' : 'Minimizar campo';
      btn.setAttribute('aria-label', shouldCollapse ? 'Expandir campo' : 'Minimizar campo');
      btn.setAttribute('aria-expanded', shouldCollapse ? 'false' : 'true');
    }
  });
  camposUpdateMobileBulkToggle();
}

function camposRenderTable() {
  const search = (document.getElementById('campos-search').value || '').toLowerCase();
  let fields = camposAllFields.filter(f => {
    if (camposTab === 'native' && f.origin !== 'native') return false;
    if (camposTab === 'custom' && f.origin !== 'custom') return false;
    if (search) {
      return f.label.toLowerCase().includes(search) ||
             f.id.toLowerCase().includes(search) ||
             f.type.toLowerCase().includes(search);
    }
    return true;
  });

  // Sort
  fields.sort((a, b) => {
    let av = a[camposSortKey] || '';
    let bv = b[camposSortKey] || '';
    const cmp = String(av).localeCompare(String(bv), 'pt-BR');
    return camposSortAsc ? cmp : -cmp;
  });

  // Update sort headers
  ['label','id','type','origin','vis','section'].forEach(k => {
    const th = document.getElementById(`campos-th-${k}`);
    if (!th) return;
    th.className = camposSortKey === k ? (camposSortAsc ? 'sort-asc' : 'sort-desc') : '';
  });

  document.getElementById('campos-count').textContent = `${fields.length} campo${fields.length !== 1 ? 's' : ''}`;

  const tbody = document.getElementById('campos-tbody');
  if (!fields.length) {
    tbody.innerHTML = '<tr><td colspan="8" style="color:#aaa;font-size:12px;padding:12px;">Nenhum campo encontrado.</td></tr>';
    camposApplyColumnWidths([]);
    camposUpdateMobileBulkToggle();
    return;
  }

  tbody.innerHTML = fields.map(f => {
    const globalIdx = camposAllFields.indexOf(f);
    const typeLabel  = TIPO_LABEL[f.type] || f.type;
    const visClass   = f.vis;
    const visLabel   = f.vis === 'vis' ? 'Visível' : f.vis === 'fixed' ? 'Fixo' : 'Oculto';
    const oriClass   = f.origin;
    const oriLabel   = f.origin === 'native' ? 'Nativo' : 'Personalizado';
    const opts       = f.options && f.options.length ? f.options : [];
    const optsHtml   = opts.length
      ? `<div class="campos-opts-cell">${opts.map(o => `<span class="campos-opt-tag">${escHtml(o)}</span>`).join('')}</div>`
      : '<span style="color:#ccc;font-size:11px;"></span>';
    const rawIdHtml = f.origin === 'custom' && f.rawId
      ? `<div style="font-size:10px;color:#999;margin-top:2px;">num: ${escHtml(f.rawId)}</div>`
      : '';
    const acoesHtml  = f.origin === 'custom' && f.rawId
      ? `<div style="display:flex;gap:5px;">
           <button class="campos-action-btn edit" onclick="camposOpenEdit(${globalIdx})">Editar</button>
           <button class="campos-action-btn del"  onclick="camposDeleteField(${globalIdx})">Excluir</button>
         </div>`
      : '<span style="color:#ccc;font-size:11px;"></span>';
    const collapsed = camposCollapsedFieldIds.has(f.id);
    return `
      <tr data-field-id="${escHtml(f.id)}" class="${collapsed ? 'campos-mobile-collapsed' : ''}">
        <td>
          <div class="campos-mobile-field-head">
            <div class="campos-mobile-field-summary">
              <span class="campos-mobile-field-title">${escHtml(f.label)}</span>
              <span class="campos-mobile-field-id">${escHtml(f.id)}</span>
            </div>
            <button type="button" class="icon-btn section-collapse-btn campos-mobile-row-toggle" onclick="camposToggleMobileRow(this.closest('tr').dataset.fieldId)" title="${collapsed ? 'Expandir campo' : 'Minimizar campo'}" aria-label="${collapsed ? 'Expandir campo' : 'Minimizar campo'}" aria-expanded="${collapsed ? 'false' : 'true'}">${collapsed ? '+' : '-'}</button>
          </div>
          <span class="campos-field-label-desktop">${escHtml(f.label)}</span>
        </td>
        <td style="font-family:monospace;font-size:10px;color:#555;">${escHtml(f.id)}${rawIdHtml}</td>
        <td>${escHtml(typeLabel)}</td>
        <td><span class="origin-badge ${escHtml(oriClass)}">${escHtml(oriLabel)}</span></td>
        <td><span class="vis-badge ${escHtml(visClass)}">${escHtml(visLabel)}</span></td>
        <td style="font-size:11px;color:#888;">${escHtml(f.section)}</td>
        <td>${optsHtml}</td>
        <td>${acoesHtml}</td>
      </tr>
    `;
  }).join('');
  camposApplyColumnWidths(fields);
  camposUpdateMobileBulkToggle();
}

function camposExportCSV() {
  if (!camposAllFields.length) { toast('Carregue os campos antes de exportar.', 'wn'); return; }
  const BOM = '\uFEFF';
  const header = ['Label','ID','Tipo','Origem','Visibilidade','Seção','Opções'];
  const rows = camposAllFields.map(f => [
    f.label, f.id, TIPO_LABEL[f.type] || f.type,
    f.origin === 'native' ? 'Nativo' : 'Personalizado',
    f.vis === 'vis' ? 'Visível' : f.vis === 'fixed' ? 'Fixo' : 'Oculto',
    f.section,
    (f.options && f.options.length) ? f.options.join(' | ') : '',
  ].map(v => `"${String(v).replace(/"/g,'""')}"`).join(';'));

  const csv = BOM + [header.join(';'), ...rows].join('\r\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  const ts   = new Date().toISOString().slice(0,19).replace(/[T:]/g,'-');
  const entityNames = { 4:'empresa', 3:'contato', 2:'negocio', 1:'lead' };
  const activeSpa = window.EntityContext ? window.EntityContext.getActiveSpa() : null;
  const exportName = camposMode() === 'spa' && activeSpa
    ? `spa-${activeSpa.id}`
    : (entityNames[camposEntityTypeId] || 'entidade');
  a.href = url;
  a.download = `campos-${exportName}-${ts}.csv`;
  a.click();
  URL.revokeObjectURL(url);
  toast('CSV exportado!', 'ok');
}

/*  Editar / Excluir campo  */
function camposOpenEdit(idx) {
  const f = camposAllFields[idx];
  if (!f || f.origin !== 'custom') return;
  camposEditFieldIndex = idx;

  document.getElementById('campos-edit-id-info').textContent = `ID: ${f.id}  |  Tipo: ${TIPO_LABEL[f.type] || f.type}`;
  document.getElementById('campos-edit-label').value = f.label;
  document.getElementById('campos-edit-mandatory').checked = !!f.mandatory;
  document.getElementById('campos-edit-multiple').checked  = !!f.multiple;

  if (f.type === 'enumeration') {
    document.getElementById('campos-edit-enum-options').value = (f.options || []).join('\n');
  }
  camposSetEditConditional(f.type, f.settings, f.helpMessage);

  document.getElementById('campos-edit-overlay').classList.add('open');
  document.getElementById('campos-edit-label').focus();
}

function camposCloseEdit() {
  document.getElementById('campos-edit-overlay').classList.remove('open');
  camposEditFieldIndex = -1;
}

async function camposSaveEditLegacy() {
  const f = camposAllFields[camposEditFieldIndex];
  if (!f) return;

  const label = document.getElementById('campos-edit-label').value.trim();
  if (!label) { toast('Informe o nome do campo.', 'wn'); return; }

  const mandatory = document.getElementById('campos-edit-mandatory').checked ? 'Y' : 'N';
  const multiple  = document.getElementById('campos-edit-multiple').checked  ? 'Y' : 'N';

  const fields = {
    EDIT_FORM_LABEL:    label,
    LIST_COLUMN_LABEL:  label,
    MANDATORY: mandatory,
    MULTIPLE:  multiple,
  };

  if (f.type === 'enumeration') {
    const optsText = document.getElementById('campos-edit-enum-options').value;
    const opts = optsText.split('\n').map(v => v.trim()).filter(Boolean);
    fields.LIST = opts.map((v, i) => ({ SORT: (i + 1) * 10, VALUE: v, DEF: 'N' }));
  }

  if (camposMode() === 'spa') {
    const spaField = {
      editFormLabel: { br: label },
      listColumnLabel: { br: label },
      mandatory,
      multiple,
    };
    if (f.type === 'enumeration') {
      const optsText = document.getElementById('campos-edit-enum-options').value;
      const opts = optsText.split('\n').map(v => v.trim()).filter(Boolean);
      spaField.enum = opts.map((v, i) => ({ sort: (i + 1) * 10, value: v, def: 'N' }));
    }

    const saveBtn = document.getElementById('campos-edit-save-btn');
    saveBtn.textContent = 'Salvando...';
    saveBtn.disabled = true;

    try {
      const data = await call('userfieldconfig.update', { moduleId: 'crm', id: f.rawId, field: spaField });
      if (data.error) throw new Error(data.error_description || data.error);

      camposAllFields[camposEditFieldIndex].label = label;
      camposAllFields[camposEditFieldIndex].mandatory = mandatory === 'Y';
      camposAllFields[camposEditFieldIndex].multiple = multiple === 'Y';
      if (f.type === 'enumeration') camposAllFields[camposEditFieldIndex].options = spaField.enum.map(o => o.value);

      camposCloseEdit();
      camposRenderTable();
      toast('Campo atualizado com sucesso!', 'ok');
    } catch (e) {
      toast('Erro ao salvar: ' + e.message, 'er');
    } finally {
      saveBtn.textContent = 'Salvar';
      saveBtn.disabled = false;
    }
    return;
  }

  const methods = ENTITY_METHODS[camposEntityTypeId];
  if (!methods || !methods.update) { toast('Método de atualização não disponível.', 'wn'); return; }

  const saveBtn = document.getElementById('campos-edit-save-btn');
  saveBtn.textContent = 'Salvando';
  saveBtn.disabled = true;

  try {
    const data = await call(methods.update, { id: f.rawId, fields });
    if (data.error) throw new Error(data.error_description || data.error);

    camposAllFields[camposEditFieldIndex].label     = label;
    camposAllFields[camposEditFieldIndex].mandatory = mandatory === 'Y';
    camposAllFields[camposEditFieldIndex].multiple  = multiple  === 'Y';
    if (f.type === 'enumeration') {
      const opts = document.getElementById('campos-edit-enum-options').value
        .split('\n').map(v => v.trim()).filter(Boolean);
      camposAllFields[camposEditFieldIndex].options = opts;
    }

    camposCloseEdit();
    camposRenderTable();
    toast('Campo atualizado com sucesso!', 'ok');
  } catch (e) {
    toast('Erro ao salvar: ' + e.message, 'er');
  } finally {
    saveBtn.textContent = 'Salvar';
    saveBtn.disabled = false;
  }
}

async function camposDeleteFieldLegacy(idx) {
  const f = camposAllFields[idx];
  if (!f || f.origin !== 'custom') return;

  if (!confirm(`Excluir o campo "${f.label}" (${f.id})?\n\nEsta ação não pode ser desfeita.`)) return;

  if (camposMode() === 'spa') {
    try {
      const data = await call('userfieldconfig.delete', { moduleId: 'crm', id: f.rawId });
      if (data.error) throw new Error(data.error_description || data.error);

      camposAllFields.splice(idx, 1);
      camposRenderTable();
      toast('Campo excluído com sucesso.', 'ok');
    } catch (e) {
      toast('Erro ao excluir: ' + e.message, 'er');
    }
    return;
  }

  const methods = ENTITY_METHODS[camposEntityTypeId];
  if (!methods || !methods.delete) { toast('Método de exclusão não disponível.', 'wn'); return; }

  try {
    const data = await call(methods.delete, { id: f.rawId });
    if (data.error) throw new Error(data.error_description || data.error);

    camposAllFields.splice(idx, 1);
    camposRenderTable();
    toast('Campo excludo com sucesso.', 'ok');
  } catch (e) {
    toast('Erro ao excluir: ' + e.message, 'er');
  }
}

async function camposSaveEdit() {
  const f = camposAllFields[camposEditFieldIndex];
  if (!f) return;

  const label = document.getElementById('campos-edit-label').value.trim();
  if (!label) { toast('Informe o nome do campo.', 'wn'); return; }

  const mandatory = document.getElementById('campos-edit-mandatory').checked ? 'Y' : 'N';
  const multiple = document.getElementById('campos-edit-multiple').checked ? 'Y' : 'N';
  const options = f.type === 'enumeration'
    ? document.getElementById('campos-edit-enum-options').value.split('\n').map(v => v.trim()).filter(Boolean)
    : (f.options || []);
  if (f.type === 'enumeration' && !options.length) {
    toast('Adicione ao menos uma opção à lista.', 'wn');
    return;
  }

  const settings = camposReadEditSettings(f.type, f.settings);
  const helpMessage = (document.getElementById('campos-edit-help') || { value: '' }).value.trim();
  const adapter = window.EntityContext ? window.EntityContext.getCurrentAdapter() : null;
  if (!adapter || !adapter.updateField || !window.EntityContext.buildUserFieldUpdatePayload) {
    toast('Método de atualização não disponível.', 'wn');
    return;
  }

  const saveBtn = document.getElementById('campos-edit-save-btn');
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
    }, camposMode());
    const data = camposMode() === 'spa'
      ? await adapter.updateField(f.rawId, payload)
      : await adapter.updateField(f.rawId, payload, camposEntityTypeId);
    if (data.error) throw new Error(data.error_description || data.error);

    camposAllFields[camposEditFieldIndex].label = label;
    camposAllFields[camposEditFieldIndex].mandatory = mandatory === 'Y';
    camposAllFields[camposEditFieldIndex].multiple = multiple === 'Y';
    camposAllFields[camposEditFieldIndex].settings = settings;
    camposAllFields[camposEditFieldIndex].helpMessage = helpMessage;
    if (f.type === 'enumeration') camposAllFields[camposEditFieldIndex].options = options;

    camposCloseEdit();
    camposRenderTable();
    toast('Campo atualizado com sucesso!', 'ok');
  } catch (e) {
    toast('Erro ao salvar: ' + e.message, 'er');
  } finally {
    saveBtn.textContent = 'Salvar';
    saveBtn.disabled = false;
  }
}

async function camposDeleteField(idx) {
  const f = camposAllFields[idx];
  if (!f || f.origin !== 'custom') return;

  if (!confirm(`Excluir o campo "${f.label}" (${f.id})?\n\nEsta ação não pode ser desfeita.`)) return;
  const adapter = window.EntityContext ? window.EntityContext.getCurrentAdapter() : null;
  if (!adapter || !adapter.deleteField) { toast('Método de exclusão não disponível.', 'wn'); return; }

  try {
    const data = camposMode() === 'spa'
      ? await adapter.deleteField(f.rawId)
      : await adapter.deleteField(f.rawId, camposEntityTypeId);
    if (data.error) throw new Error(data.error_description || data.error);

    camposAllFields.splice(idx, 1);
    camposRemoveFromLoadedCard(f.id);
    camposRenderTable();
    toast('Campo excludo com sucesso.', 'ok');
  } catch (e) {
    toast('Erro ao excluir: ' + e.message, 'er');
  }
}

window.camposSyncContextUI = camposSyncContextUI;
window.camposToggleMobileRow = camposToggleMobileRow;
window.camposToggleAllMobileRows = camposToggleAllMobileRows;
