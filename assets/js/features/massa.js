/* 
   MDULO: CRIAR EM MASSA (Tpico 5)
    */

let massaMode      = 'form';
let massaRows      = [];
let massaCsvData   = [];
let massaRunning   = false;
let massaCancelled = false;
let massaRowId     = 0;

const MASSA_ENTITIES = [
  { id: 4, key: 'empresa', label: 'Empresa', method: 'crm.company' },
  { id: 3, key: 'contato', label: 'Contato', method: 'crm.contact' },
  { id: 2, key: 'negocio', label: 'Negócio', method: 'crm.deal'    },
  { id: 1, key: 'lead',    label: 'Lead',    method: 'crm.lead'    },
];

const MASSA_TIPOS = [
  'string','integer','double','boolean','date','datetime',
  'money','url','enumeration','file','employee','crm_status','crm',
];

function massaInit() {
  massaAddRow();
  massaSyncContextUI();
}

function massaModeContext() {
  return window.EntityContext ? window.EntityContext.getCurrentMode() : 'crm';
}

function massaIsSpa() {
  return massaModeContext() === 'spa';
}

function massaGetActiveSpa(requireSelection = false) {
  if (!window.EntityContext) return null;
  return requireSelection ? window.EntityContext.requireActiveSpa() : window.EntityContext.getActiveSpa();
}

function massaGetSpaFieldPrefix() {
  const spa = massaGetActiveSpa(false);
  return spa && spa.id ? `UF_CRM_${spa.id}_` : 'UF_CRM_{SPA}_';
}

function massaFullFieldName(suffix) {
  if (massaIsSpa()) {
    const spa = massaGetActiveSpa(true);
    return `UF_CRM_${spa.id}_${suffix}`;
  }
  return `UF_CRM_${suffix}`;
}

function massaPreviewFieldName(suffix) {
  if (!suffix) return massaIsSpa() ? `${massaGetSpaFieldPrefix()}...` : 'UF_CRM_...';
  if (!massaIsSpa()) return `UF_CRM_${suffix}`;
  const spa = massaGetActiveSpa(false);
  return spa && spa.id ? `UF_CRM_${spa.id}_${suffix}` : `UF_CRM_{SPA}_${suffix}`;
}

function massaGetTargetLabel() {
  if (!massaIsSpa()) return 'Entidades';
  const spa = massaGetActiveSpa(false);
  return spa
    ? `SPA ativo: ${spa.title} (id ${spa.id})`
    : 'SPA ativo não selecionado';
}

function massaEnsureContextBanner() {
  const panel = document.getElementById('panel-massa');
  if (!panel) return null;
  let banner = document.getElementById('massa-context-banner');
  if (banner) return banner;
  banner = document.createElement('div');
  banner.id = 'massa-context-banner';
  banner.className = 'massa-context-banner';
  const firstCard = panel.querySelector('.card');
  if (firstCard && firstCard.parentNode) firstCard.parentNode.insertBefore(banner, firstCard.nextSibling);
  return banner;
}

function massaSyncContextUI() {
  const isSpa = massaIsSpa();
  const banner = massaEnsureContextBanner();
  const spa = massaGetActiveSpa(false);

  if (banner) {
    banner.innerHTML = isSpa
      ? (spa
        ? `Modo SPA: os campos serão criados no processo ativo <strong>${escHtml(spa.title)}</strong> (id ${escHtml(String(spa.id))} / entityTypeId ${escHtml(String(spa.entityTypeId))}).`
        : 'Modo SPA: selecione um processo inteligente na Visão geral antes de criar campos em massa.')
      : 'Modo CRM: marque uma ou mais entidades por linha. O template traz as colunas Empresa, Contato, Negócio e Lead.';
    banner.style.display = '';
  }

  const firstTh = document.querySelector('#panel-massa .massa-table thead th:first-child');
  if (firstTh) firstTh.textContent = isSpa ? 'Processo alvo' : 'Entidades';
  document.querySelectorAll('.massa-target-cell').forEach(cell => {
    const rid = cell.dataset.row;
    cell.innerHTML = massaBuildTargetCellHtml(rid);
  });
  document.querySelectorAll('.id-preview').forEach(preview => {
    const rid = preview.id.replace('midp-', '');
    if (rid) massaOnIdInput(rid);
  });
}

function massaSwitchMode(mode) {
  massaMode = mode;
  document.querySelectorAll('.massa-tab').forEach((t, i) => {
    t.classList.toggle('active', (i === 0 && mode === 'form') || (i === 1 && mode === 'csv'));
  });
  document.getElementById('massa-form-section').style.display = mode === 'form' ? 'block' : 'none';
  document.getElementById('massa-csv-section').style.display  = mode === 'csv'  ? 'block' : 'none';
  massaSyncContextUI();
}

function massaBuildTargetCellHtml(rid) {
  if (massaIsSpa()) {
    const spa = massaGetActiveSpa(false);
    return `
      <div class="massa-spa-target">
        <div class="massa-spa-target-label">${escHtml(spa ? spa.title : 'Nenhum SPA selecionado')}</div>
        <div class="massa-spa-target-meta">${spa ? `id ${escHtml(String(spa.id))} / entityTypeId ${escHtml(String(spa.entityTypeId))}` : 'Use a Visão geral do módulo SPA'}</div>
      </div>
    `;
  }

  return MASSA_ENTITIES.map(e => `
    <label class="entity-check-label">
      <input type="checkbox" class="massa-ent-check" data-row="${rid}" data-key="${e.key}" data-method="${e.method}" checked />
      ${escHtml(e.label)}
    </label>
  `).join('');
}

function massaAddRow() {
  massaRowId++;
  const rid = massaRowId;
  massaRows.push(rid);
  const tbody = document.getElementById('massa-tbody');
  const tr = document.createElement('tr');
  tr.id = `massa-row-${rid}`;

  const tipoOptions = MASSA_TIPOS.map(t =>
    `<option value="${t}">${escHtml(TIPO_LABEL[t] || t)}</option>`
  ).join('');

  tr.innerHTML = `
    <td><div class="entity-checks massa-target-cell" data-row="${rid}">${massaBuildTargetCellHtml(rid)}</div></td>
    <td>
      <input class="massa-input" id="ml-${rid}" placeholder="Label do campo"
        oninput="massaOnLabelInput(${rid})" />
    </td>
    <td>
      <input class="massa-input" id="mid-${rid}" placeholder="SUFIXO" maxlength="13" style="text-transform:uppercase;"
        oninput="massaOnIdInput(${rid})" />
      <div class="id-preview" id="midp-${rid}">UF_CRM_</div>
    </td>
    <td>
      <select class="massa-select" id="mtype-${rid}">
        ${tipoOptions}
      </select>
    </td>
    <td>
      <input class="massa-input" id="menum-${rid}" placeholder="Op1;Op2;Op3" title="Opções separadas por ; (s para Lista)" />
    </td>
    <td>
      <input class="massa-input" id="mhint-${rid}" placeholder="Dica (opcional)" />
    </td>
    <td style="text-align:center;">
      <input type="checkbox" id="mmand-${rid}" title="Obrigatrio" />
    </td>
    <td style="text-align:center;">
      <input type="checkbox" id="mmult-${rid}" title="Mltiplo" />
    </td>
    <td>
      <button class="icon-btn del" onclick="massaRemoveRow(${rid})" title="Remover linha"></button>
    </td>
  `;
  tbody.appendChild(tr);
}

function massaRemoveRow(rid) {
  const tr = document.getElementById(`massa-row-${rid}`);
  if (tr) tr.remove();
  massaRows = massaRows.filter(r => r !== rid);
}

function massaOnLabelInput(rid) {
  const label = document.getElementById(`ml-${rid}`).value;
  const suggested = massaIdSuggest(label);
  const idInput = document.getElementById(`mid-${rid}`);
  if (!idInput.dataset.manual) {
    idInput.value = suggested;
    massaOnIdInput(rid);
  }
}

function massaOnIdInput(rid) {
  const input = document.getElementById(`mid-${rid}`);
  let val = input.value.toUpperCase().replace(/[^A-Z0-9_]/g, '');
  if (val !== input.value) input.value = val;
  input.dataset.manual = val ? '1' : '';
  const preview = document.getElementById(`midp-${rid}`);
  const v = massaValidateId(val);
  preview.textContent = massaPreviewFieldName(val);
  preview.className   = 'id-preview ' + (val ? (v.ok ? 'ok' : 'err') : '');
}

function massaIdSuggest(label) {
  if (!label) return '';
  const stopwords = new Set(['DE','DO','DA','DAS','DOS','UM','UMA','O','A','OS','AS','E','OU','EM','NO','NA','POR','PARA','COM','SEM']);
  return label
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .toUpperCase()
    .replace(/[^A-Z0-9\s]/g,'')
    .trim()
    .split(/\s+/)
    .filter(w => !stopwords.has(w))
    .join('_')
    .replace(/__+/g,'_')
    .slice(0,13)
    .replace(/_+$/,'');
}

function massaValidateId(suffix) {
  if (!suffix) return { ok: false, msg: 'ID vazio' };
  if (/^[0-9]/.test(suffix)) return { ok: false, msg: 'Não pode começar com número' };
  if (/__/.test(suffix))     return { ok: false, msg: 'Sem __ duplo' };
  if (suffix.length < 2)     return { ok: false, msg: 'Mnimo 2 chars' };
  if (!/^[A-Z0-9_]+$/.test(suffix)) return { ok: false, msg: 'Apenas A-Z, 0-9, _' };
  return { ok: true, msg: '' };
}

function massaGetActiveRows() {
  return massaRows.filter(rid => document.getElementById(`massa-row-${rid}`));
}

function massaValidateAll() {
  let valid = true;
  if (massaIsSpa()) {
    try {
      massaGetActiveSpa(true);
    } catch(e) {
      toast(e.message || 'Selecione um processo inteligente antes de criar campos em massa.', 'wn', 5000);
      return false;
    }
  }
  for (const rid of massaGetActiveRows()) {
    const label = (document.getElementById(`ml-${rid}`).value || '').trim();
    const suffix = (document.getElementById(`mid-${rid}`).value || '').trim();
    const labelEl = document.getElementById(`ml-${rid}`);
    const idEl    = document.getElementById(`mid-${rid}`);
    labelEl.classList.remove('invalid','valid');
    idEl.classList.remove('invalid','valid');
    if (!label) { labelEl.classList.add('invalid'); valid = false; }
    else labelEl.classList.add('valid');
    const v = massaValidateId(suffix);
    if (!v.ok) { idEl.classList.add('invalid'); valid = false; }
    else idEl.classList.add('valid');

    if (!massaIsSpa()) {
      const hasEntity = MASSA_ENTITIES.some(e => {
        const cb = document.querySelector(`[data-row="${rid}"][data-key="${e.key}"]`);
        return cb && cb.checked;
      });
      if (!hasEntity) valid = false;
    }
  }
  return valid;
}

function massaLogLine(msg, type = 'ok') {
  const logId = massaMode === 'form' ? 'massa-log' : 'massa-log-csv';
  const log = document.getElementById(logId);
  if (!log) return;
  log.style.display = 'flex';
  const line = document.createElement('div');
  line.className = `massa-log-line ${type}`;
  line.textContent = msg;
  log.appendChild(line);
  log.scrollTop = log.scrollHeight;
}

function massaUpdateProgress(done, total) {
  const barId = massaMode === 'form' ? 'massa-progress-bar' : 'massa-progress-bar-csv';
  const txtId = massaMode === 'form' ? 'massa-progress-txt' : 'massa-progress-txt-csv';
  const wrapId= massaMode === 'form' ? 'massa-progress-wrap' : 'massa-progress-wrap-csv';
  const wrap  = document.getElementById(wrapId);
  if (wrap) wrap.style.display = 'block';
  const bar = document.getElementById(barId);
  const txt = document.getElementById(txtId);
  if (bar) bar.style.width = total > 0 ? `${Math.round(done/total*100)}%` : '0%';
  if (txt) txt.textContent = `${done} / ${total}`;
}

function delay(ms) { return new Promise(r => setTimeout(r, ms)); }

function massaExtractUserfieldConfigFields(data) {
  const result = data && data.result;
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.fields)) return result.fields;
  if (result && Array.isArray(result.items)) return result.items;
  return [];
}

function massaNormalizeCsvBoolean(value) {
  const val = String(value || '').toLowerCase().trim();
  return val === 'sim' || val === 's' || val === '1' || val === 'x' || val === 'true' || val === 'yes';
}

function massaNormalizeType(rawType) {
  const normalized = String(rawType || '')
    .toLowerCase()
    .trim()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g,'');
  return ({
    'texto': 'string',
    'numero inteiro': 'integer',
    'inteiro': 'integer',
    'numero decimal': 'double',
    'decimal': 'double',
    'sim/não': 'boolean',
    'sim não': 'boolean',
    'data e hora': 'datetime',
    'data': 'date',
    'moeda': 'money',
    'url': 'url',
    'lista': 'enumeration',
    'arquivo': 'file',
    'usuario': 'employee',
    'status crm': 'crm_status',
    'vinculo crm': 'crm',
  })[normalized] || rawType || 'string';
}

function massaGetItemTargets(item) {
  if (massaIsSpa()) {
    const spa = massaGetActiveSpa(true);
    return [{
      mode: 'spa',
      label: spa.title || 'SPA',
      spa,
    }];
  }
  return item.entities || [];
}

function massaBuildSpaField(item, fieldName) {
  const field = {
    fieldName,
    userTypeId: item.type,
    multiple: item.multiple ? 'Y' : 'N',
    mandatory: item.mandatory ? 'Y' : 'N',
    showFilter: 'Y',
    editInList: 'Y',
    sort: 100,
    editFormLabel: { br: item.label },
    listColumnLabel: { br: item.label },
    listFilterLabel: { br: item.label },
    settings: {},
  };

  if (item.hint) field.helpMessage = { br: item.hint };
  if (item.type === 'enumeration' && item.enumStr) {
    field.settings = { DISPLAY: 'UI', LIST_HEIGHT: 1 };
    field.enum = massaBuildEnumList(item.enumStr).map(opt => ({
      value: opt.VALUE,
      xmlId: opt.XML_ID,
      sort: opt.SORT,
      def: opt.DEF,
    }));
  }

  return field;
}

async function massaCheckDuplicateForTarget(fieldName, target) {
  if (target.mode === 'spa') {
    const entityId = `CRM_${target.spa.id}`;
    let start = 0;
    while (true) {
      const r = await call('userfieldconfig.list', {
        moduleId: 'crm',
        select: { 0: '*', language: 'br' },
        filter: { entityId },
        start,
      });
      const arr = massaExtractUserfieldConfigFields(r);
      if (arr.some(f => String(f.fieldName || f.FIELD_NAME || '').toUpperCase() === fieldName.toUpperCase())) {
        return true;
      }
      if (arr.length < 50) return false;
      start += 50;
    }
  }

  const r = await call(`${target.method}.userfield.list`, { filter: { FIELD_NAME: fieldName } });
  const arr = Array.isArray(r.result) ? r.result : [];
  return arr.length > 0;
}

async function massaCheckDuplicates() {
  const rows = massaMode === 'form' ? massaGetActiveRows() : massaCsvData;
  if (!rows.length) { toast('Nenhuma linha para verificar.', 'wn'); return; }
  toast('Verificando duplicatas', 'wn', 2000);
  let found = 0;
  if (massaMode === 'form') {
    for (const rid of rows) {
      const suffix = (document.getElementById(`mid-${rid}`).value || '').trim().toUpperCase();
      if (!suffix) continue;
      const fieldName = `UF_CRM_${suffix}`;
      for (const ent of MASSA_ENTITIES) {
        const checks = document.querySelectorAll(`[data-row="${rid}"][data-key="${ent.key}"]`);
        if (checks.length && !checks[0].checked) continue;
        try {
          const r = await call(`${ent.method}.userfield.list`, { filter: { FIELD_NAME: fieldName } });
          const arr = Array.isArray(r.result) ? r.result : [];
          if (arr.length > 0) {
            massaLogLine(` Duplicata: ${fieldName} já existe em ${ent.label}`, 'wn');
            found++;
          }
        } catch(e) { /* silencioso */ }
      }
    }
  }
  if (found === 0) toast('Nenhuma duplicata encontrada!', 'ok');
  else toast(`${found} duplicata(s) encontrada(s). Veja o log.`, 'wn');
}

async function massaCheckDuplicates() {
  const rows = massaMode === 'form' ? massaGetActiveRows() : massaCsvData;
  if (!rows.length) { toast('Nenhuma linha para verificar.', 'wn'); return; }
  if (massaIsSpa()) {
    try {
      massaGetActiveSpa(true);
    } catch(e) {
      toast(e.message || 'Selecione um processo inteligente antes de verificar duplicatas.', 'wn', 5000);
      return;
    }
  }

  toast('Verificando duplicatas...', 'wn', 2000);
  let found = 0;

  const checkItem = async (suffix, item) => {
    if (!suffix) return;
    const fieldName = massaFullFieldName(suffix);
    for (const target of massaGetItemTargets(item)) {
      try {
        if (await massaCheckDuplicateForTarget(fieldName, target)) {
          massaLogLine(`Duplicata: ${fieldName} ja existe em ${target.label}`, 'wn');
          found++;
        }
      } catch(e) { /* silencioso */ }
    }
  };

  if (massaMode === 'form') {
    for (const rid of rows) {
      await checkItem((document.getElementById(`mid-${rid}`).value || '').trim().toUpperCase(), {
        entities: MASSA_ENTITIES.filter(e => {
          const cb = document.querySelector(`[data-row="${rid}"][data-key="${e.key}"]`);
          return cb && cb.checked;
        }),
      });
    }
  } else {
    for (const row of rows) {
      await checkItem((row.id || '').trim().toUpperCase(), {
        entities: MASSA_ENTITIES.filter(e => massaNormalizeCsvBoolean(row[e.key] || row[e.label])),
      });
    }
  }

  if (found === 0) toast('Nenhuma duplicata encontrada!', 'ok');
  else toast(`${found} duplicata(s) encontrada(s). Veja o log.`, 'wn');
}

async function massaRun() {
  if (massaRunning) return;
  const logId = massaMode === 'form' ? 'massa-log' : 'massa-log-csv';
  const log = document.getElementById(logId);
  if (log) { log.innerHTML = ''; log.style.display = 'flex'; }

  const cancelBtnId = massaMode === 'form' ? 'massa-cancel-btn' : 'massa-cancel-btn-csv';
  const cancelBtn   = document.getElementById(cancelBtnId);

  let items = [];
  if (massaMode === 'form') {
    if (!massaValidateAll()) { toast('Corrija os campos marcados em vermelho.', 'er'); return; }
    items = massaGetActiveRows().map(rid => ({
      label:    (document.getElementById(`ml-${rid}`).value || '').trim(),
      suffix:   (document.getElementById(`mid-${rid}`).value || '').trim().toUpperCase(),
      type:     document.getElementById(`mtype-${rid}`).value,
      enumStr:  (document.getElementById(`menum-${rid}`).value || ''),
      hint:     (document.getElementById(`mhint-${rid}`).value || '').trim(),
      mandatory: document.getElementById(`mmand-${rid}`).checked,
      multiple:  document.getElementById(`mmult-${rid}`).checked,
      entities: MASSA_ENTITIES.filter(e => {
        const cb = document.querySelector(`[data-row="${rid}"][data-key="${e.key}"]`);
        return cb && cb.checked;
      }),
    }));
  } else {
    items = massaCsvData.map(row => ({
      label:    row['nome do campo'] || row.label    || '',
      suffix:   (row.id     || '').toUpperCase(),
      type:     ({'texto':'string','numero inteiro':'integer','numero decimal':'double','sim/não':'boolean','sim/não':'boolean','data e hora':'datetime','data':'date','moeda':'money','url':'url','lista':'enumeration','arquivo':'file','usuario':'employee','status crm':'crm_status','vinculo crm':'crm'}[(row.tipo||'').toLowerCase().trim().normalize('NFD').replace(/[-]/g,'').replace('/', '/')] || row.tipo || 'string'),
      enumStr:  row.opcoes   || '',
      hint:     row['descrição do campo'] || row['descricao do campo'] || row.hint || '',
      mandatory: String(row.obrigatorio).toLowerCase() === 'sim',
      multiple:  String(row.multiplo).toLowerCase()    === 'sim',
      entities: MASSA_ENTITIES.filter(e => {
        const val = String(row[e.key] || row[e.label] || '').toLowerCase();
        return val === 'sim' || val === 's' || val === '1' || val === 'x' || val === 'true';
      }),
    }));
  }

  if (!items.length) { toast('Nenhuma linha para criar.', 'wn'); return; }

  massaRunning   = true;
  massaCancelled = false;
  if (cancelBtn) cancelBtn.style.display = 'inline-flex';

  let done  = 0;
  let total = items.reduce((acc, item) => acc + item.entities.length, 0);
  massaUpdateProgress(0, total);

  let ok = 0, err = 0, skipped = 0;

  for (const item of items) {
    if (massaCancelled) { massaLogLine('Operação cancelada pelo usuário.', 'wn'); break; }
    const fieldName = `UF_CRM_${item.suffix}`;
    for (const ent of item.entities) {
      if (massaCancelled) break;
      // Check duplicate
      try {
        const chk = await call(`${ent.method}.userfield.list`, { filter: { FIELD_NAME: fieldName } });
        if (Array.isArray(chk.result) && chk.result.length > 0) {
          massaLogLine(` SKIP  ${fieldName} já existe em ${ent.label}`, 'wn');
          skipped++;
          done++;
          massaUpdateProgress(done, total);
          await delay(300);
          continue;
        }
      } catch(e) { /* continua */ }

      // Build fields
      const fields = {
        USER_TYPE_ID: item.type,
        FIELD_NAME:   item.suffix,
        LABEL:        item.label,
        MANDATORY:    item.mandatory ? 'Y' : 'N',
        MULTIPLE:     item.multiple  ? 'Y' : 'N',
        EDIT_IN_LIST: 'Y',
        SORT:         100,
      };
      if (item.hint) fields.HELP_MESSAGE = item.hint;
      if (item.type === 'enumeration' && item.enumStr) {
        fields.LIST = massaBuildEnumList(item.enumStr);
        fields.SETTINGS = { DISPLAY: 'UI', LIST_HEIGHT: 1 };
      }

      try {
        const r = await call(`${ent.method}.userfield.add`, { fields });
        if (r.error) throw new Error(r.error_description || r.error);
        massaLogLine(` ${fieldName} criado em ${ent.label}`, 'ok');
        ok++;
      } catch(e) {
        massaLogLine(` ${fieldName} em ${ent.label}: ${e.message}`, 'er');
        err++;
      }
      done++;
      massaUpdateProgress(done, total);
      await delay(300);
    }
  }

  massaRunning = false;
  if (cancelBtn) cancelBtn.style.display = 'none';
  toast(`Concludo: ${ok} criado(s), ${skipped} ignorado(s), ${err} erro(s).`, err > 0 ? 'wn' : 'ok', 6000);
}

function massaBuildItems() {
  if (massaMode === 'form') {
    if (!massaValidateAll()) return null;
    return massaGetActiveRows().map(rid => ({
      label: (document.getElementById(`ml-${rid}`).value || '').trim(),
      suffix: (document.getElementById(`mid-${rid}`).value || '').trim().toUpperCase(),
      type: document.getElementById(`mtype-${rid}`).value,
      enumStr: (document.getElementById(`menum-${rid}`).value || ''),
      hint: (document.getElementById(`mhint-${rid}`).value || '').trim(),
      mandatory: document.getElementById(`mmand-${rid}`).checked,
      multiple: document.getElementById(`mmult-${rid}`).checked,
      entities: MASSA_ENTITIES.filter(e => {
        const cb = document.querySelector(`[data-row="${rid}"][data-key="${e.key}"]`);
        return cb && cb.checked;
      }),
    }));
  }

  if (massaIsSpa()) {
    try {
      massaGetActiveSpa(true);
    } catch(e) {
      toast(e.message || 'Selecione um processo inteligente antes de criar campos em massa.', 'wn', 5000);
      return null;
    }
  }

  return massaCsvData.map(row => ({
    label: row['nome do campo'] || row.label || '',
    suffix: (row.id || '').toUpperCase(),
    type: massaNormalizeType(row.tipo),
    enumStr: row.opcoes || '',
    hint: row['descrição do campo'] || row['descricao do campo'] || row.hint || '',
    mandatory: massaNormalizeCsvBoolean(row.obrigatorio),
    multiple: massaNormalizeCsvBoolean(row.multiplo),
    spaColumn: row.spa || row.processo || '',
    entities: massaIsSpa()
      ? []
      : MASSA_ENTITIES.filter(e => massaNormalizeCsvBoolean(row[e.key] || row[e.label])),
  })).filter(item => item.label && item.suffix);
}

async function massaRun() {
  if (massaRunning) return;
  const logId = massaMode === 'form' ? 'massa-log' : 'massa-log-csv';
  const log = document.getElementById(logId);
  if (log) { log.innerHTML = ''; log.style.display = 'flex'; }

  const cancelBtnId = massaMode === 'form' ? 'massa-cancel-btn' : 'massa-cancel-btn-csv';
  const cancelBtn = document.getElementById(cancelBtnId);
  const items = massaBuildItems();
  if (!items) return;
  if (!items.length) { toast('Nenhuma linha para criar.', 'wn'); return; }

  massaRunning = true;
  massaCancelled = false;
  if (cancelBtn) cancelBtn.style.display = 'inline-flex';

  let done = 0;
  const total = items.reduce((acc, item) => acc + massaGetItemTargets(item).length, 0);
  massaUpdateProgress(0, total);

  let ok = 0, err = 0, skipped = 0;

  for (const item of items) {
    if (massaCancelled) { massaLogLine('Operacao cancelada pelo usuario.', 'wn'); break; }
    const fieldName = massaFullFieldName(item.suffix);
    const targets = massaGetItemTargets(item);

    if (!targets.length) {
      massaLogLine(`SKIP - ${fieldName}: nenhuma entidade alvo selecionada`, 'wn');
      continue;
    }

    for (const target of targets) {
      if (massaCancelled) break;

      try {
        if (await massaCheckDuplicateForTarget(fieldName, target)) {
          massaLogLine(`SKIP - ${fieldName} ja existe em ${target.label}`, 'wn');
          skipped++;
          done++;
          massaUpdateProgress(done, total);
          await delay(300);
          continue;
        }
      } catch(e) { /* continua */ }

      try {
        let r;
        if (target.mode === 'spa') {
          r = await call('userfieldconfig.add', {
            moduleId: 'crm',
            field: {
              entityId: `CRM_${target.spa.id}`,
              ...massaBuildSpaField(item, fieldName),
            },
          });
        } else {
          const fields = {
            USER_TYPE_ID: item.type,
            FIELD_NAME: item.suffix,
            LABEL: item.label,
            MANDATORY: item.mandatory ? 'Y' : 'N',
            MULTIPLE: item.multiple ? 'Y' : 'N',
            EDIT_IN_LIST: 'Y',
            SORT: 100,
          };
          if (item.hint) fields.HELP_MESSAGE = item.hint;
          if (item.type === 'enumeration' && item.enumStr) {
            fields.LIST = massaBuildEnumList(item.enumStr);
            fields.SETTINGS = { DISPLAY: 'UI', LIST_HEIGHT: 1 };
          }
          r = await call(`${target.method}.userfield.add`, { fields });
        }

        if (r.error) throw new Error(r.error_description || r.error);
        massaLogLine(`${fieldName} criado em ${target.label}`, 'ok');
        ok++;
      } catch(e) {
        massaLogLine(`${fieldName} em ${target.label}: ${e.message}`, 'er');
        err++;
      }

      done++;
      massaUpdateProgress(done, total);
      await delay(300);
    }
  }

  massaRunning = false;
  if (cancelBtn) cancelBtn.style.display = 'none';
  toast(`Concluido: ${ok} criado(s), ${skipped} ignorado(s), ${err} erro(s).`, err > 0 ? 'wn' : 'ok', 6000);
}

function massaCancel() {
  massaCancelled = true;
  massaRunning   = false;
}

function massaBuildEnumList(str) {
  return str.split(';').map((s, i) => {
    const v = s.trim();
    return {
      VALUE:  v,
      XML_ID: v.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9_]/g,'_').replace(/__+/g,'_').slice(0,50) || `OPT_${i+1}`,
      SORT:   (i+1)*10,
      DEF:    'N',
    };
  }).filter(o => o.VALUE);
}

function massaDownloadTemplate() {
  const b64 = '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';
  const bin  = atob(b64);
  const arr  = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) arr[i] = bin.charCodeAt(i);
  const blob = new Blob([arr], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  const url  = URL.createObjectURL(blob);
  const a    = document.createElement('a');
  a.href = url;
  a.download = 'template-massa.xlsx';
  a.click();
  URL.revokeObjectURL(url);
  toast('Template XLS baixado!', 'ok');
}

function massaDownloadTemplate() {
  const isSpa = massaIsSpa();
  const spa = massaGetActiveSpa(false);
  const rows = isSpa
    ? [
      {
        spa: spa ? `${spa.id} - ${spa.title}` : 'usar processo ativo',
        'nome do campo': 'Exemplo SPA',
        id: 'EXEMPLO_SPA',
        tipo: 'texto',
        opcoes: '',
        'descricao do campo': 'Criado no processo ativo',
        obrigatorio: 'não',
        multiplo: 'não',
      },
    ]
    : [
      {
        empresa: 'sim',
        contato: 'sim',
        negocio: 'não',
        lead: 'não',
        'nome do campo': 'Exemplo CRM',
        id: 'EXEMPLO_CRM',
        tipo: 'texto',
        opcoes: '',
        'descricao do campo': 'Criado nas entidades marcadas',
        obrigatorio: 'não',
        multiplo: 'não',
      },
    ];

  if (typeof XLSX !== 'undefined') {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.json_to_sheet(rows);
    XLSX.utils.book_append_sheet(wb, ws, isSpa ? 'SPA ativo' : 'CRM entidades');
    XLSX.writeFile(wb, isSpa ? 'template-massa-spa.xlsx' : 'template-massa-crm.xlsx');
    toast(`Template ${isSpa ? 'SPA' : 'CRM'} baixado!`, 'ok');
    return;
  }

  const headers = Object.keys(rows[0]);
  const csv = [
    headers.join(';'),
    headers.map(h => `"${String(rows[0][h]).replace(/"/g, '""')}"`).join(';'),
  ].join('\n');
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = isSpa ? 'template-massa-spa.csv' : 'template-massa-crm.csv';
  a.click();
  URL.revokeObjectURL(url);
  toast(`Template ${isSpa ? 'SPA' : 'CRM'} baixado em CSV!`, 'ok');
}

function massaHandleFileUpload(file) {
  if (!file) return;
  const ext = file.name.split('.').pop().toLowerCase();
  if (ext === 'xlsx' || ext === 'xls') {
    massaHandleXlsxUpload(file);
  } else {
    massaHandleCsvUpload(file);
  }
}

function massaHandleXlsxUpload(file) {
  if (!file) return;
  if (typeof XLSX === 'undefined') {
    toast('Biblioteca XLS não carregada. Verifique a conexão.', 'err');
    return;
  }
  const reader = new FileReader();
  reader.onload = e => {
    try {
      const wb   = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
      const ws   = wb.Sheets[wb.SheetNames[0]];
      const data = XLSX.utils.sheet_to_json(ws, { defval: '' });
      // Normaliza as chaves para lowercase
      const rows = data.map(row => {
        const norm = {};
        Object.keys(row).forEach(k => { norm[k.toLowerCase().trim()] = String(row[k]).trim(); });
        return norm;
      }).filter(r => Object.values(r).some(v => v));
      massaCsvData = rows;
      massaRenderCsvPreview(rows);
      toast(rows.length + ' linha(s) carregada(s) do XLS', 'ok');
    } catch(err) {
      toast('Erro ao ler o arquivo XLS: ' + err.message, 'err');
    }
  };
  reader.readAsArrayBuffer(file);
}

function massaHandleCsvUpload(file) {
  if (!file) return;
  const reader = new FileReader();
  reader.onload = e => {
    let text = e.target.result;
    // Remove BOM if present
    if (text.charCodeAt(0) === 0xFEFF) text = text.slice(1);
    const rows = massaParseCsv(text);
    massaCsvData = rows;
    massaRenderCsvPreview(rows);
  };
  reader.readAsText(file, 'utf-8');
}

function massaParseCsv(text) {
  const lines = text.split(/\r?\n/).filter(l => l.trim());
  if (lines.length < 2) return [];
  const headers = massaParseCsvLine(lines[0]).map(h => h.trim().toLowerCase());
  return lines.slice(1).map(line => {
    const vals = massaParseCsvLine(line);
    const obj = {};
    headers.forEach((h, i) => { obj[h] = (vals[i] || '').trim(); });
    return obj;
  }).filter(r => Object.values(r).some(v => v));
}

function massaParseCsvLine(line) {
  const result = [];
  let cur = '', inQ = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      if (inQ && line[i+1] === '"') { cur += '"'; i++; }
      else inQ = !inQ;
    } else if ((ch === ';' || ch === ',') && !inQ) {
      result.push(cur); cur = '';
    } else {
      cur += ch;
    }
  }
  result.push(cur);
  return result;
}

function massaRenderCsvPreview(rows) {
  const container = document.getElementById('massa-csv-preview');
  if (!rows.length) { container.innerHTML = '<div style="font-size:12px;color:#aaa;">Nenhum dado encontrado no CSV.</div>'; return; }
  const headers = Object.keys(rows[0]);
  const html = `
    <div style="font-size:11px;color:#888;margin-bottom:6px;">${rows.length} linha(s) carregada(s)</div>
    <table class="csv-preview-table">
      <thead><tr>${headers.map(h => `<th>${escHtml(h)}</th>`).join('')}</tr></thead>
      <tbody>
        ${rows.slice(0,10).map(r => `<tr>${headers.map(h => `<td>${escHtml(r[h]||'')}</td>`).join('')}</tr>`).join('')}
        ${rows.length > 10 ? `<tr><td colspan="${headers.length}" style="color:#aaa;text-align:center;"> e mais ${rows.length-10} linha(s)</td></tr>` : ''}
      </tbody>
    </table>
  `;
  container.innerHTML = html;
}


