/* 
   MDULO: CRIAR CAMPO (Tpico 4)
    */

/*  Estado  */
let criarEntityTypeId  = 4;      // entidade selecionada
let criarTipoSelecionado = 'string'; // tipo selecionado
let criarLastCreated   = null;   // { fieldId, entityTypeId, label } do ltimo campo criado
let criarSuggestTimer  = null;   // debounce timer
const CRIAR_COLLAPSE_STORAGE = 'b24panel_cards_collapsed';

const ENTITY_METHODS = {
  4: { add: 'crm.company.userfield.add',  update: 'crm.company.userfield.update',  delete: 'crm.company.userfield.delete',  confGet: 'crm.item.details.configuration.get',  confSet: 'crm.item.details.configuration.set'  },
  3: { add: 'crm.contact.userfield.add',  update: 'crm.contact.userfield.update',  delete: 'crm.contact.userfield.delete',  confGet: 'crm.item.details.configuration.get',  confSet: 'crm.item.details.configuration.set'  },
  2: { add: 'crm.deal.userfield.add',     update: 'crm.deal.userfield.update',     delete: 'crm.deal.userfield.delete',     confGet: 'crm.item.details.configuration.get',  confSet: 'crm.item.details.configuration.set'  },
  1: { add: 'crm.lead.userfield.add',     update: 'crm.lead.userfield.update',     delete: 'crm.lead.userfield.delete',     confGet: 'crm.lead.details.configuration.get',  confSet: 'crm.lead.details.configuration.set'  },
};

const TIPOS = [
  { id: 'string',        name: 'Texto',         desc: 'Linha nica ou multilinha' },
  { id: 'integer',       name: 'Inteiro',        desc: 'Número sem decimais' },
  { id: 'double',        name: 'Decimal',        desc: 'Número com casas decimais' },
  { id: 'boolean',       name: 'Sim / No',      desc: 'Campo verdadeiro/falso' },
  { id: 'date',          name: 'Data',           desc: 'Somente data' },
  { id: 'datetime',      name: 'Data e hora',    desc: 'Data + horrio' },
  { id: 'money',         name: 'Moeda',          desc: 'Valor com smbolo monetrio' },
  { id: 'url',           name: 'URL / Link',     desc: 'Endereço web' },
  { id: 'address',       name: 'Endereço',       desc: 'Campos estruturados de endereo' },
  { id: 'enumeration',   name: 'Lista',          desc: 'Opções pr-definidas' },
  { id: 'file',          name: 'Arquivo',        desc: 'Anexo / upload' },
  { id: 'employee',      name: 'Usuário',        desc: 'Usuário do Bitrix24' },
  { id: 'crm_status',    name: 'Status CRM',     desc: 'Referncia de status/CRM' },
  { id: 'crm',           name: 'Vínculo CRM',    desc: 'Link com outra entidade' },
  { id: 'iblock_section',name: 'Seção IBlock',   desc: 'Seção de bloco de informação' },
  { id: 'iblock_element',name: 'Elem. IBlock',   desc: 'Elemento de bloco de informação' },
];

/* Tipos que tm painel condicional */
const TIPOS_COM_CONF = new Set(['string','double','boolean','enumeration']);

function criarCollapseIcon(collapsed) {
  const common = 'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"';
  return collapsed
    ? `<svg ${common}><path d="m9 18 6-6-6-6"/></svg>`
    : `<svg ${common}><path d="m6 9 6 6 6-6"/></svg>`;
}

function criarLoadCollapseState() {
  try { return JSON.parse(localStorage.getItem(CRIAR_COLLAPSE_STORAGE)) || {}; } catch (_) { return {}; }
}

function criarSaveCollapseState(state) {
  try { localStorage.setItem(CRIAR_COLLAPSE_STORAGE, JSON.stringify(state)); } catch (_) {}
}

function criarSetCardCollapsed(card, collapsed, persist = true) {
  const body = card.querySelector(':scope > .card-collapse-body');
  const btn = card.querySelector(':scope > .card-title .card-collapse-toggle');
  const title = card.querySelector(':scope > .card-title');
  card.classList.toggle('collapsed', collapsed);
  if (body) body.hidden = collapsed;
  if (btn) {
    btn.innerHTML = criarCollapseIcon(collapsed);
    btn.title = collapsed ? 'Expandir caixa' : 'Minimizar caixa';
    btn.setAttribute('aria-label', collapsed ? 'Expandir caixa' : 'Minimizar caixa');
    btn.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
  }
  if (title) title.setAttribute('aria-expanded', collapsed ? 'false' : 'true');
  if (persist && card.dataset.collapseKey) {
    const state = criarLoadCollapseState();
    state[card.dataset.collapseKey] = collapsed;
    criarSaveCollapseState(state);
  }
}

function criarToggleCardCollapse(card) {
  criarSetCardCollapsed(card, !card.classList.contains('collapsed'));
}

function criarInitCollapsibleCards(root = document) {
  const state = criarLoadCollapseState();
  root.querySelectorAll('.panel > .card, .panel > * > .card').forEach((card, idx) => {
    if (card.closest('.metric-grid')) return;
    if (!card.dataset.collapseKey) {
      const panel = card.closest('.panel');
      card.dataset.collapseKey = card.id || `${panel ? panel.id : 'panel'}-card-${idx}`;
    }
    if (!card.classList.contains('collapsible-card')) {
      const title = Array.from(card.children).find(el => el.classList && el.classList.contains('card-title'));
      if (!title) return;
      if (title.querySelector('.hidden-fields-panel-toggle')) return;

      const body = document.createElement('div');
      body.className = 'card-collapse-body';
      while (title.nextSibling) body.appendChild(title.nextSibling);
      card.appendChild(body);

      const toggle = document.createElement('button');
      toggle.type = 'button';
      toggle.className = 'icon-btn section-collapse-btn card-collapse-toggle';
      toggle.onclick = event => {
        event.stopPropagation();
        criarToggleCardCollapse(card);
      };

      title.classList.add('collapsible-card-title');
      title.setAttribute('role', 'button');
      title.setAttribute('tabindex', '0');
      title.appendChild(toggle);
      title.addEventListener('click', event => {
        if (event.target.closest('button, input, select, textarea, a')) return;
        criarToggleCardCollapse(card);
      });
      title.addEventListener('keydown', event => {
        if (event.key !== 'Enter' && event.key !== ' ') return;
        event.preventDefault();
        criarToggleCardCollapse(card);
      });

      card.classList.add('collapsible-card');
    }
    criarSetCardCollapsed(card, !!state[card.dataset.collapseKey], false);
  });
}

function criarMode() {
  return window.EntityContext ? window.EntityContext.getCurrentMode() : 'crm';
}

function criarGetActiveSpa(requireSelection = true) {
  if (!window.EntityContext) return null;
  return requireSelection ? window.EntityContext.requireActiveSpa() : window.EntityContext.getActiveSpa();
}

function criarGetFieldPrefix() {
  if (criarMode() === 'spa') {
    const spa = criarGetActiveSpa(false);
    return spa && spa.id ? `UF_CRM_${spa.id}_` : 'UF_CRM_{SPA}_';
  }
  return 'UF_CRM_';
}

function criarFullFieldName(suffix) {
  return criarGetFieldPrefix().replace('{SPA}', criarGetActiveSpa(true).id) + suffix;
}

function criarSyncContextUI() {
  const isSpa = criarMode() === 'spa';
  const entityGroup = document.getElementById('criar-entity-tabs')?.closest('.form-group');
  if (entityGroup) entityGroup.style.display = isSpa ? 'none' : 'block';

  document.querySelectorAll('#panel-criar .id-prefix').forEach(el => {
    el.textContent = criarGetFieldPrefix();
  });

  const help = document.getElementById('criar-help-msg');
  if (help) {
    help.placeholder = isSpa
      ? 'Texto de ajuda exibido ao lado do campo no card do SPA...'
      : 'Texto de ajuda exibido ao lado do campo no card do CRM...';
  }

  criarOnSuffixInput();
}

/*  Inicializa grid de tipos  */
function criarInitTipos() {
  criarInitCollapsibleCards();
  const grid = document.getElementById('criar-tipo-grid');
  grid.innerHTML = TIPOS.map(t => `
    <div class="tipo-card ${t.id === 'string' ? 'active' : ''}"
         id="tipo-card-${t.id}"
         onclick="criarSelectTipo('${t.id}')">
      <div class="tipo-card-name">${t.name}</div>
      <div class="tipo-card-desc">${t.desc}</div>
    </div>
  `).join('');

  // Inicializa a lista de enum vazia com 2 linhas de exemplo
  enumAddRow('Opção 1');
  enumAddRow('Opção 2');

  criarUpdateCondBlock();
}

/*  Seleciona entidade  */
function criarSelectEntity(entityTypeId, btn) {
  criarEntityTypeId = entityTypeId;
  if (window.EntityContext) window.EntityContext.setCurrentCrmEntity(entityTypeId);
  document.querySelectorAll('#criar-entity-tabs .etab').forEach(t => t.classList.remove('active'));
  btn.classList.add('active');
  // Limpa banner de sucesso ao trocar entidade
  document.getElementById('criar-success-banner').style.display = 'none';
}

/*  Seleciona tipo  */
function criarSelectTipo(tipoId) {
  criarTipoSelecionado = tipoId;
  document.querySelectorAll('.tipo-card').forEach(c => c.classList.remove('active'));
  const card = document.getElementById(`tipo-card-${tipoId}`);
  if (card) card.classList.add('active');
  criarUpdateCondBlock();
}

/*  Mostra/oculta bloco condicional por tipo  */
function criarUpdateCondBlock() {
  const card = document.getElementById('criar-cond-card');
  document.querySelectorAll('.cond-block').forEach(b => b.classList.remove('visible'));

  if (TIPOS_COM_CONF.has(criarTipoSelecionado)) {
    card.style.display = 'block';
    const block = document.getElementById(`cond-${criarTipoSelecionado}`);
    if (block) block.classList.add('visible');
  } else {
    // Tipos sem conf: mostra mensagem genrica
    card.style.display = 'block';
    document.getElementById('cond-none').classList.add('visible');
  }
}

/*  Label  sugesto de ID (debounced 300ms)  */
function criarOnLabelInput() {
  clearTimeout(criarSuggestTimer);
  criarSuggestTimer = setTimeout(criarSuggestId, 300);
}

function criarSuggestId() {
  const label = document.getElementById('criar-label').value;
  if (!label.trim()) return;

  // Normaliza: remove acentos, minsculas  maisculas, espaos  _
  const normalized = label
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // remove acentos
    .toUpperCase()
    .replace(/[^A-Z0-9\s]/g, '')    // s alfanum + espao
    .trim()
    .split(/\s+/)
    .filter(w => !['DE','DO','DA','DAS','DOS','UM','UMA','O','A','OS','AS','E','OU','EM','NO','NA','POR','PARA','COM','SEM'].includes(w))
    .join('_');

  // Trunca em 13 chars, garante que não termina em _
  let suffix = normalized.slice(0, 13).replace(/_+$/, '');

  // Remove __ duplos
  suffix = suffix.replace(/__+/g, '_');

  document.getElementById('criar-suffix').value = suffix;
  criarOnSuffixInput();
}

/*  Validao em tempo real do sufixo  */
function criarOnSuffixInput() {
  const input = document.getElementById('criar-suffix');
  // Fora maisculas e remove chars inválidos imediatamente
  let val = input.value.toUpperCase().replace(/[^A-Z0-9_]/g, '');
  if (val !== input.value) input.value = val;

  const count = document.getElementById('criar-char-count');
  count.textContent = `${val.length} / 13`;
  count.style.color = val.length >= 13 ? '#ef4444' : '#aaa';

  const validEl = document.getElementById('criar-id-validation');
  if (!val) { validEl.style.display = 'none'; return; }

  const erros = [];
  if (/^[0-9]/.test(val))   erros.push('não pode começar com número');
  if (/__/.test(val))        erros.push('não pode ter __ duplo');
  if (val.length < 2)        erros.push('mnimo 2 caracteres');

  validEl.style.display = 'flex';
  if (erros.length === 0) {
    validEl.className = 'id-validation ok';
    let preview = '';
    try {
      preview = criarFullFieldName(val);
    } catch (_) {
      preview = criarGetFieldPrefix() + val;
    }
    validEl.innerHTML = ` ID válido  campo será criado como <strong>${preview}</strong>`;
  } else {
    validEl.className = 'id-validation err';
    validEl.innerHTML = ` ${erros.join('  ')}`;
  }
}

/*  Toggle dica de campo  */
function criarToggleHelp() {
  const on = document.getElementById('criar-help-toggle').checked;
  document.getElementById('criar-help-block').style.display = on ? 'block' : 'none';
  if (on) document.getElementById('criar-help-msg').focus();
}

/*  Editor de opções (enumeration)  */
let enumCounter = 0;

function enumAddRow(valorInicial = '') {
  enumCounter++;
  const id = `enum-row-${enumCounter}`;
  const list = document.getElementById('enum-list');
  const row = document.createElement('div');
  row.className = 'enum-row';
  row.id = id;
  row.innerHTML = `
    <span class="enum-drag" title="Arrastar para reordenar"></span>
    <input class="enum-name-input" placeholder="Nome da opção" value="${escHtml(valorInicial)}" />
    <label class="enum-default-check">
      <input type="radio" name="enum-default" value="${enumCounter}" />
      Padrão
    </label>
    <button class="enum-del-btn" onclick="enumDelRow('${id}')" title="Remover"></button>
  `;
  list.appendChild(row);
}

function enumDelRow(id) {
  const el = document.getElementById(id);
  const list = document.getElementById('enum-list');
  if (list.children.length <= 1) { toast('Enumeration precisa de ao menos uma opção.', 'wn'); return; }
  if (el) el.remove();
}

function enumGetOptions() {
  const rows = document.querySelectorAll('#enum-list .enum-row');
  const options = [];
  let sort = 10;
  rows.forEach((row, i) => {
    const nameInput    = row.querySelector('.enum-name-input');
    const defaultRadio = row.querySelector('input[type=radio]');
    const val = nameInput.value.trim();
    if (!val) return;
    options.push({
      VALUE:  val,
      XML_ID: val.toUpperCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^A-Z0-9_]/g,'_').replace(/__+/g,'_').slice(0,50) || `OPT_${i+1}`,
      SORT:   sort,
      DEF:    defaultRadio && defaultRadio.checked ? 'Y' : 'N',
    });
    sort += 10;
  });
  return options;
}

/*  Monta o payload SETTINGS por tipo  */
function criarGetSettings() {
  switch (criarTipoSelecionado) {
    case 'string': {
      const rows = parseInt(document.getElementById('cond-string-rows').value) || 1;
      return { DEFAULT_VALUE: '', ROWS: rows };
    }
    case 'double': {
      const prec = parseInt(document.getElementById('cond-double-precision').value) || 2;
      return { PRECISION: prec };
    }
    case 'boolean': {
      const def = document.getElementById('cond-boolean-default').value;
      return { DEFAULT_VALUE: def === '0' ? '' : def === '1' ? 'Y' : 'N' };
    }
    default: return {};
  }
}

/*  Validao final antes de criar  */
function criarValidar() {
  const label  = document.getElementById('criar-label').value.trim();
  const suffix = document.getElementById('criar-suffix').value.trim();

  if (!label)  { toast('Informe o label do campo.', 'wn'); return false; }
  if (!suffix) { toast('Informe o ID do campo.', 'wn'); return false; }
  if (/^[0-9]/.test(suffix)) { toast('O ID não pode começar com número.', 'wn'); return false; }
  if (/__/.test(suffix))     { toast('O ID não pode ter __ duplo.', 'wn'); return false; }
  if (suffix.length < 2)     { toast('O ID precisa ter ao menos 2 caracteres.', 'wn'); return false; }
  if (!/^[A-Z0-9_]+$/.test(suffix)) { toast('O ID s pode ter A-Z, 0-9 e _.', 'wn'); return false; }

  if (criarMode() === 'spa') {
    try {
      criarGetActiveSpa(true);
    } catch (e) {
      toast(e.message || 'Selecione um processo inteligente antes de criar o campo.', 'wn', 5000);
      return false;
    }
  }

  if (criarTipoSelecionado === 'enumeration') {
    const opts = enumGetOptions();
    if (opts.length === 0) { toast('Adicione ao menos uma opção à lista.', 'wn'); return false; }
  }

  return true;
}

function criarBuildCrmFields(label, suffix, helpOn, helpMsg) {
  const fields = {
    USER_TYPE_ID:  criarTipoSelecionado,
    FIELD_NAME:    suffix,
    LABEL:         label,
    MULTIPLE:      document.getElementById('criar-multiple').checked  ? 'Y' : 'N',
    MANDATORY:     document.getElementById('criar-mandatory').checked ? 'Y' : 'N',
    SHOW_FILTER:   document.getElementById('criar-filter').checked    ? 'Y' : 'N',
    EDIT_IN_LIST:  'Y',
    SORT:          100,
    SETTINGS:      criarGetSettings(),
  };

  if (helpOn && helpMsg) fields.HELP_MESSAGE = helpMsg;

  if (criarTipoSelecionado === 'enumeration') {
    fields.LIST = enumGetOptions();
    fields.SETTINGS = { DISPLAY: 'UI', LIST_HEIGHT: 1 };
  }

  return fields;
}

function criarBuildSpaField(label, suffix, helpOn, helpMsg) {
  const settings = criarTipoSelecionado === 'enumeration'
    ? { DISPLAY: 'UI', LIST_HEIGHT: 1 }
    : criarGetSettings();

  const field = {
    fieldName: criarFullFieldName(suffix),
    userTypeId: criarTipoSelecionado,
    multiple: document.getElementById('criar-multiple').checked ? 'Y' : 'N',
    mandatory: document.getElementById('criar-mandatory').checked ? 'Y' : 'N',
    showFilter: document.getElementById('criar-filter').checked ? 'Y' : 'N',
    editInList: 'Y',
    sort: 100,
    editFormLabel: { br: label },
    listColumnLabel: { br: label },
    listFilterLabel: { br: label },
    settings,
  };

  if (helpOn && helpMsg) field.helpMessage = { br: helpMsg };

  if (criarTipoSelecionado === 'enumeration') {
    field.enum = enumGetOptions().map(opt => ({
      value: opt.VALUE,
      xmlId: opt.XML_ID,
      sort: opt.SORT,
      def: opt.DEF,
    }));
  }

  return field;
}

function criarExtractApiError(data, fieldId) {
  if (!data || !data.error) return '';
  if (data.error === 'ERROR_FIELD_ALREADY_EXISTS') {
    return `Já existe um campo com o ID ${fieldId} nesta entidade.`;
  }
  return data.error_description || data.error;
}

/*  Cria o campo via API  */
async function criarCampo() {
  if (!criarValidar()) return;

  const btn    = document.getElementById('criar-btn');
  const label  = document.getElementById('criar-label').value.trim();
  const suffix = document.getElementById('criar-suffix').value.trim();
  const helpOn = document.getElementById('criar-help-toggle').checked;
  const helpMsg = document.getElementById('criar-help-msg').value.trim();

  btn.textContent = 'Criando';
  btn.disabled = true;
  document.getElementById('criar-success-banner').style.display = 'none';

  try {
    const isSpa = criarMode() === 'spa';
    const fieldId = isSpa ? criarFullFieldName(suffix) : `UF_CRM_${suffix}`;
    const adapter = window.EntityContext ? window.EntityContext.getCurrentAdapter() : null;
    const data = isSpa
      ? await adapter.addField(criarBuildSpaField(label, suffix, helpOn, helpMsg))
      : await (adapter
        ? adapter.addField(criarBuildCrmFields(label, suffix, helpOn, helpMsg), criarEntityTypeId)
        : call(ENTITY_METHODS[criarEntityTypeId].add, { fields: criarBuildCrmFields(label, suffix, helpOn, helpMsg) }));

    if (data.error) {
      throw new Error(criarExtractApiError(data, fieldId));
    }

    const createdEntityTypeId = isSpa ? criarGetActiveSpa(true).entityTypeId : criarEntityTypeId;
    criarLastCreated = { fieldId, entityTypeId: createdEntityTypeId, label, mode: isSpa ? 'spa' : 'crm' };

    // Mostra banner de sucesso
    document.getElementById('criar-success-title').textContent = `"${label}" criado com sucesso!`;
    document.getElementById('criar-success-id').textContent    = fieldId;
    document.getElementById('criar-success-banner').style.display = 'block';
    document.getElementById('criar-success-banner').scrollIntoView({ behavior: 'smooth', block: 'nearest' });

    toast(`Campo ${fieldId} criado!`, 'ok');

  } catch (e) {
    toast('Erro ao criar campo: ' + e.message, 'er', 7000);
  } finally {
    btn.textContent = 'Criar campo ';
    btn.disabled = false;
  }
}

/*  Adicionar campo ao card aps criar  */
async function criarAddToCard() {
  if (!criarLastCreated) return;

  const { fieldId, entityTypeId, mode } = criarLastCreated;
  const addBtn = document.getElementById('criar-add-card-btn');
  addBtn.textContent = 'Adicionando';
  addBtn.disabled = true;

  try {
    const isSpa = mode === 'spa';
    const methods = isSpa
      ? { confGet: 'crm.item.details.configuration.get', confSet: 'crm.item.details.configuration.set' }
      : ENTITY_METHODS[entityTypeId];

    // Para Lead, usa método especfico; para outros, usa item.details
    const getParams = { entityTypeId };
    if (!isSpa && entityTypeId === 1) {
      // Lead: scope C
      getParams.scope = 'C';
    }

    const confData = await call(methods.confGet, getParams);

    if (confData.error) throw new Error(confData.error_description || confData.error);

    // A configuração vem em result (array de seções)
    let sections = confData.result || [];

    // Se vier como objeto aninhado (algumas verses da API), tenta extrair
    if (!Array.isArray(sections) && sections.data) sections = sections.data;

    if (!Array.isArray(sections) || sections.length === 0) {
      // Configuração vazia - cria uma seção padrão
      sections = [{ name: 'main', title: 'Informaes', type: 'section', elements: [] }];
    }

    // Verifica se o campo já está em alguma seção
    const jaExiste = sections.some(sec =>
      Array.isArray(sec.elements) && sec.elements.some(el => el.name === fieldId)
    );

    if (jaExiste) {
      toast(`O campo ${fieldId} já está visível no card.`, 'wn');
      addBtn.textContent = 'Já está no card ';
      return;
    }

    // Adiciona o campo na primeira seção com optionFlags: 0 (recolhível)
    sections[0].elements = sections[0].elements || [];
    sections[0].elements.push({ name: fieldId, optionFlags: 0 });

    // Salva
    const setParams = { entityTypeId, scope: 'C', data: sections };
    const setData = await call(methods.confSet, setParams);

    if (setData.error) throw new Error(setData.error_description || setData.error);

    toast(`${fieldId} adicionado ao card (visível para todos).`, 'ok');
    addBtn.textContent = 'Adicionado ao card ';
    addBtn.disabled = true;

  } catch (e) {
    toast(`Não foi possível adicionar ao card: ${e.message}. Use o módulo Config. do Card para ajuste manual.`, 'er', 8000);
    addBtn.textContent = '+ Adicionar ao card (visível)';
    addBtn.disabled = false;
  }
}

/*  Reset do formulrio  */
function criarReset() {
  document.getElementById('criar-label').value  = '';
  document.getElementById('criar-suffix').value = '';
  document.getElementById('criar-char-count').textContent = '0 / 13';
  document.getElementById('criar-id-validation').style.display = 'none';
  document.getElementById('criar-mandatory').checked = false;
  document.getElementById('criar-multiple').checked  = false;
  document.getElementById('criar-filter').checked    = false;
  document.getElementById('criar-help-toggle').checked = false;
  document.getElementById('criar-help-block').style.display = 'none';
  document.getElementById('criar-help-msg').value = '';
  document.getElementById('criar-success-banner').style.display = 'none';
  document.getElementById('cond-string-rows').value = '1';
  document.getElementById('cond-double-precision').value = '2';
  document.getElementById('cond-boolean-default').value  = '0';
  // Reseta enum list
  document.getElementById('enum-list').innerHTML = '';
  enumCounter = 0;
  enumAddRow('Opção 1');
  enumAddRow('Opção 2');
  criarLastCreated = null;
  toast('Formulrio limpo.', 'wn', 2000);
}
