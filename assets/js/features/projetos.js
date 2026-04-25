/* MODULO: PROJETOS */

let projetosList = [];
let projSelectedId = null;
let projSelectedName = '';
let projCurrentTab = 'projetos';
const projCheckedIds = new Set();
let projImportPlan = null;
const PROJ_IMPORT_SHEETS = [
  {
    key: 'projects',
    title: 'Projetos',
    columns: [
      ['chave_projeto', 'Chave'],
      ['nome_projeto', 'Nome'],
      ['descricao', 'Descricao'],
      ['visibilidade', 'Visibilidade'],
      ['adesao', 'Adesao'],
      ['data_inicio', 'Data inicio'],
      ['data_termino', 'Data termino'],
      ['permissao_convite', 'Permissao convite'],
    ],
  },
  {
    key: 'members',
    title: 'Membros',
    columns: [
      ['chave_projeto', 'Chave projeto'],
      ['nome', 'Nome, e-mail ou ID'],
      ['papel', 'Papel'],
    ],
  },
  {
    key: 'stages',
    title: 'Etapas Kanban',
    columns: [
      ['chave_projeto', 'Chave projeto'],
      ['chave_etapa', 'Chave etapa'],
      ['nome_etapa', 'Nome etapa'],
      ['depois_de_etapa', 'Depois de'],
      ['cor', 'Cor'],
    ],
  },
];
let projKanbanStages = [];
let projMembers = [];
let projEditingStageId = null;
const projUserCache = new Map();
let projAllUsers = [];
let projUsersLoadedFor = '';
let projUsersLoading = false;
const projNewInvites = [];
let projCurrentPage = 1;
let projLastQuery = '';
const projPageSize = 25;

function projGetId(proj) {
  return proj && (proj.ID ?? proj.id);
}

function projGetName(proj) {
  return proj && (proj.NAME || proj.name || '');
}

function projGetById(id) {
  return projetosList.find(p => String(projGetId(p)) === String(id)) || null;
}

function projNormalizeSearch(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .trim();
}

function projSearchHaystack(proj) {
  const id = projGetId(proj);
  const ownerId = projOwnerId(proj);
  return projNormalizeSearch([
    projGetName(proj),
    proj.DESCRIPTION || proj.description,
    id ? `ID ${id}` : '',
    id,
    ownerId ? `OWNER ${ownerId}` : '',
    ownerId ? projUserNameById(ownerId) : '',
  ].filter(Boolean).join(' '));
}

function projOwnerId(proj) {
  return proj && (proj.OWNER_ID || proj.ownerId || proj.OWNER || proj.owner || '');
}

function projUserDisplayName(user) {
  if (!user) return '';
  const first = user.NAME || user.name || '';
  const last = user.LAST_NAME || user.lastName || '';
  const full = [first, last].filter(Boolean).join(' ').trim();
  return user.FULL_NAME || user.fullName || full || user.LOGIN || user.login || user.EMAIL || user.email || '';
}

function projUserNameById(id) {
  const key = String(id || '');
  if (!key) return '-';
  const cached = projUserCache.get(key);
  return cached || `Usuario ${key}`;
}

function projUserId(user) {
  return user && (user.ID || user.id || user.USER_ID || user.userId);
}

function projUserEmail(user) {
  return user && (user.EMAIL || user.email || user.LOGIN || user.login || '');
}

function projUserOptionLabel(user) {
  const id = projUserId(user);
  const name = projUserDisplayName(user) || `Usuario ${id}`;
  const email = projUserEmail(user);
  return email && email !== name ? `${name} - ${email}` : name;
}

function projUserLookupLabel(user) {
  const id = projUserId(user);
  return `${projUserOptionLabel(user)} (ID: ${id})`;
}

function projCurrentEnvKey() {
  const env = envs && envs[activeIdx];
  return env && env.url ? env.url : '';
}

function projSortUsers(users) {
  return [...users].sort((a, b) => projUserOptionLabel(a).localeCompare(projUserOptionLabel(b), 'pt-BR'));
}

function projSetUserSelectState(message) {
  ['proj-new-owner', 'proj-invite-user', 'proj-edit-owner'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.innerHTML = `<option value="">${escHtml(message)}</option>`;
  });
  const memberInput = document.getElementById('proj-member-user');
  if (memberInput) {
    memberInput.placeholder = message;
    memberInput.disabled = projUsersLoading;
  }
  const memberList = document.getElementById('proj-member-users');
  if (memberList) memberList.innerHTML = '';
}

function projRenderUserSelects() {
  const owner = document.getElementById('proj-new-owner');
  const invite = document.getElementById('proj-invite-user');
  const editOwner = document.getElementById('proj-edit-owner');
  const memberInput = document.getElementById('proj-member-user');
  const memberList = document.getElementById('proj-member-users');
  if (!owner && !invite && !editOwner && !memberInput && !memberList) return;
  const options = projAllUsers.map(user => {
    const id = projUserId(user);
    return `<option value="${escHtml(id)}">${escHtml(projUserOptionLabel(user))} (ID: ${escHtml(id)})</option>`;
  }).join('');
  if (owner) owner.innerHTML = `<option value="">Selecione o proprietario</option>${options}`;
  if (invite) invite.innerHTML = `<option value="">Selecione uma pessoa</option>${options}`;
  if (editOwner) editOwner.innerHTML = `<option value="">Manter proprietario atual</option>${options}`;
  if (memberList) {
    memberList.innerHTML = projAllUsers.map(user => {
      const id = projUserId(user);
      return `<option value="${escHtml(projUserLookupLabel(user))}" data-user-id="${escHtml(id)}"></option>`;
    }).join('');
  }
  if (memberInput) {
    memberInput.disabled = false;
    memberInput.placeholder = projAllUsers.length ? 'Digite o nome para buscar...' : 'Nenhum usuario encontrado';
  }
}

async function projLoadUsers(force = false) {
  const envKey = projCurrentEnvKey();
  if (!envKey) {
    projSetUserSelectState('Configure um ambiente primeiro');
    return;
  }
  if (!force && projUsersLoadedFor === envKey && projAllUsers.length) {
    projRenderUserSelects();
    return;
  }
  if (projUsersLoading) return;
  projUsersLoading = true;
  projSetUserSelectState('Carregando usuarios...');
  try {
    const users = [];
    let start = 0;
    let guard = 0;
    do {
      const data = await call('user.get', { FILTER: { ACTIVE: 'Y' }, SORT: 'LAST_NAME', ORDER: 'ASC', start });
      if (data.error) throw new Error(data.error_description || data.error);
      users.push(...projExtractList(data));
      start = data && data.next !== undefined && data.next !== null ? Number(data.next) : null;
      guard += 1;
    } while (start !== null && Number.isFinite(start) && guard < 100);
    projAllUsers = projSortUsers(users.filter(user => projUserId(user)));
    projAllUsers.forEach(user => {
      const id = projUserId(user);
      projUserCache.set(String(id), projUserDisplayName(user) || `Usuario ${id}`);
    });
    projUsersLoadedFor = envKey;
    projRenderUserSelects();
  } catch (e) {
    projSetUserSelectState('Erro ao carregar usuarios');
    toast('Erro ao carregar usuarios: ' + e.message, 'er');
  } finally {
    projUsersLoading = false;
  }
}

function projNormalizeHex(value, fallback = '#59abed') {
  const raw = String(value || '').trim();
  const hex = raw.startsWith('#') ? raw : `#${raw}`;
  return /^#[0-9a-f]{6}$/i.test(hex) ? hex : fallback;
}

function projApiColor(value, fallback = '#59abed') {
  return projNormalizeHex(value, fallback).replace('#', '').toUpperCase();
}

function projIsHexColor(value) {
  const raw = String(value || '').trim();
  if (!raw) return false;
  const hex = raw.startsWith('#') ? raw : `#${raw}`;
  return /^#[0-9a-f]{6}$/i.test(hex);
}

function projStageId(stage) {
  return stage && (stage.ID ?? stage.id);
}

function projStageSort(stage) {
  return parseInt(stage?.SORT ?? stage?.sort ?? 0, 10) || 0;
}

function projSortStages(stages) {
  return [...(stages || [])].sort((a, b) => {
    const sortDiff = projStageSort(a) - projStageSort(b);
    if (sortDiff) return sortDiff;
    return String(projStageId(a) || '').localeCompare(String(projStageId(b) || ''), undefined, { numeric: true });
  });
}

async function projResolveUsers(ids) {
  const missing = [...new Set((ids || []).map(id => String(id || '').trim()).filter(Boolean))]
    .filter(id => !projUserCache.has(id));
  if (!missing.length) return;
  try {
    const data = await call('user.get', { FILTER: { ID: missing } });
    if (data.error) throw new Error(data.error_description || data.error);
    const users = projExtractList(data);
    users.forEach(user => {
      const id = user.ID || user.id;
      if (id) projUserCache.set(String(id), projUserDisplayName(user) || `Usuario ${id}`);
    });
  } catch (_) {
    /* A chamada individual abaixo tambem cobre portais que nao aceitam lista no filtro. */
  }
  const unresolved = missing.filter(id => !projUserCache.has(id));
  await Promise.all(unresolved.map(async id => {
    try {
      const data = await call('user.get', { FILTER: { ID: id } });
      if (data.error) throw new Error(data.error_description || data.error);
      const user = projExtractList(data)[0];
      projUserCache.set(String(id), projUserDisplayName(user) || `Usuario ${id}`);
    } catch (e) {
      projUserCache.set(String(id), `Usuario ${id}`);
    }
  }));
  missing.forEach(id => {
    if (!projUserCache.has(id)) projUserCache.set(id, `Usuario ${id}`);
  });
}

function projJsArg(value) {
  return escHtml(JSON.stringify(String(value ?? '')));
}

function projActionIcon(name) {
  const common = 'viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true" focusable="false"';
  const icons = {
    edit: `<svg ${common}><path d="M12 20h9"/><path d="M16.5 3.5a2.12 2.12 0 0 1 3 3L7 19l-4 1 1-4Z"/></svg>`,
    delete: `<svg ${common}><path d="M3 6h18"/><path d="M8 6V4h8v2"/><path d="M19 6l-1 14H6L5 6"/><path d="M10 11v5"/><path d="M14 11v5"/></svg>`,
  };
  return icons[name] || '';
}

function projOpenModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.add('open');
  if (id === 'proj-modal-import') projBindImportDrop();
}

function projCloseModal(id) {
  const el = document.getElementById(id);
  if (el) el.classList.remove('open');
}

function projDateOnly(value) {
  if (!value) return '';
  return String(value).slice(0, 10);
}

function projDateToApi(value) {
  return value ? `${value}T00:00:00` : undefined;
}

function projRoleLabel(role) {
  const r = String(role || '').toUpperCase();
  if (r === 'A' || r === 'OWNER') return { text: 'Proprietario', cls: 'owner' };
  if (r === 'E' || r === 'MODERATOR') return { text: 'Moderador', cls: 'mod' };
  return { text: 'Participante', cls: 'member' };
}

function projInviteRoleLabel(role) {
  return String(role || '').toUpperCase() === 'E'
    ? { text: 'Moderador', cls: 'mod' }
    : { text: 'Equipe', cls: 'member' };
}

function projRenderInvites() {
  const list = document.getElementById('proj-invite-list');
  if (!list) return;
  if (!projNewInvites.length) {
    list.innerHTML = '<div class="proj-invite-empty">Nenhuma pessoa selecionada.</div>';
    return;
  }
  list.innerHTML = projNewInvites.map(invite => {
    const role = projInviteRoleLabel(invite.role);
    return `<div class="proj-invite-pill">
      <span>${escHtml(invite.name)} <span style="color:var(--muted);font-weight:500;">ID: ${escHtml(invite.id)}</span></span>
      <span class="member-role ${role.cls}">${escHtml(role.text)}</span>
      <button type="button" title="Remover" aria-label="Remover convite" onclick="projRemoveInvite(${projJsArg(invite.id)})">x</button>
    </div>`;
  }).join('');
}

function projAddInvite() {
  const userSelect = document.getElementById('proj-invite-user');
  const roleSelect = document.getElementById('proj-invite-role');
  const id = String(userSelect?.value || '').trim();
  const role = roleSelect?.value || 'K';
  if (!id) {
    toast('Selecione uma pessoa para convidar.', 'wn');
    return;
  }
  const ownerId = String(document.getElementById('proj-new-owner')?.value || '').trim();
  if (ownerId && ownerId === id) {
    toast('O proprietario ja entra no projeto como dono.', 'wn');
    return;
  }
  const existing = projNewInvites.find(invite => String(invite.id) === id);
  const selectedText = userSelect.selectedOptions && userSelect.selectedOptions[0]
    ? userSelect.selectedOptions[0].textContent.replace(/\s+\(ID:\s*\d+\)\s*$/, '')
    : projUserNameById(id);
  if (existing) {
    existing.role = role;
    existing.name = selectedText;
  } else {
    projNewInvites.push({ id, role, name: selectedText });
  }
  userSelect.value = '';
  projRenderInvites();
}

function projRemoveInvite(id) {
  const index = projNewInvites.findIndex(invite => String(invite.id) === String(id));
  if (index >= 0) projNewInvites.splice(index, 1);
  projRenderInvites();
}

function projClearInvites() {
  projNewInvites.splice(0, projNewInvites.length);
  projRenderInvites();
}

function projExtractList(data) {
  const result = data && data.result;
  if (Array.isArray(result)) return result;
  if (result && Array.isArray(result.items)) return result.items;
  if (result && Array.isArray(result.groups)) return result.groups;
  if (result && typeof result === 'object') return Object.values(result);
  return [];
}

function projSyncContextUI() {
  if (document.getElementById('proj-new-owner')) projLoadUsers();
  projRenderInvites();
  const proj = projSelectedId ? projGetById(projSelectedId) : null;
  if (proj) projSelectedName = projGetName(proj);
  document.getElementById('proj-tab-count')?.replaceChildren(document.createTextNode(String(projetosList.length)));
  document.getElementById('proj-kanban-count')?.replaceChildren(document.createTextNode(String(projKanbanStages.length)));
  document.getElementById('proj-members-count')?.replaceChildren(document.createTextNode(String(projMembers.length)));
  const label = projSelectedId ? `${projSelectedName || 'Projeto'} (ID: ${projSelectedId})` : 'Selecione um projeto';
  const kt = document.getElementById('proj-kanban-title');
  const mt = document.getElementById('proj-members-title');
  if (kt) kt.textContent = label;
  if (mt) mt.textContent = label;
  if (!projSelectedId) {
    const body = document.getElementById('proj-kanban-tbody');
    const members = document.getElementById('proj-members-list');
    if (body) body.innerHTML = '<tr><td colspan="5">Selecione um projeto na aba Projetos para gerenciar o Kanban.</td></tr>';
    if (members) members.innerHTML = '<div class="proj-empty">Selecione um projeto na aba Projetos para gerenciar membros.</div>';
  }
}

function projSwitchTab(tabId, element) {
  projCurrentTab = tabId;
  document.querySelectorAll('#panel-projetos .main-tab').forEach(t => t.classList.remove('active'));
  document.querySelectorAll('#panel-projetos .tab-panel').forEach(p => p.classList.remove('active'));
  if (element) element.classList.add('active');
  const panel = document.getElementById(`tab-${tabId}`);
  if (panel) panel.classList.add('active');
  projSyncContextUI();
  if (tabId === 'kanban' && projSelectedId) projLoadKanban();
  if (tabId === 'membros' && projSelectedId) projLoadMembers();
}

function projToggleCollapse() {
  const list = document.getElementById('proj-list');
  const btn = document.getElementById('proj-collapse-btn');
  if (!list) return;
  const collapsed = list.style.display !== 'none';
  list.style.display = collapsed ? 'none' : 'flex';
  if (btn) btn.classList.toggle('collapsed', collapsed);
}

async function projLoad() {
  const list = document.getElementById('proj-list');
  if (list) list.innerHTML = '<div style="font-size:12px;color:#aaa;">Carregando...</div>';
  try {
    const allProjects = [];
    let start = 0;
    let guard = 0;
    do {
      const data = await call('sonet_group.get', { FILTER: { PROJECT: 'Y' }, ORDER: { NAME: 'ASC' }, start });
      if (data.error) throw new Error(data.error_description || data.error);
      allProjects.push(...projExtractList(data));
      if (list) list.innerHTML = `<div style="font-size:12px;color:#aaa;">Carregando... ${allProjects.length} projeto(s)</div>`;
      start = data && data.next !== undefined && data.next !== null ? Number(data.next) : null;
      guard += 1;
    } while (start !== null && Number.isFinite(start) && guard < 100);
    projetosList = allProjects;
    projCurrentPage = 1;
    const validIds = new Set(projetosList.map(p => String(projGetId(p))));
    [...projCheckedIds].forEach(id => { if (!validIds.has(id)) projCheckedIds.delete(id); });
    if (projSelectedId && !validIds.has(String(projSelectedId))) {
      projSelectedId = null;
      projSelectedName = '';
      projKanbanStages = [];
      projMembers = [];
    }
    await projResolveUsers(projetosList.map(projOwnerId));
    projRenderList();
  } catch (e) {
    if (list) list.innerHTML = `<div style="font-size:12px;color:#ef4444;">Erro: ${escHtml(e.message)}</div>`;
  }
}

function projRenderList() {
  const list = document.getElementById('proj-list');
  if (!list) return;
  const query = projNormalizeSearch(document.getElementById('proj-search')?.value || '');
  if (query !== projLastQuery) {
    projCurrentPage = 1;
    projLastQuery = query;
  }
  const filtered = query ? projetosList.filter(p => projSearchHaystack(p).includes(query)) : projetosList;
  const totalPages = Math.max(1, Math.ceil(filtered.length / projPageSize));
  projCurrentPage = Math.min(Math.max(1, projCurrentPage), totalPages);
  const pageStart = (projCurrentPage - 1) * projPageSize;
  const pageItems = filtered.slice(pageStart, pageStart + projPageSize);
  const count = document.getElementById('proj-search-count');
  if (count) count.textContent = `${filtered.length} projeto${filtered.length === 1 ? '' : 's'}`;
  projRenderPager(filtered.length, pageStart, pageItems.length, totalPages);
  if (!filtered.length) {
    list.innerHTML = '<div style="font-size:12px;color:#aaa;">Nenhum projeto encontrado. Crie um novo projeto abaixo.</div>';
    projUpdateSelectionBar();
    projSyncContextUI();
    return;
  }
  list.innerHTML = pageItems.map(proj => {
    const id = projGetId(proj);
    const name = projGetName(proj);
    const desc = proj.DESCRIPTION || proj.description || '';
    const ownerId = projOwnerId(proj);
    const owner = ownerId ? `${projUserNameById(ownerId)} (ID: ${ownerId})` : '-';
    const closed = String(proj.CLOSED || proj.closed || 'N') === 'Y';
    const privateBadge = String(proj.VISIBLE || proj.visible || 'Y') === 'N' ? '<span class="proj-badge-private">Privado</span>' : '';
    const active = String(id) === String(projSelectedId) ? ' active' : '';
    const selected = projCheckedIds.has(String(id)) ? ' selected' : '';
    const metaBits = [`Proprietario: ${owner}`];
    if (proj.PROJECT_DATE_START || proj.projectDateStart) metaBits.push(`Inicio: ${projDateOnly(proj.PROJECT_DATE_START || proj.projectDateStart)}`);
    if (proj.PROJECT_DATE_FINISH || proj.projectDateFinish) metaBits.push(`Termino: ${projDateOnly(proj.PROJECT_DATE_FINISH || proj.projectDateFinish)}`);
    if (desc) metaBits.push(String(desc).slice(0, 120));
    return `
      <div class="proj-row${active}${selected}" id="proj-row-${escHtml(id)}">
        <div class="proj-check"><input type="checkbox" class="proj-cb" ${projCheckedIds.has(String(id)) ? 'checked' : ''} onchange="projToggleChecked(${projJsArg(id)},this.checked)" aria-label="Selecionar projeto" /></div>
        <div class="proj-arrow" onclick="projOpenContext(${projJsArg(id)},${projJsArg(name)})" title="Usar como contexto">&#9656;</div>
        <div class="proj-body">
          <div class="proj-name">${escHtml(name || '(sem nome)')}</div>
          <div class="proj-meta">${escHtml(metaBits.join(' - '))}</div>
          <div class="proj-badges">
            <span class="${closed ? 'proj-badge-archived' : 'proj-badge-active'}">${closed ? 'Arquivado' : 'Ativo'}</span>
            ${privateBadge}
            <span class="proj-badge-id">ID: ${escHtml(id)}</span>
          </div>
        </div>
        <div class="proj-actions">
          <button class="icon-btn" title="Editar" aria-label="Editar" onclick="projOpenEdit(${projJsArg(id)})">${projActionIcon('edit')}</button>
          <button class="icon-btn del" title="Excluir" aria-label="Excluir" onclick="projDelete(${projJsArg(id)},${projJsArg(name)})">${projActionIcon('delete')}</button>
        </div>
      </div>
    `;
  }).join('');
  projUpdateSelectionBar();
  projSyncContextUI();
}

function projRenderPager(total, pageStart, pageCount, totalPages) {
  const pager = document.getElementById('proj-pager');
  if (!pager) return;
  if (total <= projPageSize) {
    pager.innerHTML = '';
    pager.classList.remove('show');
    return;
  }
  const from = pageStart + 1;
  const to = pageStart + pageCount;
  pager.classList.add('show');
  pager.innerHTML = `
    <button class="proj-page-btn" type="button" onclick="projPagePrev()" ${projCurrentPage <= 1 ? 'disabled' : ''} aria-label="Pagina anterior">&lt;</button>
    <span class="proj-page-info">${from}-${to} de ${total} | Pagina ${projCurrentPage}/${totalPages}</span>
    <button class="proj-page-btn" type="button" onclick="projPageNext()" ${projCurrentPage >= totalPages ? 'disabled' : ''} aria-label="Proxima pagina">&gt;</button>
  `;
}

function projSetPage(page) {
  projCurrentPage = Number(page) || 1;
  projRenderList();
}

function projPagePrev() {
  projSetPage(projCurrentPage - 1);
}

function projPageNext() {
  projSetPage(projCurrentPage + 1);
}

function projToggleChecked(id, checked) {
  const key = String(id);
  if (checked) projCheckedIds.add(key);
  else projCheckedIds.delete(key);
  document.getElementById(`proj-row-${key}`)?.classList.toggle('selected', checked);
  projUpdateSelectionBar();
}

function projUpdateSelectionBar() {
  const bar = document.getElementById('proj-sel-bar');
  const count = document.getElementById('proj-sel-count');
  const n = projCheckedIds.size;
  if (count) count.textContent = String(n);
  if (bar) bar.classList.toggle('show', n > 0);
}

function projSelectAll() {
  if (!projetosList.length) {
    toast('Carregue os projetos antes de selecionar todos.', 'wn');
    return;
  }
  projetosList.forEach(proj => {
    const id = projGetId(proj);
    if (id !== null && id !== undefined) projCheckedIds.add(String(id));
  });
  document.querySelectorAll('.proj-cb').forEach(cb => { cb.checked = true; });
  document.querySelectorAll('.proj-row').forEach(row => row.classList.add('selected'));
  projUpdateSelectionBar();
}

function projClearSelection() {
  projCheckedIds.clear();
  document.querySelectorAll('.proj-cb').forEach(cb => { cb.checked = false; });
  document.querySelectorAll('.proj-row.selected').forEach(row => row.classList.remove('selected'));
  projUpdateSelectionBar();
}

function projOpenContext(id, name) {
  projSelectedId = id;
  projSelectedName = name || '';
  projKanbanStages = [];
  projMembers = [];
  projRenderList();
  projSyncContextUI();
  if (projCurrentTab === 'kanban') projLoadKanban();
  if (projCurrentTab === 'membros') projLoadMembers();
}

function projCreateFields(prefix) {
  const name = document.getElementById(`${prefix}-name`)?.value.trim();
  const fields = {
    NAME: name,
    DESCRIPTION: document.getElementById(`${prefix}-desc`)?.value.trim() || '',
    VISIBLE: document.getElementById(`${prefix}-visible`)?.checked ? 'Y' : 'N',
    OPENED: document.getElementById(`${prefix}-opened`)?.checked ? 'Y' : 'N',
    PROJECT_DATE_START: projDateToApi(document.getElementById(`${prefix}-start`)?.value),
    PROJECT_DATE_FINISH: projDateToApi(document.getElementById(`${prefix}-finish`)?.value),
  };
  Object.keys(fields).forEach(k => fields[k] === undefined && delete fields[k]);
  return fields;
}

async function projCreate() {
  const fields = projCreateFields('proj-new');
  if (!fields.NAME) {
    toast('Informe o nome do projeto.', 'wn');
    return;
  }
  fields.PROJECT = 'Y';
  fields.INITIATE_PERMS = document.getElementById('proj-new-perms')?.value || 'A';
  const owner = document.getElementById('proj-new-owner')?.value;
  if (owner) fields.OWNER_ID = owner;
  try {
    const data = await call('sonet_group.create', fields);
    if (data.error) throw new Error(data.error_description || data.error);
    const groupId = data.result && (data.result.ID || data.result.id) ? (data.result.ID || data.result.id) : data.result;
    if (groupId && projNewInvites.length) {
      for (const invite of projNewInvites) {
        const add = await call('sonet_group.user.add', { GROUP_ID: groupId, USER_ID: invite.id });
        if (add.error) throw new Error(add.error_description || add.error);
        if (invite.role === 'E') {
          const up = await call('sonet_group.user.update', { GROUP_ID: groupId, USER_ID: invite.id, ROLE: 'E' });
          if (up.error) throw new Error(up.error_description || up.error);
        }
      }
    }
    toast(`Projeto "${escHtml(fields.NAME)}" criado.`, 'ok');
    ['name', 'owner', 'desc', 'start', 'finish'].forEach(s => { const el = document.getElementById(`proj-new-${s}`); if (el) el.value = ''; });
    projClearInvites();
    projLoad();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function projOpenEdit(id) {
  const proj = projGetById(id);
  if (!proj) return;
  document.getElementById('proj-edit-id').value = id;
  document.getElementById('proj-edit-name').value = projGetName(proj);
  document.getElementById('proj-edit-desc').value = proj.DESCRIPTION || '';
  document.getElementById('proj-edit-start').value = projDateOnly(proj.PROJECT_DATE_START || '');
  document.getElementById('proj-edit-finish').value = projDateOnly(proj.PROJECT_DATE_FINISH || '');
  document.getElementById('proj-edit-visible').checked = String(proj.VISIBLE || 'Y') !== 'N';
  document.getElementById('proj-edit-opened').checked = String(proj.OPENED || 'N') === 'Y';
  document.getElementById('proj-edit-owner').value = '';
  projOpenModal('proj-modal-edit');
}

async function projSaveEdit() {
  const id = document.getElementById('proj-edit-id')?.value;
  const fields = projCreateFields('proj-edit');
  if (!id || !fields.NAME) {
    toast('Informe o nome do projeto.', 'wn');
    return;
  }
  try {
    const data = await call('sonet_group.update', { GROUP_ID: id, ...fields });
    if (data.error) throw new Error(data.error_description || data.error);
    const owner = document.getElementById('proj-edit-owner')?.value;
    if (owner) {
      const ownerData = await call('sonet_group.setowner', { GROUP_ID: id, USER_ID: owner });
      if (ownerData.error) throw new Error(ownerData.error_description || ownerData.error);
    }
    toast('Projeto atualizado.', 'ok');
    projCloseModal('proj-modal-edit');
    projLoad();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projDelete(id, name) {
  if (!confirm(`Excluir projeto "${name}"? Esta acao nao pode ser desfeita.`)) return;
  try {
    const data = await call('sonet_group.delete', { GROUP_ID: id });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Projeto excluido.', 'wn');
    if (String(projSelectedId) === String(id)) {
      projSelectedId = null;
      projSelectedName = '';
    }
    projCheckedIds.delete(String(id));
    projLoad();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projDeleteSelected() {
  const ids = [...projCheckedIds];
  if (!ids.length) return;
  if (!confirm(`Excluir ${ids.length} projeto(s) selecionado(s)?`)) return;
  for (const id of ids) {
    const data = await call('sonet_group.delete', { GROUP_ID: id });
    if (data.error) {
      toast(`Erro ao excluir ${id}: ${data.error_description || data.error}`, 'er');
      break;
    }
    projCheckedIds.delete(String(id));
  }
  projLoad();
}

async function projLoadKanban() {
  const body = document.getElementById('proj-kanban-tbody');
  if (!projSelectedId) {
    if (body) body.innerHTML = '<tr><td colspan="5">Selecione um projeto na aba Projetos para gerenciar o Kanban.</td></tr>';
    return;
  }
  if (body) body.innerHTML = '<tr><td colspan="5">Carregando...</td></tr>';
  try {
    const data = await call('task.stages.get', { entityId: projSelectedId });
    if (data.error) throw new Error(data.error_description || data.error);
    projKanbanStages = projExtractList(data);
    projRenderKanban();
  } catch (e) {
    if (body) body.innerHTML = `<tr><td colspan="5" style="color:#ef4444;">Erro: ${escHtml(e.message)}</td></tr>`;
  }
  projSyncContextUI();
}

function projStageType(st) {
  return String(st.SYSTEM_TYPE || st.systemType || st.TYPE || st.type || 'CUSTOM').toUpperCase();
}

function projStageTitle(st) {
  return st && (st.TITLE || st.title || st.NAME || st.name || '');
}

function projRenderKanban() {
  const body = document.getElementById('proj-kanban-tbody');
  if (!body) return;
  if (!projKanbanStages.length) {
    body.innerHTML = '<tr><td colspan="5">Nenhuma etapa encontrada.</td></tr>';
    return;
  }
  body.innerHTML = projKanbanStages.map(st => {
    const id = st.ID ?? st.id;
    const title = st.TITLE || st.title || st.NAME || st.name || '-';
    const color = projNormalizeHex(st.COLOR || st.color, '#cccccc');
    const type = projStageType(st);
    const sort = st.SORT || st.sort || '-';
    if (String(id) === String(projEditingStageId)) {
      return `<tr class="proj-stage-edit-row">
        <td><input id="proj-stage-edit-title-${escHtml(id)}" class="form-input" value="${escHtml(title)}" /></td>
        <td><span class="stage-badge WORK">${escHtml(type || 'CUSTOM')}</span></td>
        <td style="font-family:monospace;font-size:10px;color:#aaa;">${escHtml(id)}</td>
        <td>${escHtml(sort)}</td>
        <td>
          <div class="proj-stage-edit-actions">
            <label class="proj-stage-color-edit" title="Editar cor da etapa">
              <span class="stage-dot" style="background:${escHtml(color)};"></span>
              <input id="proj-stage-edit-color-${escHtml(id)}" type="color" value="${escHtml(color)}" oninput="this.previousElementSibling.style.background=this.value" />
            </label>
            <button class="icon-btn" title="Salvar" onclick="projSaveStage(${projJsArg(id)})">OK</button>
            <button class="icon-btn" title="Cancelar" onclick="projCancelEditStage()">x</button>
          </div>
        </td>
      </tr>`;
    }
    return `<tr>
      <td><span class="stage-dot" style="background:${escHtml(color)};"></span>${escHtml(title)}</td>
      <td><span class="stage-badge WORK">${escHtml(type || 'CUSTOM')}</span></td>
      <td style="font-family:monospace;font-size:10px;color:#aaa;">${escHtml(id)}</td>
      <td>${escHtml(sort)}</td>
      <td><div style="display:flex;gap:4px;"><button class="icon-btn" title="Editar" onclick="projEditStage(${projJsArg(id)})">${projActionIcon('edit')}</button><button class="icon-btn del" title="Excluir" onclick="projDeleteStage(${projJsArg(id)},${projJsArg(title)})">${projActionIcon('delete')}</button></div></td>
    </tr>`;
  }).join('');
}

async function projCreateStage() {
  if (!projSelectedId) {
    toast('Selecione um projeto antes de criar etapa.', 'wn');
    return;
  }
  const title = document.getElementById('proj-stage-name')?.value.trim();
  if (!title) {
    toast('Informe o nome da etapa.', 'wn');
    return;
  }
  const fields = {
    TITLE: title,
    COLOR: projApiColor(document.getElementById('proj-stage-color')?.value, '#59abed'),
    AFTER_ID: document.getElementById('proj-stage-after')?.value || 0,
    ENTITY_ID: projSelectedId,
  };
  try {
    const data = await call('task.stages.add', { fields });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Etapa criada.', 'ok');
    document.getElementById('proj-stage-name').value = '';
    projLoadKanban();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function projEditStage(id) {
  const stage = projKanbanStages.find(st => String(st.ID ?? st.id) === String(id));
  if (!stage) return;
  projEditingStageId = id;
  projRenderKanban();
}

function projCancelEditStage() {
  projEditingStageId = null;
  projRenderKanban();
}

async function projSaveStage(id) {
  const title = document.getElementById(`proj-stage-edit-title-${id}`)?.value.trim();
  const color = projApiColor(document.getElementById(`proj-stage-edit-color-${id}`)?.value, '#59abed');
  if (!title) {
    toast('Informe o nome da etapa.', 'wn');
    return;
  }
  try {
    const data = await call('task.stages.update', { id, fields: { TITLE: title, COLOR: color } });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Etapa atualizada.', 'ok');
    projEditingStageId = null;
    projLoadKanban();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projDeleteStage(id, title) {
  if (!confirm(`Excluir etapa "${title}"?`)) return;
  try {
    const data = await call('task.stages.delete', { id });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Etapa excluida.', 'wn');
    projLoadKanban();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projLoadMembers() {
  const list = document.getElementById('proj-members-list');
  if (!projSelectedId) {
    if (list) list.innerHTML = '<div class="proj-empty">Selecione um projeto na aba Projetos para gerenciar membros.</div>';
    return;
  }
  if (list) list.innerHTML = '<div class="proj-empty">Carregando...</div>';
  try {
    const data = await call('sonet_group.user.get', { ID: projSelectedId });
    if (data.error) throw new Error(data.error_description || data.error);
    projMembers = projExtractList(data);
    await projResolveUsers(projMembers.map(projMemberId));
    projRenderMembers();
  } catch (e) {
    if (list) list.innerHTML = `<div class="proj-empty" style="color:#ef4444;">Erro: ${escHtml(e.message)}</div>`;
  }
  projSyncContextUI();
}

function projMemberId(m) {
  return m.USER_ID || m.ID || m.userId || m.id;
}

function projMemberName(m) {
  return m.NAME || m.FULL_NAME || m.USER_NAME || m.LOGIN || projUserNameById(projMemberId(m));
}

function projRenderMembers() {
  const list = document.getElementById('proj-members-list');
  if (!list) return;
  if (!projMembers.length) {
    list.innerHTML = '<div class="proj-empty">Nenhum membro encontrado.</div>';
    return;
  }
  list.innerHTML = projMembers.map(m => {
    const id = projMemberId(m);
    const name = projMemberName(m);
    const role = projRoleLabel(m.ROLE || m.role);
    const initials = String(name).split(/\s+/).filter(Boolean).slice(0, 2).map(x => x[0]).join('').toUpperCase() || String(id).slice(0, 2);
    const nextRole = role.cls === 'mod' ? 'K' : 'E';
    return `<div class="member-row">
      <div class="member-avatar">${escHtml(initials)}</div>
      <div class="member-info"><div class="member-name">${escHtml(name)}</div><div class="member-email">ID: ${escHtml(id)}</div></div>
      <span class="member-role ${role.cls}">${escHtml(role.text)}</span>
      ${role.cls === 'owner' ? '' : `<button class="icon-btn" title="Alterar papel" onclick="projUpdateMemberRole(${projJsArg(id)},'${nextRole}')">${projActionIcon('edit')}</button><button class="icon-btn del" title="Remover" onclick="projRemoveMember(${projJsArg(id)})">${projActionIcon('delete')}</button>`}
    </div>`;
  }).join('');
}

function projFindUserFromMemberInput(value) {
  const raw = String(value || '').trim();
  if (!raw) return { id: '', matches: [] };
  const directId = raw.match(/(?:^|\b)ID:\s*(\d+)(?:\b|\))/i);
  if (directId) return { id: directId[1], matches: [] };
  if (/^\d+$/.test(raw) && projAllUsers.some(user => String(projUserId(user)) === raw)) {
    return { id: raw, matches: [] };
  }
  const exact = projAllUsers.find(user => {
    const id = String(projUserId(user));
    return [
      projUserLookupLabel(user),
      projUserOptionLabel(user),
      projUserDisplayName(user),
      projUserEmail(user),
      id,
    ].some(label => String(label || '').trim().toLowerCase() === raw.toLowerCase());
  });
  if (exact) return { id: String(projUserId(exact)), matches: [] };

  const query = projNormalizeSearch(raw);
  const matches = projAllUsers.filter(user => projNormalizeSearch([
    projUserLookupLabel(user),
    projUserDisplayName(user),
    projUserEmail(user),
    projUserId(user),
  ].filter(Boolean).join(' ')).includes(query));
  if (matches.length === 1) return { id: String(projUserId(matches[0])), matches };
  return { id: '', matches };
}

async function projAddMember() {
  if (!projSelectedId) {
    toast('Selecione um projeto antes de adicionar membros.', 'wn');
    return;
  }
  await projLoadUsers();
  const memberInput = document.getElementById('proj-member-user');
  const lookup = projFindUserFromMemberInput(memberInput?.value || '');
  const ids = lookup.id ? [lookup.id] : [];
  const role = document.getElementById('proj-member-role')?.value || 'K';
  if (!ids.length) {
    const message = lookup.matches.length > 1
      ? 'Selecione um usuario da lista para evitar duplicidade.'
      : 'Digite e selecione um usuario existente.';
    toast(message, 'wn');
    return;
  }
  try {
    const payloadId = ids.length === 1 ? ids[0] : ids;
    const data = await call('sonet_group.user.add', { GROUP_ID: projSelectedId, USER_ID: payloadId });
    if (data.error) throw new Error(data.error_description || data.error);
    if (role === 'E') {
      for (const id of ids) {
        const up = await call('sonet_group.user.update', { GROUP_ID: projSelectedId, USER_ID: id, ROLE: 'E' });
        if (up.error) throw new Error(up.error_description || up.error);
      }
    }
    toast('Membro(s) adicionado(s).', 'ok');
    if (memberInput) memberInput.value = '';
    projLoadMembers();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projUpdateMemberRole(userId, role) {
  try {
    const data = await call('sonet_group.user.update', { GROUP_ID: projSelectedId, USER_ID: userId, ROLE: role });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Papel atualizado.', 'ok');
    projLoadMembers();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projRemoveMember(userId) {
  if (!confirm(`Remover usuario ${userId} do projeto?`)) return;
  try {
    const data = await call('sonet_group.user.delete', { GROUP_ID: projSelectedId, USER_ID: userId });
    if (data.error) throw new Error(data.error_description || data.error);
    toast('Membro removido.', 'wn');
    projLoadMembers();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function projRequireXlsx() {
  if (!window.XLSX) throw new Error('SheetJS nao foi carregado.');
}

function projDownloadWorkbook(wb, name) {
  projRequireXlsx();
  XLSX.writeFile(wb, name);
}

function projDownloadModel() {
  try {
    projRequireXlsx();
    const wb = XLSX.utils.book_new();
    const projectRows = [
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', nome_projeto: 'Implantacao CRM - Cliente X', descricao: 'Descricao opcional', visibilidade: 'PUBLICO', adesao: 'CONVITE', data_inicio: '2026-05-01', data_termino: '2026-06-30', permissao_convite: 'DONO' },
    ];
    const memberRows = [
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', nome: 'Maria Silva', papel: 'DONO' },
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', nome: 'Joao Souza', papel: 'MODERADOR' },
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', nome: 'Ana Costa', papel: 'PARTICIPANTE' },
    ];
    const stageRows = [
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', chave_etapa: 'BACKLOG', nome_etapa: 'Backlog', cor: '#59abed', depois_de_etapa: '' },
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', chave_etapa: 'EM_ANDAMENTO', nome_etapa: 'Em andamento', cor: '#f59e0b', depois_de_etapa: 'BACKLOG' },
      { chave_projeto: 'PROJ_IMPLANTACAO_CLIENTE_X', chave_etapa: 'CONCLUIDO', nome_etapa: 'Concluido', cor: '#22c55e', depois_de_etapa: 'EM_ANDAMENTO' },
    ];
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(projectRows), 'Projetos');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(memberRows), 'Membros');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stageRows), 'Etapas Kanban');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet([
      ['Coluna', 'Obrigatoria', 'Descricao'],
      ['chave_projeto', 'Sim', 'Codigo livre para relacionar Projetos, Membros e Etapas Kanban'],
      ['nome_projeto', 'Sim', 'Nome do projeto'],
      ['descricao', 'Nao', 'Descricao'],
      ['visibilidade', 'Nao', 'PUBLICO ou PRIVADO'],
      ['adesao', 'Nao', 'LIVRE ou CONVITE'],
      ['data_inicio', 'Nao', 'YYYY-MM-DD'],
      ['data_termino', 'Nao', 'YYYY-MM-DD'],
      ['permissao_convite', 'Nao', 'DONO, MODERADORES ou TODOS'],
      ['', '', ''],
      ['Aba Membros', '', 'Use chave_projeto para vincular ao projeto criado'],
      ['nome', 'Sim', 'Nome ou e-mail do usuario no Bitrix24'],
      ['papel', 'Nao', 'DONO, MODERADOR ou PARTICIPANTE'],
      ['', '', ''],
      ['Aba Etapas Kanban', '', 'Use chave_projeto para vincular ao projeto criado'],
      ['chave_etapa', 'Nao', 'Codigo livre para ordenar etapas pela coluna depois_de_etapa'],
      ['nome_etapa', 'Sim', 'Nome da etapa do Kanban'],
      ['cor', 'Nao', 'Cor hexadecimal, exemplo #59abed'],
      ['depois_de_etapa', 'Nao', 'chave_etapa anterior; deixe em branco para criar no inicio'],
    ]), 'Instrucoes');
    projDownloadWorkbook(wb, `modelo-projetos-${new Date().toISOString().slice(0, 10)}.xlsx`);
    projCloseModal('proj-modal-model');
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

async function projExportXlsx(selectedOnly = false) {
  try {
    projRequireXlsx();
    const ids = selectedOnly || projCheckedIds.size ? new Set([...projCheckedIds]) : null;
    const projects = ids ? projetosList.filter(p => ids.has(String(projGetId(p)))) : projetosList;
    if (!projects.length) {
      toast('Nao ha projetos para exportar.', 'wn');
      return;
    }
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(projects.map(p => ({
      id: projGetId(p),
      nome: projGetName(p),
      descricao: p.DESCRIPTION || '',
      visivel: p.VISIBLE || '',
      aberto: p.OPENED || '',
      arquivado: p.CLOSED || '',
      proprietario: projUserNameById(projOwnerId(p)),
      owner_id: projOwnerId(p),
      data_inicio: projDateOnly(p.PROJECT_DATE_START || ''),
      data_termino: projDateOnly(p.PROJECT_DATE_FINISH || ''),
    }))), 'Projetos');
    const memberRows = [];
    const stageRows = [];
    for (const p of projects) {
      const pid = projGetId(p);
      try {
        const data = await call('sonet_group.user.get', { ID: pid });
        const members = projExtractList(data);
        await projResolveUsers(members.map(projMemberId));
        members.forEach(m => memberRows.push({
          projeto_id: pid,
          projeto: projGetName(p),
          nome: projMemberName(m),
          papel: projRoleLabel(m.ROLE || m.role).text,
        }));
      } catch (_) {}
      try {
        const data = await call('task.stages.get', { entityId: pid });
        if (data.error) throw new Error(data.error_description || data.error);
        const stages = projExtractList(data);
        stages.forEach(st => stageRows.push({
          projeto_id: pid,
          projeto: projGetName(p),
          etapa_id: st.ID ?? st.id ?? '',
          nome: projStageTitle(st),
          tipo: projStageType(st),
          ordem: st.SORT ?? st.sort ?? '',
          cor: projNormalizeHex(st.COLOR || st.color, ''),
        }));
      } catch (e) {
        stageRows.push({
          projeto_id: pid,
          projeto: projGetName(p),
          etapa_id: '',
          nome: '',
          tipo: '',
          ordem: '',
          cor: '',
          erro: e.message,
        });
      }
    }
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(memberRows), 'Membros');
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(stageRows), 'Etapas Kanban');
    projDownloadWorkbook(wb, `projetos-${new Date().toISOString().slice(0, 10)}.xlsx`);
    toast('Exportacao gerada.', 'ok');
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  }
}

function projCleanRow(row) {
  const out = {};
  Object.entries(row || {}).forEach(([k, v]) => { out[String(k).trim()] = typeof v === 'string' ? v.trim() : v; });
  return out;
}

function projImportKey(row) {
  const clean = projCleanRow(row);
  return String(clean.chave_projeto || clean.projeto_chave || clean.codigo_projeto || clean.nome_projeto || '').trim();
}

function projMapImportRow(row) {
  const clean = projCleanRow(row);
  const visibility = String(clean.visibilidade || 'PUBLICO').toUpperCase();
  const adesao = String(clean.adesao || 'CONVITE').toUpperCase();
  const perms = String(clean.permissao_convite || 'DONO').toUpperCase();
  const permMap = { DONO: 'A', MODERADORES: 'E', TODOS: 'K' };
  return {
    source: clean,
    key: projImportKey(clean),
    fields: {
      NAME: clean.nome_projeto,
      DESCRIPTION: clean.descricao || '',
      PROJECT: 'Y',
      VISIBLE: visibility === 'PRIVADO' ? 'N' : 'Y',
      OPENED: adesao === 'LIVRE' ? 'Y' : 'N',
      PROJECT_DATE_START: projDateToApi(clean.data_inicio),
      PROJECT_DATE_FINISH: projDateToApi(clean.data_termino),
      INITIATE_PERMS: permMap[perms] || 'A',
    },
  };
}

function projMapMemberImportRow(row) {
  const clean = projCleanRow(row);
  const role = String(clean.papel || clean.role || 'PARTICIPANTE').toUpperCase();
  const roleMap = { DONO: 'A', OWNER: 'A', PROPRIETARIO: 'A', MODERADOR: 'E', MODERATOR: 'E', E: 'E' };
  return {
    source: clean,
    projectKey: projImportKey(clean),
    name: String(clean.nome || clean.usuario || clean.email || clean.usuario_id || clean.user_id || clean.id_usuario || '').trim(),
    role: roleMap[role] || 'K',
  };
}

function projMapStageImportRow(row) {
  const clean = projCleanRow(row);
  const color = projStageImportColor(clean);
  return {
    source: clean,
    projectKey: projImportKey(clean),
    stageKey: String(clean.chave_etapa || clean.etapa_chave || clean.codigo_etapa || clean.nome_etapa || '').trim(),
    afterKey: String(clean.depois_de_etapa || clean.after_key || clean.after_id || '').trim(),
    fields: {
      TITLE: clean.nome_etapa || clean.etapa || clean.nome || '',
      COLOR: projApiColor(color, '#59abed'),
    },
  };
}

function projStageImportColor(clean) {
  return clean.cor || clean.cor_hex || clean.codigo_cor || clean.color || clean.colour;
}

function projSheetRows(wb, sheetName) {
  const sheet = wb.Sheets[sheetName];
  if (!sheet) return [];
  return XLSX.utils.sheet_to_json(sheet, { defval: '' })
    .filter(row => Object.values(row).some(value => String(value).trim() !== ''));
}

function projNormalizeImportRows(rows, columns) {
  return rows.map(row => {
    const clean = projCleanRow(row);
    const normalized = {};
    columns.forEach(([key]) => { normalized[key] = clean[key] ?? ''; });
    Object.entries(clean).forEach(([key, value]) => {
      if (!(key in normalized)) normalized[key] = value ?? '';
    });
    return normalized;
  });
}

function projRebuildImportPlanFromRaw() {
  if (!projImportPlan?.raw) return;
  projImportPlan.projects = (projImportPlan.raw.projects || []).map(projMapImportRow);
  projImportPlan.members = (projImportPlan.raw.members || []).map(projMapMemberImportRow);
  projImportPlan.stages = (projImportPlan.raw.stages || []).map(projMapStageImportRow);
}

function projValidateImportPlan() {
  projRebuildImportPlanFromRaw();
  const errors = [];
  const projects = projImportPlan?.projects || [];
  const members = projImportPlan?.members || [];
  const stages = projImportPlan?.stages || [];
  if (!projects.length) errors.push('Nenhuma linha valida encontrada na aba Projetos.');
  const projectKeys = new Set();
  projects.forEach((row, i) => {
    if (!row.key) errors.push(`Projetos linha ${i + 2}: chave_projeto obrigatoria.`);
    if (!row.fields.NAME) errors.push(`Projetos linha ${i + 2}: nome_projeto obrigatorio.`);
    if (row.key && projectKeys.has(row.key)) errors.push(`Projetos linha ${i + 2}: chave_projeto duplicada (${row.key}).`);
    if (row.key) projectKeys.add(row.key);
  });
  members.forEach((row, i) => {
    if (!row.projectKey) errors.push(`Membros linha ${i + 2}: chave_projeto obrigatoria.`);
    else if (!projectKeys.has(row.projectKey)) errors.push(`Membros linha ${i + 2}: chave_projeto nao encontrada (${row.projectKey}).`);
    if (!row.name) errors.push(`Membros linha ${i + 2}: nome obrigatorio.`);
  });
  stages.forEach((row, i) => {
    if (!row.projectKey) errors.push(`Etapas Kanban linha ${i + 2}: chave_projeto obrigatoria.`);
    else if (!projectKeys.has(row.projectKey)) errors.push(`Etapas Kanban linha ${i + 2}: chave_projeto nao encontrada (${row.projectKey}).`);
    if (!row.fields.TITLE) errors.push(`Etapas Kanban linha ${i + 2}: nome_etapa obrigatorio.`);
    const color = projStageImportColor(row.source || {});
    if (color && !projIsHexColor(color)) errors.push(`Etapas Kanban linha ${i + 2}: cor invalida (${color}). Use formato #RRGGBB.`);
  });
  return { errors, projects, members, stages };
}

function projUpdateImportValidation() {
  const box = document.getElementById('proj-import-validation');
  const btn = document.getElementById('proj-import-run-btn');
  if (!box || !btn) return;
  const { errors, projects, members, stages } = projValidateImportPlan();
  box.innerHTML = errors.length
    ? `<div class="err">Erros encontrados.</div><ul>${errors.map(e => `<li>${escHtml(e)}</li>`).join('')}</ul>`
    : `<div class="ok">Pronto para importar ${projects.length} projeto(s), ${members.length} membro(s) e ${stages.length} etapa(s).</div>`;
  btn.disabled = errors.length > 0;
}

function projUpdateImportCell(sheetKey, rowIndex, columnKey, value) {
  if (!projImportPlan?.raw?.[sheetKey]?.[rowIndex]) return;
  projImportPlan.raw[sheetKey][rowIndex][columnKey] = value;
  projUpdateImportValidation();
}

function projRenderImportSheet(sheet) {
  const rows = projImportPlan?.raw?.[sheet.key] || [];
  if (!rows.length) {
    return `<div class="proj-import-sheet"><div class="proj-import-sheet-title">${escHtml(sheet.title)} <span>0 linhas</span></div><div class="proj-import-empty">Nenhuma linha nesta aba.</div></div>`;
  }
  return `
    <div class="proj-import-sheet">
      <div class="proj-import-sheet-title">${escHtml(sheet.title)} <span>${rows.length} linha(s)</span></div>
      <div class="proj-import-table-wrap">
        <table class="proj-import-table">
          <thead><tr>${sheet.columns.map(([, label]) => `<th>${escHtml(label)}</th>`).join('')}</tr></thead>
          <tbody>
            ${rows.map((row, rowIndex) => `
              <tr>
                ${sheet.columns.map(([key, label]) => `<td><input aria-label="${escHtml(label)}" value="${escHtml(row[key] ?? '')}" oninput="projUpdateImportCell('${sheet.key}',${rowIndex},'${key}',this.value)" /></td>`).join('')}
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>`;
}

function projCreatedGroupId(data) {
  const result = data && data.result;
  if (result && typeof result === 'object') return result.ID || result.id || result.GROUP_ID || result.groupId;
  return result;
}

async function projFetchKanbanStages(groupId) {
  const data = await call('task.stages.get', { entityId: groupId });
  if (data.error) throw new Error(data.error_description || data.error);
  return projSortStages(projExtractList(data));
}

function projImportStagesByProject(stages) {
  const grouped = new Map();
  (stages || []).forEach(stage => {
    if (!grouped.has(stage.projectKey)) grouped.set(stage.projectKey, []);
    grouped.get(stage.projectKey).push(stage);
  });
  return grouped;
}

async function projApplyImportedStages(projectKey, groupId, stages, results) {
  const existingStages = await projFetchKanbanStages(groupId);
  const stageIdsByKey = new Map();
  const usedStageIds = new Set();
  let lastStageId = 0;

  for (let index = 0; index < stages.length; index += 1) {
    const item = stages[index];
    const reusable = existingStages[index] || null;
    const explicitAfterId = item.afterKey ? (stageIdsByKey.get(item.afterKey) || 0) : 0;
    const afterId = item.afterKey ? explicitAfterId : lastStageId;
    const fields = { ...item.fields, AFTER_ID: afterId };

    if (reusable) {
      const id = projStageId(reusable);
      const data = await call('task.stages.update', { id, fields });
      if (data.error) {
        results.push(`Falha ao atualizar etapa ${fields.TITLE}: ${data.error_description || data.error}`);
        continue;
      }
      if (id && item.stageKey) stageIdsByKey.set(item.stageKey, id);
      if (id) usedStageIds.add(String(id));
      lastStageId = id || lastStageId;
      results.push(`Etapa ajustada: ${fields.TITLE}`);
      continue;
    }

    const data = await call('task.stages.add', { fields: { ...fields, ENTITY_ID: groupId } });
    if (data.error) {
      results.push(`Falha ao criar etapa ${fields.TITLE}: ${data.error_description || data.error}`);
      continue;
    }
    const stageId = projCreatedGroupId(data);
    if (stageId && item.stageKey) stageIdsByKey.set(item.stageKey, stageId);
    if (stageId) usedStageIds.add(String(stageId));
    lastStageId = stageId || lastStageId;
    results.push(`Etapa criada: ${fields.TITLE}`);
  }

  for (const stage of existingStages) {
    const id = projStageId(stage);
    if (!id || usedStageIds.has(String(id))) continue;
    const title = projStageTitle(stage) || id;
    const data = await call('task.stages.delete', { id });
    if (data.error) {
      results.push(`Etapa padrao mantida (${title}): ${data.error_description || data.error}`);
      continue;
    }
    results.push(`Etapa padrao removida: ${title}`);
  }
}

async function projResolveImportUserId(name) {
  const target = projNormalizeSearch(name);
  if (!target) return '';
  await projLoadUsers();
  const user = projAllUsers.find(item => {
    const id = String(projUserId(item) || '').trim();
    const display = projNormalizeSearch(projUserDisplayName(item));
    const option = projNormalizeSearch(projUserOptionLabel(item));
    const email = projNormalizeSearch(projUserEmail(item));
    return String(id) === String(name).trim() || display === target || option === target || email === target;
  });
  return user ? projUserId(user) : '';
}

function projRenderImportPreview() {
  const box = document.getElementById('proj-import-preview');
  const btn = document.getElementById('proj-import-run-btn');
  if (!box || !btn) return;
  box.innerHTML = `
    <div class="proj-import-editor">
      <div class="proj-import-validation" id="proj-import-validation"></div>
      ${PROJ_IMPORT_SHEETS.map(projRenderImportSheet).join('')}
    </div>`;
  projUpdateImportValidation();
}

async function projHandleImportFile(file) {
  if (!file) return;
  try {
    projRequireXlsx();
    const wb = XLSX.read(await file.arrayBuffer(), { type: 'array' });
    const sheetName = wb.SheetNames.includes('Projetos') ? 'Projetos' : wb.SheetNames[0];
    const projectColumns = PROJ_IMPORT_SHEETS.find(sheet => sheet.key === 'projects').columns;
    const memberColumns = PROJ_IMPORT_SHEETS.find(sheet => sheet.key === 'members').columns;
    const stageColumns = PROJ_IMPORT_SHEETS.find(sheet => sheet.key === 'stages').columns;
    const raw = {
      projects: projNormalizeImportRows(projSheetRows(wb, sheetName), projectColumns),
      members: projNormalizeImportRows(projSheetRows(wb, 'Membros'), memberColumns),
      stages: projNormalizeImportRows(projSheetRows(wb, 'Etapas Kanban'), stageColumns),
    };
    projImportPlan = {
      raw,
      projects: raw.projects.map(projMapImportRow),
      members: raw.members.map(projMapMemberImportRow),
      stages: raw.stages.map(projMapStageImportRow),
    };
    projRenderImportPreview();
  } catch (e) {
    projImportPlan = null;
    const box = document.getElementById('proj-import-preview');
    if (box) box.innerHTML = `<span class="err">Erro: ${escHtml(e.message)}</span>`;
  }
}

async function projRunImport() {
  if (!projImportPlan?.projects?.length) return;
  const validation = projValidateImportPlan();
  if (validation.errors.length) {
    projUpdateImportValidation();
    return;
  }
  const btn = document.getElementById('proj-import-run-btn');
  if (btn) btn.disabled = true;
  const results = [];
  const groupIdsByKey = new Map();
  try {
    for (const item of projImportPlan.projects) {
      const fields = { ...item.fields };
      Object.keys(fields).forEach(k => fields[k] === undefined && delete fields[k]);
      const data = await call('sonet_group.create', fields);
      if (data.error) results.push(`Falha: ${fields.NAME} - ${data.error_description || data.error}`);
      else {
        const groupId = projCreatedGroupId(data);
        if (groupId) groupIdsByKey.set(item.key, groupId);
        results.push(`Criado: ${fields.NAME}`);
      }
    }
    for (const item of projImportPlan.members || []) {
      const groupId = groupIdsByKey.get(item.projectKey);
      if (!groupId) {
        results.push(`Membro ignorado: projeto ${item.projectKey} nao foi criado.`);
        continue;
      }
      const userId = await projResolveImportUserId(item.name);
      if (!userId) {
        results.push(`Falha ao adicionar ${item.name}: usuario nao encontrado.`);
        continue;
      }
      if (item.role === 'A') {
        const owner = await call('sonet_group.setowner', { GROUP_ID: groupId, USER_ID: userId });
        if (owner.error) results.push(`Falha ao definir dono ${item.name}: ${owner.error_description || owner.error}`);
        else results.push(`Dono definido: ${item.name}`);
        continue;
      }
      const data = await call('sonet_group.user.add', { GROUP_ID: groupId, USER_ID: userId });
      if (data.error) {
        results.push(`Falha ao adicionar ${item.name}: ${data.error_description || data.error}`);
        continue;
      }
      if (item.role === 'E') {
        const up = await call('sonet_group.user.update', { GROUP_ID: groupId, USER_ID: userId, ROLE: 'E' });
        if (up.error) results.push(`${item.name} adicionado, mas papel falhou: ${up.error_description || up.error}`);
      }
      results.push(`Membro adicionado: ${item.name}`);
    }
    const stagesByProject = projImportStagesByProject(projImportPlan.stages || []);
    for (const [projectKey, stages] of stagesByProject.entries()) {
      const groupId = groupIdsByKey.get(projectKey);
      if (!groupId) {
        results.push(`Etapas ignoradas: projeto ${projectKey} nao foi criado.`);
        continue;
      }
      await projApplyImportedStages(projectKey, groupId, stages, results);
    }
    const box = document.getElementById('proj-import-preview');
    if (box) box.innerHTML = `<div class="ok">Importacao concluida.</div><ul>${results.map(x => `<li>${escHtml(x)}</li>`).join('')}</ul>`;
    toast('Importacao concluida.', 'ok');
    projLoad();
  } catch (e) {
    toast('Erro: ' + e.message, 'er');
  } finally {
    if (btn) btn.disabled = false;
  }
}

function projBindImportDrop() {
  const zone = document.getElementById('proj-import-drop');
  if (!zone || zone.dataset.bound === '1') return;
  zone.dataset.bound = '1';
  ['dragenter', 'dragover'].forEach(evt => zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.add('dragover'); }));
  ['dragleave', 'drop'].forEach(evt => zone.addEventListener(evt, e => { e.preventDefault(); zone.classList.remove('dragover'); }));
  zone.addEventListener('drop', e => projHandleImportFile(e.dataTransfer && e.dataTransfer.files && e.dataTransfer.files[0]));
}

function projOnEnvironmentChanged() {
  projetosList = [];
  projSelectedId = null;
  projSelectedName = '';
  projKanbanStages = [];
  projMembers = [];
  projAllUsers = [];
  projUsersLoadedFor = '';
  projUserCache.clear();
  projClearInvites();
  projEditingStageId = null;
  projCurrentPage = 1;
  projLastQuery = '';
  projCheckedIds.clear();
  projSetUserSelectState('Carregando usuarios...');
  projRenderList();
}

window.projSyncContextUI = projSyncContextUI;
window.projOnEnvironmentChanged = projOnEnvironmentChanged;
window.projSwitchTab = projSwitchTab;
window.projToggleCollapse = projToggleCollapse;
window.projLoad = projLoad;
window.projRenderList = projRenderList;
window.projSetPage = projSetPage;
window.projPagePrev = projPagePrev;
window.projPageNext = projPageNext;
window.projToggleChecked = projToggleChecked;
window.projSelectAll = projSelectAll;
window.projClearSelection = projClearSelection;
window.projOpenContext = projOpenContext;
window.projLoadUsers = projLoadUsers;
window.projAddInvite = projAddInvite;
window.projRemoveInvite = projRemoveInvite;
window.projCreate = projCreate;
window.projOpenEdit = projOpenEdit;
window.projSaveEdit = projSaveEdit;
window.projDelete = projDelete;
window.projDeleteSelected = projDeleteSelected;
window.projLoadKanban = projLoadKanban;
window.projCreateStage = projCreateStage;
window.projEditStage = projEditStage;
window.projCancelEditStage = projCancelEditStage;
window.projSaveStage = projSaveStage;
window.projDeleteStage = projDeleteStage;
window.projLoadMembers = projLoadMembers;
window.projAddMember = projAddMember;
window.projUpdateMemberRole = projUpdateMemberRole;
window.projRemoveMember = projRemoveMember;
window.projOpenModal = projOpenModal;
window.projCloseModal = projCloseModal;
window.projDownloadModel = projDownloadModel;
window.projExportXlsx = projExportXlsx;
window.projHandleImportFile = projHandleImportFile;
window.projUpdateImportCell = projUpdateImportCell;
window.projRunImport = projRunImport;
