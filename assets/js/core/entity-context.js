/* -- CONTEXTO DE ENTIDADES CRM / SPA --------------------------------------- */
const CRM_ENTITY_TYPES = Object.freeze({
  LEAD: 1,
  DEAL: 2,
  CONTACT: 3,
  COMPANY: 4,
});

const CRM_ENTITY_META = Object.freeze({
  1: {
    key: 'lead',
    label: 'Lead',
    entityTypeId: CRM_ENTITY_TYPES.LEAD,
    userFieldEntityId: 'CRM_LEAD',
    fieldsMethod: 'crm.lead.fields',
    userFieldListMethod: 'crm.lead.userfield.list',
    userFieldAddMethod: 'crm.lead.userfield.add',
    userFieldUpdateMethod: 'crm.lead.userfield.update',
    userFieldDeleteMethod: 'crm.lead.userfield.delete',
    detailsConfigGetMethod: 'crm.lead.details.configuration.get',
    detailsConfigSetMethod: 'crm.lead.details.configuration.set',
  },
  2: {
    key: 'deal',
    label: 'Negócio',
    entityTypeId: CRM_ENTITY_TYPES.DEAL,
    userFieldEntityId: 'CRM_DEAL',
    fieldsMethod: 'crm.deal.fields',
    userFieldListMethod: 'crm.deal.userfield.list',
    userFieldAddMethod: 'crm.deal.userfield.add',
    userFieldUpdateMethod: 'crm.deal.userfield.update',
    userFieldDeleteMethod: 'crm.deal.userfield.delete',
    detailsConfigGetMethod: 'crm.item.details.configuration.get',
    detailsConfigSetMethod: 'crm.item.details.configuration.set',
  },
  3: {
    key: 'contact',
    label: 'Contato',
    entityTypeId: CRM_ENTITY_TYPES.CONTACT,
    userFieldEntityId: 'CRM_CONTACT',
    fieldsMethod: 'crm.contact.fields',
    userFieldListMethod: 'crm.contact.userfield.list',
    userFieldAddMethod: 'crm.contact.userfield.add',
    userFieldUpdateMethod: 'crm.contact.userfield.update',
    userFieldDeleteMethod: 'crm.contact.userfield.delete',
    detailsConfigGetMethod: 'crm.item.details.configuration.get',
    detailsConfigSetMethod: 'crm.item.details.configuration.set',
  },
  4: {
    key: 'company',
    label: 'Empresa',
    entityTypeId: CRM_ENTITY_TYPES.COMPANY,
    userFieldEntityId: 'CRM_COMPANY',
    fieldsMethod: 'crm.company.fields',
    userFieldListMethod: 'crm.company.userfield.list',
    userFieldAddMethod: 'crm.company.userfield.add',
    userFieldUpdateMethod: 'crm.company.userfield.update',
    userFieldDeleteMethod: 'crm.company.userfield.delete',
    detailsConfigGetMethod: 'crm.item.details.configuration.get',
    detailsConfigSetMethod: 'crm.item.details.configuration.set',
  },
});

let activeCrmEntity = CRM_ENTITY_TYPES.COMPANY;
let activeSpa = null;

function getCurrentMode() {
  return typeof appMode !== 'undefined' && appMode === 'spa' ? 'spa' : 'crm';
}

function setCurrentMode(mode) {
  if (mode !== 'crm' && mode !== 'spa') return getCurrentMode();
  appMode = mode;
  return getCurrentMode();
}

function setCurrentCrmEntity(entityTypeId) {
  const eid = parseInt(entityTypeId, 10);
  if (CRM_ENTITY_META[eid]) activeCrmEntity = eid;
  return activeCrmEntity;
}

function getCurrentCrmEntityTypeId() {
  if (typeof currentPanel !== 'undefined') {
    if (currentPanel === 'pipelines' && typeof pipEntityTypeId !== 'undefined') return pipEntityTypeId;
    if (currentPanel === 'campos' && typeof camposEntityTypeId !== 'undefined') return camposEntityTypeId;
    if (currentPanel === 'criar' && typeof criarEntityTypeId !== 'undefined') return criarEntityTypeId;
    if (currentPanel === 'card' && typeof cardEntityTypeId !== 'undefined') return cardEntityTypeId;
  }
  return activeCrmEntity;
}

function setActiveSpa(spa) {
  if (!spa) {
    activeSpa = null;
    return activeSpa;
  }
  activeSpa = {
    id: parseInt(spa.id, 10),
    entityTypeId: parseInt(spa.entityTypeId, 10),
    title: spa.title || spa.name || '',
    raw: spa,
  };
  return activeSpa;
}

function getActiveSpa() {
  return activeSpa;
}

function requireActiveSpa() {
  if (!activeSpa || !activeSpa.id || !activeSpa.entityTypeId) {
    throw new Error('Selecione um processo inteligente antes de usar este módulo.');
  }
  return activeSpa;
}

function getCurrentEntityTypeId() {
  if (getCurrentMode() === 'spa') return requireActiveSpa().entityTypeId;
  return getCurrentCrmEntityTypeId();
}

function getCurrentUserFieldEntityId() {
  if (getCurrentMode() === 'spa') return 'CRM_' + requireActiveSpa().id;
  const meta = CRM_ENTITY_META[getCurrentCrmEntityTypeId()];
  return meta ? meta.userFieldEntityId : '';
}

function getStageEntityId(entityTypeId, categoryId) {
  const eid = parseInt(entityTypeId, 10);
  const cid = parseInt(categoryId, 10) || 0;

  if (getCurrentMode() === 'spa' || eid > CRM_ENTITY_TYPES.COMPANY) {
    return `DYNAMIC_${eid}_STAGE_${cid}`;
  }
  if (eid === CRM_ENTITY_TYPES.LEAD) return 'STATUS';
  return cid === 0 ? 'DEAL_STAGE' : `DEAL_STAGE_${cid}`;
}

function unsupportedSpaMethod(name) {
  return async function() {
    throw new Error(`SPA: ${name} ainda não foi conectado nesta etapa.`);
  };
}

function buildUserFieldUpdatePayload(input, mode = getCurrentMode()) {
  const label = input.label || '';
  const mandatory = input.mandatory === true ? 'Y' : (input.mandatory || 'N');
  const multiple = input.multiple === true ? 'Y' : (input.multiple || 'N');
  const type = input.type || '';
  const settings = input.settings || {};
  const options = Array.isArray(input.options) ? input.options : [];
  const helpMessage = input.helpMessage || '';

  if (mode === 'spa') {
    const field = {
      editFormLabel: { br: label },
      listColumnLabel: { br: label },
      listFilterLabel: { br: label },
      mandatory,
      multiple,
      settings,
      helpMessage: { br: helpMessage },
    };
    if (type === 'enumeration') {
      field.userTypeId = 'enumeration';
      field.enum = options.map((value, i) => ({
        sort: (i + 1) * 10,
        value,
        def: 'N',
      }));
    }
    return field;
  }

  const fields = {
    EDIT_FORM_LABEL: label,
    LIST_COLUMN_LABEL: label,
    LIST_FILTER_LABEL: label,
    MANDATORY: mandatory,
    MULTIPLE: multiple,
    SETTINGS: settings,
    HELP_MESSAGE: helpMessage,
  };
  if (type === 'enumeration') {
    fields.LIST = options.map((value, i) => ({
      SORT: (i + 1) * 10,
      VALUE: value,
      DEF: 'N',
    }));
  }
  return fields;
}

const EntityAdapter = {
  crm: {
    getEntityMeta(entityTypeId = getCurrentCrmEntityTypeId()) {
      return CRM_ENTITY_META[entityTypeId] || CRM_ENTITY_META[CRM_ENTITY_TYPES.COMPANY];
    },
    getEntityTypeId() {
      return getCurrentCrmEntityTypeId();
    },
    getUserFieldEntityId() {
      return this.getEntityMeta().userFieldEntityId;
    },
    getFieldApi(entityTypeId = getCurrentCrmEntityTypeId()) {
      const meta = this.getEntityMeta(entityTypeId);
      return {
        entityTypeId: meta.entityTypeId,
        entityId: meta.userFieldEntityId,
        fields: meta.fieldsMethod,
        uf: meta.userFieldListMethod,
        add: meta.userFieldAddMethod,
        update: meta.userFieldUpdateMethod,
        delete: meta.userFieldDeleteMethod,
        configGet: meta.detailsConfigGetMethod,
        configSet: meta.detailsConfigSetMethod,
      };
    },
    listPipelines(entityTypeId = getCurrentCrmEntityTypeId()) {
      return call('crm.category.list', { entityTypeId });
    },
    addPipeline(entityTypeId, fields) {
      return call('crm.category.add', { entityTypeId, fields });
    },
    updatePipeline(entityTypeId, id, fields) {
      return call('crm.category.update', { entityTypeId, id, fields });
    },
    deletePipeline(entityTypeId, id) {
      return call('crm.category.delete', { entityTypeId, id });
    },
    listStages(entityTypeId, categoryId) {
      return call('crm.status.list', { filter: { ENTITY_ID: getStageEntityId(entityTypeId, categoryId) } });
    },
    addStage(fields) {
      return call('crm.status.add', { fields });
    },
    updateStage(id, fields) {
      return call('crm.status.update', { id, fields });
    },
    deleteStage(id) {
      return call('crm.status.delete', { id });
    },
    addField(fields, entityTypeId = getCurrentCrmEntityTypeId()) {
      const api = this.getFieldApi(entityTypeId);
      return call(api.add, { fields });
    },
    updateField(id, fields, entityTypeId = getCurrentCrmEntityTypeId()) {
      const api = this.getFieldApi(entityTypeId);
      return call(api.update, { id, fields });
    },
    deleteField(id, entityTypeId = getCurrentCrmEntityTypeId()) {
      const api = this.getFieldApi(entityTypeId);
      return call(api.delete, { id });
    },
  },
  spa: {
    getActive() {
      return requireActiveSpa();
    },
    getEntityTypeId() {
      return requireActiveSpa().entityTypeId;
    },
    getUserFieldEntityId() {
      return 'CRM_' + requireActiveSpa().id;
    },
    getFieldApi() {
      const spa = requireActiveSpa();
      return {
        entityTypeId: spa.entityTypeId,
        entityId: 'CRM_' + spa.id,
        fields: 'crm.item.fields',
        uf: 'userfieldconfig.list',
        add: 'userfieldconfig.add',
        update: 'userfieldconfig.update',
        delete: 'userfieldconfig.delete',
        configGet: 'crm.item.details.configuration.get',
        configSet: 'crm.item.details.configuration.set',
      };
    },
    listTypes() {
      return call('crm.type.list', { order: { title: 'ASC' } });
    },
    addType(fields) {
      return call('crm.type.add', { fields });
    },
    updateType(id, fields) {
      return call('crm.type.update', { id, fields });
    },
    deleteType(id) {
      return call('crm.type.delete', { id });
    },
    listPipelines() {
      return call('crm.category.list', { entityTypeId: requireActiveSpa().entityTypeId });
    },
    addPipeline(fields) {
      return call('crm.category.add', { entityTypeId: requireActiveSpa().entityTypeId, fields });
    },
    updatePipeline(id, fields) {
      return call('crm.category.update', { entityTypeId: requireActiveSpa().entityTypeId, id, fields });
    },
    deletePipeline(id) {
      return call('crm.category.delete', { entityTypeId: requireActiveSpa().entityTypeId, id });
    },
    listStages(categoryId) {
      const spa = requireActiveSpa();
      return call('crm.status.list', { filter: { ENTITY_ID: getStageEntityId(spa.entityTypeId, categoryId) } });
    },
    addStage(fields) {
      return call('crm.status.add', { fields });
    },
    updateStage(id, fields) {
      return call('crm.status.update', { id, fields });
    },
    deleteStage(id) {
      return call('crm.status.delete', { id });
    },
    addField(field) {
      return call('userfieldconfig.add', {
        moduleId: 'crm',
        field: {
          entityId: this.getUserFieldEntityId(),
          ...field,
        },
      });
    },
    listFields: unsupportedSpaMethod('listFields'),
    updateField(id, field) {
      return call('userfieldconfig.update', { moduleId: 'crm', id, field });
    },
    deleteField(id) {
      return call('userfieldconfig.delete', { moduleId: 'crm', id });
    },
  },
};

function getCurrentAdapter() {
  return EntityAdapter[getCurrentMode()];
}

function getFieldApiForCurrentContext() {
  return getCurrentAdapter().getFieldApi();
}

window.CRM_ENTITY_TYPES = CRM_ENTITY_TYPES;
window.CRM_ENTITY_META = CRM_ENTITY_META;
window.EntityAdapter = EntityAdapter;
window.EntityContext = {
  getCurrentMode,
  setCurrentMode,
  setCurrentCrmEntity,
  getCurrentCrmEntityTypeId,
  setActiveSpa,
  getActiveSpa,
  requireActiveSpa,
  getCurrentEntityTypeId,
  getCurrentUserFieldEntityId,
  getStageEntityId,
  getCurrentAdapter,
  getFieldApiForCurrentContext,
  buildUserFieldUpdatePayload,
};
