// ===============================================
// SISTEMA DE CERTIFICACIONES PRESUPUESTALES
// Google Apps Script Backend - COMPLETO Y LIMPIO
// Cáritas Lima - Versión Final
// ===============================================

// Configuración global
const CONFIG = {
  SPREADSHEET_ID: SpreadsheetApp.getActiveSpreadsheet().getId(),
  DRIVE_FOLDER_NAME: 'Certificaciones Presupuestales - Cáritas Lima',
  CARPETA_PLANTILLAS: '1DXyPgOvEn4o-qPT945V_y7ctSQVnV8ce',
  CARPETA_CERTIFICADOS: '1RJ4Ts7fATs_q3IINPTlLK6VHEl4l8hG3',
  AI_ENDPOINT: 'https://oi-server.onrender.com/chat/completions',
  AI_MODEL: 'openrouter/claude-sonnet-4',
  CUSTOMER_ID: 'cus_T7t8xrMoWnLOgO'
};

const SHEET_NAMES = Object.freeze({
  CERTIFICACIONES: 'Certificaciones',
  ITEMS: 'Items',
  FIRMANTES: 'Firmantes',
  CONFIG_SOLICITANTES: 'Config_Solicitantes',
  CONFIG_FIRMANTES: 'Config_Firmantes',
  CONFIG_GENERAL: 'Config_General',
  CATALOGO_INICIATIVAS: 'Cat_Iniciativas',
  CATALOGO_TIPOS: 'Cat_Tipos',
  CATALOGO_FUENTES: 'Cat_Fuentes',
  CATALOGO_FINALIDADES: 'Cat_Finalidades',
  CATALOGO_OFICINAS: 'Cat_Oficinas',
  PLANTILLAS: 'Plantillas',
  BITACORA: 'Bitacora'
});

const PLANTILLA_FIRMANTES = Object.freeze({
  plantilla_evelyn: {
    nombre: 'Evelyn Elena Huaycacllo Marin',
    cargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto'
  },
  plantilla_jorge: {
    nombre: 'Jorge Herrera',
    cargo: 'Director Ejecutivo'
  },
  plantilla_susana: {
    nombre: 'Susana Palomino',
    cargo: 'Coordinadora de Planeamiento y Presupuesto'
  },
  plantilla_otro: {
    nombre: 'Equipo Designado',
    cargo: 'Responsable según tipo de certificación'
  }
});

const DEFAULT_TIMEZONE = (() => {
  try {
    const tz = Session.getScriptTimeZone();
    return tz || 'America/Lima';
  } catch (error) {
    Logger.log('No se pudo obtener la zona horaria del script: ' + error.toString());
    return 'America/Lima';
  }
})();

const CERTIFICACIONES_HEADERS = Object.freeze([
  'Código',
  'Fecha Emisión',
  'Descripción',
  'Iniciativa',
  'Tipo',
  'Fuente',
  'Finalidad',
  'Oficina',
  'Solicitante',
  'Cargo Solicitante',
  'Email Solicitante',
  'Número Autorización',
  'Cargo Autorizador',
  'Estado',
  'Disposición/Base Legal',
  'Monto Total',
  'Monto en Letras',
  'Fecha Creación',
  'Creado Por',
  'Fecha Modificación',
  'Modificado Por',
  'Fecha Anulación',
  'Anulado Por',
  'Motivo Anulación',
  'Plantilla',
  'URL Documento',
  'URL PDF',
  'Finalidad Detallada'
]);

const ITEMS_HEADERS = Object.freeze([
  'Código Certificación',
  'Orden',
  'Descripción',
  'Cantidad',
  'Unidad',
  'Precio Unitario',
  'Subtotal',
  'Fecha Creación',
  'Creado Por'
]);

const FIRMANTES_HEADERS = Object.freeze([
  'Código Certificación',
  'Orden',
  'Nombre',
  'Cargo',
  'Obligatorio',
  'Fecha Creación',
  'Creado Por'
]);

const BITACORA_HEADERS = Object.freeze([
  'Fecha',
  'Usuario',
  'Acción',
  'Detalles',
  'Usuario Completo'
]);

const PLANTILLAS_HEADERS = Object.freeze([
  'ID',
  'Nombre',
  'Descripción',
  'Activa',
  'Firmantes',
  'Plantilla HTML',
  'Firmante ID',
  'Firmante Nombre',
  'Firmante Cargo'
]);

const PLANTILLAS_COLUMN_DEFINITIONS = Object.freeze([
  { field: 'id', defaultIndex: 0, aliases: ['id', 'codigo', 'identificador'] },
  { field: 'nombre', defaultIndex: 1, aliases: ['nombre', 'nombre plantilla'] },
  { field: 'descripcion', defaultIndex: 2, aliases: ['descripcion', 'descripción'] },
  { field: 'activa', defaultIndex: 3, aliases: ['activa', 'activo', 'habilitada'] },
  { field: 'firmantes', defaultIndex: 4, aliases: ['firmantes', 'cantidad firmantes', 'numero firmantes'] },
  { field: 'plantillaHtml', defaultIndex: 5, aliases: ['plantilla html', 'url plantilla', 'enlace plantilla', 'plantilla doc'] },
  { field: 'firmanteId', defaultIndex: 6, aliases: ['firmante id', 'id firmante', 'codigo firmante'] },
  { field: 'firmanteNombre', defaultIndex: 7, aliases: ['firmante nombre', 'nombre firmante'] },
  { field: 'firmanteCargo', defaultIndex: 8, aliases: ['firmante cargo', 'cargo firmante'] }
]);

let plantillasCache = null;
let plantillasCacheTimestamp = 0;
const PLANTILLAS_CACHE_TTL_MS = 5 * 60 * 1000;
let certificadosFolderCache = null;

function getSpreadsheet() {
  return SpreadsheetApp.getActiveSpreadsheet();
}

function getSheetOrThrow(name) {
  const sheet = getSpreadsheet().getSheetByName(name);
  if (!sheet) {
    throw new Error(`La hoja "${name}" no existe. Ejecute configurarSistema() primero.`);
  }
  return sheet;
}

function getSheetValues(sheet) {
  return sheet.getDataRange().getValues();
}

function findRowIndex(values, columnIndex, value) {
  for (let i = 1; i < values.length; i++) {
    if (values[i][columnIndex] === value) {
      return i;
    }
  }
  return -1;
}

function normalizeHeaderName(value) {
  if (value === null || value === undefined) {
    return '';
  }

  try {
    return String(value)
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(/[\u0300-\u036f]/g, '');
  } catch (error) {
    return String(value).trim().toLowerCase();
  }
}

function findColumnIndexByAliases(headers, aliases, fallbackIndex = -1) {
  const normalizedHeaders = headers.map(normalizeHeaderName);
  for (const alias of aliases) {
    const index = normalizedHeaders.indexOf(alias);
    if (index !== -1) {
      return index;
    }
  }
  return fallbackIndex;
}

const getDefaultFinalidadDetalladaAliases = (() => {
  const defaults = Object.freeze([
    'finalidad detallada',
    'finalidad detallada / justificación',
    'finalidad detallada / justificacion',
    'finalidad (detalle)',
    'detalle de la finalidad',
    'detalle finalidad',
    'justificación',
    'justificacion'
  ]);
  return () => defaults;
})();

function getFinalidadDetalladaAliases() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const rawAliases = properties.getProperty('FINALIDAD_DETALLADA_ALIASES');

    if (!rawAliases) {
      return getDefaultFinalidadDetalladaAliases();
    }

    const parsedAliases = JSON.parse(rawAliases);
    if (!Array.isArray(parsedAliases)) {
      return getDefaultFinalidadDetalladaAliases();
    }

    const normalizedAliases = parsedAliases
      .map(normalizeHeaderName)
      .filter(Boolean);

    const uniqueAliases = Array.from(new Set(normalizedAliases));
    return uniqueAliases.length > 0 ? uniqueAliases : getDefaultFinalidadDetalladaAliases();
  } catch (error) {
    Logger.log('No se pudieron obtener alias personalizados de finalidad detallada: ' + error.toString());
    return getDefaultFinalidadDetalladaAliases();
  }
}

function getFinalidadDetalladaColumnIndex(sheet) {
  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return -1;
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0] || [];
  const fallbackIndex = lastColumn >= 28 ? 27 : -1;
  const index = findColumnIndexByAliases(headers, getFinalidadDetalladaAliases(), fallbackIndex);
  return index >= lastColumn ? -1 : index;
}

function getCertificacionColumnDefinitions() {
  return [
    { field: 'codigo', defaultIndex: 0, aliases: ['codigo', 'codigo cp', 'codigo certificacion', 'codigo certificacion presupuestal', 'id'] },
    { field: 'fechaEmision', defaultIndex: 1, aliases: ['fecha emision', 'fecha certificacion', 'fecha emisión', 'fecha'] },
    { field: 'descripcion', defaultIndex: 2, aliases: ['descripcion', 'descripcion certificacion', 'detalle', 'descripcion detalle'] },
    { field: 'iniciativa', defaultIndex: 3, aliases: ['iniciativa', 'codigo iniciativa'] },
    { field: 'tipo', defaultIndex: 4, aliases: ['tipo', 'tipo certificacion'] },
    { field: 'fuente', defaultIndex: 5, aliases: ['fuente', 'fuente financiamiento', 'fuente de financiamiento'] },
    { field: 'finalidad', defaultIndex: 6, aliases: ['finalidad', 'justificacion', 'justificación'] },
    { field: 'oficina', defaultIndex: 7, aliases: ['oficina', 'unidad organica', 'unidad orgánica'] },
    { field: 'solicitante', defaultIndex: 8, aliases: ['solicitante', 'nombre solicitante'] },
    { field: 'cargoSolicitante', defaultIndex: 9, aliases: ['cargo solicitante', 'cargo del solicitante'] },
    { field: 'emailSolicitante', defaultIndex: 10, aliases: ['email solicitante', 'correo solicitante'] },
    { field: 'numeroAutorizacion', defaultIndex: 11, aliases: ['numero autorizacion', 'nro autorizacion', 'autorizacion'] },
    { field: 'cargoAutorizador', defaultIndex: 12, aliases: ['cargo autorizador'] },
    { field: 'estado', defaultIndex: 13, aliases: ['estado', 'estado certificacion'] },
    { field: 'disposicion', defaultIndex: 14, aliases: ['disposicion', 'base legal', 'disposicion/base legal'] },
    { field: 'montoTotal', defaultIndex: 15, aliases: ['monto total', 'monto'] },
    { field: 'montoLetras', defaultIndex: 16, aliases: ['monto en letras', 'monto letras'] },
    { field: 'fechaCreacion', defaultIndex: 17, aliases: ['fecha creacion', 'fecha creación'] },
    { field: 'creadoPor', defaultIndex: 18, aliases: ['creado por', 'usuario creador'] },
    { field: 'fechaModificacion', defaultIndex: 19, aliases: ['fecha modificacion', 'fecha modificación'] },
    { field: 'modificadoPor', defaultIndex: 20, aliases: ['modificado por', 'usuario modificador'] },
    { field: 'fechaAnulacion', defaultIndex: 21, aliases: ['fecha anulacion', 'fecha anulación'] },
    { field: 'anuladoPor', defaultIndex: 22, aliases: ['anulado por'] },
    { field: 'motivoAnulacion', defaultIndex: 23, aliases: ['motivo anulacion', 'motivo anulación'] },
    { field: 'plantilla', defaultIndex: 24, aliases: ['plantilla', 'plantilla certificacion'] },
    { field: 'urlDocumento', defaultIndex: 25, aliases: ['url documento', 'enlace documento'] },
    { field: 'urlPDF', defaultIndex: 26, aliases: ['url pdf', 'enlace pdf'] },
    { field: 'finalidadDetallada', defaultIndex: 27, aliases: getFinalidadDetalladaAliases() }
  ];
}

function buildCertificacionHeaderMap(headers) {
  const normalizedHeaders = headers.map(normalizeHeaderName);
  const map = {};
  const definitions = getCertificacionColumnDefinitions();

  definitions.forEach(definition => {
    const aliases = Array.isArray(definition.aliases)
      ? definition.aliases.map(normalizeHeaderName)
      : [];

    let index = -1;
    for (let i = 0; i < aliases.length; i++) {
      const aliasIndex = normalizedHeaders.indexOf(aliases[i]);
      if (aliasIndex !== -1) {
        index = aliasIndex;
        break;
      }
    }

    if (index === -1 && definition.defaultIndex >= 0 && definition.defaultIndex < headers.length) {
      index = definition.defaultIndex;
    }

    map[definition.field] = index;
  });

  return map;
}

function getCertificacionFieldValue(row, headerMap, field, fallback = '') {
  if (!headerMap || typeof headerMap[field] !== 'number') {
    return fallback;
  }

  const index = headerMap[field];
  if (index >= 0 && index < row.length) {
    return row[index];
  }

  return fallback;
}

function setCertificacionFieldValue(row, headerMap, field, value) {
  if (!headerMap || typeof headerMap[field] !== 'number') {
    return;
  }

  const index = headerMap[field];
  if (index < 0) {
    return;
  }

  if (index >= row.length) {
    for (let i = row.length; i <= index; i++) {
      row[i] = '';
    }
  }

  row[index] = value;
}

function getCertificacionesHeaderInfo(sheet) {
  if (!sheet) {
    return { headers: [], map: {} };
  }

  const lastColumn = sheet.getLastColumn();
  if (lastColumn === 0) {
    return { headers: [], map: {} };
  }

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0] || [];
  const map = buildCertificacionHeaderMap(headers);

  return { headers, map };
}

function ensureSheetStructure(name, headers) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  const requiredColumns = headers.length;
  const maxColumns = sheet.getMaxColumns();
  if (maxColumns < requiredColumns) {
    sheet.insertColumnsAfter(maxColumns, requiredColumns - maxColumns);
  }

  const headerRange = sheet.getRange(1, 1, 1, requiredColumns);
  const currentHeaders = headerRange.getValues()[0];

  let needsUpdate = false;
  const normalizedTarget = headers.map(normalizeHeaderName);
  const normalizedCurrent = currentHeaders.map(normalizeHeaderName);

  for (let i = 0; i < headers.length; i++) {
    if (normalizedCurrent[i] !== normalizedTarget[i]) {
      needsUpdate = true;
      break;
    }
  }

  if (needsUpdate) {
    headerRange.setValues([headers]);
  }

  if (sheet.getFrozenRows() < 1) {
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function ensureCertificacionesSheet() {
  return ensureSheetStructure(SHEET_NAMES.CERTIFICACIONES, CERTIFICACIONES_HEADERS);
}

function ensureItemsSheet() {
  return ensureSheetStructure(SHEET_NAMES.ITEMS, ITEMS_HEADERS);
}

function ensureFirmantesSheet() {
  return ensureSheetStructure(SHEET_NAMES.FIRMANTES, FIRMANTES_HEADERS);
}

function ensureBitacoraSheet() {
  return ensureSheetStructure(SHEET_NAMES.BITACORA, BITACORA_HEADERS);
}

function getActiveUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || 'sistema@caritaslima.org';
  } catch (error) {
    Logger.log('No se pudo obtener el correo del usuario activo: ' + error.toString());
    return 'sistema@caritaslima.org';
  }
}

function sanitizeText(value, fallback = '') {
  if (value === null || value === undefined) {
    return fallback;
  }
  return String(value).trim();
}

function toDate(value) {
  if (!value) return null;
  const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  return isNaN(date.getTime()) ? null : date;
}

function formatDateForClient(value) {
  const date = toDate(value);
  if (!date) return '';
  return Utilities.formatDate(date, DEFAULT_TIMEZONE, 'yyyy-MM-dd');
}

function formatDateTimeForClient(value) {
  const date = toDate(value);
  if (!date) return '';
  return Utilities.formatDate(date, DEFAULT_TIMEZONE, "yyyy-MM-dd'T'HH:mm:ss");
}

function parseNumber(value, fallback = 0) {
  if (value === null || value === undefined || value === '') {
    return fallback;
  }
  const numberValue = typeof value === 'number' ? value : Number(value);
  return Number.isFinite(numberValue) ? numberValue : fallback;
}

function parseDate(value, fallback = new Date()) {
  if (!value) return fallback;
  try {
    return new Date(value);
  } catch (error) {
    return fallback;
  }
}

function prepararDatosCertificacion(datos, usuarioActual) {
  const datosCompletos = { ...datos };

  if (datosCompletos.solicitanteId) {
    const solicitante = obtenerSolicitantePorId(datosCompletos.solicitanteId);
    if (solicitante) {
      datosCompletos.solicitante = solicitante.nombre;
      datosCompletos.cargoSolicitante = solicitante.cargo;
      datosCompletos.emailSolicitante = solicitante.email;
    }
  }

  const itemsNormalizados = Array.isArray(datosCompletos.items)
    ? datosCompletos.items.map(normalizarItemCertificacion).filter(Boolean)
    : [];

  return {
    descripcion: sanitizeText(datosCompletos.descripcion),
    iniciativa: sanitizeText(datosCompletos.iniciativa),
    tipo: sanitizeText(datosCompletos.tipo),
    fuente: sanitizeText(datosCompletos.fuente),
    finalidad: sanitizeText(datosCompletos.finalidad),
    finalidadDetallada: sanitizeText(datosCompletos.finalidadDetallada || datosCompletos.finalidad),
    oficina: sanitizeText(datosCompletos.oficina),
    solicitante: sanitizeText(datosCompletos.solicitante),
    cargoSolicitante: sanitizeText(datosCompletos.cargoSolicitante),
    emailSolicitante: sanitizeText(datosCompletos.emailSolicitante, usuarioActual),
    disposicion: sanitizeText(datosCompletos.disposicion),
    plantilla: sanitizeText(datosCompletos.plantilla, 'plantilla_evelyn'),
    items: itemsNormalizados,
    firmantes: Array.isArray(datosCompletos.firmantes) ? datosCompletos.firmantes : []
  };
}

function normalizarItemCertificacion(item) {
  if (!item) return null;
  const descripcion = sanitizeText(item.descripcion);
  if (!descripcion) return null;

  const cantidad = Number(item.cantidad || 0);
  const precioUnitario = Number(item.precioUnitario || item.precio || 0);
  const subtotalCalculado = cantidad * precioUnitario;

  return {
    descripcion,
    cantidad,
    unidad: sanitizeText(item.unidad),
    precioUnitario,
    subtotal: subtotalCalculado
  };
}

function mapRowToCertificacion(row, index, headerMap) {
  const fechaEmisionDate = toDate(getCertificacionFieldValue(row, headerMap, 'fechaEmision'));
  const fechaCreacionDate = toDate(getCertificacionFieldValue(row, headerMap, 'fechaCreacion'));
  const fechaModificacionDate = toDate(getCertificacionFieldValue(row, headerMap, 'fechaModificacion'));
  const fechaAnulacionDate = toDate(getCertificacionFieldValue(row, headerMap, 'fechaAnulacion'));

  const finalidad = sanitizeText(getCertificacionFieldValue(row, headerMap, 'finalidad'));
  const finalidadDetallada = sanitizeText(
    getCertificacionFieldValue(row, headerMap, 'finalidadDetallada'),
    finalidad
  ) || finalidad;

  const estado = sanitizeText(
    getCertificacionFieldValue(row, headerMap, 'estado'),
    ESTADOS.BORRADOR
  ) || ESTADOS.BORRADOR;

  const montoTotal = parseNumber(getCertificacionFieldValue(row, headerMap, 'montoTotal'), 0);

  return {
    codigo: sanitizeText(getCertificacionFieldValue(row, headerMap, 'codigo')),
    fechaEmision: formatDateForClient(fechaEmisionDate),
    fechaEmisionTimestamp: fechaEmisionDate ? fechaEmisionDate.getTime() : null,
    descripcion: sanitizeText(getCertificacionFieldValue(row, headerMap, 'descripcion')),
    iniciativa: sanitizeText(getCertificacionFieldValue(row, headerMap, 'iniciativa')),
    tipo: sanitizeText(getCertificacionFieldValue(row, headerMap, 'tipo')),
    fuente: sanitizeText(getCertificacionFieldValue(row, headerMap, 'fuente')),
    finalidad,
    oficina: sanitizeText(getCertificacionFieldValue(row, headerMap, 'oficina')),
    solicitante: sanitizeText(getCertificacionFieldValue(row, headerMap, 'solicitante')),
    cargoSolicitante: sanitizeText(getCertificacionFieldValue(row, headerMap, 'cargoSolicitante')),
    emailSolicitante: sanitizeText(getCertificacionFieldValue(row, headerMap, 'emailSolicitante')),
    numeroAutorizacion: sanitizeText(getCertificacionFieldValue(row, headerMap, 'numeroAutorizacion')),
    cargoAutorizador: sanitizeText(getCertificacionFieldValue(row, headerMap, 'cargoAutorizador')),
    estado,
    disposicion: sanitizeText(getCertificacionFieldValue(row, headerMap, 'disposicion')),
    montoTotal,
    montoLetras: sanitizeText(getCertificacionFieldValue(row, headerMap, 'montoLetras')),
    fechaCreacion: formatDateTimeForClient(fechaCreacionDate),
    fechaCreacionTimestamp: fechaCreacionDate ? fechaCreacionDate.getTime() : null,
    creadoPor: sanitizeText(getCertificacionFieldValue(row, headerMap, 'creadoPor')),
    fechaModificacion: formatDateTimeForClient(fechaModificacionDate),
    fechaModificacionTimestamp: fechaModificacionDate ? fechaModificacionDate.getTime() : null,
    modificadoPor: sanitizeText(getCertificacionFieldValue(row, headerMap, 'modificadoPor')),
    fechaAnulacion: formatDateForClient(fechaAnulacionDate),
    fechaAnulacionTimestamp: fechaAnulacionDate ? fechaAnulacionDate.getTime() : null,
    anuladoPor: sanitizeText(getCertificacionFieldValue(row, headerMap, 'anuladoPor')),
    motivoAnulacion: sanitizeText(getCertificacionFieldValue(row, headerMap, 'motivoAnulacion')),
    plantilla: sanitizeText(getCertificacionFieldValue(row, headerMap, 'plantilla')) || 'plantilla_evelyn',
    urlDocumento: sanitizeText(getCertificacionFieldValue(row, headerMap, 'urlDocumento')),
    urlPDF: sanitizeText(getCertificacionFieldValue(row, headerMap, 'urlPDF')),
    finalidadDetallada,
    fila: index + 2
  };
}

// Estados de certificación
const ESTADOS = {
  BORRADOR: 'Borrador',
  EN_REVISION: 'En revisión',
  AUTORIZACION_PENDIENTE: 'Autorización pendiente',
  ACTIVA: 'Activa',
  ANULADA: 'Anulada'
};

// Roles de usuario
const ROLES = {
  SOLICITANTE: 'Solicitante',
  REVISOR: 'Revisor/Presupuesto',
  AUTORIZADOR: 'Autorizador',
  ADMINISTRADOR: 'Administrador'
};

// ===============================================
// FUNCIONES PRINCIPALES DE LA APLICACIÓN WEB
// ===============================================

function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sistema de Certificaciones Presupuestales - Cáritas Lima')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ===============================================
// GESTIÓN DE CERTIFICACIONES
// ===============================================

function crearCertificacion(datos) {
  try {
    const sheet = ensureCertificacionesSheet();

    Logger.log('Creando certificación con datos: ' + JSON.stringify(datos));

    const codigo = generarCodigoCertificacionConsecutivo();
    const fechaActual = new Date();
    const fechaCertificacion = parseDate(datos.fechaCertificacion || datos.fecha, fechaActual);
    const usuario = getActiveUserEmail();

    const datosCompletos = prepararDatosCertificacion(datos, usuario);
    const finalidad = sanitizeText(datosCompletos.finalidad) || generarFinalidadAutomatica(datosCompletos.descripcion);
    const finalidadDetallada = sanitizeText(datosCompletos.finalidadDetallada) || finalidad;
    const disposicion = sanitizeText(datosCompletos.disposicion) || obtenerDisposicionPorDefecto();
    const plantilla = sanitizeText(datosCompletos.plantilla) || 'plantilla_evelyn';
    const headerInfo = getCertificacionesHeaderInfo(sheet);
    const headerCount = headerInfo.headers.length || CERTIFICACIONES_HEADERS.length;
    const headerMap = headerInfo.map;
    const fila = new Array(headerCount).fill('');

    setCertificacionFieldValue(fila, headerMap, 'codigo', codigo);
    setCertificacionFieldValue(fila, headerMap, 'fechaEmision', fechaCertificacion);
    setCertificacionFieldValue(fila, headerMap, 'descripcion', datosCompletos.descripcion);
    setCertificacionFieldValue(fila, headerMap, 'iniciativa', datosCompletos.iniciativa);
    setCertificacionFieldValue(fila, headerMap, 'tipo', datosCompletos.tipo);
    setCertificacionFieldValue(fila, headerMap, 'fuente', datosCompletos.fuente);
    setCertificacionFieldValue(fila, headerMap, 'finalidad', finalidad);
    setCertificacionFieldValue(fila, headerMap, 'oficina', datosCompletos.oficina);
    setCertificacionFieldValue(fila, headerMap, 'solicitante', datosCompletos.solicitante);
    setCertificacionFieldValue(fila, headerMap, 'cargoSolicitante', datosCompletos.cargoSolicitante);
    setCertificacionFieldValue(fila, headerMap, 'emailSolicitante', datosCompletos.emailSolicitante);
    setCertificacionFieldValue(fila, headerMap, 'numeroAutorizacion', '');
    setCertificacionFieldValue(fila, headerMap, 'cargoAutorizador', '');
    setCertificacionFieldValue(fila, headerMap, 'estado', ESTADOS.BORRADOR);
    setCertificacionFieldValue(fila, headerMap, 'disposicion', disposicion);
    setCertificacionFieldValue(fila, headerMap, 'montoTotal', 0);
    setCertificacionFieldValue(fila, headerMap, 'montoLetras', '');
    setCertificacionFieldValue(fila, headerMap, 'fechaCreacion', fechaActual);
    setCertificacionFieldValue(fila, headerMap, 'creadoPor', usuario);
    setCertificacionFieldValue(fila, headerMap, 'fechaModificacion', fechaActual);
    setCertificacionFieldValue(fila, headerMap, 'modificadoPor', usuario);
    setCertificacionFieldValue(fila, headerMap, 'fechaAnulacion', '');
    setCertificacionFieldValue(fila, headerMap, 'anuladoPor', '');
    setCertificacionFieldValue(fila, headerMap, 'motivoAnulacion', '');
    setCertificacionFieldValue(fila, headerMap, 'plantilla', plantilla);
    setCertificacionFieldValue(fila, headerMap, 'urlDocumento', '');
    setCertificacionFieldValue(fila, headerMap, 'urlPDF', '');
    setCertificacionFieldValue(fila, headerMap, 'finalidadDetallada', finalidadDetallada);

    sheet.appendRow(fila);
    SpreadsheetApp.flush();

    if (datosCompletos.items.length > 0) {
      crearItemsCertificacion(codigo, datosCompletos.items);
    }

    crearFirmantesBasadosEnPlantilla(codigo, plantilla);
    recalcularTotalesCertificacion(codigo);

    const resultadoGeneracion = generarCertificadoPerfecto(codigo);

    if (!resultadoGeneracion.success) {
      Logger.log(`❌ Error generando certificado: ${resultadoGeneracion.error}`);
    }

    registrarActividad('CREAR_CERTIFICACION', `Código: ${codigo}`);

    const certificacionActualizada = obtenerCertificacionPorCodigo(codigo);

    return {
      success: true,
      codigo,
      certificacion: certificacionActualizada,
      certificado: resultadoGeneracion,
      urls: {
        documento: resultadoGeneracion.success ? resultadoGeneracion.urlDocumento : null,
        pdf: resultadoGeneracion.success ? resultadoGeneracion.urlPDF : null,
        vistaPrevia: resultadoGeneracion.success ? resultadoGeneracion.urlVistaPrevia : null
      }
    };
  } catch (error) {
    Logger.log('Error en crearCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}
// ===============================================
// GENERACIÓN DE DOCUMENTOS (FUNCIÓN PRINCIPAL)
// ===============================================

function generarDocumentoCertificacion(codigoCertificacion) {
  try {
    const certificacion = obtenerCertificacionPorCodigo(codigoCertificacion);
    if (!certificacion) {
      return { success: false, error: 'Certificación no encontrada' };
    }

    // Crear documento básico que SIEMPRE funciona
    const doc = DocumentApp.create(`Certificacion_${codigoCertificacion}`);
    const body = doc.getBody();

    const folder = getCertificadosFolder();
    if (folder) {
      try {
        const docFile = DriveApp.getFileById(doc.getId());
        folder.addFile(docFile);
        const parents = docFile.getParents();
        const folderId = folder.getId();
        const padresAEliminar = [];
        while (parents.hasNext()) {
          const parent = parents.next();
          if (parent.getId() !== folderId) {
            padresAEliminar.push(parent);
          }
        }
        padresAEliminar.forEach(parent => parent.removeFile(docFile));
      } catch (folderError) {
        Logger.log('No se pudo mover el documento a la carpeta configurada: ' + folderError.toString());
      }
    }

    // Configurar márgenes
    body.setMarginTop(72);
    body.setMarginBottom(72);
    body.setMarginLeft(72);
    body.setMarginRight(72);
    
    // Header con logo y título
    const headerTable = body.appendTable();
    const headerRow = headerTable.appendTableRow();
    
    // Logo
    const logoCell = headerRow.appendTableCell();
    logoCell.appendParagraph('🍀 Cáritas').editAsText().setBold(true).setFontSize(14).setForegroundColor('#019952');
    logoCell.appendParagraph('LIMA').editAsText().setBold(true).setFontSize(12).setForegroundColor('#019952');
    logoCell.setWidth(120);
    
    // Título
    const titleCell = headerRow.appendTableCell();
    titleCell.appendParagraph('Certificación Presupuestal').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(16);
    
    headerTable.setBorderWidth(0);
    
    body.appendParagraph(''); // Espaciado
    
    // Información básica
    body.appendParagraph(`Número: ${certificacion.codigo}`).editAsText().setBold(0, 7, true).setFontSize(11);
    body.appendParagraph(`Fecha: ${formatearFechaDocumento(certificacion.fechaEmision)}`).editAsText().setBold(0, 5, true).setFontSize(11);
    body.appendParagraph(`Responsable del área solicitante: ${certificacion.solicitante}`).editAsText().setBold(0, 35, true).setFontSize(11);
    body.appendParagraph(`Oficina solicitante: ${obtenerNombreOficina(certificacion.oficina)}`).editAsText().setBold(0, 18, true).setFontSize(11);
    body.appendParagraph(`Iniciativa: ${obtenerNombreCatalogo('iniciativas', certificacion.iniciativa)} y ${certificacion.descripcion}`).editAsText().setBold(0, 10, true).setFontSize(11);
    
    body.appendParagraph(''); // Espaciado
    
    // Tabla de ítems (método simplificado)
    if (certificacion.items && certificacion.items.length > 0) {
      const itemsTable = body.appendTable();
      itemsTable.setBorderWidth(1);
      
      // Encabezados
      const headerRow = itemsTable.appendTableRow();
      headerRow.appendTableCell('Descripción').editAsText().setBold(true).setFontSize(10);
      headerRow.appendTableCell('Cant.').editAsText().setBold(true).setFontSize(10);
      headerRow.appendTableCell('C/U (S/)').editAsText().setBold(true).setFontSize(10);
      headerRow.appendTableCell('C/T(S/)').editAsText().setBold(true).setFontSize(10);
      
      // Datos
      certificacion.items.forEach(item => {
        const dataRow = itemsTable.appendTableRow();
        dataRow.appendTableCell(item.descripcion).editAsText().setFontSize(10);
        dataRow.appendTableCell(item.cantidad.toString()).editAsText().setFontSize(10);
        dataRow.appendTableCell(`S/ ${item.precioUnitario.toFixed(2)}`).editAsText().setFontSize(10);
        dataRow.appendTableCell(`S/ ${item.subtotal.toFixed(2)}`).editAsText().setFontSize(10);
      });
      
      // Total
      const totalRow = itemsTable.appendTableRow();
      totalRow.appendTableCell('Total').editAsText().setBold(true).setFontSize(10);
      totalRow.appendTableCell('1').editAsText().setBold(true).setFontSize(10);
      totalRow.appendTableCell(`S/ ${certificacion.montoTotal.toFixed(2)}`).editAsText().setBold(true).setFontSize(10);
      totalRow.appendTableCell(`S/ ${certificacion.montoTotal.toFixed(2)}`).editAsText().setBold(true).setFontSize(10);
    }
    
    body.appendParagraph(''); // Espaciado
    
    // Información adicional
    body.appendParagraph(`Base Legal: ${certificacion.disposicion || obtenerDisposicionPorDefecto()}`).editAsText().setBold(0, 10, true).setFontSize(10);
    body.appendParagraph(`Fuente de Financiamiento: ${obtenerNombreCatalogo('fuentes', certificacion.fuente)}`).editAsText().setBold(0, 24, true).setFontSize(10);
    body.appendParagraph(`Finalidad: ${certificacion.finalidad}`).editAsText().setBold(0, 9, true).setFontSize(10);
    body.appendParagraph(`Monto: S/ ${certificacion.montoTotal.toFixed(2)} | ${certificacion.montoLetras}`).editAsText().setBold(0, 6, true).setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    
    // Disposiciones
    body.appendParagraph('Disposiciones:').editAsText().setBold(true).setFontSize(11);
    body.appendParagraph('• Se ha considerado la evaluación realizada por el área de logística desde la oficina de administración y según el estudio de mercado (cuadro comparativo)').editAsText().setFontSize(10);
    body.appendParagraph('• La presente autorización presupuestal se emite en base a la disponibilidad presupuestal aprobada para la iniciativa').editAsText().setFontSize(10);
    body.appendParagraph('• El responsable de la ejecución del gasto deberá presentar la documentación sustentatoria de acuerdo a las normas vigentes.').editAsText().setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    body.appendParagraph('Adjuntos: Documento sustentatorio obligatorios (contrataciones, proformas, términos de referencia, etc.)').editAsText().setBold(0, 8, true).setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    
    // Fecha de firma
    const fechaFirma = new Date(certificacion.fechaEmision);
    body.appendParagraph(`Firmado en fecha ${fechaFirma.getDate()} de ${obtenerNombreMesCompleto(fechaFirma.getMonth())} de ${fechaFirma.getFullYear()} por:`).editAsText().setBold(0, 16, true).setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    
    // Firma según plantilla
    const firmantePorPlantilla = obtenerFirmantePorPlantilla(certificacion.plantilla);
    
    body.appendParagraph('_'.repeat(35)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    body.appendParagraph(firmantePorPlantilla.nombre).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
    body.appendParagraph(firmantePorPlantilla.cargo).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
    body.appendParagraph('Cáritas Lima').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    
    // Footer de control
    const numeroSolicitud = codigoCertificacion.split('-')[2];
    const año = fechaFirma.getFullYear();
    const mesAbrev = obtenerNombreMes(fechaFirma.getMonth()).substring(0, 4);
    
    const controlText = `*Control electrónico con asunto - Re: FP 149 Aprobación cédula Solicitud ${numeroSolicitud} de ${año} ***COMPRA ADICIONAL ACEITE*** enviado por la Administración el ${fechaFirma.getDate()} ${mesAbrev} ${año}. ${fechaFirma.getHours()}:${fechaFirma.getMinutes().toString().padStart(2, '0')} ${fechaFirma.getHours() >= 12 ? 'p.m.' : 'a.m.'}*`;
    
    body.appendParagraph(controlText).editAsText().setFontSize(8);
    
    doc.saveAndClose();
    
    // Generar PDF
    let pdf;
    try {
      const pdfBlob = doc.getAs(MimeType.PDF);
      pdfBlob.setName(`Certificacion_${codigoCertificacion}.pdf`);
      pdf = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
    } catch (pdfError) {
      Logger.log('No se pudo crear el PDF en la carpeta configurada: ' + pdfError.toString());
      const fallbackBlob = doc.getAs(MimeType.PDF);
      fallbackBlob.setName(`Certificacion_${codigoCertificacion}.pdf`);
      pdf = DriveApp.createFile(fallbackBlob);
    }

    // URLs
    const urlDocumento = `https://docs.google.com/document/d/${doc.getId()}/edit`;
    const urlPDF = `https://drive.google.com/file/d/${pdf.getId()}/view`;
    const urlVistaPrevia = `https://docs.google.com/document/d/${doc.getId()}/preview`;
    
    // Actualizar URLs en certificación
    actualizarCertificacion(codigoCertificacion, {
      urlDocumento: urlDocumento,
      urlPDF: urlPDF
    });
    
    Logger.log(`Certificado generado exitosamente: ${codigoCertificacion}`);
    Logger.log(`URL Documento: ${urlDocumento}`);
    Logger.log(`URL PDF: ${urlPDF}`);
    
    // Registrar actividad
    registrarActividad('GENERAR_CERTIFICADO', `Código: ${codigoCertificacion}`);
    
    return {
      success: true,
      urlDocumento: urlDocumento,
      urlPDF: urlPDF,
      urlVistaPrevia: urlVistaPrevia,
      documentoId: doc.getId(),
      pdfId: pdf.getId()
    };
  } catch (error) {
    Logger.log('Error en generarDocumentoCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function obtenerCertificaciones(filtros = {}) {
  try {
    const sheet = ensureCertificacionesSheet();
    const data = getSheetValues(sheet);
    if (data.length <= 1) {
      return [];
    }

    const headerInfo = getCertificacionesHeaderInfo(sheet);
    const headerMap = headerInfo.map;

    const certificaciones = data
      .slice(1)
      .map((row, index) => {
        const codigo = sanitizeText(getCertificacionFieldValue(row, headerMap, 'codigo'));
        if (!codigo) {
          return null;
        }
        return mapRowToCertificacion(row, index, headerMap);
      })
      .filter(Boolean);

    const fechaDesde = toDate(filtros.fechaDesde);
    const fechaHasta = toDate(filtros.fechaHasta);

    const filtradas = certificaciones.filter(cert => {
      if (filtros.estado && cert.estado !== filtros.estado) return false;
      if (filtros.oficina && cert.oficina !== filtros.oficina) return false;

      if (fechaDesde) {
        const fechaCert = toDate(cert.fechaEmision || cert.fechaCreacion);
        if (fechaCert && fechaCert < fechaDesde) return false;
      }

      if (fechaHasta) {
        const fechaCert = toDate(cert.fechaEmision || cert.fechaCreacion);
        if (fechaCert && fechaCert > fechaHasta) return false;
      }

      if (filtros.busqueda) {
        const busqueda = sanitizeText(filtros.busqueda).toLowerCase();
        const coincideBusqueda =
          (cert.codigo || '').toLowerCase().includes(busqueda) ||
          (cert.descripcion || '').toLowerCase().includes(busqueda) ||
          (cert.solicitante || '').toLowerCase().includes(busqueda);
        if (!coincideBusqueda) {
          return false;
        }
      }

      return true;
    });

    const ordenadas = filtradas
      .slice()
      .sort((a, b) => {
        const fechaA = a.fechaEmisionTimestamp || a.fechaCreacionTimestamp || 0;
        const fechaB = b.fechaEmisionTimestamp || b.fechaCreacionTimestamp || 0;
        if (fechaA !== fechaB) {
          return fechaB - fechaA;
        }
        return (b.codigo || '').localeCompare(a.codigo || '');
      });

    return ordenadas;
  } catch (error) {
    Logger.log('Error en obtenerCertificaciones: ' + error.toString());
    throw new Error('No se pudieron obtener certificaciones: ' + error.message);
  }
}

function obtenerCertificacionPorCodigo(codigo) {
  try {
    const certificaciones = obtenerCertificaciones();
    const certificacion = certificaciones.find(c => c.codigo === codigo);
    
    if (certificacion) {
      // Obtener ítems
      certificacion.items = obtenerItemsCertificacion(codigo);
      // Obtener firmantes
      certificacion.firmantes = obtenerFirmantesCertificacion(codigo);
    }
    
    return certificacion;
  } catch (error) {
    Logger.log('Error en obtenerCertificacionPorCodigo: ' + error.toString());
    return null;
  }
}

function actualizarCertificacion(codigo, datos) {
  try {
    const sheet = ensureCertificacionesSheet();
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const filaIndex = findRowIndex(values, 0, codigo);

    if (filaIndex === -1) {
      return { success: false, error: 'Certificación no encontrada' };
    }

    const usuario = getActiveUserEmail();
    const fechaActual = new Date();
    const headerMap = buildCertificacionHeaderMap(values[0] || []);
    const fila = values[filaIndex];
    const finalidadDetalladaIndex = headerMap.finalidadDetallada;
    const puedeActualizarFinalidadDetallada = typeof finalidadDetalladaIndex === 'number' && finalidadDetalladaIndex >= 0;

    if (datos.fechaEmision !== undefined) {
      setCertificacionFieldValue(fila, headerMap, 'fechaEmision', parseDate(datos.fechaEmision));
    }
    if (datos.descripcion !== undefined) {
      const descripcion = sanitizeText(datos.descripcion);
      setCertificacionFieldValue(fila, headerMap, 'descripcion', descripcion);
      if (!datos.finalidad) {
        const finalidadAuto = generarFinalidadAutomatica(descripcion);
        setCertificacionFieldValue(fila, headerMap, 'finalidad', finalidadAuto);
        if (puedeActualizarFinalidadDetallada) {
          setCertificacionFieldValue(fila, headerMap, 'finalidadDetallada', finalidadAuto);
        }
      }
    }
    if (datos.finalidad !== undefined) {
      const finalidadActualizada = sanitizeText(datos.finalidad);
      setCertificacionFieldValue(fila, headerMap, 'finalidad', finalidadActualizada);
      if (puedeActualizarFinalidadDetallada) {
        setCertificacionFieldValue(fila, headerMap, 'finalidadDetallada', finalidadActualizada);
      }
    }
    if (datos.finalidadDetallada !== undefined && puedeActualizarFinalidadDetallada) {
      setCertificacionFieldValue(fila, headerMap, 'finalidadDetallada', sanitizeText(datos.finalidadDetallada));
    }
    if (datos.iniciativa !== undefined) setCertificacionFieldValue(fila, headerMap, 'iniciativa', sanitizeText(datos.iniciativa));
    if (datos.tipo !== undefined) setCertificacionFieldValue(fila, headerMap, 'tipo', sanitizeText(datos.tipo));
    if (datos.fuente !== undefined) setCertificacionFieldValue(fila, headerMap, 'fuente', sanitizeText(datos.fuente));
    if (datos.oficina !== undefined) setCertificacionFieldValue(fila, headerMap, 'oficina', sanitizeText(datos.oficina));
    if (datos.solicitante !== undefined) setCertificacionFieldValue(fila, headerMap, 'solicitante', sanitizeText(datos.solicitante));
    if (datos.cargoSolicitante !== undefined) setCertificacionFieldValue(fila, headerMap, 'cargoSolicitante', sanitizeText(datos.cargoSolicitante));
    if (datos.emailSolicitante !== undefined) setCertificacionFieldValue(fila, headerMap, 'emailSolicitante', sanitizeText(datos.emailSolicitante));
    if (datos.numeroAutorizacion !== undefined) setCertificacionFieldValue(fila, headerMap, 'numeroAutorizacion', sanitizeText(datos.numeroAutorizacion));
    if (datos.cargoAutorizador !== undefined) setCertificacionFieldValue(fila, headerMap, 'cargoAutorizador', sanitizeText(datos.cargoAutorizador));
    if (datos.estado !== undefined) setCertificacionFieldValue(fila, headerMap, 'estado', sanitizeText(datos.estado));
    if (datos.disposicion !== undefined) setCertificacionFieldValue(fila, headerMap, 'disposicion', sanitizeText(datos.disposicion));
    if (datos.plantilla !== undefined) setCertificacionFieldValue(fila, headerMap, 'plantilla', sanitizeText(datos.plantilla));
    if (datos.urlDocumento !== undefined) setCertificacionFieldValue(fila, headerMap, 'urlDocumento', sanitizeText(datos.urlDocumento));
    if (datos.urlPDF !== undefined) setCertificacionFieldValue(fila, headerMap, 'urlPDF', sanitizeText(datos.urlPDF));

    setCertificacionFieldValue(fila, headerMap, 'fechaModificacion', fechaActual);
    setCertificacionFieldValue(fila, headerMap, 'modificadoPor', usuario);

    if (datos.estado === ESTADOS.ANULADA) {
      setCertificacionFieldValue(fila, headerMap, 'fechaAnulacion', fechaActual);
      setCertificacionFieldValue(fila, headerMap, 'anuladoPor', usuario);
      setCertificacionFieldValue(fila, headerMap, 'motivoAnulacion', sanitizeText(datos.motivoAnulacion));
    }

    dataRange.setValues(values);
    SpreadsheetApp.flush();

    if (datos.items) {
      eliminarItemsCertificacion(codigo);
      crearItemsCertificacion(codigo, datos.items);
    }

    if (datos.firmantes) {
      eliminarFirmantesCertificacion(codigo);
      crearFirmantesCertificacion(codigo, datos.firmantes);
    }

    recalcularTotalesCertificacion(codigo);
    registrarActividad('ACTUALIZAR_CERTIFICACION', `Código: ${codigo}`);

    return { success: true };
  } catch (error) {
    Logger.log('Error en actualizarCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ===============================================
// FUNCIONES DE CONFIGURACIÓN
// ===============================================

function obtenerSolicitantes() {
  try {
    let sheet;
    try {
      sheet = getSheetOrThrow(SHEET_NAMES.CONFIG_SOLICITANTES);
    } catch (error) {
      crearHojaConfigSolicitantes();
      sheet = getSheetOrThrow(SHEET_NAMES.CONFIG_SOLICITANTES);
    }
    const data = getSheetValues(sheet);
    if (data.length <= 1) return [];

    const solicitantes = data.slice(1)
      .filter(row => row[0])
      .map(row => ({
        id: row[0],
        nombre: row[1],
        cargo: row[2],
        email: row[3],
        activo: row[4] !== false
      }));

    return solicitantes.filter(s => s.activo);
  } catch (error) {
    Logger.log('Error en obtenerSolicitantes: ' + error.toString());
    return [];
  }
}

function obtenerSolicitantePorId(id) {
  try {
    const solicitantes = obtenerSolicitantes();
    return solicitantes.find(s => s.id === id) || null;
  } catch (error) {
    Logger.log('Error en obtenerSolicitantePorId: ' + error.toString());
    return null;
  }
}

function obtenerConfiguracionGeneral() {
  try {
    let sheet;
    try {
      sheet = getSheetOrThrow(SHEET_NAMES.CONFIG_GENERAL);
    } catch (error) {
      crearHojaConfigGeneral();
      sheet = getSheetOrThrow(SHEET_NAMES.CONFIG_GENERAL);
    }

    const data = getSheetValues(sheet);
    const config = {};

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] && row[1]) {
        config[row[0]] = row[1];
      }
    }
    
    return config;
  } catch (error) {
    Logger.log('Error en obtenerConfiguracionGeneral: ' + error.toString());
    return {};
  }
}

function obtenerDisposicionPorDefecto() {
  const config = obtenerConfiguracionGeneral();
  return config.disposicion_base_legal || 'Directiva 003-2023-SG/CARITASLIMA, Directiva de contratación de bienes y servicios de la Vicaría de Pastoral Social y Dignidad Humana - Caritas Lima';
}

function obtenerFirmantes() {
  try {
    let sheet;
    try {
      sheet = getSheetOrThrow(SHEET_NAMES.CONFIG_FIRMANTES);
    } catch (error) {
      crearHojaConfigFirmantes();
      sheet = getSheetOrThrow(SHEET_NAMES.CONFIG_FIRMANTES);
    }

    const data = getSheetValues(sheet);
    if (data.length <= 1) return [];

    const firmantes = data.slice(1)
      .filter(row => row[0])
      .map(row => ({
        id: row[0],
        nombre: row[1],
        cargo: row[2],
        orden: row[3] || 1,
        activo: row[4] !== false
      }));

    return firmantes.filter(f => f.activo).sort((a, b) => a.orden - b.orden);
  } catch (error) {
    Logger.log('Error en obtenerFirmantes: ' + error.toString());
    return [];
  }
}

function obtenerFirmantePorDefecto() {
  const firmantes = obtenerFirmantes();
  const firmantePorDefecto = firmantes.find(f => f.orden === 1);
  
  if (firmantePorDefecto) {
    return {
      nombre: firmantePorDefecto.nombre,
      cargo: firmantePorDefecto.cargo,
      obligatorio: true
    };
  }
  
  return {
    nombre: 'Evelyn Elena Huaycacllo Marin',
    cargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto',
    obligatorio: true
  };
}

// ===============================================
// GENERACIÓN DE FINALIDAD CON IA
// ===============================================

function generarFinalidadConIA(payload) {
  let descripcionTexto = '';
  try {
    const esObjeto = typeof payload === 'object' && payload !== null;
    descripcionTexto = esObjeto ? payload.descripcion : payload;
    if (!descripcionTexto) {
      return {
        success: false,
        finalidad: 'Complementar necesidades operativas de la institución.'
      };
    }

    const detalles = [];
    if (esObjeto) {
      if (payload.iniciativa) detalles.push(`Iniciativa: ${payload.iniciativa}`);
      if (payload.tipo) detalles.push(`Tipo de gasto: ${payload.tipo}`);
      if (payload.fuente) detalles.push(`Fuente de financiamiento: ${payload.fuente}`);
      if (payload.oficina) detalles.push(`Oficina solicitante: ${payload.oficina}`);
      if (payload.montoEstimado) detalles.push(`Monto estimado: S/ ${payload.montoEstimado}`);
    }

    const contextoAdicional = detalles.length ? `\nDATOS ADICIONALES:\n- ${detalles.join('\n- ')}` : '';

    const prompt = `Basándote en la siguiente descripción de una certificación presupuestal de Cáritas Lima, genera una FINALIDAD concisa y específica:

DESCRIPCIÓN: "${descripcionTexto}"
${contextoAdicional}

EJEMPLOS de finalidades correctas:
- "Complementar con productos adicionales la conformación de los kits de ollas"
- "Contar con implementos adecuados que faciliten el desarrollo de las actividades"
- "Fortalecer el área de comunicaciones mediante la implementación de recursos tecnológicos"
- "Garantizar el traslado oportuno y seguro de las donaciones"
- "Garantizar que las personas beneficiarias reciban una nutrición adecuada y oportuna"

INSTRUCCIONES:
1. La finalidad debe ser específica al tipo de gasto descrito
2. Debe iniciar con un verbo (Complementar, Contar, Fortalecer, Garantizar, Mejorar, etc.)
3. Debe explicar el propósito específico, no ser genérica
4. Debe estar alineada con la misión social de Cáritas Lima
5. Máximo 2 líneas de texto
6. Sin puntuación final

Responde SOLO con la finalidad, sin explicaciones adicionales.`;

    const requestBody = {
      model: CONFIG.AI_MODEL,
      messages: [
        {
          role: "system",
          content: "Eres un experto en redacción de finalidades para certificaciones presupuestales de organizaciones sin fines de lucro católicas."
        },
        {
          role: "user",
          content: prompt
        }
      ],
      max_tokens: 150,
      temperature: 0.3
    };

    const options = {
      method: 'POST',
      headers: {
        'CustomerId': CONFIG.CUSTOMER_ID,
        'Content-Type': 'application/json',
        'Authorization': 'Bearer xxx'
      },
      payload: JSON.stringify(requestBody)
    };

    const response = UrlFetchApp.fetch(CONFIG.AI_ENDPOINT, options);
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.choices && responseData.choices.length > 0) {
      let finalidad = responseData.choices[0].message.content.trim();
      finalidad = finalidad.replace(/^["']|["']$/g, '');
      finalidad = finalidad.replace(/\.$/, '');
      
      return {
        success: true,
        finalidad: finalidad
      };
    } else {
      throw new Error('Respuesta inválida de la API de AI');
    }
  } catch (error) {
    Logger.log('Error en generarFinalidadConIA: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      finalidad: generarFinalidadAutomatica(descripcionTexto)
    };
  }
}

function generarFinalidadAutomatica(descripcion) {
  if (!descripcion) return 'Complementar necesidades operativas de la institución';
  
  const desc = descripcion.toLowerCase();
  
  if (desc.includes('kit') && (desc.includes('olla') || desc.includes('alimento'))) {
    return 'Complementar con productos adicionales la conformación de los kits de ollas';
  } else if (desc.includes('adquisic') && desc.includes('productos adicionales')) {
    return 'Complementar con productos adicionales la conformación de los kits de ollas';
  } else if (desc.includes('alimento') || desc.includes('nutrición') || desc.includes('comida')) {
    return 'Garantizar que las personas beneficiarias de las actividades programadas reciban una nutrición adecuada y oportuna';
  } else if (desc.includes('transporte') || desc.includes('traslado')) {
    return 'Garantizar el traslado oportuno y seguro de las donaciones provenientes del CIFO';
  } else if (desc.includes('equipo') || desc.includes('implemento')) {
    return 'Contar con implementos adecuados que faciliten el desarrollo de las actividades';
  } else if (desc.includes('tecnolog') || desc.includes('comunicacion')) {
    return 'Fortalecer el área de comunicaciones mediante la implementación de recursos tecnológicos';
  } else if (desc.includes('mantenimiento') || desc.includes('reparación')) {
    return 'Garantizar su adecuado funcionamiento, prolongar su vida útil y asegurar condiciones óptimas de seguridad';
  } else if (desc.includes('capacitación') || desc.includes('formación')) {
    return 'Fortalecer las capacidades del personal para mejorar la calidad de atención';
  } else if (desc.includes('mobiliario') || desc.includes('silla')) {
    return 'Mejorar las condiciones de trabajo, promover el cuidado de la salud postural del personal';
  } else if (desc.includes('inventario') || desc.includes('registro')) {
    return 'Garantizar un adecuado registro, verificación y actualización de los bienes';
  }
  
  return 'Complementar las necesidades operativas para el cumplimiento efectivo de la misión institucional';
}

// ===============================================
// CÓDIGO CONSECUTIVO
// ===============================================

function generarCodigoCertificacionConsecutivo() {
  try {
    const año = new Date().getFullYear();
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.CERTIFICACIONES);

    if (!sheet) {
      return `CP-${año}-0001`;
    }

    const data = sheet.getDataRange().getValues();
    
    // Buscar el último número consecutivo del año
    let ultimoNumero = 0;
    for (let i = 1; i < data.length; i++) {
      const codigo = data[i][0];
      if (codigo && codigo.toString().includes(`CP-${año}-`)) {
        const partes = codigo.split('-');
        if (partes.length >= 3) {
          const numero = parseInt(partes[2]);
          if (numero > ultimoNumero) {
            ultimoNumero = numero;
          }
        }
      }
    }
    
    const siguienteNumero = ultimoNumero + 1;
    const numeroFormateado = siguienteNumero.toString().padStart(4, '0');
    
    return `CP-${año}-${numeroFormateado}`;
  } catch (error) {
    Logger.log('Error en generarCodigoCertificacionConsecutivo: ' + error.toString());
    const año = new Date().getFullYear();
    return `CP-${año}-0001`;
  }
}

// ===============================================
// GESTIÓN DE ÍTEMS
// ===============================================

function crearItemsCertificacion(codigoCertificacion, items) {
  try {
    const sheet = ensureItemsSheet();

    items.forEach((item, index) => {
      const subtotal = (item.cantidad || 0) * (item.precioUnitario || 0);
      const fila = [
        codigoCertificacion,
        index + 1,
        item.descripcion || '',
        item.cantidad || 0,
        item.unidad || 'Unidad',
        item.precioUnitario || 0,
        subtotal,
        new Date(),
        getActiveUserEmail()
      ];
      sheet.appendRow(fila);
    });

    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log('Error en crearItemsCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function obtenerItemsCertificacion(codigoCertificacion) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.ITEMS);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    const items = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === codigoCertificacion) {
        items.push({
          codigoCertificacion: row[0],
          orden: row[1],
          descripcion: row[2],
          cantidad: row[3],
          unidad: row[4],
          precioUnitario: row[5],
          subtotal: row[6],
          fechaCreacion: row[7],
          creadoPor: row[8]
        });
      }
    }
    
    return items.sort((a, b) => a.orden - b.orden);
  } catch (error) {
    Logger.log('Error en obtenerItemsCertificacion: ' + error.toString());
    return [];
  }
}

function eliminarItemsCertificacion(codigoCertificacion) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.ITEMS);

    if (!sheet) return { success: true };

    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === codigoCertificacion) {
        sheet.deleteRow(i + 1);
      }
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('Error en eliminarItemsCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ===============================================
// GESTIÓN DE FIRMANTES
// ===============================================

function crearFirmantesCertificacion(codigoCertificacion, firmantes) {
  try {
    const sheet = ensureFirmantesSheet();

    firmantes.forEach((firmante, index) => {
      const fila = [
        codigoCertificacion,
        index + 1,
        firmante.nombre || '',
        firmante.cargo || '',
        firmante.obligatorio || false,
        new Date(),
        getActiveUserEmail()
      ];
      sheet.appendRow(fila);
    });

    SpreadsheetApp.flush();

    return { success: true };
  } catch (error) {
    Logger.log('Error en crearFirmantesCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function obtenerFirmantesCertificacion(codigoCertificacion) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.FIRMANTES);

    if (!sheet) return [];

    const data = sheet.getDataRange().getValues();
    
    if (data.length <= 1) return [];
    
    const firmantes = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (row[0] === codigoCertificacion) {
        firmantes.push({
          codigoCertificacion: row[0],
          orden: row[1],
          nombre: row[2],
          cargo: row[3],
          obligatorio: row[4],
          fechaCreacion: row[5],
          creadoPor: row[6]
        });
      }
    }
    
    return firmantes.sort((a, b) => a.orden - b.orden);
  } catch (error) {
    Logger.log('Error en obtenerFirmantesCertificacion: ' + error.toString());
    return [];
  }
}

function eliminarFirmantesCertificacion(codigoCertificacion) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.FIRMANTES);

    if (!sheet) return { success: true };

    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === codigoCertificacion) {
        sheet.deleteRow(i + 1);
      }
    }
    
    return { success: true };
  } catch (error) {
    Logger.log('Error en eliminarFirmantesCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function crearFirmantesBasadosEnPlantilla(codigoCertificacion, plantillaId) {
  try {
    const firmantePrincipal = obtenerFirmantePorPlantilla(plantillaId);
    const firmantes = [];

    if (firmantePrincipal && (firmantePrincipal.nombre || firmantePrincipal.cargo)) {
      firmantes.push({
        nombre: firmantePrincipal.nombre || '',
        cargo: firmantePrincipal.cargo || '',
        obligatorio: true
      });
    }

    if (firmantes.length === 0) {
      return { success: true };
    }

    return crearFirmantesCertificacion(codigoCertificacion, firmantes);
  } catch (error) {
    Logger.log('Error en crearFirmantesBasadosEnPlantilla: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function obtenerFirmantePorPlantilla(plantillaId) {
  const plantilla = getPlantillaConfigurada(plantillaId);
  if (plantilla) {
    const nombre = sanitizeText(plantilla.firmanteNombre);
    const cargo = sanitizeText(plantilla.firmanteCargo);

    if (nombre || cargo) {
      return {
        nombre: nombre || (PLANTILLA_FIRMANTES[plantillaId] && PLANTILLA_FIRMANTES[plantillaId].nombre) || 'Firmante Principal',
        cargo: cargo || (PLANTILLA_FIRMANTES[plantillaId] && PLANTILLA_FIRMANTES[plantillaId].cargo) || ''
      };
    }
  }

  return PLANTILLA_FIRMANTES[plantillaId] || PLANTILLA_FIRMANTES['plantilla_evelyn'];
}

// ===============================================
// CATÁLOGOS
// ===============================================

function obtenerCatalogo(tipo) {
  try {
    const ss = getSpreadsheet();
    const nombreHoja = {
      iniciativas: SHEET_NAMES.CATALOGO_INICIATIVAS,
      tipos: SHEET_NAMES.CATALOGO_TIPOS,
      fuentes: SHEET_NAMES.CATALOGO_FUENTES,
      finalidades: SHEET_NAMES.CATALOGO_FINALIDADES,
      oficinas: SHEET_NAMES.CATALOGO_OFICINAS,
      plantillas: SHEET_NAMES.PLANTILLAS
    }[tipo];

    if (tipo === 'solicitantes') return obtenerSolicitantes();
    if (tipo === 'firmantes') return obtenerFirmantes();
    if (tipo === 'plantillas') return obtenerPlantillasConfiguradas({ soloActivas: false });
    if (!nombreHoja) return [];

    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) {
      Logger.log(`La hoja "${nombreHoja}" no existe.`);
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    
    const catalogo = [];
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) continue;

      catalogo.push({
        codigo: row[0],
        nombre: row[1],
        descripcion: row[2] || '',
        activo: row[3] !== false
      });
    }

    return catalogo.filter(item => item.activo);
  } catch (error) {
    Logger.log('Error en obtenerCatalogo: ' + error.toString());
    return [];
  }
}

function buildPlantillaHeaderMap(headers) {
  const normalizedHeaders = headers.map(normalizeHeaderName);
  const map = {};

  PLANTILLAS_COLUMN_DEFINITIONS.forEach(definition => {
    const aliases = Array.isArray(definition.aliases)
      ? definition.aliases.map(normalizeHeaderName)
      : [];

    let index = -1;
    for (let i = 0; i < aliases.length; i++) {
      const aliasIndex = normalizedHeaders.indexOf(aliases[i]);
      if (aliasIndex !== -1) {
        index = aliasIndex;
        break;
      }
    }

    if (index === -1 && definition.defaultIndex >= 0 && definition.defaultIndex < headers.length) {
      index = definition.defaultIndex;
    }

    map[definition.field] = index;
  });

  return map;
}

function getPlantillaFieldValue(row, headerMap, field, fallback = '') {
  if (!headerMap || typeof headerMap[field] !== 'number') {
    return fallback;
  }

  const index = headerMap[field];
  if (index >= 0 && index < row.length) {
    return row[index];
  }

  return fallback;
}

function extractGoogleResourceId(value) {
  const texto = value === null || value === undefined ? '' : String(value).trim();
  if (!texto) {
    return '';
  }

  const match = texto.match(/[\w-]{25,}/);
  return match ? match[0] : texto;
}

function parseBoolean(value, fallback = false) {
  if (typeof value === 'boolean') {
    return value;
  }

  if (value === null || value === undefined) {
    return fallback;
  }

  const normalized = String(value).trim().toLowerCase();
  if (!normalized) return fallback;

  return ['true', '1', 'si', 'sí', 'activo', 'activa', 'yes'].indexOf(normalized) !== -1;
}

function obtenerPlantillasConfiguradas({ soloActivas = false } = {}) {
  try {
    const ahora = Date.now();
    if (
      plantillasCache &&
      ahora - plantillasCacheTimestamp < PLANTILLAS_CACHE_TTL_MS &&
      !soloActivas
    ) {
      return plantillasCache.slice();
    }

    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.PLANTILLAS);
    if (!sheet) {
      Logger.log('La hoja "Plantillas" no existe.');
      return [];
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      plantillasCache = [];
      plantillasCacheTimestamp = ahora;
      return [];
    }

    const headers = data[0] || [];
    const headerMap = buildPlantillaHeaderMap(headers);

    const plantillas = data.slice(1).map(row => {
      const id = sanitizeText(getPlantillaFieldValue(row, headerMap, 'id'));
      if (!id) {
        return null;
      }

      const nombre = sanitizeText(getPlantillaFieldValue(row, headerMap, 'nombre'), id);
      const descripcion = sanitizeText(getPlantillaFieldValue(row, headerMap, 'descripcion'));
      const activa = parseBoolean(getPlantillaFieldValue(row, headerMap, 'activa'), true);
      const firmantesRaw = parseNumber(getPlantillaFieldValue(row, headerMap, 'firmantes'), 1);
      const firmantes = firmantesRaw > 0 ? Math.min(Math.round(firmantesRaw), 5) : 1;
      const plantillaHtml = sanitizeText(getPlantillaFieldValue(row, headerMap, 'plantillaHtml'));
      const docId = extractGoogleResourceId(plantillaHtml);
      const firmanteId = sanitizeText(getPlantillaFieldValue(row, headerMap, 'firmanteId'));
      const firmanteNombre = sanitizeText(getPlantillaFieldValue(row, headerMap, 'firmanteNombre'));
      const firmanteCargo = sanitizeText(getPlantillaFieldValue(row, headerMap, 'firmanteCargo'));

      return {
        id,
        nombre,
        descripcion,
        activa,
        firmantes,
        docId,
        plantillaHtml,
        firmanteId,
        firmanteNombre,
        firmanteCargo
      };
    }).filter(Boolean);

    const plantillasOrdenadas = plantillas
      .slice()
      .sort((a, b) => {
        if (a.activa !== b.activa) {
          return a.activa ? -1 : 1;
        }
        return a.nombre.localeCompare(b.nombre || '', 'es', { sensitivity: 'base' });
      });

    plantillasCache = plantillasOrdenadas;
    plantillasCacheTimestamp = ahora;

    return soloActivas
      ? plantillasOrdenadas.filter(plantilla => plantilla.activa)
      : plantillasOrdenadas;
  } catch (error) {
    Logger.log('Error en obtenerPlantillasConfiguradas: ' + error.toString());
    return [];
  }
}

function getPlantillaConfigurada(plantillaId) {
  if (!plantillaId) {
    return null;
  }

  const plantillas = obtenerPlantillasConfiguradas();
  return plantillas.find(plantilla => plantilla.id === plantillaId) || null;
}

function invalidarCachePlantillas() {
  plantillasCache = null;
  plantillasCacheTimestamp = 0;
}

function getCertificadosFolder() {
  if (certificadosFolderCache) {
    return certificadosFolderCache;
  }

  try {
    certificadosFolderCache = DriveApp.getFolderById(CONFIG.CARPETA_CERTIFICADOS);
  } catch (error) {
    Logger.log('No se pudo acceder a la carpeta de certificados: ' + error.toString());
    certificadosFolderCache = null;
  }

  return certificadosFolderCache;
}

// ===============================================
// ESTADÍSTICAS DEL DASHBOARD
// ===============================================

function obtenerEstadisticasDashboard() {
  try {
    const certificaciones = obtenerCertificaciones();

    if (!certificaciones || certificaciones.length === 0) {
      return {
        success: true,
        data: {
          total: 0,
          montoTotal: 0,
          porEstado: {
            'Borrador': 0,
            'En revisión': 0,
            'Autorización pendiente': 0,
            'Activa': 0,
            'Anulada': 0
          },
          porOficina: {},
          certificacionesRecientes: []
        }
      };
    }

    const estadisticas = {
      total: certificaciones.length,
      montoTotal: certificaciones.reduce((sum, cert) => sum + parseNumber(cert.montoTotal, 0), 0),
      porEstado: {
        'Borrador': 0,
        'En revisión': 0,
        'Autorización pendiente': 0,
        'Activa': 0,
        'Anulada': 0
      },
      porOficina: {},
      certificacionesRecientes: certificaciones.slice(0, 10).map(cert => ({ ...cert }))
    };

    // Contar por estado
    certificaciones.forEach(cert => {
      if (estadisticas.porEstado.hasOwnProperty(cert.estado)) {
        estadisticas.porEstado[cert.estado]++;
      }

      // Contar por oficina
      const nombreOficina = obtenerNombreOficina(cert.oficina);
      estadisticas.porOficina[nombreOficina] = (estadisticas.porOficina[nombreOficina] || 0) + 1;
    });

    return {
      success: true,
      data: estadisticas
    };
  } catch (error) {
    Logger.log('Error en obtenerEstadisticasDashboard: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      data: {
        total: 0,
        montoTotal: 0,
        porEstado: {},
        porOficina: {},
        certificacionesRecientes: []
      }
    };
  }
}

// ===============================================
// FUNCIONES DE CÁLCULO Y UTILIDADES
// ===============================================

function recalcularTotalesCertificacion(codigoCertificacion) {
  try {
    const items = obtenerItemsCertificacion(codigoCertificacion);
    const total = items.reduce((sum, item) => sum + parseNumber(item.subtotal, 0), 0);
    const montoLetras = convertirNumeroALetrasTexto(total);

    const sheet = ensureCertificacionesSheet();
    const headerInfo = getCertificacionesHeaderInfo(sheet);
    const headerMap = headerInfo.map;
    const data = sheet.getDataRange().getValues();

    const montoTotalCol = typeof headerMap.montoTotal === 'number' ? headerMap.montoTotal + 1 : 16;
    const montoLetrasCol = typeof headerMap.montoLetras === 'number' ? headerMap.montoLetras + 1 : 17;

    for (let i = 1; i < data.length; i++) {
      const fila = data[i];
      const codigoFila = sanitizeText(getCertificacionFieldValue(fila, headerMap, 'codigo'));
      if (codigoFila === codigoCertificacion) {
        sheet.getRange(i + 1, montoTotalCol).setValue(total);
        sheet.getRange(i + 1, montoLetrasCol).setValue(montoLetras);
        break;
      }
    }

    SpreadsheetApp.flush();

    return { success: true, total: total, montoLetras: montoLetras };
  } catch (error) {
    Logger.log('Error en recalcularTotalesCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function convertirNumeroALetrasTexto(numero) {
  const cantidad = parseNumber(numero, 0);
  if (cantidad === 0) {
    return 'CERO CON 00/100 SOLES';
  }

  const unidades = ['', 'UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE'];
  const decenas = ['', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
  const especiales = ['DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE'];
  const centenas = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];

  const entero = Math.floor(cantidad);
  const decimales = Math.round((cantidad - entero) * 100);

  function convertirGrupo(num) {
    if (num === 0) return '';

    let resultado = '';
    const c = Math.floor(num / 100);
    const d = Math.floor((num % 100) / 10);
    const u = num % 10;

    if (c > 0) {
      if (c === 1 && d === 0 && u === 0) {
        resultado += 'CIEN';
      } else {
        resultado += centenas[c];
      }
    }

    if (d > 0) {
      if (resultado) resultado += ' ';

      if (d === 1) {
        resultado += especiales[u];
        return resultado;
      }

      if (d === 2 && u > 0) {
        resultado += 'VEINTI';
      } else {
        resultado += decenas[d];
      }
    }

    if (u > 0 && d !== 1) {
      if (resultado) resultado += ' ';
      if (d === 2) {
        resultado += unidades[u].toLowerCase();
      } else {
        resultado += unidades[u];
      }
    }

    if (!resultado) {
      resultado = unidades[u];
    }

    return resultado.trim();
  }

  function convertirMiles(num) {
    if (num === 0) return '';

    const miles = Math.floor(num / 1000);
    const resto = num % 1000;
    let resultado = '';

    if (miles > 0) {
      if (miles === 1) {
        resultado += 'MIL';
      } else {
        resultado += `${convertirGrupo(miles)} MIL`;
      }
    }

    if (resto > 0) {
      if (resultado) resultado += ' ';
      resultado += convertirGrupo(resto);
    }

    return resultado.trim();
  }

  function convertirMillones(num) {
    if (num === 0) return '';

    const millones = Math.floor(num / 1000000);
    const resto = num % 1000000;
    let resultado = '';

    if (millones > 0) {
      if (millones === 1) {
        resultado += 'UN MILLÓN';
      } else {
        resultado += `${convertirMiles(millones)} MILLONES`;
      }
    }

    if (resto > 0) {
      if (resultado) resultado += ' ';
      resultado += convertirMiles(resto);
    }

    return resultado.trim();
  }

  const letras = convertirMillones(entero);
  const decimalesTexto = decimales.toString().padStart(2, '0');

  return `${letras || 'CERO'} CON ${decimalesTexto}/100 SOLES`;
}

function convertirNumeroALetras(numero) {
  try {
    const montoEnLetras = convertirNumeroALetrasTexto(numero);
    return {
      success: true,
      montoLetras: montoEnLetras
    };
  } catch (error) {
    Logger.log('Error en convertirNumeroALetras: ' + error.toString());
    return {
      success: false,
      error: error.toString(),
      montoLetras: 'CERO CON 00/100 SOLES'
    };
  }
}

// ===============================================
// FUNCIONES DE UTILIDAD
// ===============================================

function registrarActividad(accion, detalles = '') {
  try {
    const sheet = ensureBitacoraSheet();

    const usuario = getActiveUserEmail();
    const fecha = new Date();

    const fila = [
      fecha,
      usuario,
      accion,
      detalles,
      usuario
    ];
    
    sheet.appendRow(fila);
    SpreadsheetApp.flush();
  } catch (error) {
    Logger.log('Error en registrarActividad: ' + error.toString());
  }
}

function formatearFechaDocumento(fecha) {
  const f = new Date(fecha);
  return `${f.getDate().toString().padStart(2, '0')}/${(f.getMonth() + 1).toString().padStart(2, '0')}/${f.getFullYear()}`;
}

function obtenerNombreMes(numeroMes) {
  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return meses[numeroMes] || '';
}

function obtenerNombreMesCompleto(numeroMes) {
  const meses = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio',
                 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  return meses[numeroMes] || '';
}

function obtenerNombreCatalogo(tipo, codigo) {
  try {
    const catalogo = obtenerCatalogo(tipo);
    const item = catalogo.find(i => i.codigo === codigo);
    return item ? item.nombre : codigo;
  } catch (error) {
    return codigo || '';
  }
}

function obtenerNombreOficina(codigo) {
  try {
    const oficinas = obtenerCatalogo('oficinas');
    const oficina = oficinas.find(o => o.codigo === codigo);
    return oficina ? oficina.nombre : codigo;
  } catch (error) {
    return codigo || 'Sin oficina';
  }
}

// ===============================================
// FUNCIONES PARA CREAR HOJAS DE CONFIGURACIÓN
// ===============================================

function crearHojaConfigSolicitantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config_Solicitantes');
  
  if (sheet) return sheet;
  
  sheet = ss.insertSheet('Config_Solicitantes');
  
  const headers = ['ID', 'Nombre Completo', 'Cargo', 'Email', 'Activo'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  
  const solicitantesDefecto = [
    ['SOL001', 'Evelyn Elena Huaycacllo Marin', 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', 'evelyn.huaycacllo@caritaslima.org', true],
    ['SOL002', 'Guadalupe Susana Callupe Pacheco', 'Coordinadora de Logística', 'guadalupe.callupe@caritaslima.org', true]
  ];
  
  sheet.getRange(2, 1, solicitantesDefecto.length, 5).setValues(solicitantesDefecto);
  
  Logger.log('Hoja Config_Solicitantes creada');
  return sheet;
}

function crearHojaConfigFirmantes() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config_Firmantes');
  
  if (sheet) return sheet;
  
  sheet = ss.insertSheet('Config_Firmantes');
  
  const headers = ['ID', 'Nombre Completo', 'Cargo', 'Orden', 'Activo'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  
  const firmantesDefecto = [
    ['FIR001', 'Evelyn Elena Huaycacllo Marin', 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', 1, true],
    ['FIR002', 'Padre Miguel Ángel Castillo Seminario', 'Director Ejecutivo', 2, true]
  ];
  
  sheet.getRange(2, 1, firmantesDefecto.length, 5).setValues(firmantesDefecto);
  
  Logger.log('Hoja Config_Firmantes creada');
  return sheet;
}

function crearHojaConfigGeneral() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Config_General');
  
  if (sheet) return sheet;
  
  sheet = ss.insertSheet('Config_General');
  
  const headers = ['Configuración', 'Valor'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  
  sheet.setFrozenRows(1);
  
  const configuracionDefecto = [
    ['disposicion_base_legal', 'Directiva 003-2023-SG/CARITASLIMA, Directiva de contratación de bienes y servicios de la Vicaría de Pastoral Social y Dignidad Humana - Caritas Lima']
  ];
  
  sheet.getRange(2, 1, configuracionDefecto.length, 2).setValues(configuracionDefecto);
  
  Logger.log('Hoja Config_General creada');
  return sheet;
}

// ===============================================
// FUNCIONES DE GESTIÓN DE CONFIGURACIÓN
// ===============================================

function actualizarSolicitante(id, datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Config_Solicitantes');
    
    if (!sheet) {
      crearHojaConfigSolicitantes();
      sheet = ss.getSheetByName('Config_Solicitantes');
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let filaIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        filaIndex = i;
        break;
      }
    }
    
    if (filaIndex === -1) {
      const fila = [id, datos.nombre, datos.cargo, datos.email, datos.activo !== false];
      sheet.appendRow(fila);
    } else {
      values[filaIndex][1] = datos.nombre;
      values[filaIndex][2] = datos.cargo;
      values[filaIndex][3] = datos.email;
      values[filaIndex][4] = datos.activo !== false;
      dataRange.setValues(values);
    }
    
    registrarActividad('ACTUALIZAR_SOLICITANTE', `ID: ${id}, Nombre: ${datos.nombre}`);
    return { success: true };
  } catch (error) {
    Logger.log('Error en actualizarSolicitante: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function actualizarFirmante(id, datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Config_Firmantes');
    
    if (!sheet) {
      crearHojaConfigFirmantes();
      sheet = ss.getSheetByName('Config_Firmantes');
    }
    
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    
    let filaIndex = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        filaIndex = i;
        break;
      }
    }
    
    if (filaIndex === -1) {
      const fila = [id, datos.nombre, datos.cargo, datos.orden || 1, datos.activo !== false];
      sheet.appendRow(fila);
    } else {
      values[filaIndex][1] = datos.nombre;
      values[filaIndex][2] = datos.cargo;
      values[filaIndex][3] = datos.orden || 1;
      values[filaIndex][4] = datos.activo !== false;
      dataRange.setValues(values);
    }
    
    registrarActividad('ACTUALIZAR_FIRMANTE', `ID: ${id}, Nombre: ${datos.nombre}`);
    return { success: true };
  } catch (error) {
    Logger.log('Error en actualizarFirmante: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function eliminarSolicitante(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config_Solicitantes');
    
    if (!sheet) {
      return { success: false, error: 'Hoja de solicitantes no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    registrarActividad('ELIMINAR_SOLICITANTE', `ID: ${id}`);
    return { success: true };
  } catch (error) {
    Logger.log('Error en eliminarSolicitante: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function eliminarFirmante(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Config_Firmantes');
    
    if (!sheet) {
      return { success: false, error: 'Hoja de firmantes no encontrada' };
    }
    
    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === id) {
        sheet.deleteRow(i + 1);
        break;
      }
    }
    
    registrarActividad('ELIMINAR_FIRMANTE', `ID: ${id}`);
    return { success: true };
  } catch (error) {
    Logger.log('Error en eliminarFirmante: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ===============================================
// FUNCIONES DE TESTING
// ===============================================

function testCrearCertificacionCompleta() {
  try {
    const datosTest = {
      fechaCertificacion: '2025-01-27',
      descripcion: 'Adquisición de productos adicionales para completar los kits de ollas (AZÚCAR y ACEITE)',
      iniciativa: 'INI001',
      tipo: 'TIP001',
      fuente: 'FUE001',
      oficina: 'OFI001',
      solicitanteId: 'SOL002',
      plantilla: 'plantilla_evelyn',
      finalidad: 'Complementar con productos adicionales la conformación de los kits de ollas',
      items: [
        {
          descripcion: 'AZÚCAR CARTAVIO RUBIA GRANEL',
          cantidad: 25,
          unidad: 'Kilogramos',
          precioUnitario: 4.20
        },
        {
          descripcion: 'ACEITE VEGA BOTELLA 1 LITRO',
          cantidad: 30,
          unidad: 'Unidades',
          precioUnitario: 8.50
        }
      ]
    };
    
    Logger.log('Datos de prueba preparados');
    const resultado = crearCertificacion(datosTest);
    
    if (resultado.success) {
      Logger.log('=== CERTIFICACIÓN CREADA EXITOSAMENTE ===');
      Logger.log('Código: ' + resultado.codigo);
      
      if (resultado.urls && resultado.urls.documento) {
        Logger.log('URL Documento: ' + resultado.urls.documento);
        Logger.log('URL PDF: ' + resultado.urls.pdf);
      }
    } else {
      Logger.log('=== ERROR EN CREACIÓN ===');
      Logger.log('Error: ' + resultado.error);
    }
    
    return resultado;
  } catch (error) {
    Logger.log('Error en testCrearCertificacionCompleta: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

// ===============================================
// GENERADOR DE CERTIFICADO PERFECTO (COMO EL QUE MOSTRASTE)
// ===============================================

function generarCertificadoPerfecto(codigoCertificacion) {
  try {
    const certificacion = obtenerCertificacionPorCodigo(codigoCertificacion);
    if (!certificacion) {
      return { success: false, error: 'Certificación no encontrada' };
    }
    
    // Crear documento con el formato EXACTO del que me mostraste
    const doc = DocumentApp.create(`Certificacion_${codigoCertificacion}`);
    const body = doc.getBody();

    const folder = getCertificadosFolder();
    if (folder) {
      try {
        const docFile = DriveApp.getFileById(doc.getId());
        folder.addFile(docFile);
        const parents = docFile.getParents();
        const folderId = folder.getId();
        const padresAEliminar = [];
        while (parents.hasNext()) {
          const parent = parents.next();
          if (parent.getId() !== folderId) {
            padresAEliminar.push(parent);
          }
        }
        padresAEliminar.forEach(parent => parent.removeFile(docFile));
      } catch (folderError) {
        Logger.log('No se pudo mover el documento perfecto a la carpeta configurada: ' + folderError.toString());
      }
    }
    
    // Configurar márgenes
    body.setMarginTop(50);
    body.setMarginBottom(50);
    body.setMarginLeft(60);
    body.setMarginRight(60);
    
    // Header con logo real de Cáritas Lima
    const headerTable = body.appendTable();
    const headerRow = headerTable.appendTableRow();
    
    // Logo (usar el logo real como en tu imagen)
    const logoCell = headerRow.appendTableCell();
    logoCell.appendParagraph('🍀 Cáritas').editAsText().setBold(true).setFontSize(16).setForegroundColor('#019952');
    logoCell.appendParagraph('LIMA').editAsText().setBold(true).setFontSize(14).setForegroundColor('#019952');
    logoCell.setWidth(100);
    
    // Título centrado
    const titleCell = headerRow.appendTableCell();
    titleCell.appendParagraph('Certificación Presupuestal').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(18);
    
    headerTable.setBorderWidth(0);
    
    body.appendParagraph(''); // Espaciado
    
    // Información básica (formato exacto de tu imagen)
    const infoTable = body.appendTable();
    infoTable.setBorderWidth(0);
    
    // Fila 1: Número y Fecha (como en tu imagen)
    const row1 = infoTable.appendTableRow();
    row1.appendTableCell(`Número: ${certificacion.codigo}`).editAsText().setBold(0, 7, true).setFontSize(11);
    row1.appendTableCell(`Fecha: ${formatearFechaDocumento(certificacion.fechaEmision)}`).editAsText().setBold(0, 5, true).setFontSize(11);
    
    // Responsable del área solicitante
    const row2 = infoTable.appendTableRow();
    const responsableCell = row2.appendTableCell(`Responsable del área solicitante: ${certificacion.solicitante}`);
    responsableCell.setWidth(400);
    row2.appendTableCell(''); // Celda vacía para mantener estructura
    responsableCell.editAsText().setBold(0, 35, true).setFontSize(11);
    
    // Oficina solicitante
    const row3 = infoTable.appendTableRow();
    const oficinaCell = row3.appendTableCell(`Oficina solicitante: ${obtenerNombreOficina(certificacion.oficina)}`);
    row3.appendTableCell(''); // Celda vacía
    oficinaCell.editAsText().setBold(0, 18, true).setFontSize(11);
    
    // Iniciativa (formato exacto)
    const row4 = infoTable.appendTableRow();
    const iniciativaText = `Iniciativa: ${obtenerNombreCatalogo('iniciativas', certificacion.iniciativa)} y ${certificacion.descripcion}`;
    const iniciativaCell = row4.appendTableCell(iniciativaText);
    row4.appendTableCell(''); // Celda vacía
    iniciativaCell.editAsText().setBold(0, 10, true).setFontSize(11);
    
    body.appendParagraph(''); // Espaciado
    
    // Tabla de ítems con formato EXACTO (como tu imagen)
    if (certificacion.items && certificacion.items.length > 0) {
      const itemsTable = body.appendTable();
      itemsTable.setBorderWidth(1);
      itemsTable.setBorderColor('#000000');
      
      // Encabezados con fondo gris (EXACTO como tu imagen)
      const headerItemsRow = itemsTable.appendTableRow();
      const headerCells = [
        headerItemsRow.appendTableCell('Descripción'),
        headerItemsRow.appendTableCell('Cant.'),
        headerItemsRow.appendTableCell('C/U (S/)'),
        headerItemsRow.appendTableCell('C/T(S/)')
      ];
      
      // Aplicar estilo a encabezados
      headerCells.forEach(cell => {
        cell.setBackgroundColor('#D3D3D3');
        cell.editAsText().setBold(true).setFontSize(10);
        cell.setPaddingTop(8);
        cell.setPaddingBottom(8);
        cell.setPaddingLeft(8);
        cell.setPaddingRight(8);
      });
      
      // Filas de datos
      certificacion.items.forEach(item => {
        const dataRow = itemsTable.appendTableRow();
        dataRow.appendTableCell(item.descripcion).editAsText().setFontSize(10);
        dataRow.appendTableCell(item.cantidad.toString()).editAsText().setFontSize(10);
        dataRow.appendTableCell(`S/ ${item.precioUnitario.toFixed(2)}`).editAsText().setFontSize(10);
        dataRow.appendTableCell(`S/ ${item.subtotal.toFixed(2)}`).editAsText().setFontSize(10);
      });
      
      // Fila de total con fondo gris (EXACTO como tu imagen)
      const totalRow = itemsTable.appendTableRow();
      const totalCells = [
        totalRow.appendTableCell('Total'),
        totalRow.appendTableCell('1'),
        totalRow.appendTableCell(`S/ ${certificacion.montoTotal.toFixed(2)}`),
        totalRow.appendTableCell(`S/ ${certificacion.montoTotal.toFixed(2)}`)
      ];
      
      totalCells.forEach(cell => {
        cell.setBackgroundColor('#E5E5E5');
        cell.editAsText().setBold(true).setFontSize(10);
        cell.setPaddingTop(8);
        cell.setPaddingBottom(8);
        cell.setPaddingLeft(8);
        cell.setPaddingRight(8);
      });
    }
    
    body.appendParagraph(''); // Espaciado
    
    // Información adicional (formato exacto)
    body.appendParagraph(`Base Legal: ${certificacion.disposicion || obtenerDisposicionPorDefecto()}`).editAsText().setBold(0, 10, true).setFontSize(10);
    body.appendParagraph(`Fuente de Financiamiento: ${obtenerNombreCatalogo('fuentes', certificacion.fuente)}`).editAsText().setBold(0, 24, true).setFontSize(10);
    body.appendParagraph(`Finalidad: ${certificacion.finalidad}`).editAsText().setBold(0, 9, true).setFontSize(10);
    body.appendParagraph(`Monto: S/ ${certificacion.montoTotal.toFixed(2)} | ${certificacion.montoLetras}`).editAsText().setBold(0, 6, true).setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    
    // Disposiciones (formato exacto)
    body.appendParagraph('Disposiciones:').editAsText().setBold(true).setFontSize(11);
    
    const disposicionesTexto = [
      'Se ha considerado la evaluación realizada por el área de logística desde la oficina de administración y según el estudio de mercado (cuadro comparativo)',
      'La presente autorización presupuestal se emite en base a la disponibilidad presupuestal aprobada para la iniciativa',
      'El responsable de la ejecución del gasto deberá presentar la documentación sustentatoria de acuerdo a las normas vigentes.'
    ];
    
    disposicionesTexto.forEach(disposicion => {
      const bulletPara = body.appendParagraph(`• ${disposicion}`);
      bulletPara.editAsText().setFontSize(10);
      bulletPara.setIndentFirstLine(20);
    });
    
    body.appendParagraph(''); // Espaciado
    
    // Adjuntos
    body.appendParagraph('Adjuntos: Documento sustentatorio obligatorios (contrataciones, proformas, términos de referencia, etc.)').editAsText().setBold(0, 8, true).setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    
    // Fecha de firma (formato exacto)
    const fechaFirma = new Date(certificacion.fechaEmision);
    body.appendParagraph(`Firmado en fecha ${fechaFirma.getDate()} de ${obtenerNombreMesCompleto(fechaFirma.getMonth())} de ${fechaFirma.getFullYear()} por:`).editAsText().setBold(0, 16, true).setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    
    // Firma REAL según plantilla (EXACTO como tu imagen)
    const firmantePorPlantilla = obtenerFirmantePorPlantilla(certificacion.plantilla);
    
    // Línea de firma
    body.appendParagraph('_'.repeat(35)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    
    // Aquí iría la imagen real de la firma si la tienes
    // Por ahora usamos el texto con el formato exacto
    body.appendParagraph(firmantePorPlantilla.nombre).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
    body.appendParagraph(firmantePorPlantilla.cargo).setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
    body.appendParagraph('Cáritas Lima').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
    
    body.appendParagraph(''); // Espaciado
    body.appendParagraph(''); // Espaciado
    
    // Footer de control electrónico (EXACTO como tu imagen)
    const numeroSolicitud = codigoCertificacion.split('-')[2];
    const año = fechaFirma.getFullYear();
    const mesAbrev = obtenerNombreMes(fechaFirma.getMonth()).substring(0, 4);
    
    const controlText = `*Control electrónico con asunto - Re: FP 149 Aprobación cédula Solicitud ${numeroSolicitud} de ${año} ***COMPRA ADICIONAL ACEITE*** enviado por la Administración el ${fechaFirma.getDate()} ${mesAbrev} ${año}. ${fechaFirma.getHours()}:${fechaFirma.getMinutes().toString().padStart(2, '0')} ${fechaFirma.getHours() >= 12 ? 'p.m.' : 'a.m.'}*`;
    
    body.appendParagraph(controlText).editAsText().setFontSize(7);
    
    doc.saveAndClose();
    
    // Generar PDF
    let pdf;
    try {
      const pdfBlob = doc.getAs(MimeType.PDF);
      pdfBlob.setName(`Certificacion_${codigoCertificacion}.pdf`);
      pdf = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
    } catch (pdfError) {
      Logger.log('No se pudo crear el PDF perfecto en la carpeta configurada: ' + pdfError.toString());
      const fallbackBlob = doc.getAs(MimeType.PDF);
      fallbackBlob.setName(`Certificacion_${codigoCertificacion}.pdf`);
      pdf = DriveApp.createFile(fallbackBlob);
    }
    
    // URLs
    const urlDocumento = `https://docs.google.com/document/d/${doc.getId()}/edit`;
    const urlPDF = `https://drive.google.com/file/d/${pdf.getId()}/view`;
    const urlVistaPrevia = `https://docs.google.com/document/d/${doc.getId()}/preview`;
    
    // Actualizar URLs en certificación
    actualizarCertificacion(codigoCertificacion, {
      urlDocumento: urlDocumento,
      urlPDF: urlPDF
    });
    
    Logger.log(`📄 Certificado generado: ${codigoCertificacion}`);
    
    registrarActividad('GENERAR_CERTIFICADO_PERFECTO', `Código: ${codigoCertificacion}`);
    
    return {
      success: true,
      urlDocumento: urlDocumento,
      urlPDF: urlPDF,
      urlVistaPrevia: urlVistaPrevia,
      documentoId: doc.getId(),
      pdfId: pdf.getId()
    };
  } catch (error) {
    Logger.log('Error en generarCertificadoPerfecto: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}