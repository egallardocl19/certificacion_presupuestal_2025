// ===============================================
// CONFIGURACIÓN INICIAL DEL SISTEMA - ACTUALIZADO
// Google Apps Script Setup
// ===============================================

// Helpers defensivos -------------------------------------------------------
//
// Cuando el archivo de configuración se ejecuta sin haber cargado todavía
// // Código.js (por ejemplo al depurar funciones individuales en Apps Script)
// // algunas utilidades como `ensureSheetStructure` o `ensureCertificacionesSheet`
// // no existen en el ámbito global.  Para evitar ReferenceError se definen
// // implementaciones básicas que respetan la firma esperada únicamente cuando
// // aún no hay una versión oficial disponible.

if (typeof normalizeHeaderName !== 'function') {
  var normalizeHeaderName = function (value) {
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
  };
}

if (typeof ensureSheetStructure !== 'function') {
  var ensureSheetStructure = function (name, headers) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    if (Array.isArray(headers) && headers.length > 0) {
      const range = sheet.getRange(1, 1, 1, headers.length);
      const current = range.getValues()[0];
      const target = headers.map(normalizeHeaderName);
      const normalizedCurrent = current.map(normalizeHeaderName);

      let needsUpdate = normalizedCurrent.length !== target.length;
      if (!needsUpdate) {
        for (let i = 0; i < target.length; i++) {
          if (normalizedCurrent[i] !== target[i]) {
            needsUpdate = true;
            break;
          }
        }
      }

      if (needsUpdate) {
        range.setValues([headers]);
      }
      if (sheet.getFrozenRows() < 1) {
        sheet.setFrozenRows(1);
      }
    }

    return sheet;
  };
}

if (typeof SHEET_NAMES === 'undefined') {
  var SHEET_NAMES = {
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
  };
}

if (typeof CERTIFICACIONES_HEADERS === 'undefined') {
  var CERTIFICACIONES_HEADERS = [
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
  ];
}

if (typeof ITEMS_HEADERS === 'undefined') {
  var ITEMS_HEADERS = [
    'Código Certificación',
    'Orden',
    'Descripción',
    'Cantidad',
    'Unidad',
    'Precio Unitario',
    'Subtotal',
    'Fecha Creación',
    'Creado Por'
  ];
}

if (typeof FIRMANTES_HEADERS === 'undefined') {
  var FIRMANTES_HEADERS = [
    'Código Certificación',
    'Orden',
    'Nombre',
    'Cargo',
    'Obligatorio',
    'Fecha Creación',
    'Creado Por'
  ];
}

if (typeof BITACORA_HEADERS === 'undefined') {
  var BITACORA_HEADERS = [
    'Fecha',
    'Usuario',
    'Acción',
    'Detalles',
    'Usuario Completo'
  ];
}

if (typeof PLANTILLAS_HEADERS === 'undefined') {
  var PLANTILLAS_HEADERS = [
    'ID',
    'Nombre',
    'Descripción',
    'Activa',
    'Firmantes',
    'Plantilla HTML',
    'Firmante ID',
    'Firmante Nombre',
    'Firmante Cargo'
  ];
}

if (typeof ensureCertificacionesSheet !== 'function') {
  var ensureCertificacionesSheet = function () {
    return ensureSheetStructure(
      SHEET_NAMES && SHEET_NAMES.CERTIFICACIONES ? SHEET_NAMES.CERTIFICACIONES : 'Certificaciones',
      typeof CERTIFICACIONES_HEADERS !== 'undefined' ? CERTIFICACIONES_HEADERS : [
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
      ]
    );
  };
}

if (typeof ensureItemsSheet !== 'function') {
  var ensureItemsSheet = function () {
    return ensureSheetStructure(
      SHEET_NAMES && SHEET_NAMES.ITEMS ? SHEET_NAMES.ITEMS : 'Items',
      typeof ITEMS_HEADERS !== 'undefined'
        ? ITEMS_HEADERS
        : ['Código Certificación', 'Orden', 'Descripción', 'Cantidad', 'Unidad', 'Precio Unitario', 'Subtotal', 'Fecha Creación', 'Creado Por']
    );
  };
}

if (typeof ensureFirmantesSheet !== 'function') {
  var ensureFirmantesSheet = function () {
    return ensureSheetStructure(
      SHEET_NAMES && SHEET_NAMES.FIRMANTES ? SHEET_NAMES.FIRMANTES : 'Firmantes',
      typeof FIRMANTES_HEADERS !== 'undefined'
        ? FIRMANTES_HEADERS
        : ['Código Certificación', 'Orden', 'Nombre', 'Cargo', 'Obligatorio', 'Fecha Creación', 'Creado Por']
    );
  };
}

if (typeof ensureBitacoraSheet !== 'function') {
  var ensureBitacoraSheet = function () {
    return ensureSheetStructure(
      SHEET_NAMES && SHEET_NAMES.BITACORA ? SHEET_NAMES.BITACORA : 'Bitacora',
      typeof BITACORA_HEADERS !== 'undefined'
        ? BITACORA_HEADERS
        : ['Fecha', 'Usuario', 'Acción', 'Detalles', 'Usuario Completo']
    );
  };
}

function configurarSistema() {
  try {
    Logger.log('Iniciando configuración del sistema...');

    asegurarPropiedadesDeScript();

    // Crear estructura de hojas
    crearEstructuraHojas();
    
    // Crear datos de ejemplo
    crearDatosEjemplo();
    
    Logger.log('Configuración del sistema completada exitosamente');
    return { success: true, message: 'Sistema configurado correctamente' };
  } catch (error) {
    Logger.log('Error en configurarSistema: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

const getSetupDefaultFinalidadDetalladaAliases = (() => {
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

function asegurarPropiedadesDeScript() {
  const properties = PropertiesService.getScriptProperties();
  if (!properties) {
    Logger.log('No se pudo acceder a las propiedades del script para inicializar alias de finalidad detallada.');
    return;
  }

  try {
    const raw = properties.getProperty('FINALIDAD_DETALLADA_ALIASES');
    if (raw) {
      const parsed = JSON.parse(raw);
      if (Array.isArray(parsed) && parsed.length > 0) {
        return;
      }
    }
  } catch (error) {
    Logger.log('Alias de finalidad detallada inválidos, se restablecerán a los valores por defecto: ' + error.toString());
  }

  properties.setProperty(
    'FINALIDAD_DETALLADA_ALIASES',
    JSON.stringify(getSetupDefaultFinalidadDetalladaAliases())
  );
  Logger.log('Alias de finalidad detallada inicializados en las propiedades del script.');
}

function crearEstructuraHojas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Hojas principales
  crearHojaCertificaciones(ss);
  crearHojaItems(ss);
  crearHojaFirmantes(ss);
  
  // Hojas de catálogos
  crearHojaCatalogoIniciativas(ss);
  crearHojaCatalogoTipos(ss);
  crearHojaCatalogoFuentes(ss);
  crearHojaCatalogoFinalidades(ss);
  crearHojaCatalogoOficinas(ss);
  crearHojaPlantillas(ss);
  
  // Hojas de configuración (NUEVAS)
  crearHojaConfigSolicitantes(ss);
  crearHojaConfigFirmantes(ss);
  crearHojaConfigGeneral(ss);
  
  // Hojas de sistema
  crearHojaUsuarios(ss);
  crearHojaBitacora(ss);
  
  Logger.log('Estructura de hojas creada exitosamente');
}

// ===============================================
// HOJAS DE CONFIGURACIÓN (NUEVAS)
// ===============================================

function crearHojaConfigSolicitantes(ss) {
  const headers = ['ID', 'Nombre Completo', 'Cargo', 'Email', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CONFIG_SOLICITANTES ? SHEET_NAMES.CONFIG_SOLICITANTES : 'Config_Solicitantes',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 200);
  sheet.setColumnWidth(5, 80);

  Logger.log('Hoja Config_Solicitantes creada');
}

function crearHojaConfigFirmantes(ss) {
  const headers = ['ID', 'Nombre Completo', 'Cargo', 'Orden', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CONFIG_FIRMANTES ? SHEET_NAMES.CONFIG_FIRMANTES : 'Config_Firmantes',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 80);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);

  Logger.log('Hoja Config_Firmantes creada');
}

function crearHojaConfigGeneral(ss) {
  const headers = ['Configuración', 'Valor'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CONFIG_GENERAL ? SHEET_NAMES.CONFIG_GENERAL : 'Config_General',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 400);

  Logger.log('Hoja Config_General creada');
}

// ===============================================
// HOJAS PRINCIPALES (actualizadas para Cáritas)
// ===============================================

function crearHojaCertificaciones(ss) {
  let sheet = ss.getSheetByName('Certificaciones');
  if (!sheet) {
    sheet = ss.insertSheet('Certificaciones');
  }

  const headers = [
    'Código', // A
    'Fecha Emisión', // B - PERMITIR MODIFICAR
    'Descripción', // C
    'Iniciativa', // D
    'Tipo', // E
    'Fuente', // F
    'Finalidad', // G - AUTOMÁTICA
    'Oficina', // H
    'Solicitante', // I - AUTOMÁTICO
    'Cargo Solicitante', // J - AUTOMÁTICO
    'Email Solicitante', // K - AUTOMÁTICO
    'Número Autorización', // L
    'Cargo Autorizador', // M
    'Estado', // N
    'Disposición/Base Legal', // O - CONFIGURACIÓN
    'Monto Total', // P
    'Monto en Letras', // Q
    'Fecha Creación', // R
    'Creado Por', // S
    'Fecha Modificación', // T
    'Modificado Por', // U
    'Fecha Anulación', // V
    'Anulado Por', // W
    'Motivo Anulación', // X
    'Plantilla', // Y
    'URL Documento', // Z
    'URL PDF', // AA
    'Finalidad Detallada' // AB
  ];

  asegurarColumnasMinimas(sheet, headers.length);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Formatear encabezados
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#019952');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');
  headerRange.setWrap(true);
  
  // Ajustar anchos de columna
  sheet.setColumnWidth(1, 100); // Código
  sheet.setColumnWidth(2, 100); // Fecha Emisión
  sheet.setColumnWidth(3, 200); // Descripción
  sheet.setColumnWidth(4, 120); // Iniciativa
  sheet.setColumnWidth(5, 100); // Tipo
  sheet.setColumnWidth(6, 120); // Fuente
  sheet.setColumnWidth(7, 150); // Finalidad
  sheet.setColumnWidth(8, 120); // Oficina
  sheet.setColumnWidth(9, 150); // Solicitante
  sheet.setColumnWidth(10, 120); // Cargo Solicitante

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Certificaciones actualizada');
}

function asegurarColumnasMinimas(sheet, cantidadColumnas) {
  const columnasActuales = sheet.getMaxColumns();
  if (columnasActuales < cantidadColumnas) {
    sheet.insertColumnsAfter(columnasActuales, cantidadColumnas - columnasActuales);
  }
}

function crearHojaItems(ss) {
  const headers = [
    'Código Certificación',
    'Orden',
    'Descripción',
    'Cantidad',
    'Unidad',
    'Precio Unitario',
    'Subtotal',
    'Fecha Creación',
    'Creado Por'
  ];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.ITEMS ? SHEET_NAMES.ITEMS : 'Items',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#34a853');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 250);
  sheet.setColumnWidth(4, 80);
  sheet.setColumnWidth(5, 80);
  sheet.setColumnWidth(6, 100);
  sheet.setColumnWidth(7, 100);
  
  sheet.setFrozenRows(1);
  
  Logger.log('Hoja de Items creada');
}

function crearHojaFirmantes(ss) {
  const headers = [
    'Código Certificación',
    'Orden',
    'Nombre',
    'Cargo',
    'Obligatorio',
    'Fecha Creación',
    'Creado Por'
  ];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.FIRMANTES ? SHEET_NAMES.FIRMANTES : 'Firmantes',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#ff9800');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 180);
  sheet.setColumnWidth(4, 180);
  sheet.setColumnWidth(5, 100);
  
  sheet.setFrozenRows(1);
  
  Logger.log('Hoja de Firmantes creada');
}

// ===============================================
// HOJAS DE CATÁLOGOS (actualizadas para Cáritas)
// ===============================================

function crearHojaCatalogoIniciativas(ss) {
  const headers = ['Código', 'Nombre', 'Descripción', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CATALOGO_INICIATIVAS ? SHEET_NAMES.CATALOGO_INICIATIVAS : 'Cat_Iniciativas',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Catálogo Iniciativas creada');
}

function crearHojaCatalogoTipos(ss) {
  const headers = ['Código', 'Nombre', 'Descripción', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CATALOGO_TIPOS ? SHEET_NAMES.CATALOGO_TIPOS : 'Cat_Tipos',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Catálogo Tipos creada');
}

function crearHojaCatalogoFuentes(ss) {
  const headers = ['Código', 'Nombre', 'Descripción', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CATALOGO_FUENTES ? SHEET_NAMES.CATALOGO_FUENTES : 'Cat_Fuentes',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);
  
  Logger.log('Hoja de Catálogo Fuentes creada');
}

function crearHojaCatalogoFinalidades(ss) {
  const headers = ['Código', 'Nombre', 'Descripción', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CATALOGO_FINALIDADES ? SHEET_NAMES.CATALOGO_FINALIDADES : 'Cat_Finalidades',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Catálogo Finalidades creada');
}

function crearHojaCatalogoOficinas(ss) {
  const headers = ['Código', 'Nombre', 'Descripción', 'Activo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.CATALOGO_OFICINAS ? SHEET_NAMES.CATALOGO_OFICINAS : 'Cat_Oficinas',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#9c27b0');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);
  
  Logger.log('Hoja de Catálogo Oficinas creada');
}

function crearHojaPlantillas(ss) {
  const headers = typeof PLANTILLAS_HEADERS !== 'undefined'
    ? PLANTILLAS_HEADERS
    : ['ID', 'Nombre', 'Descripción', 'Activa', 'Firmantes', 'Plantilla HTML', 'Firmante ID', 'Firmante Nombre', 'Firmante Cargo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.PLANTILLAS ? SHEET_NAMES.PLANTILLAS : 'Plantillas',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#f44336');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Plantillas creada o actualizada');
}

function crearHojaUsuarios(ss) {
  const headers = ['Email', 'Nombre', 'Rol', 'Oficina', 'Activo', 'Fecha Creación'];
  const sheet = ensureSheetStructure('Usuarios', headers);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#607d8b');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Usuarios creada');
}

function crearHojaBitacora(ss) {
  const headers = ['Fecha', 'Usuario', 'Acción', 'Detalles', 'Usuario Completo'];
  const sheet = ensureSheetStructure(
    SHEET_NAMES && SHEET_NAMES.BITACORA ? SHEET_NAMES.BITACORA : 'Bitacora',
    headers
  );

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#795548');
  headerRange.setFontColor('white');
  headerRange.setFontWeight('bold');

  sheet.setFrozenRows(1);

  Logger.log('Hoja de Bitácora creada');
}

// ===============================================
// DATOS DE EJEMPLO ACTUALIZADOS PARA CÁRITAS
// ===============================================

function crearDatosEjemplo() {
  crearCatalogosEjemplo();
  crearPlantillasEjemplo();
  crearConfiguracionEjemplo(); // IMPORTANTE: Crear configuración
  crearUsuariosEjemplo();
  crearCertificacionesEjemplo();
  
  Logger.log('Datos de ejemplo creados exitosamente');
}

function crearConfiguracionEjemplo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Configuración General
  const sheetConfigGeneral = ss.getSheetByName('Config_General');
  if (sheetConfigGeneral) {
    if (sheetConfigGeneral.getLastRow() > 1) {
      sheetConfigGeneral.getRange(2, 1, sheetConfigGeneral.getLastRow() - 1, sheetConfigGeneral.getLastColumn()).clearContent();
    }

    const configuracionGeneral = [
      ['disposicion_base_legal', 'Directiva 003-2023-SG/CARITASLIMA, Directiva de contratación de bienes y servicios de la Vicaría de Pastoral Social y Dignidad Humana - Caritas Lima'],
      ['codigo_formato', 'CP-{YEAR}-{NUMBER}'],
      ['moneda_por_defecto', 'SOLES'],
      ['timezone', 'America/Lima']
    ];
    sheetConfigGeneral.getRange(2, 1, configuracionGeneral.length, 2).setValues(configuracionGeneral);
  }
  
  // Solicitantes
  const sheetSolicitantes = ss.getSheetByName('Config_Solicitantes');
  if (sheetSolicitantes) {
    const solicitantes = [
      ['SOL001', 'Evelyn Elena Huaycacllo Marin', 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', 'evelyn.huaycacllo@caritaslima.org', true],
      ['SOL002', 'Guadalupe Susana Callupe Pacheco', 'Coordinadora de Logística', 'guadalupe.callupe@caritaslima.org', true],
      ['SOL003', 'José Luis Mendoza Vargas', 'Jefe de Administración', 'jose.mendoza@caritaslima.org', true],
      ['SOL004', 'Ana Sofía Quispe Mamani', 'Coordinadora de Programas Sociales', 'ana.quispe@caritaslima.org', true]
    ];
    sheetSolicitantes.getRange(2, 1, solicitantes.length, 5).setValues(solicitantes);
  }
  
  // Firmantes
  const sheetFirmantes = ss.getSheetByName('Config_Firmantes');
  if (sheetFirmantes) {
    const firmantes = [
      ['FIR_EVELYN', 'Evelyn Elena Huaycacllo Marin', 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', 1, true],
      ['FIR_JORGE', 'Jorge Herrera', 'Director Ejecutivo', 1, true],
      ['FIR_SUSANA', 'Susana Palomino', 'Coordinadora de Planeamiento y Presupuesto', 1, true],
      ['FIR_OTRO', 'Equipo Designado', 'Responsable según tipo de certificación', 1, true]
    ];
    sheetFirmantes.getRange(2, 1, firmantes.length, 5).setValues(firmantes);
  }
  
  Logger.log('Configuración de ejemplo creada');
}

function crearCatalogosEjemplo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Iniciativas específicas de Cáritas
  const sheetIniciativas = ss.getSheetByName('Cat_Iniciativas');
  if (sheetIniciativas.getLastRow() > 1) {
    sheetIniciativas.getRange(2, 1, sheetIniciativas.getLastRow() - 1, sheetIniciativas.getLastColumn()).clearContent();
  }
  const iniciativas = [
    ['INI001', 'Provisión de Alimentos para los Servicios de Alimentación Comunitaria', 'Programa de alimentación para comunidades vulnerables', true],
    ['INI002', 'Fortalecimiento Institucional y apoyo para la formalización', 'Mejoras en capacidad operativa institucional', true],
    ['INI003', 'Implementación de Programas Sociales', 'Desarrollo de programas de asistencia social', true],
    ['INI004', 'Modernización Tecnológica', 'Actualización de sistemas y equipos', true],
    ['INI005', 'Capacitación y Formación del Personal', 'Programas de formación del personal', true]
  ];
  sheetIniciativas.getRange(2, 1, iniciativas.length, 4).setValues(iniciativas);
  
  // Tipos
  const sheetTipos = ss.getSheetByName('Cat_Tipos');
  if (sheetTipos.getLastRow() > 1) {
    sheetTipos.getRange(2, 1, sheetTipos.getLastRow() - 1, sheetTipos.getLastColumn()).clearContent();
  }
  const tipos = [
    ['TIP001', 'Bienes', 'Adquisición de productos y materiales', true],
    ['TIP002', 'Servicios', 'Contratación de servicios profesionales', true],
    ['TIP003', 'Obras', 'Ejecución de obras y construcciones', true],
    ['TIP004', 'Consultorías', 'Servicios de consultoría especializada', true]
  ];
  sheetTipos.getRange(2, 1, tipos.length, 4).setValues(tipos);
  
  // Fuentes específicas de Cáritas
  const sheetFuentes = ss.getSheetByName('Cat_Fuentes');
  if (sheetFuentes.getLastRow() > 1) {
    sheetFuentes.getRange(2, 1, sheetFuentes.getLastRow() - 1, sheetFuentes.getLastColumn()).clearContent();
  }
  const fuentes = [
    ['FUE001', 'Otros Gastos', 'Recursos propios de la institución', true],
    ['FUE002', 'Donaciones Internacionales', 'Fondos de cooperación internacional', true],
    ['FUE003', 'Transferencias del Estado', 'Recursos del gobierno peruano', true],
    ['FUE004', 'Autogestión', 'Recursos generados por actividades propias', true]
  ];
  sheetFuentes.getRange(2, 1, fuentes.length, 4).setValues(fuentes);
  
  // Finalidades
  const sheetFinalidades = ss.getSheetByName('Cat_Finalidades');
  if (sheetFinalidades.getLastRow() > 1) {
    sheetFinalidades.getRange(2, 1, sheetFinalidades.getLastRow() - 1, sheetFinalidades.getLastColumn()).clearContent();
  }
  const finalidades = [
    ['FIN001', 'Administración y Gestión', 'Gastos administrativos y de gestión', true],
    ['FIN002', 'Programas Sociales', 'Actividades de asistencia social', true],
    ['FIN003', 'Infraestructura', 'Mejoras en infraestructura', true],
    ['FIN004', 'Capacitación', 'Formación y desarrollo de capacidades', true],
    ['FIN005', 'Alimentación Comunitaria', 'Programas de alimentación', true]
  ];
  sheetFinalidades.getRange(2, 1, finalidades.length, 4).setValues(finalidades);
  
  // Oficinas específicas de Cáritas
  const sheetOficinas = ss.getSheetByName('Cat_Oficinas');
  if (sheetOficinas.getLastRow() > 1) {
    sheetOficinas.getRange(2, 1, sheetOficinas.getLastRow() - 1, sheetOficinas.getLastColumn()).clearContent();
  }
  const oficinas = [
    ['OFI001', 'Oficina de Administración', 'Área administrativa general', true],
    ['OFI002', 'Oficina de Política, Planeamiento y Presupuesto', 'Área de planificación y presupuesto', true],
    ['OFI003', 'Oficina de Recursos Humanos', 'Gestión del talento humano', true],
    ['OFI004', 'Oficina de Logística', 'Área de adquisiciones y suministros', true],
    ['OFI005', 'Oficina de Programas Sociales', 'Coordinación de programas sociales', true]
  ];
  sheetOficinas.getRange(2, 1, oficinas.length, 4).setValues(oficinas);
  
  Logger.log('Catálogos de ejemplo creados');
}

// Agregar/reemplazar en setup.gs - función crearPlantillasEjemplo:

function crearPlantillasEjemplo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ensureSheetStructure(SHEET_NAMES.PLANTILLAS, PLANTILLAS_HEADERS);
  const headersLength = PLANTILLAS_HEADERS.length;

  const data = sheet.getDataRange().getValues();
  const existentes = new Map();

  for (let i = 1; i < data.length; i++) {
    const id = (data[i][0] || '').toString().trim();
    if (id) {
      existentes.set(id, {
        index: i + 1,
        values: data[i].slice(0, headersLength)
      });
    }
  }

  const plantillasBase = [
    {
      id: 'plantilla_evelyn',
      nombre: 'Certificación - Evelyn Huaycacllo',
      descripcion: 'Formato institucional firmado por la Jefa de Presupuesto',
      activa: true,
      firmantes: 1,
      plantillaHtml: 'https://docs.google.com/document/d/PLANTILLA_EVELYN/edit',
      firmanteId: 'FIR_EVELYN',
      firmanteNombre: 'Evelyn Elena Huaycacllo Marin',
      firmanteCargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto'
    },
    {
      id: 'plantilla_jorge',
      nombre: 'Certificación - Jorge Herrera',
      descripcion: 'Formato cuando el Director Ejecutivo autoriza la certificación',
      activa: true,
      firmantes: 1,
      plantillaHtml: 'https://docs.google.com/document/d/PLANTILLA_JORGE/edit',
      firmanteId: 'FIR_JORGE',
      firmanteNombre: 'Jorge Herrera',
      firmanteCargo: 'Director Ejecutivo'
    },
    {
      id: 'plantilla_susana',
      nombre: 'Certificación - Susana Palomino',
      descripcion: 'Formato para coordinaciones de planeamiento y presupuesto',
      activa: true,
      firmantes: 1,
      plantillaHtml: 'https://docs.google.com/document/d/PLANTILLA_SUSANA/edit',
      firmanteId: 'FIR_SUSANA',
      firmanteNombre: 'Susana Palomino',
      firmanteCargo: 'Coordinadora de Planeamiento y Presupuesto'
    },
    {
      id: 'plantilla_otro',
      nombre: 'Certificación - Equipo Designado',
      descripcion: 'Formato genérico para otras áreas responsables',
      activa: true,
      firmantes: 1,
      plantillaHtml: 'https://docs.google.com/document/d/PLANTILLA_OTRO/edit',
      firmanteId: 'FIR_OTRO',
      firmanteNombre: 'Equipo Designado',
      firmanteCargo: 'Responsable según tipo de certificación'
    }
  ];

  const nuevasFilas = [];

  plantillasBase.forEach(plantilla => {
    const filaExistente = existentes.get(plantilla.id);
    const valores = filaExistente ? filaExistente.values : new Array(headersLength).fill('');

    valores[0] = plantilla.id;
    valores[1] = plantilla.nombre;
    valores[2] = plantilla.descripcion;
    valores[3] = true;
    valores[4] = plantilla.firmantes;
    valores[5] = plantilla.plantillaHtml;
    valores[6] = plantilla.firmanteId;
    valores[7] = plantilla.firmanteNombre;
    valores[8] = plantilla.firmanteCargo;

    if (filaExistente) {
      sheet.getRange(filaExistente.index, 1, 1, headersLength).setValues([valores]);
    } else {
      nuevasFilas.push(valores);
    }

    existentes.set(plantilla.id, {
      index: filaExistente ? filaExistente.index : null,
      values: valores
    });
  });

  if (nuevasFilas.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, nuevasFilas.length, headersLength).setValues(nuevasFilas);
  }

  const placeholders = [];
  for (let i = 1; i <= 5; i++) {
    const id = `plantilla_personalizada_${i}`;
    if (!existentes.has(id)) {
      placeholders.push([
        id,
        `Plantilla Personalizada ${i}`,
        'Personaliza esta plantilla con tu propio formato y firmante.',
        false,
        1,
        '',
        '',
        '',
        ''
      ]);
    }
  }

  if (placeholders.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, placeholders.length, headersLength).setValues(placeholders);
  }

  SpreadsheetApp.flush();

  invalidarCachePlantillas();

  Logger.log('Plantillas base y espacios personalizados asegurados');
}

function crearUsuariosEjemplo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Usuarios');
  
  const usuarios = [
    [Session.getActiveUser().getEmail(), 'Usuario Administrador', 'Administrador', 'OFI001', true, new Date()],
    ['evelyn.huaycacllo@caritaslima.org', 'Evelyn Elena Huaycacllo Marin', 'Solicitante', 'OFI002', true, new Date()],
    ['guadalupe.callupe@caritaslima.org', 'Guadalupe Susana Callupe Pacheco', 'Revisor/Presupuesto', 'OFI002', true, new Date()],
    ['director@caritaslima.org', 'Padre Miguel Ángel Castillo Seminario', 'Autorizador', 'OFI001', true, new Date()]
  ];
  
  sheet.getRange(2, 1, usuarios.length, 6).setValues(usuarios);
  
  Logger.log('Usuarios de ejemplo creados');
}

function crearCertificacionesEjemplo() {
  // Crear certificación de ejemplo idéntica a la imagen
  const certificacionEjemplo = {
    fecha: new Date('2025-01-24'), // Fecha específica
    descripcion: 'Adquisición de productos adicionales para completar los kits de ollas (AZÚCAR CARTAVIO RUBIA GRANEL y ACEITE VEGA BOTELLA)',
    iniciativa: 'INI001',
    tipo: 'TIP001',
    fuente: 'FUE001',
    oficina: 'OFI001',
    solicitante: 'Guadalupe Susana Callupe Pacheco',
    cargoSolicitante: 'Coordinadora de Logística',
    emailSolicitante: 'guadalupe.callupe@caritaslima.org',
    finalidad: 'Complementar con productos alimentarios la conformación de los kits de ollas para atender a familias en situación de vulnerabilidad, garantizando una asistencia oportuna y completa.',
    finalidadDetallada: 'Complementar con productos alimentarios la conformación de los kits de ollas para atender a familias en situación de vulnerabilidad, garantizando una asistencia oportuna y completa.',
    items: [
      {
        descripcion: 'Adquisición de productos adicionales para completar los kits de ollas (AZÚCAR CARTAVIO RUBIA GRANEL y ACEITE VEGA BOTELLA)',
        cantidad: 1,
        unidad: 'Lote',
        precioUnitario: 525.00
      }
    ],
    firmantes: [
      {
        nombre: 'Evelyn Elena Huaycacllo Marin',
        cargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto',
        obligatorio: true
      }
    ],
    plantilla: 'plantilla_evelyn'
  };
  
  try {
    const resultado = crearCertificacion(certificacionEjemplo);
    if (resultado.success) {
      Logger.log(`Certificación de ejemplo creada: ${resultado.codigo}`);
    }
  } catch (error) {
    Logger.log(`Error creando certificación de ejemplo: ${error.toString()}`);
  }
  
  Logger.log('Certificaciones de ejemplo creadas');
}

// Función para resetear completamente el sistema
function resetearSistema() {
  const confirmacion = Browser.msgBox(
    'Confirmación',
    '¿Está seguro que desea resetear completamente el sistema? Esta acción eliminará todos los datos existentes.',
    Browser.Buttons.YES_NO
  );
  
  if (confirmacion === 'yes') {
    configurarSistema();
    Browser.msgBox('Sistema reseteado y configurado exitosamente');
    return { success: true };
  }
  
  return { success: false, message: 'Operación cancelada por el usuario' };
}

// Función para crear solo la estructura (sin datos de ejemplo)
function crearSoloEstructura() {
  try {
    Logger.log('Creando solo estructura de hojas...');
    crearEstructuraHojas();
    crearCatalogosEjemplo(); // Los catálogos son necesarios para el funcionamiento
    crearPlantillasEjemplo(); // Las plantillas son necesarias para el funcionamiento
    crearConfiguracionEjemplo(); // La configuración es necesaria
    Logger.log('Estructura básica creada exitosamente');
    return { success: true, message: 'Estructura básica creada correctamente' };
  } catch (error) {
    Logger.log('Error en crearSoloEstructura: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function normalizarEstructuraSistema() {
  try {
    Logger.log('Normalizando estructura del sistema...');

    asegurarPropiedadesDeScript();

    ensureCertificacionesSheet();
    ensureItemsSheet();
    ensureFirmantesSheet();
    ensureBitacoraSheet();

    ensureSheetStructure(SHEET_NAMES.CATALOGO_INICIATIVAS, ['Código', 'Nombre', 'Descripción', 'Activo']);
    ensureSheetStructure(SHEET_NAMES.CATALOGO_TIPOS, ['Código', 'Nombre', 'Descripción', 'Activo']);
    ensureSheetStructure(SHEET_NAMES.CATALOGO_FUENTES, ['Código', 'Nombre', 'Descripción', 'Activo']);
    ensureSheetStructure(SHEET_NAMES.CATALOGO_FINALIDADES, ['Código', 'Nombre', 'Descripción', 'Activo']);
    ensureSheetStructure(SHEET_NAMES.CATALOGO_OFICINAS, ['Código', 'Nombre', 'Descripción', 'Activo']);

    ensureSheetStructure(SHEET_NAMES.CONFIG_SOLICITANTES, ['ID', 'Nombre Completo', 'Cargo', 'Email', 'Activo']);
    ensureSheetStructure(SHEET_NAMES.CONFIG_FIRMANTES, ['ID', 'Nombre Completo', 'Cargo', 'Orden', 'Activo']);
    ensureSheetStructure(SHEET_NAMES.CONFIG_GENERAL, ['Configuración', 'Valor']);
    ensureSheetStructure('Usuarios', ['Email', 'Nombre', 'Rol', 'Oficina', 'Activo', 'Fecha Creación']);

    ensureSheetStructure(SHEET_NAMES.PLANTILLAS, PLANTILLAS_HEADERS);

    crearPlantillasEjemplo();

    invalidarCachePlantillas();

    Logger.log('Estructura del sistema normalizada correctamente');
    return { success: true, message: 'Estructura y plantillas normalizadas correctamente' };
  } catch (error) {
    Logger.log('Error en normalizarEstructuraSistema: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}


// ===============================================
// NUEVOS DATOS CON FIRMANTES CORRECTOS
// ===============================================

function crearConfiguracionEjemplo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Configuración General
  const sheetConfigGeneral = ss.getSheetByName('Config_General');
  if (sheetConfigGeneral) {
    const configuracionGeneral = [
      ['disposicion_base_legal', 'Directiva 003-2023-SG/CARITASLIMA, Directiva de contratación de bienes y servicios de la Vicaría de Pastoral Social y Dignidad Humana - Caritas Lima'],
      ['codigo_formato', 'CP-{YEAR}-{NUMBER}'],
      ['moneda_por_defecto', 'SOLES'],
      ['timezone', 'America/Lima']
    ];
    sheetConfigGeneral.getRange(2, 1, configuracionGeneral.length, 2).setValues(configuracionGeneral);
  }
  
  // Solicitantes ACTUALIZADOS
  const sheetSolicitantes = ss.getSheetByName('Config_Solicitantes');
  if (sheetSolicitantes) {
    if (sheetSolicitantes.getLastRow() > 1) {
      sheetSolicitantes.getRange(2, 1, sheetSolicitantes.getLastRow() - 1, sheetSolicitantes.getLastColumn()).clearContent();
    }

    const solicitantes = [
      ['SOL001', 'Evelyn Elena Huaycacllo Marin', 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', 'evelyn.huaycacllo@caritaslima.org', true],
      ['SOL002', 'Guadalupe Susana Callupe Pacheco', 'Coordinadora de Logística', 'guadalupe.callupe@caritaslima.org', true],
      ['SOL003', 'José Luis Mendoza Vargas', 'Jefe de Administración', 'jose.mendoza@caritaslima.org', true],
      ['SOL004', 'Ana Sofía Quispe Mamani', 'Coordinadora de Programas Sociales', 'ana.quispe@caritaslima.org', true]
    ];
    sheetSolicitantes.getRange(2, 1, solicitantes.length, 5).setValues(solicitantes);
  }
  
  // Firmantes ACTUALIZADOS
  const sheetFirmantes = ss.getSheetByName('Config_Firmantes');
  if (sheetFirmantes) {
    if (sheetFirmantes.getLastRow() > 1) {
      sheetFirmantes.getRange(2, 1, sheetFirmantes.getLastRow() - 1, sheetFirmantes.getLastColumn()).clearContent();
    }

    const firmantes = [
      ['FIR_EVELYN', 'Evelyn Elena Huaycacllo Marin', 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', 1, true],
      ['FIR_JORGE', 'Jorge Herrera', 'Director Ejecutivo', 1, true],
      ['FIR_SUSANA', 'Susana Palomino', 'Coordinadora de Planeamiento y Presupuesto', 1, true],
      ['FIR_OTRO', 'Equipo Designado', 'Responsable según tipo de certificación', 1, true]
    ];
    sheetFirmantes.getRange(2, 1, firmantes.length, 5).setValues(firmantes);
  }
  
  Logger.log('Configuración actualizada con nuevos firmantes');
}

function crearPlantillasConURLs() {
  try {
    Logger.log('Creando plantillas con URLs válidas...');
    
    // Crear plantilla 1: Evelyn Travezaño
    const plantilla1 = crearPlantillaEvelyn();
    
    // Crear plantilla 2: Jorge Herrera  
    const plantilla2 = crearPlantillaJorge();
    
    // Actualizar hoja de plantillas con URLs
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Plantillas');
    
    if (sheet) {
      // Limpiar plantillas existentes
      const dataRange = sheet.getDataRange();
      if (dataRange.getNumRows() > 1) {
        sheet.deleteRows(2, dataRange.getNumRows() - 1);
      }
      
      // Agregar nuevas plantillas CON URLs
      const plantillas = [
        [
          'plantilla_evelyn', 
          'Certificación - Evelyn Travezaño', 
          'Certificación firmada por la Directora de Administración y Finanzas', 
          true, 
          1, 
          `https://docs.google.com/document/d/${plantilla1.docId}/edit`
        ],
        [
          'plantilla_jorge', 
          'Certificación - Jorge Herrera', 
          'Certificación firmada por el Director Ejecutivo', 
          true, 
          1, 
          `https://docs.google.com/document/d/${plantilla2.docId}/edit`
        ]
      ];
      
      sheet.getRange(2, 1, plantillas.length, 6).setValues(plantillas);
      
      Logger.log('✅ Plantillas creadas con URLs:');
      Logger.log(`📄 Plantilla Evelyn: https://docs.google.com/document/d/1-_brt0nTwy8oDgA_ZENNZst6-bTYgtgqansBjNxU0oQ/edit`);
      Logger.log(`📄 Plantilla Jorge: https://docs.google.com/document/d/${plantilla2.docId}/edit`);
    }
    
    return { 
      success: true, 
      plantillas: [
        { ...plantilla1, url: `https://docs.google.com/document/d/${plantilla1.docId}/edit` },
        { ...plantilla2, url: `https://docs.google.com/document/d/${plantilla2.docId}/edit` }
      ]
    };
  } catch (error) {
    Logger.log('Error en crearPlantillasConURLs: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function crearPlantillaEvelyn() {
  const doc = DocumentApp.create('Plantilla_Evelyn_Travezaño');
  const body = doc.getBody();
  
  // Configurar márgenes
  body.setMarginTop(50);
  body.setMarginBottom(50);
  body.setMarginLeft(60);
  body.setMarginRight(60);
  
  // Header
  const headerTable = body.appendTable();
  const headerRow = headerTable.appendTableRow();
  
  // Logo
  const logoCell = headerRow.appendTableCell();
  logoCell.appendParagraph('🍀 Cáritas').editAsText().setBold(true).setFontSize(16).setForegroundColor('#019952');
  logoCell.appendParagraph('LIMA').editAsText().setBold(true).setFontSize(14).setForegroundColor('#019952');
  logoCell.setWidth(100);
  
  // Título
  const titleCell = headerRow.appendTableCell();
  titleCell.appendParagraph('Certificación Presupuestal').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(18);
  
  headerTable.setBorderWidth(0);
  
  body.appendParagraph(''); // Espaciado
  
  // Placeholders para datos variables
  body.appendParagraph('Número: {{CODIGO}}').editAsText().setBold(0, 7, true);
  body.appendParagraph('Fecha: {{FECHA}}').editAsText().setBold(0, 5, true);
  body.appendParagraph('Responsable del área solicitante: {{SOLICITANTE}}').editAsText().setBold(0, 35, true);
  body.appendParagraph('Oficina solicitante: {{OFICINA}}').editAsText().setBold(0, 18, true);
  body.appendParagraph('Iniciativa: {{INICIATIVA}} y {{DESCRIPCION}}').editAsText().setBold(0, 10, true);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('{{TABLA_ITEMS}}'); // Placeholder para tabla
  body.appendParagraph(''); // Espaciado
  
  // Información fija
  body.appendParagraph('Base Legal: {{BASE_LEGAL}}').editAsText().setBold(0, 10, true).setFontSize(10);
  body.appendParagraph('Fuente de Financiamiento: {{FUENTE}}').editAsText().setBold(0, 24, true).setFontSize(10);
  body.appendParagraph('Finalidad: {{FINALIDAD}}').editAsText().setBold(0, 9, true).setFontSize(10);
  body.appendParagraph('Monto: {{MONTO}}').editAsText().setBold(0, 6, true).setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  
  // Disposiciones estándar
  body.appendParagraph('Disposiciones:').editAsText().setBold(true).setFontSize(11);
  body.appendParagraph('• Se ha considerado la evaluación realizada por el área de logística desde la oficina de administración y según el estudio de mercado (cuadro comparativo)').editAsText().setFontSize(10);
  body.appendParagraph('• La presente autorización presupuestal se emite en base a la disponibilidad presupuestal aprobada para la iniciativa').editAsText().setFontSize(10);
  body.appendParagraph('• El responsable de la ejecución del gasto deberá presentar la documentación sustentatoria de acuerdo a las normas vigentes.').editAsText().setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('Adjuntos: Documento sustentatorio obligatorios (contrataciones, proformas, términos de referencia, etc.)').editAsText().setBold(0, 8, true).setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('Firmado en fecha {{FECHA_FIRMA}} por:').editAsText().setBold(0, 16, true).setFontSize(10);
  
  body.appendParagraph(''); // Espaciado para firma
  body.appendParagraph(''); 
  body.appendParagraph(''); 
  body.appendParagraph(''); 
  
  // FIRMA DE EVELYN TRAVEZAÑO
  body.appendParagraph('_'.repeat(35)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Evelyn Travezaño').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
  body.appendParagraph('Directora de Administración y Finanzas').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
  body.appendParagraph('Cáritas Lima').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('*Control electrónico con asunto - Re: FP 149 Aprobación cédula Solicitud {{NUMERO_SOLICITUD}} de {{ANNO}} ***COMPRA ADICIONAL ACEITE*** enviado por la Administración el {{FECHA_CONTROL}}*').editAsText().setFontSize(7);
  
  doc.saveAndClose();
  
  return {
    docId: doc.getId(),
    nombre: 'Plantilla_Evelyn_Travezaño',
    firmante: 'Evelyn Travezaño',
    cargo: 'Directora de Administración y Finanzas'
  };
}

function crearPlantillaJorge() {
  const doc = DocumentApp.create('Plantilla_Jorge_Herrera');
  const body = doc.getBody();
  
  // Configurar márgenes
  body.setMarginTop(50);
  body.setMarginBottom(50);
  body.setMarginLeft(60);
  body.setMarginRight(60);
  
  // Header
  const headerTable = body.appendTable();
  const headerRow = headerTable.appendTableRow();
  
  // Logo
  const logoCell = headerRow.appendTableCell();
  logoCell.appendParagraph('🍀 Cáritas').editAsText().setBold(true).setFontSize(16).setForegroundColor('#019952');
  logoCell.appendParagraph('LIMA').editAsText().setBold(true).setFontSize(14).setForegroundColor('#019952');
  logoCell.setWidth(100);
  
  // Título
  const titleCell = headerRow.appendTableCell();
  titleCell.appendParagraph('Certificación Presupuestal').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(18);
  
  headerTable.setBorderWidth(0);
  
  body.appendParagraph(''); // Espaciado
  
  // Placeholders para datos variables
  body.appendParagraph('Número: {{CODIGO}}').editAsText().setBold(0, 7, true);
  body.appendParagraph('Fecha: {{FECHA}}').editAsText().setBold(0, 5, true);
  body.appendParagraph('Responsable del área solicitante: {{SOLICITANTE}}').editAsText().setBold(0, 35, true);
  body.appendParagraph('Oficina solicitante: {{OFICINA}}').editAsText().setBold(0, 18, true);
  body.appendParagraph('Iniciativa: {{INICIATIVA}} y {{DESCRIPCION}}').editAsText().setBold(0, 10, true);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('{{TABLA_ITEMS}}'); // Placeholder para tabla
  body.appendParagraph(''); // Espaciado
  
  // Información fija
  body.appendParagraph('Base Legal: {{BASE_LEGAL}}').editAsText().setBold(0, 10, true).setFontSize(10);
  body.appendParagraph('Fuente de Financiamiento: {{FUENTE}}').editAsText().setBold(0, 24, true).setFontSize(10);
  body.appendParagraph('Finalidad: {{FINALIDAD}}').editAsText().setBold(0, 9, true).setFontSize(10);
  body.appendParagraph('Monto: {{MONTO}}').editAsText().setBold(0, 6, true).setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  
  // Disposiciones estándar
  body.appendParagraph('Disposiciones:').editAsText().setBold(true).setFontSize(11);
  body.appendParagraph('• Se ha considerado la evaluación realizada por el área de logística desde la oficina de administración y según el estudio de mercado (cuadro comparativo)').editAsText().setFontSize(10);
  body.appendParagraph('• La presente autorización presupuestal se emite en base a la disponibilidad presupuestal aprobada para la iniciativa').editAsText().setFontSize(10);
  body.appendParagraph('• El responsable de la ejecución del gasto deberá presentar la documentación sustentatoria de acuerdo a las normas vigentes.').editAsText().setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('Adjuntos: Documento sustentatorio obligatorios (contrataciones, proformas, términos de referencia, etc.)').editAsText().setBold(0, 8, true).setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('Firmado en fecha {{FECHA_FIRMA}} por:').editAsText().setBold(0, 16, true).setFontSize(10);
  
  body.appendParagraph(''); // Espaciado para firma
  body.appendParagraph(''); 
  body.appendParagraph(''); 
  body.appendParagraph(''); 
  
  // FIRMA DE JORGE HERRERA
  body.appendParagraph('_'.repeat(35)).setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  body.appendParagraph('Jorge Herrera').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setBold(true).setFontSize(11);
  body.appendParagraph('Director Ejecutivo').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
  body.appendParagraph('Cáritas Lima').setAlignment(DocumentApp.HorizontalAlignment.CENTER).editAsText().setFontSize(10);
  
  body.appendParagraph(''); // Espaciado
  body.appendParagraph('*Control electrónico con asunto - Re: FP 149 Aprobación cédula Solicitud {{NUMERO_SOLICITUD}} de {{ANNO}} ***COMPRA ADICIONAL ACEITE*** enviado por la Administración el {{FECHA_CONTROL}}*').editAsText().setFontSize(7);
  
  doc.saveAndClose();
  
  return {
    docId: doc.getId(),
    nombre: 'Plantilla_Jorge_Herrera',
    firmante: 'Jorge Herrera',
    cargo: 'Director Ejecutivo'
  };
}

// Función para ejecutar y actualizar todo
function actualizarSistemaCompleto() {
  try {
    Logger.log('🔄 Actualizando sistema completo...');
    
    // 1. Actualizar configuración con nuevos firmantes
    crearConfiguracionEjemplo();
    
    // 2. Crear plantillas con URLs
    const resultadoPlantillas = crearPlantillasConURLs();
    
    if (resultadoPlantillas.success) {
      Logger.log('✅ Sistema actualizado exitosamente');
      Logger.log('📋 Nuevos firmantes:');
      Logger.log('  - Evelyn Travezaño (Directora de Administración y Finanzas)');
      Logger.log('  - Jorge Herrera (Director Ejecutivo)');
      Logger.log('🔗 URLs de plantillas generadas');
      
      return { 
        success: true, 
        message: 'Sistema actualizado con nuevos firmantes y plantillas con URLs',
        plantillas: resultadoPlantillas.plantillas
      };
    } else {
      return { success: false, error: 'Error creando plantillas: ' + resultadoPlantillas.error };
    }
  } catch (error) {
    Logger.log('Error en actualizarSistemaCompleto: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}