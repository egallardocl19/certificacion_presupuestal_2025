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
  plantilla_director: {
    nombre: 'Padre Miguel Ángel Castillo Seminario',
    cargo: 'Director Ejecutivo'
  },
  plantilla_1_firmante: {
    nombre: 'Evelyn Elena Huaycacllo Marin',
    cargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto'
  }
});

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
  return (value === null || value === undefined)
    ? ''
    : String(value).trim().toLowerCase();
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

function mapRowToCertificacion(row, index, finalidadDetalladaIndex = -1) {
  const tieneFinalidadDetallada = finalidadDetalladaIndex >= 0 && finalidadDetalladaIndex < row.length;
  return {
    codigo: row[0],
    fechaEmision: row[1],
    descripcion: row[2],
    iniciativa: row[3],
    tipo: row[4],
    fuente: row[5],
    finalidad: row[6],
    oficina: row[7],
    solicitante: row[8],
    cargoSolicitante: row[9],
    emailSolicitante: row[10],
    numeroAutorizacion: row[11],
    cargoAutorizador: row[12],
    estado: row[13],
    disposicion: row[14],
    montoTotal: row[15] || 0,
    montoLetras: row[16],
    fechaCreacion: row[17],
    creadoPor: row[18],
    fechaModificacion: row[19],
    modificadoPor: row[20],
    fechaAnulacion: row[21],
    anuladoPor: row[22],
    motivoAnulacion: row[23],
    plantilla: row[24],
    urlDocumento: row[25],
    urlPDF: row[26],
    finalidadDetallada: tieneFinalidadDetallada ? row[finalidadDetalladaIndex] : '',
    fila: index + 1
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

  function normalizeString(value) {
    if (value === null || value === undefined) {
      return '';
    }
    return String(value).trim();
  }

  function toNumber(value) {
    if (value === null || value === undefined || value === '') {
      return 0;
    }
    const number = Number(value);
    return Number.isFinite(number) ? number : 0;
  }

function crearCertificacion(datos) {
  try {
    const sheet = getSheetOrThrow(SHEET_NAMES.CERTIFICACIONES);

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

    const fila = [
      codigo,
      fechaCertificacion,
      datosCompletos.descripcion,
      datosCompletos.iniciativa,
      datosCompletos.tipo,
      datosCompletos.fuente,
      finalidad,
      datosCompletos.oficina,
      datosCompletos.solicitante,
      datosCompletos.cargoSolicitante,
      datosCompletos.emailSolicitante,
      '',
      '',
      ESTADOS.BORRADOR,
      disposicion,
      0,
      '',
      fechaActual,
      usuario,
      fechaActual,
      usuario,
      '',
      '',
      '',
      plantilla,
      '',
      '',
      finalidadDetallada
    ];

    sheet.appendRow(fila);

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

    return {
      success: true,
      codigo,
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

  function formatDate(value, timezone) {
    const date = toDate(value);
    if (!date) {
      return '';
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
  }

function obtenerCertificaciones(filtros = {}) {
  try {
    const sheet = getSheetOrThrow(SHEET_NAMES.CERTIFICACIONES);
    const data = getSheetValues(sheet);

    if (data.length <= 1) return [];

    const headers = data[0] || [];
    const finalidadDetalladaIndex = findColumnIndexByAliases(
      headers,
      getFinalidadDetalladaAliases(),
      headers.length > 27 ? 27 : -1
    );

    const certificaciones = data
      .slice(1)
      .map((row, index) => {
        if (!row[0]) {
          return null;
        }
        return mapRowToCertificacion(row, index + 1, finalidadDetalladaIndex);
      })
      .filter(Boolean)
      .filter(cert => {
        if (filtros.estado && cert.estado !== filtros.estado) return false;
        if (filtros.oficina && cert.oficina !== filtros.oficina) return false;
        if (filtros.busqueda) {
          const busqueda = filtros.busqueda.toLowerCase();
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

    return certificaciones;
  } catch (error) {
    Logger.log('Error en obtenerCertificaciones: ' + error.toString());
    throw new Error('No se pudieron obtener certificaciones: ' + error.message);
  }

  function ensureBaseStructure() {
    ensureSheet(SHEETS.CERTIFICACIONES, HEADERS.CERTIFICACIONES);
    ensureSheet(SHEETS.ITEMS, HEADERS.ITEMS);
    ensureSheet(SHEETS.FIRMANTES, HEADERS.FIRMANTES);
    ensureSheet(SHEETS.CONFIG_SOLICITANTES, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.CONFIG_FIRMANTES, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.CONFIG_GENERAL, HEADERS.CONFIG_GENERAL);
    ensureSheet(SHEETS.CATALOGO_INICIATIVAS, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.CATALOGO_TIPOS, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.CATALOGO_FUENTES, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.CATALOGO_FINALIDADES, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.CATALOGO_OFICINAS, HEADERS.CATALOGOS);
    ensureSheet(SHEETS.PLANTILLAS, HEADERS.PLANTILLAS);
    ensureSheet(SHEETS.BITACORA, HEADERS.BITACORA);
  }

function actualizarCertificacion(codigo, datos) {
  try {
    const sheet = getSheetOrThrow(SHEET_NAMES.CERTIFICACIONES);
    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const filaIndex = findRowIndex(values, 0, codigo);

    if (filaIndex === -1) {
      return { success: false, error: 'Certificación no encontrada' };
    }

    const usuario = getActiveUserEmail();
    const fechaActual = new Date();
    const finalidadDetalladaIndex = getFinalidadDetalladaColumnIndex(sheet);
    const puedeActualizarFinalidadDetallada = finalidadDetalladaIndex >= 0 && finalidadDetalladaIndex < values[0].length;

    if (datos.fechaEmision !== undefined) {
      values[filaIndex][1] = parseDate(datos.fechaEmision);
    }
    if (datos.descripcion !== undefined) {
      values[filaIndex][2] = sanitizeText(datos.descripcion);
      if (!datos.finalidad) {
        const finalidadAuto = generarFinalidadAutomatica(datos.descripcion);
        values[filaIndex][6] = finalidadAuto;
        if (puedeActualizarFinalidadDetallada) {
          values[filaIndex][finalidadDetalladaIndex] = finalidadAuto;
        }
      }
    }
    if (datos.finalidad !== undefined) {
      const finalidadActualizada = sanitizeText(datos.finalidad);
      values[filaIndex][6] = finalidadActualizada;
      if (puedeActualizarFinalidadDetallada) {
        values[filaIndex][finalidadDetalladaIndex] = finalidadActualizada;
      }
    }
    if (datos.iniciativa !== undefined) values[filaIndex][3] = sanitizeText(datos.iniciativa);
    if (datos.tipo !== undefined) values[filaIndex][4] = sanitizeText(datos.tipo);
    if (datos.fuente !== undefined) values[filaIndex][5] = sanitizeText(datos.fuente);
    if (datos.oficina !== undefined) values[filaIndex][7] = sanitizeText(datos.oficina);
    if (datos.solicitante !== undefined) values[filaIndex][8] = sanitizeText(datos.solicitante);
    if (datos.cargoSolicitante !== undefined) values[filaIndex][9] = sanitizeText(datos.cargoSolicitante);
    if (datos.emailSolicitante !== undefined) values[filaIndex][10] = sanitizeText(datos.emailSolicitante);
    if (datos.numeroAutorizacion !== undefined) values[filaIndex][11] = sanitizeText(datos.numeroAutorizacion);
    if (datos.cargoAutorizador !== undefined) values[filaIndex][12] = sanitizeText(datos.cargoAutorizador);
    if (datos.estado !== undefined) values[filaIndex][13] = sanitizeText(datos.estado);
    if (datos.disposicion !== undefined) values[filaIndex][14] = sanitizeText(datos.disposicion);
    if (datos.urlDocumento !== undefined) values[filaIndex][25] = sanitizeText(datos.urlDocumento);
    if (datos.urlPDF !== undefined) values[filaIndex][26] = sanitizeText(datos.urlPDF);

    values[filaIndex][19] = fechaActual;
    values[filaIndex][20] = usuario;

    if (datos.estado === ESTADOS.ANULADA) {
      values[filaIndex][21] = fechaActual;
      values[filaIndex][22] = usuario;
      values[filaIndex][23] = sanitizeText(datos.motivoAnulacion);
    }

    dataRange.setValues(values);

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

    return { index: -1, row: null };
  }

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

  function ensureFolder(id) {
    return getFolderById(id);
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
  }

  namespace.Drive = Object.freeze({
    getFolderById,
    ensureFolder,
    moveFileToFolder
  });
})(CP);

// =============================================================
// Repositorios de datos
// =============================================================
(function (namespace) {
  const { SHEETS, HEADERS, DEFAULT_TEMPLATES } = namespace.Constants;
  const { normalizeString, toBoolean, mapRowToObject } = namespace.Utils;
  const { readTable, writeTable, appendRow, ensureSheet, findRow } = namespace.Sheets;

  const CatalogRepository = {
    ensureStructure: function () {
      ensureSheet(SHEETS.CONFIG_SOLICITANTES, HEADERS.CATALOGOS);
      ensureSheet(SHEETS.CONFIG_FIRMANTES, HEADERS.CATALOGOS);
      ensureSheet(SHEETS.CATALOGO_INICIATIVAS, HEADERS.CATALOGOS);
      ensureSheet(SHEETS.CATALOGO_TIPOS, HEADERS.CATALOGOS);
      ensureSheet(SHEETS.CATALOGO_FUENTES, HEADERS.CATALOGOS);
      ensureSheet(SHEETS.CATALOGO_FINALIDADES, HEADERS.CATALOGOS);
      ensureSheet(SHEETS.CATALOGO_OFICINAS, HEADERS.CATALOGOS);
    },

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

    saveAll: function (sheetName, records) {
      const headers = HEADERS.CATALOGOS;
      const rows = records.map(function (record) {
        return [
          normalizeString(record.id),
          normalizeString(record.nombre),
          normalizeString(record.descripcion),
          record.activo === false ? false : true,
          normalizeString(record.extra1),
          normalizeString(record.extra2)
        ];
      });
      writeTable(sheetName, headers, rows);
    },

    upsert: function (sheetName, id, payload) {
      const headers = HEADERS.CATALOGOS;
      const result = findRow(sheetName, headers, 'ID', id);
      const data = {
        ID: id,
        Nombre: payload.nombre || '',
        Descripción: payload.descripcion || '',
        Activo: payload.activo === false ? false : true,
        'Extra 1': payload.extra1 || '',
        'Extra 2': payload.extra2 || ''
      };
      if (result.index === -1) {
        appendRow(sheetName, headers, data);
      } else {
        updateRow(sheetName, headers, result.index, data);
      }
      SpreadsheetApp.flush();
      return { success: true };
    },

    remove: function (sheetName, id) {
      const headers = HEADERS.CATALOGOS;
      const result = findRow(sheetName, headers, 'ID', id);
      if (result.index === -1) {
        return { success: false, error: 'Registro no encontrado' };
      }
      const sheet = ensureSheet(sheetName, headers);
      sheet.deleteRow(result.index);
      SpreadsheetApp.flush();
      return { success: true };
    }
  };

  const PlantillaRepository = {
    list: function () {
      const table = readTable(SHEETS.PLANTILLAS);
      if (!table.headers.length) {
        return [];
      }
      return table.rows.map(function (row) {
        const item = mapRowToObject(table.headers, row);
        return {
          id: normalizeString(item.ID || item.id),
          nombre: normalizeString(item.Nombre || item.nombre),
          descripcion: normalizeString(item.Descripción || item.descripcion),
          activa: toBoolean(item.Activa || item.activa),
          firmantes: Number(item.Firmantes || item.firmantes || 1),
          plantillaHtml: normalizeString(item['Plantilla HTML'] || item.plantillaHtml),
          firmanteId: normalizeString(item['Firmante ID'] || item.firmanteId),
          firmanteNombre: normalizeString(item['Firmante Nombre'] || item.firmanteNombre),
          firmanteCargo: normalizeString(item['Firmante Cargo'] || item.firmanteCargo)
        };
      });
    },

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
      if (result.index === -1) {
        appendRow(SHEETS.PLANTILLAS, headers, data);
      } else {
        updateRow(SHEETS.PLANTILLAS, headers, result.index, data);
      }
      SpreadsheetApp.flush();
    }
  };

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

  const CertificacionRepository = {
    listCertificaciones: function () {
      const table = readTable(SHEETS.CERTIFICACIONES);
      if (!table.headers.length) {
        return [];
      }
      return table.rows.map(function (row) {
        return mapRowToObject(table.headers, row);
      });
    },

    listItems: function () {
      const table = readTable(SHEETS.ITEMS);
      return table.rows.map(function (row) {
        return mapRowToObject(table.headers, row);
      });
    },

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

    update: function (codigo, payload) {
      const headers = HEADERS.CERTIFICACIONES;
      const result = findRow(SHEETS.CERTIFICACIONES, headers, 'Código', codigo);
      if (result.index === -1) {
        throw new Error('Certificación no encontrada');
      }
      updateRow(SHEETS.CERTIFICACIONES, headers, result.index, payload);
      SpreadsheetApp.flush();
    },

    replaceItems: function (codigo, items) {
      const headers = HEADERS.ITEMS;
      const sheet = ensureSheet(SHEETS.ITEMS, headers);
      const table = readTable(SHEETS.ITEMS);
      const columnIndex = table.headers
        .map(namespace.Utils.normalizeHeaderName)
        .indexOf(namespace.Utils.normalizeHeaderName('Código Certificación'));

      if (columnIndex !== -1) {
        for (let i = table.rows.length; i >= 1; i--) {
          if (table.rows[i - 1][columnIndex] === codigo) {
            sheet.deleteRow(i + 1);
          }
        }
      }

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

      firmantes.forEach(function (firmante) {
        appendRow(SHEETS.FIRMANTES, headers, firmante);
      });
      SpreadsheetApp.flush();
    }
  };

  namespace.Repositories = Object.freeze({
    Catalog: CatalogRepository,
    Plantilla: PlantillaRepository,
    Config: ConfigRepository,
    Certificacion: CertificacionRepository
  });
})(CP);

// =============================================================
// Servicios de negocio
// =============================================================
(function (namespace) {
  const { SHEETS, HEADERS, DEFAULT_TEMPLATES, FOLDERS } = namespace.Constants;
  const utils = namespace.Utils;
  const sheets = namespace.Sheets;
  const repos = namespace.Repositories;
  const drive = namespace.Drive;

  function buildCorrelative() {
    const certificaciones = repos.Certificacion.listCertificaciones();
    const year = new Date().getFullYear();
    const prefix = 'CP-' + year;
    const correlatives = certificaciones
      .map(function (cert) {
        const code = utils.normalizeString(cert['Código'] || cert.codigo);
        if (!code || code.indexOf(prefix) !== 0) {
          return 0;
        }
        const numberPart = Number(code.split('-').pop());
        return Number.isFinite(numberPart) ? numberPart : 0;
      })
      .filter(function (value) {
        return value > 0;
      });
    const next = correlatives.length ? Math.max.apply(null, correlatives) + 1 : 1;
    return prefix + '-' + String(next).padStart(4, '0');
  }

function crearItemsCertificacion(codigoCertificacion, items) {
  try {
    const sheet = getSheetOrThrow(SHEET_NAMES.ITEMS);

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
      return base;
    });

function eliminarItemsCertificacion(codigoCertificacion) {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.ITEMS);

    if (!sheet) return { success: true };

    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === codigoCertificacion) {
        sheet.deleteRow(i + 1);
      }
    });

    repos.Plantilla.saveAll(plantillas);
    return plantillas;
  }

  function seedCatalogs() {
    repos.Catalog.saveAll(SHEETS.CATALOGO_INICIATIVAS, [
      { id: 'INI-001', nombre: 'Programa de Ayuda Social', descripcion: 'Iniciativa de apoyo social', activo: true },
      { id: 'INI-002', nombre: 'Proyecto de Infraestructura', descripcion: 'Mejoras de infraestructura', activo: true }
    ]);
    repos.Catalog.saveAll(SHEETS.CATALOGO_TIPOS, [
      { id: 'TIP-001', nombre: 'Bienes', descripcion: 'Adquisición de bienes', activo: true },
      { id: 'TIP-002', nombre: 'Servicios', descripcion: 'Contratación de servicios', activo: true }
    ]);
    repos.Catalog.saveAll(SHEETS.CATALOGO_FUENTES, [
      { id: 'FUE-001', nombre: 'Recursos Ordinarios', descripcion: '', activo: true },
      { id: 'FUE-002', nombre: 'Recursos Directamente Recaudados', descripcion: '', activo: true }
    ]);
    repos.Catalog.saveAll(SHEETS.CATALOGO_FINALIDADES, [
      { id: 'FIN-001', nombre: 'Atención a comunidades vulnerables', descripcion: '', activo: true },
      { id: 'FIN-002', nombre: 'Mejoras institucionales', descripcion: '', activo: true }
    ]);
    repos.Catalog.saveAll(SHEETS.CATALOGO_OFICINAS, [
      { id: 'OFI-001', nombre: 'Oficina de Planeamiento', descripcion: '', activo: true },
      { id: 'OFI-002', nombre: 'Oficina de Logística', descripcion: '', activo: true }
    ]);
    repos.Catalog.saveAll(SHEETS.CONFIG_SOLICITANTES, [
      { id: 'SOL-001', nombre: 'Carlos Rivera', descripcion: 'Coordinador de Logística', activo: true, extra1: 'carlos@caritas.pe' },
      { id: 'SOL-002', nombre: 'María Gonzales', descripcion: 'Analista de Planeamiento', activo: true, extra1: 'maria@caritas.pe' }
    ]);
    repos.Catalog.saveAll(SHEETS.CONFIG_FIRMANTES, [
      { id: 'FIR-001', nombre: 'Evelyn Elena Huaycacllo Marín', descripcion: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto', activo: true, extra1: '1' },
      { id: 'FIR-002', nombre: 'Jorge Herrera', descripcion: 'Director Ejecutivo', activo: true, extra1: '2' },
      { id: 'FIR-003', nombre: 'Susana Palomino', descripcion: 'Coordinadora de Planeamiento y Presupuesto', activo: true, extra1: '3' }
    ]);
  }

function crearFirmantesCertificacion(codigoCertificacion, firmantes) {
  try {
    const sheet = getSheetOrThrow(SHEET_NAMES.FIRMANTES);

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

    return { success: true };
  } catch (error) {
    Logger.log('Error en crearFirmantesCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
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
    const codigo = buildCorrelative();
    const fecha = utils.formatDate(new Date());
    const payload = {
      'Código': codigo,
      'Fecha Emisión': fecha,
      Descripción: 'Adquisición de kits de abrigo para comunidades vulnerables',
      Iniciativa: 'INI-001',
      Tipo: 'TIP-001',
      Fuente: 'FUE-001',
      Finalidad: 'Atender a familias afectadas por bajas temperaturas',
      Oficina: 'OFI-001',
      Solicitante: 'SOL-001',
      'Cargo Solicitante': 'Coordinador de Logística',
      'Email Solicitante': 'carlos@caritas.pe',
      'Número Autorización': '',
      'Cargo Autorizador': 'Director Ejecutivo',
      Estado: 'Activa',
      'Disposición/Base Legal': 'Directiva 003-2023-SG/CARITASLIMA',
      'Monto Total': 15000,
      'Monto en Letras': numeroALetras(15000),
      'Fecha Creación': utils.formatDateTime(new Date()),
      'Creado Por': Session.getActiveUser().getEmail(),
      'Fecha Modificación': '',
      'Modificado Por': '',
      'Fecha Anulación': '',
      'Anulado Por': '',
      'Motivo Anulación': '',
      Plantilla: 'plantilla_evelyn',
      'URL Documento': '',
      'URL PDF': '',
      'Finalidad Detallada': 'Atender a familias afectadas por bajas temperaturas con kits de abrigo.'
    };
    repos.Certificacion.create(payload);

    repos.Certificacion.replaceItems(codigo, [
      {
        'Código Certificación': codigo,
        Orden: 1,
        Descripción: 'Kit de abrigo completo',
        Cantidad: 150,
        Unidad: 'Unidades',
        'Precio Unitario': 100,
        Subtotal: 15000,
        'Fecha Creación': utils.formatDateTime(new Date()),
        'Creado Por': Session.getActiveUser().getEmail()
      }
    ]);

    repos.Certificacion.replaceFirmantes(codigo, [
      {
        'Código Certificación': codigo,
        Orden: 1,
        Nombre: 'Evelyn Elena Huaycacllo Marín',
        Cargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto',
        Obligatorio: true,
        'Fecha Creación': utils.formatDateTime(new Date()),
        'Creado Por': Session.getActiveUser().getEmail()
      },
      {
        'Código Certificación': codigo,
        Orden: 2,
        Nombre: 'Jorge Herrera',
        Cargo: 'Director Ejecutivo',
        Obligatorio: true,
        'Fecha Creación': utils.formatDateTime(new Date()),
        'Creado Por': Session.getActiveUser().getEmail()
      }
    ]);
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
      const millones = Math.floor(n / 1000000);
      const resto = n % 1000000;
      const millonesTexto = millones === 1 ? 'un millón' : convertir(millones) + ' millones';
      if (resto === 0) return millonesTexto;
      return millonesTexto + ' ' + convertir(resto);
    }

    const entero = Math.floor(numero);
    const decimales = Math.round((numero - entero) * 100);
    let resultado = convertir(entero).replace(/\buno\b/g, 'un');
    resultado = resultado.charAt(0).toUpperCase() + resultado.slice(1);
    if (decimales > 0) {
      resultado += ' con ' + convertir(decimales) + ' céntimos';
    }
    return resultado + ' soles';
  }

function crearFirmantesBasadosEnPlantilla(codigoCertificacion, plantillaId) {
  try {
    const firmante = PLANTILLA_FIRMANTES[plantillaId] || PLANTILLA_FIRMANTES['plantilla_evelyn'];

    return crearFirmantesCertificacion(codigoCertificacion, [{
      nombre: firmante.nombre,
      cargo: firmante.cargo,
      obligatorio: true
    }]);
  } catch (error) {
    Logger.log('Error en crearFirmantesBasadosEnPlantilla: ' + error.toString());
    return { success: false, error: error.toString() };
  }

function obtenerFirmantePorPlantilla(plantillaId) {
  return PLANTILLA_FIRMANTES[plantillaId] || PLANTILLA_FIRMANTES['plantilla_evelyn'];
}

  function crearCertificacion(payload) {
    sheets.ensureBaseStructure();
    const codigo = buildCorrelative();
    const ahora = utils.now();
    const fechaEmision = payload.fechaCertificacion || payload.fecha || utils.formatDate(ahora);
    const configuracionGeneral = repos.Config.list();
    const solicitante = repos.Catalog.list(SHEETS.CONFIG_SOLICITANTES).find(function (item) {
      return item.id === payload.solicitante;
    });

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
          porEstado: {},
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

    return {
      success: true,
      data: {
        total: certificaciones.length,
        montoTotal,
        porEstado,
        certificacionesRecientes: recientes
      }

      // Contar por oficina
      const nombreOficina = obtenerNombreOficina(cert.oficina);
      estadisticas.porOficina[nombreOficina] = (estadisticas.porOficina[nombreOficina] || 0) + 1;
    });
    const firmantes = repos.Certificacion.listFirmantes().filter(function (firmante) {
      return firmante['Código Certificación'] === codigo;
    });
    const resultado = buildDocument(certificacion, items, firmantes);
    certificacion['URL Documento'] = resultado.urlDocumento;
    certificacion['URL PDF'] = resultado.urlPDF;
    repos.Certificacion.update(codigo, certificacion);
    return { success: true, urlDocumento: resultado.urlDocumento, urlPDF: resultado.urlPDF };
  }

  function generarFinalidadConIA(payload) {
    try {
      const prompt = 'Genera una finalidad para una certificación presupuestal con la siguiente información: ' + JSON.stringify(payload);
      const response = UrlFetchApp.fetch('https://oi-server.onrender.com/chat/completions', {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        payload: JSON.stringify({
          model: 'openrouter/claude-sonnet-4',
          messages: [
            {
              role: 'system',
              content: 'Eres un asistente que resume finalidades presupuestales en español peruano.'
            },
            { role: 'user', content: prompt }
          ]
        })
      });
      const json = JSON.parse(response.getContentText());
      const texto = json.choices && json.choices.length ? json.choices[0].message.content.trim() : '';
      if (texto) {
        return { success: true, finalidad: texto };
      }
    } catch (error) {
      Logger.log('Fallo generando finalidad con IA: ' + error);
    }
    return {
      success: false,
      finalidad: 'Finalidad: ' + (payload.descripcion || 'Atender la necesidad presupuestal descrita.')
    };
  }

  function obtenerConfiguracionGeneral() {
    const config = repos.Config.list();
    return {
      disposicion_base_legal: config.disposicion_base_legal || '',
      codigo_formato: config.codigo_formato || 'CP-{YEAR}-{NUMBER}',
      timezone: config.timezone || namespace.Constants.DEFAULT_TIMEZONE,
      moneda_por_defecto: config.moneda_por_defecto || 'PEN'
    };
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
    const data = sheet.getDataRange().getValues();

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

// ===============================================
// FUNCIONES DE UTILIDAD
// ===============================================

function registrarActividad(accion, detalles = '') {
  try {
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAMES.BITACORA);
    if (!sheet) return;

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

  function configurarSistema() {
    sheets.ensureBaseStructure();
    seedCatalogs();
    seedPlantillas();
    seedConfiguracionGeneral();
    seedExampleCertificacion();
    SpreadsheetApp.flush();
    return { success: true, message: 'Sistema configurado correctamente.' };
  }

  function crearSoloEstructura() {
    sheets.ensureBaseStructure();
    seedPlantillas();
    SpreadsheetApp.flush();
    return { success: true, message: 'Estructura base creada.' };
  }

  function normalizarEstructuraSistema() {
    sheets.ensureBaseStructure();
    seedPlantillas();
    SpreadsheetApp.flush();
    return { success: true, message: 'Estructura normalizada.' };
  }

  function resetearSistema() {
    sheets.ensureBaseStructure();
    const { HEADERS } = namespace.Constants;

    sheets.writeTable(namespace.Constants.SHEETS.CERTIFICACIONES, HEADERS.CERTIFICACIONES, []);
    sheets.writeTable(namespace.Constants.SHEETS.ITEMS, HEADERS.ITEMS, []);
    sheets.writeTable(namespace.Constants.SHEETS.FIRMANTES, HEADERS.FIRMANTES, []);
    sheets.writeTable(namespace.Constants.SHEETS.CONFIG_SOLICITANTES, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CONFIG_FIRMANTES, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CATALOGO_INICIATIVAS, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CATALOGO_TIPOS, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CATALOGO_FUENTES, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CATALOGO_FINALIDADES, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CATALOGO_OFICINAS, HEADERS.CATALOGOS, []);
    sheets.writeTable(namespace.Constants.SHEETS.PLANTILLAS, HEADERS.PLANTILLAS, []);
    sheets.writeTable(namespace.Constants.SHEETS.CONFIG_GENERAL, HEADERS.CONFIG_GENERAL, []);
    sheets.writeTable(namespace.Constants.SHEETS.BITACORA, HEADERS.BITACORA, []);

    seedCatalogs();
    seedPlantillas();
    seedConfiguracionGeneral();
    seedExampleCertificacion();
    return { success: true, message: 'Sistema reiniciado correctamente.' };
  }

  namespace.Services = Object.freeze({
    crearCertificacion,
    listarCertificaciones,
    obtenerEstadisticasDashboard,
    generarDocumentoCertificacion,
    convertirNumeroALetras,
    generarFinalidadConIA,
    obtenerConfiguracionGeneral,
    obtenerCatalogo,
    actualizarSolicitante,
    eliminarSolicitante,
    actualizarFirmante,
    eliminarFirmante,
    actualizarConfiguracionGeneral,
    configurarSistema,
    crearSoloEstructura,
    normalizarEstructuraSistema,
    resetearSistema
  });
})(CP);

// =============================================================
// Controladores expuestos a la interfaz y al despliegue web
// =============================================================
(function (namespace) {
  const services = namespace.Services;

  function doGet() {
    return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Certificaciones Presupuestales');
  }

  function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
  }

  function obtenerCertificaciones() {
    return services.listarCertificaciones();
  }

  function obtenerEstadisticasDashboard() {
    return services.obtenerEstadisticasDashboard();
  }

  function crearCertificacion(payload) {
    return services.crearCertificacion(payload);
  }

  function generarDocumentoCertificacion(codigo) {
    return services.generarDocumentoCertificacion(codigo);
  }

  function convertirNumeroALetras(numero) {
    return services.convertirNumeroALetras(numero);
  }

  function generarFinalidadConIA(payload) {
    return services.generarFinalidadConIA(payload);
  }

  function obtenerConfiguracionGeneral() {
    return services.obtenerConfiguracionGeneral();
  }

  function obtenerCatalogo(tipo) {
    return services.obtenerCatalogo(tipo);
  }

  function actualizarSolicitante(id, payload) {
    return services.actualizarSolicitante(id, payload);
  }

  function eliminarSolicitante(id) {
    return services.eliminarSolicitante(id);
  }

  function actualizarFirmante(id, payload) {
    return services.actualizarFirmante(id, payload);
  }

  function eliminarFirmante(id) {
    return services.eliminarFirmante(id);
  }

  function actualizarConfiguracionGeneral(payload) {
    return services.actualizarConfiguracionGeneral(payload);
  }

  function configurarSistema() {
    return services.configurarSistema();
  }

  function crearSoloEstructura() {
    return services.crearSoloEstructura();
  }

  function normalizarEstructuraSistema() {
    return services.normalizarEstructuraSistema();
  }

  function resetearSistema() {
    return services.resetearSistema();
  }

  namespace.Controllers = Object.freeze({
    doGet,
    include,
    obtenerCertificaciones,
    obtenerEstadisticasDashboard,
    crearCertificacion,
    generarDocumentoCertificacion,
    convertirNumeroALetras,
    generarFinalidadConIA,
    obtenerConfiguracionGeneral,
    obtenerCatalogo,
    actualizarSolicitante,
    eliminarSolicitante,
    actualizarFirmante,
    eliminarFirmante,
    actualizarConfiguracionGeneral,
    configurarSistema,
    crearSoloEstructura,
    normalizarEstructuraSistema,
    resetearSistema
  });
})(CP);

// =============================================================
// Exposición global para Apps Script
// =============================================================
function doGet(e) {
  return CP.Controllers.doGet(e);
}

function include(filename) {
  return CP.Controllers.include(filename);
}

function obtenerCertificaciones() {
  return CP.Controllers.obtenerCertificaciones();
}

function obtenerEstadisticasDashboard() {
  return CP.Controllers.obtenerEstadisticasDashboard();
}

function crearCertificacion(payload) {
  return CP.Controllers.crearCertificacion(payload);
}

function generarDocumentoCertificacion(codigo) {
  return CP.Controllers.generarDocumentoCertificacion(codigo);
}

function convertirNumeroALetras(numero) {
  return CP.Controllers.convertirNumeroALetras(numero);
}

function generarFinalidadConIA(payload) {
  return CP.Controllers.generarFinalidadConIA(payload);
}

function obtenerConfiguracionGeneral() {
  return CP.Controllers.obtenerConfiguracionGeneral();
}

function obtenerCatalogo(tipo) {
  return CP.Controllers.obtenerCatalogo(tipo);
}

function actualizarSolicitante(id, payload) {
  return CP.Controllers.actualizarSolicitante(id, payload);
}

function eliminarSolicitante(id) {
  return CP.Controllers.eliminarSolicitante(id);
}

function actualizarFirmante(id, payload) {
  return CP.Controllers.actualizarFirmante(id, payload);
}

function eliminarFirmante(id) {
  return CP.Controllers.eliminarFirmante(id);
}

function actualizarConfiguracionGeneral(payload) {
  return CP.Controllers.actualizarConfiguracionGeneral(payload);
}

function configurarSistema() {
  return CP.Controllers.configurarSistema();
}

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
