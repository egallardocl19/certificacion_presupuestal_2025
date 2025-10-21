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

const FINALIDAD_DETALLADA_ALIASES_PROPERTY = 'FINALIDAD_DETALLADA_ALIASES';
const DEFAULT_FINALIDAD_DETALLADA_ALIASES = Object.freeze([
  'finalidad detallada',
  'finalidad detallada / justificación',
  'finalidad detallada / justificacion',
  'finalidad (detalle)',
  'detalle de la finalidad',
  'detalle finalidad',
  'justificación',
  'justificacion'
]);

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

function getFinalidadDetalladaAliases() {
  try {
    const properties = PropertiesService.getScriptProperties();
    const rawAliases = properties.getProperty(FINALIDAD_DETALLADA_ALIASES_PROPERTY);

    if (!rawAliases) {
      return DEFAULT_FINALIDAD_DETALLADA_ALIASES;
    }

    const parsedAliases = JSON.parse(rawAliases);
    if (!Array.isArray(parsedAliases)) {
      return DEFAULT_FINALIDAD_DETALLADA_ALIASES;
    }

    const normalizedAliases = parsedAliases
      .map(normalizeHeaderName)
      .filter(Boolean);

    const uniqueAliases = Array.from(new Set(normalizedAliases));
    return uniqueAliases.length > 0 ? uniqueAliases : DEFAULT_FINALIDAD_DETALLADA_ALIASES;
  } catch (error) {
    Logger.log('No se pudieron obtener alias personalizados de finalidad detallada: ' + error.toString());
    return DEFAULT_FINALIDAD_DETALLADA_ALIASES;
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
    const pdf = DriveApp.createFile(
      doc.getAs(MimeType.PDF).setName(`Certificacion_${codigoCertificacion}.pdf`)
    );
    
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
    return [];
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
}

function obtenerFirmantePorPlantilla(plantillaId) {
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
      
      if (tipo === 'plantillas') {
        catalogo.push({
          id: row[0],
          nombre: row[1],
          descripcion: row[2],
          activa: row[3] !== false,
          firmantes: row[4] || 1,
          docId: row[5] || ''
        });
      } else {
        catalogo.push({
          codigo: row[0],
          nombre: row[1],
          descripcion: row[2] || '',
          activo: row[3] !== false
        });
      }
    }
    
    return catalogo.filter(item => tipo === 'plantillas' ? item.activa : item.activo);
  } catch (error) {
    Logger.log('Error en obtenerCatalogo: ' + error.toString());
    return [];
  }
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
      montoTotal: certificaciones.reduce((sum, cert) => sum + (cert.montoTotal || 0), 0),
      porEstado: {
        'Borrador': 0,
        'En revisión': 0,
        'Autorización pendiente': 0,
        'Activa': 0,
        'Anulada': 0
      },
      porOficina: {},
      certificacionesRecientes: certificaciones.slice(0, 10)
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
    const total = items.reduce((sum, item) => sum + (item.subtotal || 0), 0);
    const montoLetras = convertirNumeroALetras(total);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Certificaciones');
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === codigoCertificacion) {
        sheet.getRange(i + 1, 16).setValue(total);
        sheet.getRange(i + 1, 17).setValue(montoLetras);
        break;
      }
    }
    
    return { success: true, total: total, montoLetras: montoLetras };
  } catch (error) {
    Logger.log('Error en recalcularTotalesCertificacion: ' + error.toString());
    return { success: false, error: error.toString() };
  }
}

function convertirNumeroALetras(numero) {
  try {
    if (numero === 0) return 'CERO CON 00/100 SOLES';
    
    const unidades = ['', 'UNO', 'DOS', 'TRES', 'CUATRO', 'CINCO', 'SEIS', 'SIETE', 'OCHO', 'NUEVE'];
    const decenas = ['', '', 'VEINTE', 'TREINTA', 'CUARENTA', 'CINCUENTA', 'SESENTA', 'SETENTA', 'OCHENTA', 'NOVENTA'];
    const especiales = ['DIEZ', 'ONCE', 'DOCE', 'TRECE', 'CATORCE', 'QUINCE', 'DIECISÉIS', 'DIECISIETE', 'DIECIOCHO', 'DIECINUEVE'];
    const centenas = ['', 'CIENTO', 'DOSCIENTOS', 'TRESCIENTOS', 'CUATROCIENTOS', 'QUINIENTOS', 'SEISCIENTOS', 'SETECIENTOS', 'OCHOCIENTOS', 'NOVECIENTOS'];
    
    const entero = Math.floor(numero);
    const decimales = Math.round((numero - entero) * 100);
    
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
        if (d === 1 && u > 0) {
          resultado += (resultado ? ' ' : '') + especiales[u];
        } else {
          resultado += (resultado ? ' ' : '') + decenas[d];
          if (u > 0) {
            resultado += (d === 2 ? '' : ' Y ') + unidades[u];
          }
        }
      } else if (u > 0) {
        resultado += (resultado ? ' ' : '') + unidades[u];
      }
      
      return resultado;
    }
    
    let resultado = '';
    
    if (entero >= 1000000) {
      const millones = Math.floor(entero / 1000000);
      resultado += convertirGrupo(millones);
      if (millones === 1) {
        resultado += ' MILLÓN';
      } else {
        resultado += ' MILLONES';
      }
      entero = entero % 1000000;
    }
    
    if (entero >= 1000) {
      const miles = Math.floor(entero / 1000);
      if (miles > 1) {
        resultado += (resultado ? ' ' : '') + convertirGrupo(miles) + ' MIL';
      } else {
        resultado += (resultado ? ' ' : '') + 'MIL';
      }
      entero = entero % 1000;
    }
    
    if (entero > 0) {
      resultado += (resultado ? ' ' : '') + convertirGrupo(entero);
    }
    
    const decimalesStr = decimales.toString().padStart(2, '0');
    return resultado + ` CON ${decimalesStr}/100 SOLES`;
  } catch (error) {
    Logger.log('Error en convertirNumeroALetras: ' + error.toString());
    return 'ERROR EN CONVERSIÓN';
  }
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
    const pdf = DriveApp.createFile(
      doc.getAs(MimeType.PDF).setName(`Certificacion_${codigoCertificacion}.pdf`)
    );
    
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