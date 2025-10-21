var CP = typeof CP !== 'undefined' ? CP : {};

// =============================================================
// Núcleo de configuración y constantes
// =============================================================
(function (namespace) {
  const SHEETS = Object.freeze({
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

  const HEADERS = Object.freeze({
    CERTIFICACIONES: [
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
    ],
    ITEMS: [
      'Código Certificación',
      'Orden',
      'Descripción',
      'Cantidad',
      'Unidad',
      'Precio Unitario',
      'Subtotal',
      'Fecha Creación',
      'Creado Por'
    ],
    FIRMANTES: [
      'Código Certificación',
      'Orden',
      'Nombre',
      'Cargo',
      'Obligatorio',
      'Fecha Creación',
      'Creado Por'
    ],
    CATALOGOS: ['ID', 'Nombre', 'Descripción', 'Activo', 'Extra 1', 'Extra 2'],
    PLANTILLAS: [
      'ID',
      'Nombre',
      'Descripción',
      'Activa',
      'Firmantes',
      'Plantilla HTML',
      'Firmante ID',
      'Firmante Nombre',
      'Firmante Cargo'
    ],
    CONFIG_GENERAL: ['Clave', 'Valor', 'Descripción'],
    BITACORA: ['Fecha', 'Usuario', 'Acción', 'Detalles', 'Usuario Completo']
  });

  const FOLDERS = Object.freeze({
    PLANTILLAS: '1DXyPgOvEn4o-qPT945V_y7ctSQVnV8ce',
    CERTIFICACIONES: '1RJ4Ts7fATs_q3IINPTlLK6VHEl4l8hG3'
  });

  const SCRIPT_PROPERTIES = Object.freeze({
    FINALIDAD_ALIASES: 'FINALIDAD_DETALLADA_ALIASES',
    PLANTILLA_CACHE: 'CP_PLANTILLAS_CACHE'
  });

  const DEFAULT_TIMEZONE = (function () {
    try {
      return Session.getScriptTimeZone() || 'America/Lima';
    } catch (error) {
      Logger.log('Fallo obteniendo zona horaria, se usa America/Lima: ' + error);
      return 'America/Lima';
    }
  })();

  const DEFAULT_TEMPLATES = Object.freeze([
    {
      id: 'plantilla_evelyn',
      nombre: 'Certificación Evelyn',
      descripcion: 'Formato oficial con firma de Evelyn Elena Huaycacllo Marín',
      firmantes: 3,
      activa: true,
      firmanteId: 'firmante_evelyn',
      firmanteNombre: 'Evelyn Elena Huaycacllo Marín',
      firmanteCargo: 'Jefa de la Oficina de Política, Planeamiento y Presupuesto'
    },
    {
      id: 'plantilla_jorge',
      nombre: 'Certificación Jorge',
      descripcion: 'Formato oficial con firma de Jorge Herrera',
      firmantes: 3,
      activa: true,
      firmanteId: 'firmante_jorge',
      firmanteNombre: 'Jorge Herrera',
      firmanteCargo: 'Director Ejecutivo'
    },
    {
      id: 'plantilla_susana',
      nombre: 'Certificación Susana',
      descripcion: 'Formato oficial con firma de Susana Palomino',
      firmantes: 3,
      activa: true,
      firmanteId: 'firmante_susana',
      firmanteNombre: 'Susana Palomino',
      firmanteCargo: 'Coordinadora de Planeamiento y Presupuesto'
    },
    {
      id: 'plantilla_equipo',
      nombre: 'Certificación Equipo Designado',
      descripcion: 'Formato oficial para equipos designados',
      firmantes: 3,
      activa: true,
      firmanteId: 'firmante_equipo',
      firmanteNombre: 'Equipo Designado',
      firmanteCargo: 'Responsable según certificación'
    }
  ]);

  namespace.Constants = Object.freeze({
    SHEETS,
    HEADERS,
    FOLDERS,
    SCRIPT_PROPERTIES,
    DEFAULT_TIMEZONE,
    DEFAULT_TEMPLATES
  });
})(CP);

// =============================================================
// Utilidades generales
// =============================================================
(function (namespace) {
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

  function toBoolean(value) {
    if (typeof value === 'boolean') {
      return value;
    }
    if (typeof value === 'number') {
      return value !== 0;
    }
    if (typeof value === 'string') {
      return ['true', '1', 'si', 'sí', 'y', 'yes'].indexOf(value.toLowerCase()) !== -1;
    }
    return Boolean(value);
  }

  function now() {
    return new Date();
  }

  function toDate(value) {
    if (!value) {
      return null;
    }
    if (value instanceof Date) {
      return value;
    }
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }

  function formatDate(value, timezone) {
    const date = toDate(value);
    if (!date) {
      return '';
    }
    const tz = timezone || namespace.Constants.DEFAULT_TIMEZONE;
    return Utilities.formatDate(date, tz, 'yyyy-MM-dd');
  }

  function formatDateTime(value, timezone) {
    const date = toDate(value);
    if (!date) {
      return '';
    }
    const tz = timezone || namespace.Constants.DEFAULT_TIMEZONE;
    return Utilities.formatDate(date, tz, 'yyyy-MM-dd HH:mm:ss');
  }

  function formatCurrency(value) {
    const amount = toNumber(value);
    return amount.toLocaleString('es-PE', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function mapRowToObject(headers, row) {
    return headers.reduce(function (acc, header, index) {
      acc[header] = index < row.length ? row[index] : '';
      return acc;
    }, {});
  }

  function mapObjectToRow(headers, data) {
    return headers.map(function (header) {
      return data.hasOwnProperty(header) ? data[header] : '';
    });
  }

  function slugify(value) {
    return normalizeHeaderName(value).replace(/[^a-z0-9]+/g, '_');
  }

  function uniqueId(prefix) {
    return prefix + '_' + Utilities.getUuid().split('-')[0];
  }

  namespace.Utils = Object.freeze({
    normalizeHeaderName,
    normalizeString,
    toNumber,
    toBoolean,
    now,
    toDate,
    formatDate,
    formatDateTime,
    formatCurrency,
    mapRowToObject,
    mapObjectToRow,
    slugify,
    uniqueId
  });
})(CP);

// =============================================================
// Utilidades para hojas de cálculo
// =============================================================
(function (namespace) {
  const { SHEETS, HEADERS } = namespace.Constants;
  const { normalizeHeaderName } = namespace.Utils;

  function getSpreadsheet() {
    return SpreadsheetApp.getActiveSpreadsheet();
  }

  function getSheet(name) {
    return getSpreadsheet().getSheetByName(name);
  }

  function ensureSheet(name, headers) {
    const ss = getSpreadsheet();
    let sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
    }

    const requiredHeaders = headers && headers.length ? headers : null;

    if (requiredHeaders) {
      const range = sheet.getRange(1, 1, 1, requiredHeaders.length);
      const current = range.getValues()[0];
      const normalizedCurrent = current.map(normalizeHeaderName);
      const normalizedTarget = requiredHeaders.map(normalizeHeaderName);

      let needsUpdate = normalizedCurrent.length !== normalizedTarget.length;
      if (!needsUpdate) {
        for (let i = 0; i < normalizedTarget.length; i++) {
          if (normalizedCurrent[i] !== normalizedTarget[i]) {
            needsUpdate = true;
            break;
          }
        }
      }

      if (needsUpdate) {
        range.clearContent();
        range.offset(0, 0, 1, requiredHeaders.length).setValues([requiredHeaders]);
      }

      if (sheet.getFrozenRows() < 1) {
        sheet.setFrozenRows(1);
      }
    }

    return sheet;
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

  function readTable(name) {
    const sheet = ensureSheet(name);
    const values = sheet.getDataRange().getValues();
    if (!values.length) {
      return { headers: [], rows: [] };
    }
    const headers = values[0];
    const rows = values.slice(1).filter(function (row) {
      return row.some(function (cell) {
        return cell !== '' && cell !== null;
      });
    });
    return { headers, rows };
  }

  function writeTable(name, headers, rows) {
    const sheet = ensureSheet(name, headers);
    sheet.clearContents();
    const allValues = [headers].concat(rows || []);
    if (allValues.length) {
      sheet.getRange(1, 1, allValues.length, headers.length).setValues(allValues);
    }
  }

  function appendRow(name, headers, data) {
    const sheet = ensureSheet(name, headers);
    const row = namespace.Utils.mapObjectToRow(headers, data);
    sheet.appendRow(row);
    return sheet.getLastRow();
  }

  function updateRow(name, headers, rowIndex, data) {
    const sheet = ensureSheet(name, headers);
    const row = namespace.Utils.mapObjectToRow(headers, data);
    const range = sheet.getRange(rowIndex, 1, 1, headers.length);
    range.setValues([row]);
  }

  function findRow(name, headers, columnKey, value) {
    const table = readTable(name);
    if (!table.headers.length) {
      return { index: -1, row: null };
    }
    const normalizedHeaders = table.headers.map(normalizeHeaderName);
    const normalizedColumn = normalizeHeaderName(columnKey);
    const columnIndex = normalizedHeaders.indexOf(normalizedColumn);
    if (columnIndex === -1) {
      return { index: -1, row: null };
    }

    for (let i = 0; i < table.rows.length; i++) {
      if (table.rows[i][columnIndex] === value) {
        return { index: i + 2, row: table.rows[i] };
      }
    }

    return { index: -1, row: null };
  }

  namespace.Sheets = Object.freeze({
    getSpreadsheet,
    getSheet,
    ensureSheet,
    ensureBaseStructure,
    readTable,
    writeTable,
    appendRow,
    updateRow,
    findRow
  });
})(CP);

// =============================================================
// Utilidades para Drive y caché simple
// =============================================================
(function (namespace) {
  const { FOLDERS } = namespace.Constants;

  function getFolderById(id) {
    if (!id) {
      throw new Error('Se requiere un ID de carpeta válido');
    }
    return DriveApp.getFolderById(id);
  }

  function ensureFolder(id) {
    return getFolderById(id);
  }

  function moveFileToFolder(file, folderId) {
    const folder = ensureFolder(folderId);
    folder.addFile(file);
    const parents = file.getParents();
    while (parents.hasNext()) {
      const parent = parents.next();
      if (parent.getId() !== folderId) {
        parent.removeFile(file);
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

    list: function (sheetName) {
      const table = readTable(sheetName);
      return table.rows.map(function (row) {
        const item = mapRowToObject(table.headers, row);
        return {
          id: normalizeString(item.ID || item.Id || item.id),
          nombre: normalizeString(item.Nombre || item.nombre || ''),
          descripcion: normalizeString(item.Descripción || item.descripcion || item['Descripción'] || ''),
          activo: toBoolean(item.Activo || item.activo || true),
          extra1: normalizeString(item['Extra 1'] || item.extra1 || ''),
          extra2: normalizeString(item['Extra 2'] || item.extra2 || '')
        };
      }).filter(function (item) {
        return item.id;
      });
    },

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

    saveAll: function (plantillas) {
      const headers = HEADERS.PLANTILLAS;
      const rows = plantillas.map(function (p) {
        return [
          p.id,
          p.nombre,
          p.descripcion,
          p.activa !== false,
          p.firmantes || 1,
          p.plantillaHtml || '',
          p.firmanteId || '',
          p.firmanteNombre || '',
          p.firmanteCargo || ''
        ];
      });
      writeTable(SHEETS.PLANTILLAS, headers, rows);
    },

    upsert: function (plantilla) {
      const headers = HEADERS.PLANTILLAS;
      const result = findRow(SHEETS.PLANTILLAS, headers, 'ID', plantilla.id);
      const data = {
        ID: plantilla.id,
        Nombre: plantilla.nombre,
        Descripción: plantilla.descripcion,
        Activa: plantilla.activa !== false,
        Firmantes: plantilla.firmantes || 1,
        'Plantilla HTML': plantilla.plantillaHtml || '',
        'Firmante ID': plantilla.firmanteId || '',
        'Firmante Nombre': plantilla.firmanteNombre || '',
        'Firmante Cargo': plantilla.firmanteCargo || ''
      };
      if (result.index === -1) {
        appendRow(SHEETS.PLANTILLAS, headers, data);
      } else {
        updateRow(SHEETS.PLANTILLAS, headers, result.index, data);
      }
      SpreadsheetApp.flush();
    }
  };

  const ConfigRepository = {
    list: function () {
      const table = readTable(SHEETS.CONFIG_GENERAL);
      const result = {};
      table.rows.forEach(function (row) {
        const entry = mapRowToObject(table.headers, row);
        if (entry.Clave) {
          result[entry.Clave] = entry.Valor;
        }
      });
      return result;
    },

    saveAll: function (configMap) {
      const headers = HEADERS.CONFIG_GENERAL;
      const rows = Object.keys(configMap).map(function (key) {
        return [key, configMap[key], ''];
      });
      writeTable(SHEETS.CONFIG_GENERAL, headers, rows);
    },

    upsert: function (key, value, description) {
      const headers = HEADERS.CONFIG_GENERAL;
      const result = findRow(SHEETS.CONFIG_GENERAL, headers, 'Clave', key);
      const data = {
        Clave: key,
        Valor: value,
        Descripción: description || ''
      };
      if (result.index === -1) {
        appendRow(SHEETS.CONFIG_GENERAL, headers, data);
      } else {
        updateRow(SHEETS.CONFIG_GENERAL, headers, result.index, data);
      }
      SpreadsheetApp.flush();
    }
  };

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

    listFirmantes: function () {
      const table = readTable(SHEETS.FIRMANTES);
      return table.rows.map(function (row) {
        return mapRowToObject(table.headers, row);
      });
    },

    getByCodigo: function (codigo) {
      const headers = HEADERS.CERTIFICACIONES;
      const result = findRow(SHEETS.CERTIFICACIONES, headers, 'Código', codigo);
      if (result.index === -1) {
        return null;
      }
      return mapRowToObject(headers, result.row);
    },

    create: function (certificacion) {
      const headers = HEADERS.CERTIFICACIONES;
      const rowIndex = appendRow(SHEETS.CERTIFICACIONES, headers, certificacion);
      SpreadsheetApp.flush();
      return rowIndex;
    },

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

      items.forEach(function (item) {
        appendRow(SHEETS.ITEMS, headers, item);
      });
      SpreadsheetApp.flush();
    },

    replaceFirmantes: function (codigo, firmantes) {
      const headers = HEADERS.FIRMANTES;
      const sheet = ensureSheet(SHEETS.FIRMANTES, headers);
      const table = readTable(SHEETS.FIRMANTES);
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

  function ensurePlaceholders(count) {
    const placeholders = [];
    for (let i = 1; i <= count; i++) {
      placeholders.push({
        id: 'custom_' + i,
        nombre: 'Plantilla Personalizada ' + i,
        descripcion: 'Espacio para personalizar un formato adicional',
        firmantes: 3,
        activa: i === 1,
        firmanteId: '',
        firmanteNombre: '',
        firmanteCargo: ''
      });
    }
    return placeholders;
  }

  function seedPlantillas() {
    const existentes = repos.Plantilla.list();
    const mapa = {};
    existentes.forEach(function (plantilla) {
      mapa[plantilla.id] = plantilla;
    });

    const plantillas = DEFAULT_TEMPLATES.map(function (base) {
      if (mapa[base.id]) {
        return Object.assign({}, mapa[base.id], base);
      }
      return base;
    });

    ensurePlaceholders(5).forEach(function (placeholder) {
      if (!mapa[placeholder.id]) {
        plantillas.push(placeholder);
      } else {
        plantillas.push(Object.assign({}, mapa[placeholder.id]));
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

  function seedConfiguracionGeneral() {
    repos.Config.saveAll({
      disposicion_base_legal: 'Directiva 003-2023-SG/CARITASLIMA',
      codigo_formato: 'CP-{YEAR}-{NUMBER}',
      timezone: namespace.Constants.DEFAULT_TIMEZONE,
      moneda_por_defecto: 'PEN'
    });
  }

  function seedExampleCertificacion() {
    const certificaciones = repos.Certificacion.listCertificaciones();
    if (certificaciones.length) {
      return;
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

  function numeroALetras(numero) {
    const unidades = ['cero', 'uno', 'dos', 'tres', 'cuatro', 'cinco', 'seis', 'siete', 'ocho', 'nueve'];
    const especiales = ['diez', 'once', 'doce', 'trece', 'catorce', 'quince', 'dieciséis', 'diecisiete', 'dieciocho', 'diecinueve'];
    const decenas = ['', '', 'veinte', 'treinta', 'cuarenta', 'cincuenta', 'sesenta', 'setenta', 'ochenta', 'noventa'];
    const centenas = ['', 'cien', 'doscientos', 'trescientos', 'cuatrocientos', 'quinientos', 'seiscientos', 'setecientos', 'ochocientos', 'novecientos'];

    function convertirMenor100(n) {
      if (n < 10) return unidades[n];
      if (n < 20) return especiales[n - 10];
      const dec = Math.floor(n / 10);
      const uni = n % 10;
      if (uni === 0) return decenas[dec];
      if (dec === 2) return 'veinti' + unidades[uni];
      return decenas[dec] + ' y ' + unidades[uni];
    }

    function convertirMenor1000(n) {
      if (n < 100) return convertirMenor100(n);
      const cen = Math.floor(n / 100);
      const resto = n % 100;
      if (cen === 1 && resto === 0) return 'cien';
      const nombreCen = cen === 1 ? 'ciento' : centenas[cen];
      if (resto === 0) return nombreCen;
      return nombreCen + ' ' + convertirMenor100(resto);
    }

    function convertir(n) {
      if (n < 1000) return convertirMenor1000(n);
      if (n < 1000000) {
        const miles = Math.floor(n / 1000);
        const resto = n % 1000;
        const milesTexto = miles === 1 ? 'mil' : convertirMenor1000(miles) + ' mil';
        if (resto === 0) return milesTexto;
        return milesTexto + ' ' + convertirMenor1000(resto);
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

  function convertirNumeroALetras(numero) {
    return numeroALetras(utils.toNumber(numero));
  }

  function buildDocument(certificacion, items, firmantes) {
    const folder = drive.ensureFolder(FOLDERS.CERTIFICACIONES);
    const fileName = certificacion['Código'] + ' - ' + certificacion.Descripción;
    const document = DocumentApp.create(fileName);
    const docFile = DriveApp.getFileById(document.getId());
    drive.moveFileToFolder(docFile, FOLDERS.CERTIFICACIONES);

    const body = document.getBody();
    body.appendParagraph('CERTIFICACIÓN PRESUPUESTAL').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Código: ' + certificacion['Código']);
    body.appendParagraph('Fecha: ' + certificacion['Fecha Emisión']);
    body.appendParagraph('Descripción: ' + certificacion.Descripción);
    body.appendParagraph('Finalidad: ' + certificacion['Finalidad Detallada']);
    body.appendParagraph('Monto total: S/ ' + utils.formatCurrency(certificacion['Monto Total']));
    body.appendParagraph('Monto en letras: ' + certificacion['Monto en Letras']);

    body.appendParagraph('Items').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    items.forEach(function (item) {
      body.appendParagraph('- ' + item.Descripción + ' (' + item.Cantidad + ' ' + item.Unidad + ') - S/ ' + utils.formatCurrency(item.Subtotal));
    });

    body.appendParagraph('Firmantes').setHeading(DocumentApp.ParagraphHeading.HEADING2);
    firmantes.forEach(function (firmante) {
      body.appendParagraph(firmante.Nombre + ' - ' + firmante.Cargo);
    });

    document.saveAndClose();

    const pdfBlob = docFile.getAs(MimeType.PDF);
    const pdfFile = folder.createFile(pdfBlob).setName(fileName + '.pdf');

    return {
      urlDocumento: document.getUrl(),
      urlPDF: pdfFile.getUrl()
    };
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

    const certificacion = {
      'Código': codigo,
      'Fecha Emisión': fechaEmision,
      Descripción: payload.descripcion || '',
      Iniciativa: payload.iniciativa || '',
      Tipo: payload.tipo || '',
      Fuente: payload.fuente || '',
      Finalidad: payload.finalidad || payload.finalidadDetallada || '',
      Oficina: payload.oficina || '',
      Solicitante: payload.solicitante || '',
      'Cargo Solicitante': solicitante ? solicitante.descripcion : '',
      'Email Solicitante': solicitante ? solicitante.extra1 : '',
      'Número Autorización': payload.numeroAutorizacion || '',
      'Cargo Autorizador': payload.cargoAutorizador || '',
      Estado: payload.estado || 'Borrador',
      'Disposición/Base Legal': payload.disposicion || configuracionGeneral.disposicion_base_legal || '',
      'Monto Total': utils.toNumber(payload.montoTotal || payload.total || 0),
      'Monto en Letras': payload.montoLetras || convertirNumeroALetras(payload.montoTotal || payload.total || 0),
      'Fecha Creación': utils.formatDateTime(ahora),
      'Creado Por': Session.getActiveUser().getEmail(),
      'Fecha Modificación': '',
      'Modificado Por': '',
      'Fecha Anulación': '',
      'Anulado Por': '',
      'Motivo Anulación': '',
      Plantilla: payload.plantilla || DEFAULT_TEMPLATES[0].id,
      'URL Documento': '',
      'URL PDF': '',
      'Finalidad Detallada': payload.finalidadDetallada || payload.finalidad || ''
    };

    repos.Certificacion.create(certificacion);

    const items = (payload.items || []).map(function (item, index) {
      return {
        'Código Certificación': codigo,
        Orden: index + 1,
        Descripción: item.descripcion || '',
        Cantidad: utils.toNumber(item.cantidad || item.cant),
        Unidad: item.unidad || 'Unidad',
        'Precio Unitario': utils.toNumber(item.precioUnitario || item.precio),
        Subtotal: utils.toNumber(item.subtotal || 0),
        'Fecha Creación': utils.formatDateTime(ahora),
        'Creado Por': Session.getActiveUser().getEmail()
      };
    });

    if (items.length) {
      repos.Certificacion.replaceItems(codigo, items);
    }

    const firmantes = (payload.firmantes || []).map(function (firmante, index) {
      return {
        'Código Certificación': codigo,
        Orden: index + 1,
        Nombre: firmante.nombre || '',
        Cargo: firmante.cargo || '',
        Obligatorio: utils.toBoolean(firmante.obligatorio !== false),
        'Fecha Creación': utils.formatDateTime(ahora),
        'Creado Por': Session.getActiveUser().getEmail()
      };
    });

    if (firmantes.length) {
      repos.Certificacion.replaceFirmantes(codigo, firmantes);
    }

    return {
      success: true,
      codigo,
      certificacion
    };
  }

  function listarCertificaciones() {
    const certificaciones = repos.Certificacion.listCertificaciones();
    return certificaciones.map(function (cert) {
      cert['Monto Total'] = utils.toNumber(cert['Monto Total']);
      return cert;
    });
  }

  function obtenerEstadisticasDashboard() {
    const certificaciones = listarCertificaciones();
    if (!certificaciones.length) {
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

    const porEstado = certificaciones.reduce(function (acc, cert) {
      const estado = cert.Estado || 'Sin estado';
      acc[estado] = (acc[estado] || 0) + 1;
      return acc;
    }, {});

    const montoTotal = certificaciones.reduce(function (acc, cert) {
      return acc + utils.toNumber(cert['Monto Total']);
    }, 0);

    const recientes = certificaciones
      .slice()
      .sort(function (a, b) {
        const fechaA = utils.toDate(a['Fecha Creación']);
        const fechaB = utils.toDate(b['Fecha Creación']);
        return fechaB - fechaA;
      })
      .slice(0, 5);

    return {
      success: true,
      data: {
        total: certificaciones.length,
        montoTotal,
        porEstado,
        certificacionesRecientes: recientes
      }
    };
  }

  function generarDocumentoCertificacion(codigo) {
    const certificacion = repos.Certificacion.getByCodigo(codigo);
    if (!certificacion) {
      return { success: false, error: 'Certificación no encontrada' };
    }
    const items = repos.Certificacion.listItems().filter(function (item) {
      return item['Código Certificación'] === codigo;
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

  function obtenerCatalogo(tipo) {
    const mapa = {
      iniciativas: SHEETS.CATALOGO_INICIATIVAS,
      tipos: SHEETS.CATALOGO_TIPOS,
      fuentes: SHEETS.CATALOGO_FUENTES,
      finalidades: SHEETS.CATALOGO_FINALIDADES,
      oficinas: SHEETS.CATALOGO_OFICINAS,
      plantillas: SHEETS.PLANTILLAS,
      solicitantes: SHEETS.CONFIG_SOLICITANTES,
      firmantes: SHEETS.CONFIG_FIRMANTES
    };
    const sheetName = mapa[tipo];
    if (!sheetName) {
      return [];
    }
    if (sheetName === SHEETS.PLANTILLAS) {
      return repos.Plantilla.list();
    }
    return repos.Catalog.list(sheetName);
  }

  function actualizarSolicitante(id, payload) {
    if (!id) {
      return { success: false, error: 'ID requerido' };
    }
    const registro = {
      nombre: payload.nombre,
      descripcion: payload.cargo,
      activo: payload.activo !== false,
      extra1: payload.email || '',
      extra2: ''
    };
    repos.Catalog.upsert(SHEETS.CONFIG_SOLICITANTES, id, registro);
    return { success: true };
  }

  function eliminarSolicitante(id) {
    if (!id) {
      return { success: false, error: 'ID requerido' };
    }
    return repos.Catalog.remove(SHEETS.CONFIG_SOLICITANTES, id);
  }

  function actualizarFirmante(id, payload) {
    if (!id) {
      return { success: false, error: 'ID requerido' };
    }
    const registro = {
      nombre: payload.nombre,
      descripcion: payload.cargo,
      activo: payload.activo !== false,
      extra1: payload.orden || '',
      extra2: ''
    };
    repos.Catalog.upsert(SHEETS.CONFIG_FIRMANTES, id, registro);
    return { success: true };
  }

  function eliminarFirmante(id) {
    if (!id) {
      return { success: false, error: 'ID requerido' };
    }
    return repos.Catalog.remove(SHEETS.CONFIG_FIRMANTES, id);
  }

  function actualizarConfiguracionGeneral(payload) {
    Object.keys(payload).forEach(function (key) {
      repos.Config.upsert(key, payload[key], '');
    });
    return { success: true };
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

function crearSoloEstructura() {
  return CP.Controllers.crearSoloEstructura();
}

function normalizarEstructuraSistema() {
  return CP.Controllers.normalizarEstructuraSistema();
}

function resetearSistema() {
  return CP.Controllers.resetearSistema();
}
