/**
 * TableCrafter AI
 *
 * Este archivo contiene la lógica principal del complemento TableCrafter AI para
 * Google Sheets.  Se han implementado las funciones públicas existentes
 * (onOpen, showSidebar, showConfig, showHelp) y nuevas utilidades para dar
 * soporte al formateo de tablas con banda de color y vistas de filtro
 * dedicadas.  También se gestiona un registro de tablas en una hoja oculta
 * (__TableCrafter_Meta) y el acceso opcional a un modelo de IA mediante
 * OpenAI.  Esta implementación es autocontenida y lista para copiar/pegar.
 *
 * Notas importantes:
 *   - No modifique los nombres de funciones públicas ya existentes para
 *     mantener la compatibilidad con el marketplace de Google Workspace.
 *   - Se ha habilitado el servicio avanzado de Sheets (Sheets API v4) para
 *     crear y eliminar vistas de filtro.
 *   - La API Key de OpenAI se guarda en las propiedades de usuario a
 *     través del diálogo de configuración.  Nunca se expone al frontend.
 */

/**
 * Añade el menú personalizado al abrir la hoja de cálculo.
 */
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('TableCrafter AI')
    .addItem('Mostrar TableCrafter', 'showSidebar')
    .addItem('Configurar API', 'showConfig')
    .addItem('Ayuda', 'showHelp')
    .addToUi();
}

/**
 * Muestra la barra lateral principal del complemento.
 */
function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('UI')
    .setTitle('TableCrafter AI');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Muestra el panel de configuración de la API.
 */
function showConfig() {
  var html = HtmlService.createHtmlOutputFromFile('Config')
    .setWidth(400)
    .setHeight(220)
    .setTitle('Configuración de TableCrafter');
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración de TableCrafter');
}

/**
 * Muestra la página de ayuda.  Incluye enlaces a privacidad y términos.
 */
function showHelp() {
  var html = HtmlService.createHtmlOutputFromFile('Help')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Ayuda de TableCrafter');
  SpreadsheetApp.getUi().showModalDialog(html, 'Ayuda de TableCrafter');
}

/**
 * Devuelve información sobre el rango activo.  Incluye nombre de la hoja,
 * notación A1, número de filas y columnas y los encabezados detectados en
 * la primera fila del rango.  Se utiliza para pre‑poblar la UI.
 *
 * @returns {Object} Objeto con a1Notation, sheetName, rows, cols y headers.
 */
function getActiveRangeInfo() {
  var ss = SpreadsheetApp.getActive();
  var range = ss.getActiveRange();
  if (!range) {
    return { error: 'No hay un rango seleccionado' };
  }
  var sheet = range.getSheet();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var values = range.getValues();
  var headers = [];
  if (values.length > 0) {
    var firstRow = values[0];
    for (var i = 0; i < firstRow.length; i++) {
      var cell = firstRow[i];
      var headerText = cell === null || cell === undefined ? '' : String(cell);
      headers.push(normalizeHeaderEntry(headerText));
    }
  }
  headers = ensureHeaderKeys(headers);
  return {
    a1Notation: range.getA1Notation(),
    sheetName: sheet.getName(),
    rows: numRows,
    cols: numCols,
    headers: headers
  };
}

var META_HEADERS = ['id', 'name', 'rangeA1', 'description', 'sheet', 'cols', 'rows', 'createdAt', 'updatedAt', 'style', 'headers'];
var META_INDEX = {
  id: 0,
  name: 1,
  rangeA1: 2,
  description: 3,
  sheet: 4,
  cols: 5,
  rows: 6,
  createdAt: 7,
  updatedAt: 8,
  style: 9,
  headers: 10
};

var ALL_TABLES_OPTION_VALUE = '__ALL__';

function normalizeMetaId(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function normalizeMetaName(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim().toLowerCase();
}

function normalizeHeaderKey(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function normalizeHeaderLabelValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim().toLowerCase();
}

function generateUniqueHeaderKey(usedKeys) {
  var key;
  var attempts = 0;
  do {
    key = 'hdr_' + Utilities.getUuid();
    attempts++;
  } while (usedKeys[key] && attempts < 10);
  while (usedKeys[key]) {
    key = 'hdr_' + Utilities.getUuid();
  }
  return key;
}

function parseJsonValue(value, fallback) {
  if (!value) {
    return fallback;
  }
  try {
    return JSON.parse(value);
  } catch (err) {
    return fallback;
  }
}

function stringifyJsonValue(value) {
  if (value === undefined) {
    return '';
  }
  try {
    return JSON.stringify(value);
  } catch (err) {
    return '';
  }
}

function normalizeHeaderEntry(entry) {
  var normalized = {
    label: '',
    description: '',
    hasDescription: false,
    key: ''
  };
  if (entry && typeof entry === 'object') {
    var labelValue = '';
    if (Object.prototype.hasOwnProperty.call(entry, 'label')) {
      labelValue = entry.label;
    } else if (Object.prototype.hasOwnProperty.call(entry, 'text')) {
      labelValue = entry.text;
    }
    var descriptionValue = Object.prototype.hasOwnProperty.call(entry, 'description') ? entry.description : '';
    var hasDescription = Object.prototype.hasOwnProperty.call(entry, 'hasDescription')
      ? !!entry.hasDescription
      : String(descriptionValue || '').trim() !== '';
    var keyValue = '';
    if (Object.prototype.hasOwnProperty.call(entry, 'key')) {
      keyValue = entry.key;
    } else if (Object.prototype.hasOwnProperty.call(entry, 'id')) {
      keyValue = entry.id;
    }
    normalized.label = String(labelValue === undefined || labelValue === null ? '' : labelValue).trim();
    normalized.description = hasDescription ? String(descriptionValue || '').trim() : '';
    normalized.hasDescription = hasDescription;
    normalized.key = normalizeHeaderKey(keyValue);
    return normalized;
  }
  if (entry !== undefined && entry !== null) {
    normalized.label = String(entry).trim();
  }
  return normalized;
}

function normalizeHeaderArray(headers) {
  if (!Array.isArray(headers)) {
    return [];
  }
  return headers.map(function(item) {
    return normalizeHeaderEntry(item);
  });
}

function ensureHeaderKeys(headers) {
  if (!Array.isArray(headers)) {
    return [];
  }
  var used = {};
  return headers.map(function(item, index) {
    var normalized = normalizeHeaderEntry(item);
    var key = normalizeHeaderKey(normalized.key);
    if (!key) {
      key = 'col_' + (index + 1);
    }
    if (used[key]) {
      key = generateUniqueHeaderKey(used);
    }
    normalized.key = key;
    used[key] = true;
    return normalized;
  });
}

function buildColumnMapping(newHeaders, previousHeaders, previousTableData, targetColumnCount) {
  var mapping = [];
  var normalizedNewHeaders = Array.isArray(newHeaders) ? newHeaders : [];
  var normalizedPreviousHeaders = Array.isArray(previousHeaders) ? previousHeaders : [];
  var previousKeyIndexMap = {};
  var previousLabelIndexMap = {};
  for (var i = 0; i < normalizedPreviousHeaders.length; i++) {
    var prevItem = normalizeHeaderEntry(normalizedPreviousHeaders[i]);
    var prevKey = normalizeHeaderKey(prevItem.key);
    if (prevKey && !previousKeyIndexMap.hasOwnProperty(prevKey)) {
      previousKeyIndexMap[prevKey] = i;
    }
    var prevLabel = normalizeHeaderLabelValue(prevItem.label);
    if (prevLabel) {
      if (!previousLabelIndexMap[prevLabel]) {
        previousLabelIndexMap[prevLabel] = [];
      }
      previousLabelIndexMap[prevLabel].push(i);
    }
  }
  var previousHeaderRow = [];
  if (previousTableData && previousTableData.values && previousTableData.values.length > 0) {
    previousHeaderRow = previousTableData.values[0] || [];
    for (var hr = 0; hr < previousHeaderRow.length; hr++) {
      var headerLabel = normalizeHeaderLabelValue(previousHeaderRow[hr]);
      if (headerLabel) {
        if (!previousLabelIndexMap[headerLabel]) {
          previousLabelIndexMap[headerLabel] = [];
        }
        if (previousLabelIndexMap[headerLabel].indexOf(hr) === -1) {
          previousLabelIndexMap[headerLabel].push(hr);
        }
      }
    }
  }
  var previousColumnCount = 0;
  if (previousTableData && previousTableData.values && previousTableData.values.length > 0) {
    previousColumnCount = previousTableData.values[0].length;
  } else if (previousTableData && previousTableData.formulas && previousTableData.formulas.length > 0) {
    previousColumnCount = previousTableData.formulas[0].length;
  } else if (normalizedPreviousHeaders.length > 0) {
    previousColumnCount = normalizedPreviousHeaders.length;
  }
  var usedIndices = {};
  var totalColumns = typeof targetColumnCount === 'number' && targetColumnCount > 0 ? targetColumnCount : normalizedNewHeaders.length;
  for (var col = 0; col < totalColumns; col++) {
    var newHeader = normalizeHeaderEntry(normalizedNewHeaders[col]);
    var desiredKey = normalizeHeaderKey(newHeader.key);
    var desiredLabel = normalizeHeaderLabelValue(newHeader.label);
    var sourceIndex = -1;
    if (desiredKey && previousKeyIndexMap.hasOwnProperty(desiredKey)) {
      var candidateIndex = previousKeyIndexMap[desiredKey];
      if (!usedIndices[candidateIndex]) {
        sourceIndex = candidateIndex;
      }
    }
    if (sourceIndex === -1 && desiredLabel && previousLabelIndexMap[desiredLabel]) {
      var labelCandidates = previousLabelIndexMap[desiredLabel];
      while (labelCandidates.length > 0 && sourceIndex === -1) {
        var candidate = labelCandidates.shift();
        if (!usedIndices[candidate]) {
          sourceIndex = candidate;
        }
      }
      previousLabelIndexMap[desiredLabel] = labelCandidates;
    }
    if (sourceIndex === -1) {
      if (!usedIndices[col] && col < previousColumnCount) {
        sourceIndex = col;
      } else {
        for (var fallback = 0; fallback < previousColumnCount; fallback++) {
          if (!usedIndices[fallback]) {
            sourceIndex = fallback;
            break;
          }
        }
      }
    }
    if (sourceIndex !== -1) {
      usedIndices[sourceIndex] = true;
    }
    mapping.push(sourceIndex);
  }
  return mapping;
}

function reorderRowByMapping(row, mapping, defaultValue) {
  var result = [];
  var sourceRow = Array.isArray(row) ? row : [];
  for (var i = 0; i < mapping.length; i++) {
    var sourceIndex = mapping[i];
    if (sourceIndex !== -1 && sourceRow && sourceIndex < sourceRow.length) {
      result.push(sourceRow[sourceIndex]);
    } else {
      result.push(defaultValue);
    }
  }
  return result;
}

function reorderDataRows(matrix, mapping, startIndex, rowCount, defaultValue) {
  var result = [];
  var sourceMatrix = Array.isArray(matrix) ? matrix : [];
  var totalRows = typeof rowCount === 'number' && rowCount >= 0 ? rowCount : 0;
  for (var r = 0; r < totalRows; r++) {
    var sourceRowIndex = startIndex + r;
    var row = sourceRowIndex >= 0 && sourceRowIndex < sourceMatrix.length ? sourceMatrix[sourceRowIndex] : null;
    result.push(reorderRowByMapping(row, mapping, defaultValue));
  }
  return result;
}


/**
 * Devuelve el listado de tablas guardadas desde la hoja oculta
 * __TableCrafter_Meta.  Cada entrada contiene id, nombre y rango A1.
 *
 * @returns {Array} Lista de objetos con los metadatos de las tablas.
 */
function listSavedTables() {
  SpreadsheetApp.flush();
  var meta = getMetaSheet();
  var data = meta.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (entryId) {
      var style = parseJsonValue(row[META_INDEX.style], null);
      var headers = normalizeHeaderArray(parseJsonValue(row[META_INDEX.headers], null));
      result.push({
        id: entryId,
        name: row[META_INDEX.name],
        rangeA1: row[META_INDEX.rangeA1],
        description: row[META_INDEX.description],
        sheetName: row[META_INDEX.sheet],
        cols: row[META_INDEX.cols],
        rows: row[META_INDEX.rows],
        createdAt: row[META_INDEX.createdAt],
        updatedAt: row[META_INDEX.updatedAt],
        style: style,
        headers: headers
      });
    }
  }
  return result;
}

function focusSavedTableRange(tableId) {
  var id = normalizeMetaId(tableId);
  if (!id) {
    return { error: 'Tabla no encontrada.' };
  }
  var entry = findMetaById(id);
  if (!entry) {
    return { error: 'Tabla no encontrada.' };
  }
  var rowData = entry.data || [];
  var sheetName = rowData[META_INDEX.sheet];
  var rangeA1 = rowData[META_INDEX.rangeA1];
  if (!sheetName || !rangeA1) {
    return { error: 'La tabla no tiene un rango asociado.' };
  }
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { error: 'No se encontró la hoja "' + sheetName + '".' };
  }
  var range;
  try {
    range = sheet.getRange(rangeA1);
  } catch (err) {
    return { error: 'No se pudo obtener el rango: ' + err.message };
  }
  ss.setActiveSheet(sheet);
  range.activate();
  if (typeof range.getCell === 'function') {
    var firstCell = range.getCell(1, 1);
    if (firstCell && typeof firstCell.activateAsCurrentCell === 'function') {
      firstCell.activateAsCurrentCell();
    }
  }
  SpreadsheetApp.flush();
  return {
    ok: true,
    sheetName: sheetName,
    rangeA1: rangeA1,
    rows: range.getNumRows(),
    cols: range.getNumColumns()
  };
}

/**
 * Borra el rango y la vista de filtro de una tabla específica sin eliminar
 * sus metadatos.  Se utiliza desde la UI para limpiar el rastro.
 *
 * @param {string} tableId ID único de la tabla.
 * @returns {Object} Resultado con ok = true si se limpia correctamente.
 */
function clearTable(tableId) {
  var entry = findMetaById(tableId);
  if (!entry) {
    return { error: 'Tabla no encontrada' };
  }
  var row = entry.row;
  var data = entry.data;
  var rangeA1 = data[META_INDEX.rangeA1];
  var sheetName = data[META_INDEX.sheet];
  if (rangeA1) {
    try {
      var ss = SpreadsheetApp.getActive();
      var sheet = ss.getSheetByName(sheetName);
      var range = sheet.getRange(rangeA1);
      clearRangeAndFormatting(range);
      // Eliminar vista de filtro correspondiente
      var title = 'TableCrafter_' + tableId;
      deleteFilterViewsByTitle(sheet, title);
    } catch (err) {
      return { error: 'Error al limpiar la tabla: ' + err.message };
    }
  }
  return { ok: true };
}

/**
 * Borra por completo una tabla guardada: limpia el rango asociado, elimina la
 * vista de filtro y remueve los metadatos del registro.
 *
 * @param {string} tableId ID único de la tabla a borrar.
 * @returns {Object} Resultado con ok=true o un mensaje de error/advertencia.
 */
function deleteSavedTable(tableId) {
  var id = normalizeMetaId(tableId);
  if (!id) {
    return {
      ok: true,
      removed: true,
      message: 'Se ha retirado de la lista.'
    };
  }
  var entry = findMetaById(id);
  if (!entry) {
    return {
      ok: true,
      removed: true,
      message: 'Se ha retirado de la lista.'
    };
  }

  var data = entry.data;
  var rangeA1 = data[META_INDEX.rangeA1];
  var sheetName = data[META_INDEX.sheet];
  var warnings = [];

  if (rangeA1 && sheetName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      try {
        var range = sheet.getRange(rangeA1);
        clearRangeAndFormatting(range);
      } catch (err) {
        warnings.push('No se pudo limpiar el rango: ' + err.message);
      }
      try {
        deleteFilterViewsByTitle(sheet, 'TableCrafter_' + id);
      } catch (fvErr) {
        warnings.push('No se pudo eliminar la vista de filtro: ' + fvErr.message);
      }
    } else {
      warnings.push('La hoja asociada ya no existe.');
    }
  }

  var metaSheet = getMetaSheet();
  metaSheet.deleteRow(entry.row);
  SpreadsheetApp.flush();

  if (warnings.length > 0) {
    return { ok: true, warning: warnings.join(' ') };
  }
  return { ok: true };
}

/**
 * Obtiene un rango válido a partir de una notación A1. Si la notación no
 * incluye hoja, se asume la hoja activa.
 *
 * @param {string} rangeA1 Notación A1 ingresada por el usuario.
 * @returns {Range} Rango correspondiente.
 */
function resolveRangeFromNotation(rangeA1) {
  if (!rangeA1 || String(rangeA1).trim() === '') {
    throw new Error('Proporcione un rango en notación A1.');
  }
  var ss = SpreadsheetApp.getActive();
  var trimmed = String(rangeA1).trim();
  try {
    return ss.getRange(trimmed);
  } catch (err) {
    var sheet = ss.getActiveSheet();
    if (!sheet) {
      throw new Error('No se encontró la hoja activa.');
    }
    try {
      return sheet.getRange(trimmed);
    } catch (inner) {
      throw new Error('El rango indicado no es válido: ' + trimmed);
    }
  }
}

/**
 * Aplica formato de tabla al rango seleccionado y guarda/actualiza los
 * metadatos.  Si tableId se proporciona y existe, se limpiará el rango y
 * vista previos.  Devuelve el id de la tabla aplicada.
 *
 * @param {string} tableId ID existente o cadena vacía para nueva tabla.
 * @param {string} rangeA1 Notación A1 del rango seleccionado.
 * @param {string} name Nombre de la tabla.
 * @param {string} description Descripción de la tabla.
 * @param {Array} headers Lista de encabezados (fila 1) editados por el usuario.
 * @param {Object} style Objeto {headerColor, altColor1, altColor2, border, bold}.
 * @param {Object=} options Permite controlar acciones (skipFormatting, skipSaving).
 * @returns {Object} Resultado con id, mensaje y datos del rango, o error.
 */
function applyTableFormatting(tableId, rangeA1, name, description, headers, style, options) {
  options = options || {};
  var doFormat = !options.skipFormatting;
  var doSave = !options.skipSaving;
  var shouldUpdateFormulaReferences = true;
  if (options && Object.prototype.hasOwnProperty.call(options, 'updateFormulaReferences')) {
    shouldUpdateFormulaReferences = !!options.updateFormulaReferences;
  }
  if (!doFormat && !doSave) {
    return { error: 'No se especificó ninguna acción para realizar.' };
  }
  var range;
  try {
    range = resolveRangeFromNotation(rangeA1);
  } catch (err) {
    return { error: err.message };
  }
  var sheet = range.getSheet();
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  if (cols > 16 || rows > 1000) {
    return {
      error: 'El rango seleccionado (' + cols + ' columnas y ' + rows + ' filas) supera los límites permitidos (máx. 16 columnas y 1000 filas).'
    };
  }
  var id = tableId && tableId.trim() !== '' ? tableId.trim() : '';
  var isNew = false;
  if (doSave && !id) {
    id = Utilities.getUuid();
    isNew = true;
  }

  if (doSave) {
    var normalizedName = normalizeMetaName(name);
    if (normalizedName) {
      var duplicate = findMetaByName(name, { ignoreId: id });
      if (duplicate) {
        return {
          error: 'Ya existe una tabla guardada con ese nombre. Elija un nombre diferente antes de guardar.'
        };
      }
    }
  }

  var appliedStyle = {
    headerColor: style && style.headerColor ? style.headerColor : '#CFE8FC',
    altColor1: style && style.altColor1 ? style.altColor1 : '#FFFFFF',
    altColor2: style && style.altColor2 ? style.altColor2 : '#F3F4F6',
    border: style && Object.prototype.hasOwnProperty.call(style, 'border') ? !!style.border : true,
    bold: style && Object.prototype.hasOwnProperty.call(style, 'bold') ? !!style.bold : true
  };

  var normalizedRange = range.getA1Notation();
  var headerRange = range.offset(0, 0, 1, cols);
  var headerMeta = [];
  var shouldOverwriteHeaders = headers && headers.length > 0;
  if (shouldOverwriteHeaders) {
    var providedMeta = normalizeHeaderArray(headers);
    for (var i = 0; i < cols; i++) {
      headerMeta.push(providedMeta[i] || normalizeHeaderEntry(''));
    }
  } else {
    var existingHeaders = headerRange.getValues()[0];
    var derivedMeta = normalizeHeaderArray(existingHeaders);
    for (var j = 0; j < cols; j++) {
      headerMeta.push(derivedMeta[j] || normalizeHeaderEntry(''));
    }
  }
  headerMeta = ensureHeaderKeys(headerMeta);
  headerMeta = headerMeta.map(function(entry) {
    var labelValue = entry && entry.label !== undefined && entry.label !== null ? String(entry.label) : '';
    var descriptionValue = entry && entry.description !== undefined && entry.description !== null ? String(entry.description) : '';
    var trimmedLabel = labelValue.trim();
    var trimmedDescription = descriptionValue.trim();
    var explicitHasDescription = entry && Object.prototype.hasOwnProperty.call(entry, 'hasDescription') ? !!entry.hasDescription : false;
    var hasDescription = explicitHasDescription || trimmedDescription !== '';
    var keyValue = entry && entry.key !== undefined && entry.key !== null ? String(entry.key) : '';
    return {
      label: trimmedLabel,
      description: hasDescription ? trimmedDescription : '',
      hasDescription: hasDescription,
      key: normalizeHeaderKey(keyValue)
    };
  });
  var headerMetaForReturn = headerMeta.map(function(item) {
    return {
      label: item.label || '',
      description: item.hasDescription ? (item.description || '') : '',
      hasDescription: !!item.hasDescription,
      key: item.key || ''
    };
  });
  var headerValues = headerMetaForReturn.map(function(item) {
    return item.label || '';
  });
  var previousTableData = null;
  var previousTableRows = 0;
  var previousTableCols = 0;
  var storedHeaderMeta = [];
  var columnMapping = [];
  var previousRangeDetails = null;
  var shouldClearPreviousRange = false;
  var previousRangeReadError = false;
  if (doFormat) {
    if (id) {
      var entry = findMetaById(id);
      if (entry) {
        storedHeaderMeta = normalizeHeaderArray(parseJsonValue(entry.data[META_INDEX.headers], null));
        var prevRangeA1 = entry.data[META_INDEX.rangeA1];
        var prevSheetName = entry.data[META_INDEX.sheet];
        if (prevRangeA1) {
          var ss = SpreadsheetApp.getActive();
          var prevSheet = ss.getSheetByName(prevSheetName);
          if (prevSheet) {
            try {
              var prevRange = prevSheet.getRange(prevRangeA1);
              var sameSheet = prevSheet.getSheetId() === sheet.getSheetId();
              var sameRange = sameSheet && prevRange.getA1Notation() === normalizedRange;
              if (!sameRange && sameSheet && rangesIntersect(prevRange, range)) {
                return {
                  error: 'No se puede superponer el nuevo rango con el anterior.'
                };
              }
              try {
                previousTableData = {
                  values: prevRange.getValues(),
                  formulas: prevRange.getFormulas(),
                  horizontalAlignments: prevRange.getHorizontalAlignments(),
                  verticalAlignments: prevRange.getVerticalAlignments(),
                  fontWeights: prevRange.getFontWeights(),
                  numberFormats: prevRange.getNumberFormats()
                };
                previousTableRows = prevRange.getNumRows();
                previousTableCols = prevRange.getNumColumns();
              } catch (readErr) {
                previousRangeReadError = true;
                previousTableData = null;
              }
              if (!sameRange) {
                previousRangeDetails = { range: prevRange, sheet: prevSheet };
                if (previousTableData && shouldUpdateFormulaReferences) {
                  var copyRows = Math.min(previousTableRows, rows);
                  var copyCols = Math.min(previousTableCols, cols);
                  if (copyRows > 0 && copyCols > 0) {
                    try {
                      var sourceCopyRange = prevRange.offset(0, 0, copyRows, copyCols);
                      var destCopyRange = range.offset(0, 0, copyRows, copyCols);
                      sourceCopyRange.copyTo(destCopyRange, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
                      SpreadsheetApp.flush();
                      var copiedFormulas = destCopyRange.getFormulas();
                      var normalizedFormulas = [];
                      for (var nfRow = 0; nfRow < rows; nfRow++) {
                        var rowArray = [];
                        for (var nfCol = 0; nfCol < cols; nfCol++) {
                          rowArray.push('');
                        }
                        normalizedFormulas.push(rowArray);
                      }
                      for (var cfRow = 0; cfRow < copiedFormulas.length && cfRow < rows; cfRow++) {
                        var copiedRow = copiedFormulas[cfRow] || [];
                        for (var cfCol = 0; cfCol < copiedRow.length && cfCol < cols; cfCol++) {
                          if (copiedRow[cfCol]) {
                            normalizedFormulas[cfRow][cfCol] = copiedRow[cfCol];
                          }
                        }
                      }
                      previousTableData.formulas = normalizedFormulas;
                    } catch (copyErr) {
                      // Si falla la copia, continuar con las fórmulas originales.
                    }
                  }
                }
              }
              if (previousRangeReadError) {
                return {
                  error: 'No se pudo copiar la tabla original al nuevo rango. Intente aplicar el formato nuevamente antes de borrar el rango anterior.'
                };
              }
              deleteFilterViewsByTitle(prevSheet, 'TableCrafter_' + id);
            } catch (prevErr) {
              // Si el rango anterior no existe, continuar sin detener la ejecución.
            }
          }
        }
      }
    }
    if (previousRangeReadError) {
      return {
        error: 'No se pudo copiar la tabla original al nuevo rango. Intente aplicar el formato nuevamente antes de borrar el rango anterior.'
      };
    }
    if (previousTableData) {
      range.clearContent();
      columnMapping = buildColumnMapping(headerMeta, storedHeaderMeta, previousTableData, cols);
      if (!columnMapping || columnMapping.length !== cols) {
        columnMapping = [];
        for (var cm = 0; cm < cols; cm++) {
          columnMapping.push(cm);
        }
      }
    } else {
      columnMapping = [];
    }
    applyFormattingToRange(range, appliedStyle);
    if (shouldOverwriteHeaders) {
      headerRange.setValues([headerValues]);
    } else if (previousTableData && previousTableData.values && previousTableData.values.length > 0) {
      var reorderedHeaderValues = reorderRowByMapping(previousTableData.values[0], columnMapping, '');
      headerRange.setValues([reorderedHeaderValues]);
      if (previousTableData.formulas && previousTableData.formulas.length > 0) {
        var reorderedHeaderFormulas = reorderRowByMapping(previousTableData.formulas[0], columnMapping, '');
        for (var hf = 0; hf < reorderedHeaderFormulas.length; hf++) {
          var headerFormula = reorderedHeaderFormulas[hf];
          if (headerFormula && headerFormula !== '') {
            headerRange.getCell(1, hf + 1).setFormula(headerFormula);
          }
        }
      }
    }
    if (previousTableData) {
      var numberFormatsMatrix = null;
      if (previousTableData.numberFormats) {
        numberFormatsMatrix = [];
        for (var nfr = 0; nfr < rows; nfr++) {
          var numberFormatRow = nfr < previousTableData.numberFormats.length ? previousTableData.numberFormats[nfr] : null;
          numberFormatsMatrix.push(reorderRowByMapping(numberFormatRow, columnMapping, '@'));
        }
      }
      var totalTargetRows = rows - 1;
      var dataRange = null;
      var valuesMatrix = null;
      var formulasMatrix = null;
      var horizontalMatrix = null;
      var verticalMatrix = null;
      var weightMatrix = null;
      if (totalTargetRows > 0) {
        dataRange = range.offset(1, 0, totalTargetRows, cols);
        valuesMatrix = reorderDataRows(previousTableData.values, columnMapping, 1, totalTargetRows, '');
        if (previousTableData.formulas) {
          formulasMatrix = reorderDataRows(previousTableData.formulas, columnMapping, 1, totalTargetRows, '');
        }
        if (previousTableData.horizontalAlignments) {
          horizontalMatrix = reorderDataRows(previousTableData.horizontalAlignments, columnMapping, 1, totalTargetRows, null);
        }
        if (previousTableData.verticalAlignments) {
          verticalMatrix = reorderDataRows(previousTableData.verticalAlignments, columnMapping, 1, totalTargetRows, null);
        }
        if (previousTableData.fontWeights) {
          weightMatrix = reorderDataRows(previousTableData.fontWeights, columnMapping, 1, totalTargetRows, 'normal');
        }
      }
      if (numberFormatsMatrix && numberFormatsMatrix.length === rows) {
        range.setNumberFormats(numberFormatsMatrix);
      }
      if (dataRange && valuesMatrix) {
        if (horizontalMatrix) {
          dataRange.setHorizontalAlignments(horizontalMatrix);
        }
        if (verticalMatrix) {
          dataRange.setVerticalAlignments(verticalMatrix);
        }
        if (weightMatrix) {
          dataRange.setFontWeights(weightMatrix);
        }

        var hasAnyFormula = false;
        if (formulasMatrix) {
          for (var fmRow = 0; fmRow < formulasMatrix.length && !hasAnyFormula; fmRow++) {
            var fmRowValues = formulasMatrix[fmRow] || [];
            for (var fmCol = 0; fmCol < fmRowValues.length; fmCol++) {
              if (fmRowValues[fmCol]) {
                hasAnyFormula = true;
                break;
              }
            }
          }
        }

        if (hasAnyFormula) {
          var combinedMatrix = [];
          for (var cr = 0; cr < valuesMatrix.length; cr++) {
            var valueRow = valuesMatrix[cr] || [];
            var formulaRow = (formulasMatrix && formulasMatrix[cr]) || [];
            var combinedRow = [];
            var maxLength = Math.max(valueRow.length, formulaRow.length);
            for (var cc = 0; cc < maxLength; cc++) {
              var candidateFormula = formulaRow[cc];
              if (candidateFormula) {
                combinedRow.push(candidateFormula);
              } else {
                combinedRow.push(cc < valueRow.length ? valueRow[cc] : '');
              }
            }
            combinedMatrix.push(combinedRow);
          }
          dataRange.setValues(combinedMatrix);
        } else {
          dataRange.setValues(valuesMatrix);
        }
      }
      if (previousRangeDetails) {
        shouldClearPreviousRange = true;
      }
    }
    if (id) {
      deleteFilterViewsByTitle(sheet, 'TableCrafter_' + id);
      createFilterViewForRange(range, 'TableCrafter_' + id);
    }
  } else if (shouldOverwriteHeaders) {
    headerRange.setValues([headerValues]);
  }

  if (shouldClearPreviousRange && previousRangeDetails && previousRangeDetails.range) {
    SpreadsheetApp.flush();
    try {
      clearRangeAndFormatting(previousRangeDetails.range);
    } catch (cleanupErr) {
      // Ignorar errores al limpiar el rango anterior para no interrumpir la operación principal.
    }
  }

  if (doSave) {
    saveOrUpdateMeta(id, name, description, normalizedRange, sheet.getName(), cols, rows, appliedStyle, headerMetaForReturn);
    if (!doFormat && id) {
      deleteFilterViewsByTitle(sheet, 'TableCrafter_' + id);
      createFilterViewForRange(range, 'TableCrafter_' + id);
    }
  }

  var message;
  if (doFormat && doSave) {
    message = 'Formato aplicado y tabla guardada correctamente.';
  } else if (doFormat) {
    message = 'Formato aplicado correctamente.';
  } else if (doSave) {
    message = isNew ? 'Tabla guardada correctamente.' : 'Tabla actualizada correctamente.';
  } else {
    message = 'Operación realizada.';
  }

  var result = {
    id: id,
    rangeInfo: {
      sheetName: sheet.getName(),
      a1Notation: normalizedRange,
      rows: rows,
      cols: cols,
      headers: headerMetaForReturn
    },
    message: message
  };
  if (doSave) {
    result.isNew = isNew;
  }
  return result;
}

/**
 * Elimina contenidos, formato y bandas del rango proporcionado.
 *
 * @param {Range} range Rango que se desea limpiar.
 */
function clearRangeAndFormatting(range) {
  // Eliminar bandings que intersectan
  var sheet = range.getSheet();
  var bandings = sheet.getBandings();
  bandings.forEach(function(b) {
    var br = b.getRange();
    if (rangesIntersect(br, range)) {
      b.remove();
    }
  });
  // Borrar contenido y formato
  range.clear({ contentsOnly: true, formatOnly: true });
  // Eliminar filtros nativos si existiesen
  if (range.getFilter()) {
    range.getFilter().remove();
  }
}

/**
 * Limpia el contenido y formato de un rango indicado por notación A1.
 *
 * @param {string} rangeA1 Notación A1 (puede incluir la hoja).
 * @returns {Object} Resultado con ok=true o un mensaje de error.
 */
function clearRangeByNotation(rangeA1) {
  try {
    var range = resolveRangeFromNotation(rangeA1);
    clearRangeAndFormatting(range);
    return { ok: true };
  } catch (err) {
    return { error: 'No se pudo limpiar el rango: ' + err.message };
  }
}

/**
 * Aplica los estilos a un rango completo: encabezado, alternancia de filas,
 * bordes y negrita opcionales.
 *
 * @param {Range} range Rango objetivo.
 * @param {Object} style {headerColor, altColor1, altColor2, border, bold}.
 */
function applyFormattingToRange(range, style) {
  var sheet = range.getSheet();
  var cols = range.getNumColumns();
  var rows = range.getNumRows();
  // Eliminar bandings previos que intersectan
  var bandings = sheet.getBandings();
  bandings.forEach(function(b) {
    var br = b.getRange();
    if (rangesIntersect(br, range)) {
      b.remove();
    }
  });
  // Aplicar banding con colores personalizados y reforzar manualmente los colores
  var banded = range.applyRowBanding();
  banded.setHeaderRowColor(style.headerColor);
  banded.setFirstRowColor(style.altColor1);
  banded.setSecondRowColor(style.altColor2);
  // Centrar todas las celdas del rango para las tablas recién creadas
  range.setHorizontalAlignment('center');
  range.setVerticalAlignment('middle');
  // Encabezado en negrita o normal y centrado
  var headerRange = range.offset(0, 0, 1, cols);
  headerRange.setFontWeight(style.bold ? 'bold' : 'normal');
  headerRange.setHorizontalAlignment('center');
  headerRange.setVerticalAlignment('middle');
  headerRange.setBackground(style.headerColor);
  // Asegurar alternancia de colores para el resto de filas
  var dataRows = rows - 1;
  if (dataRows > 0) {
    var dataRange = range.offset(1, 0, dataRows, cols);
    var backgrounds = [];
    for (var r = 0; r < dataRows; r++) {
      var color = (r % 2 === 0) ? style.altColor1 : style.altColor2;
      var rowColors = [];
      for (var c = 0; c < cols; c++) {
        rowColors.push(color);
      }
      backgrounds.push(rowColors);
    }
    dataRange.setBackgrounds(backgrounds);
    dataRange.setFontWeight('normal');
  }
  // Bordes finos negros
  if (style.border) {
    range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    range.setBorder(false, false, false, false, false, false, null, null);
  }
}

/**
 * Crea una vista de filtro dedicada para un rango específico.  Utiliza la
 * Sheets API avanzada.
 *
 * @param {Range} range Rango donde se creará la vista de filtro.
 * @param {string} title Título único de la vista de filtro.
 */
function createFilterViewForRange(range, title) {
  var ss = SpreadsheetApp.getActive();
  var ssId = ss.getId();
  var sheet = range.getSheet();
  var sheetId = sheet.getSheetId();
  var startRowIndex = range.getRow() - 1;
  var endRowIndex = startRowIndex + range.getNumRows();
  var startColIndex = range.getColumn() - 1;
  var endColIndex = startColIndex + range.getNumColumns();
  var requests = [
    {
      addFilterView: {
        filter: {
          title: title,
          range: {
            sheetId: sheetId,
            startRowIndex: startRowIndex,
            endRowIndex: endRowIndex,
            startColumnIndex: startColIndex,
            endColumnIndex: endColIndex
          }
        }
      }
    }
  ];
  Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
}

/**
 * Elimina vistas de filtro de una hoja cuyo título coincida exactamente con
 * el especificado.  Se utiliza para borrar las vistas antes de recrearlas.
 *
 * @param {Sheet} sheet Hoja donde buscar la vista de filtro.
 * @param {string} title Título de la vista a eliminar.
 */
function deleteFilterViewsByTitle(sheet, title) {
  var ss = SpreadsheetApp.getActive();
  var ssId = ss.getId();
  var response = Sheets.Spreadsheets.get(ssId, {
    ranges: sheet.getSheetName(),
    fields: 'sheets(filterViews(filterViewId,title))'
  });
  var requests = [];
  var sheetViews = (response.sheets && response.sheets[0] && response.sheets[0].filterViews) || [];
  sheetViews.forEach(function(fv) {
    if (fv.title === title) {
      requests.push({ deleteFilterView: { filterId: fv.filterViewId } });
    }
  });
  if (requests.length > 0) {
    Sheets.Spreadsheets.batchUpdate({ requests: requests }, ssId);
  }
}

/**
 * Determina si dos rangos se intersectan.  Se utiliza para eliminar
 * correctamente bandings solapados.
 *
 * @param {Range} r1 Primer rango.
 * @param {Range} r2 Segundo rango.
 * @returns {boolean} True si los rangos se intersecan.
 */
function rangesIntersect(r1, r2) {
  if (r1.getSheet().getSheetId() !== r2.getSheet().getSheetId()) return false;
  var r1RowStart = r1.getRow();
  var r1RowEnd = r1.getLastRow();
  var r1ColStart = r1.getColumn();
  var r1ColEnd = r1.getLastColumn();
  var r2RowStart = r2.getRow();
  var r2RowEnd = r2.getLastRow();
  var r2ColStart = r2.getColumn();
  var r2ColEnd = r2.getLastColumn();
  var rowIntersect = !(r2RowStart > r1RowEnd || r2RowEnd < r1RowStart);
  var colIntersect = !(r2ColStart > r1ColEnd || r2ColEnd < r1ColStart);
  return rowIntersect && colIntersect;
}

/**
 * Obtiene (o crea si no existe) la hoja oculta para metadatos.
 *
 * @returns {Sheet} Hoja __TableCrafter_Meta.
 */
function getMetaSheet() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('__TableCrafter_Meta');
  if (!sheet) {
    sheet = ss.insertSheet('__TableCrafter_Meta');
    sheet.hideSheet();
    ensureMetaSheetSchema(sheet);
  } else {
    ensureMetaSheetSchema(sheet);
  }
  return sheet;
}

function ensureMetaSheetSchema(sheet) {
  var maxColumns = sheet.getMaxColumns();
  if (maxColumns < META_HEADERS.length) {
    sheet.insertColumnsAfter(maxColumns, META_HEADERS.length - maxColumns);
  }
  var headerRange = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), META_HEADERS.length));
  var headerValues = headerRange.getValues()[0];
  var needsUpdate = false;
  for (var i = 0; i < META_HEADERS.length; i++) {
    if (headerValues[i] !== META_HEADERS[i]) {
      needsUpdate = true;
      break;
    }
  }
  if (needsUpdate) {
    sheet.getRange(1, 1, 1, META_HEADERS.length).setValues([META_HEADERS]);
  }
}

/**
 * Busca una entrada de metadatos por ID.
 *
 * @param {string} id ID de la tabla.
 * @returns {Object|null} Objeto con {row, data} o null si no existe.
 */
function findMetaById(id) {
  var targetId = normalizeMetaId(id);
  if (!targetId) {
    return null;
  }
  var sheet = getMetaSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (normalizeMetaId(data[i][0]) === targetId) {
      return { row: i + 1, data: data[i] };
    }
  }
  return null;
}

function findMetaByName(name, options) {
  var normalizedName = normalizeMetaName(name);
  if (!normalizedName) {
    return null;
  }
  var opts = options || {};
  var ignoreId = normalizeMetaId(opts.ignoreId);
  var sheet = getMetaSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    if (ignoreId && entryId === ignoreId) {
      continue;
    }
    var entryName = normalizeMetaName(row[META_INDEX.name]);
    if (entryName && entryName === normalizedName) {
      return { row: i + 1, data: row };
    }
  }
  return null;
}

/**
 * Guarda o actualiza los metadatos en la hoja __TableCrafter_Meta.
 *
 * @param {string} id ID único de la tabla.
 * @param {string} name Nombre de la tabla.
 * @param {string} description Descripción.
 * @param {string} rangeA1 Rango A1 donde se aplicó la tabla.
 * @param {string} sheetName Nombre de la hoja.
 * @param {number} cols Número de columnas.
 * @param {number} rows Número de filas.
 */
function saveOrUpdateMeta(id, name, description, rangeA1, sheetName, cols, rows, style, headers) {
  var normalizedId = normalizeMetaId(id);
  if (!normalizedId) {
    throw new Error('No se pudo determinar el identificador de la tabla.');
  }
  var sheet = getMetaSheet();
  var meta = findMetaById(normalizedId);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var styleValue = stringifyJsonValue(style || {});
  var normalizedHeaders = normalizeHeaderArray(headers);
  var headersValue = stringifyJsonValue(normalizedHeaders);
  if (meta) {
    // Actualizar
    var row = meta.row;
    var createdAt = meta.data[META_INDEX.createdAt] || now;
    sheet
      .getRange(row, 1, 1, META_HEADERS.length)
      .setValues([[normalizedId, name, rangeA1, description, sheetName, cols, rows, createdAt, now, styleValue, headersValue]]);
  } else {
    // Crear nueva
    sheet.appendRow([normalizedId, name, rangeA1, description, sheetName, cols, rows, now, now, styleValue, headersValue]);
  }

  SpreadsheetApp.flush();
}

/**
 * Devuelve la clave de API almacenada (si existe).
 *
 * @returns {Object} Objeto con key o null.
 */
function getApiKey() {
  var key = PropertiesService.getUserProperties().getProperty('TC_API_KEY');
  return { key: key || '' };
}

/**
 * Guarda la clave de API proporcionada.
 *
 * @param {string} key Clave a almacenar.
 * @returns {Object} Resultado de guardado.
 */
function saveApiKey(key) {
  var userProps = PropertiesService.getUserProperties();
  if (!key) {
    userProps.deleteProperty('TC_API_KEY');
    return { ok: true };
  }
  userProps.setProperty('TC_API_KEY', key.trim());
  return { ok: true };
}

/**
 * Permite al usuario hacer una pregunta a la IA basada en una tabla
 * previamente guardada.  Prepara un contexto compacto con los encabezados y
 * hasta 500 filas de la tabla (o menos si hay menos datos) y realiza
 * una llamada al modelo configurado.  Se recomienda limitar el número de
 * filas para mantener la respuesta rápida.
 *
 * @param {string|string[]} tableSelection ID(s) de las tablas a consultar.
 * @param {string} question Pregunta del usuario.
 * @returns {Object} Objeto con answer o error.
 */
function askQuestion(tableSelection, question) {
  var apiKey = PropertiesService.getUserProperties().getProperty('TC_API_KEY');
  if (!apiKey) {
    return { error: 'No hay API Key configurada. Configure su clave en la sección de configuración.' };
  }

  var normalizedIds = normalizeTableSelection(tableSelection);
  if (normalizedIds.length === 0) {
    return { error: 'Seleccione al menos una tabla válida para consultar.' };
  }

  var ss = SpreadsheetApp.getActive();
  var remainingSampleRows = 500;
  var contexts = [];

  for (var i = 0; i < normalizedIds.length; i++) {
    if (remainingSampleRows <= 0) {
      break;
    }
    var tableId = normalizedIds[i];
    var metaEntry = findMetaById(tableId);
    if (!metaEntry) {
      continue;
    }

    var data = metaEntry.data;
    var rangeA1 = data[META_INDEX.rangeA1];
    var sheetName = data[META_INDEX.sheet];
    if (!rangeA1 || !sheetName) {
      continue;
    }

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      continue;
    }

    var range = sheet.getRange(rangeA1);
    var values = range.getDisplayValues();
    if (!values || values.length === 0) {
      continue;
    }

    var headers = values[0] || [];
    var dataRows = values.slice(1);
    var availableRows = dataRows.length;
    if (availableRows === 0) {
      contexts.push(buildTableContextSnippet(data, headers, [], sheetName, rangeA1));
      continue;
    }

    var sampleSize = Math.min(availableRows, remainingSampleRows);
    var samples = [];
    for (var j = 0; j < sampleSize; j++) {
      samples.push(dataRows[j].join(', '));
    }
    remainingSampleRows -= sampleSize;
    contexts.push(buildTableContextSnippet(data, headers, samples, sheetName, rangeA1));
  }

  if (contexts.length === 0) {
    return { error: 'No se encontraron datos en las tablas seleccionadas.' };
  }

  var context = contexts.join('\n\n');
  var messages = [
    {
      role: 'system',
      content: 'Eres un asistente experto en análisis de datos de Google Sheets. Responde de forma breve y clara en español.'
    },
    { role: 'user', content: context + '\n\nPregunta: ' + question }
  ];

  var payload = {
    model: 'gpt-4o-mini',
    messages: messages,
    max_tokens: 256,
    temperature: 0.4,
    n: 1
  };
  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + apiKey
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    var response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', options);
    var result = JSON.parse(response.getContentText());
    if (result.error) {
      return { error: result.error.message || 'Error al procesar la solicitud.' };
    }
    if (result.choices && result.choices.length > 0) {
      return { answer: result.choices[0].message.content };
    }
    return { error: 'No se obtuvo respuesta del modelo.' };
  } catch (err) {
    return { error: 'Error al conectar con OpenAI: ' + err.message };
  }
}

function buildTableContextSnippet(metaDataRow, headers, samples, sheetName, rangeA1) {
  var name = metaDataRow[META_INDEX.name] || '';
  var tableLabel = name ? String(name).trim() : '';
  if (!tableLabel) {
    tableLabel = normalizeMetaId(metaDataRow[META_INDEX.id] || '');
  }
  var displayName = tableLabel || 'Tabla sin nombre';
  var location = sheetName + '!' + rangeA1;
  var headerLine = 'Tabla: ' + displayName + ' (' + location + ')';
  var headersLine = 'Encabezados: ' + headers.join(', ');
  var sampleText = samples.length > 0 ? 'Muestra:\n' + samples.join('\n') : 'Sin filas de datos.';
  return headerLine + '\n' + headersLine + '\n' + sampleText;
}

function normalizeTableSelection(selection) {
  if (selection === null || selection === undefined) {
    return [];
  }

  if (selection === ALL_TABLES_OPTION_VALUE) {
    return listSavedTables()
      .map(function(entry) {
        return normalizeMetaId(entry && entry.id);
      })
      .filter(function(id) {
        return !!id;
      });
  }

  if (Array.isArray(selection)) {
    var ids = [];
    for (var i = 0; i < selection.length; i++) {
      var normalized = normalizeMetaId(selection[i]);
      if (normalized && normalized !== ALL_TABLES_OPTION_VALUE) {
        ids.push(normalized);
      }
    }
    return ids;
  }

  var value = normalizeMetaId(selection);
  if (!value) {
    return [];
  }
  if (value === ALL_TABLES_OPTION_VALUE) {
    return normalizeTableSelection(ALL_TABLES_OPTION_VALUE);
  }
  return [value];
}

/**
 * Devuelve metadatos individuales de una tabla por ID para uso en la UI.
 *
 * @param {string} tableId ID de la tabla.
 * @returns {Object|null} Objeto con metadatos o null si no existe.
 */
function getTableMeta(tableId) {
  SpreadsheetApp.flush();
  var meta = findMetaById(tableId);
  if (!meta) return null;
  var d = meta.data;
  var style = parseJsonValue(d[META_INDEX.style], null);
  var headers = normalizeHeaderArray(parseJsonValue(d[META_INDEX.headers], null));
  return {
    id: normalizeMetaId(d[META_INDEX.id]),
    name: d[META_INDEX.name],
    rangeA1: d[META_INDEX.rangeA1],
    description: d[META_INDEX.description],
    sheetName: d[META_INDEX.sheet],
    cols: d[META_INDEX.cols],
    rows: d[META_INDEX.rows],
    createdAt: d[META_INDEX.createdAt],
    updatedAt: d[META_INDEX.updatedAt],
    style: style,
    headers: headers
  };
}

/**
 * Devuelve los encabezados (primera fila) de una tabla guardada.  Se
 * utiliza en la interfaz para rellenar la lista de encabezados al editar
 * una tabla existente.
 *
 * @param {string} tableId ID de la tabla.
 * @returns {Object} Objeto con headers (array de strings) o error.
 */
function getTableHeaders(tableId) {
  var meta = findMetaById(tableId);
  if (!meta) return { error: 'Tabla no encontrada' };
  var d = meta.data;
  var storedHeaders = normalizeHeaderArray(parseJsonValue(d[META_INDEX.headers], null));
  if (storedHeaders && storedHeaders.length) {
    return { headers: storedHeaders };
  }
  var sheetName = d[META_INDEX.sheet];
  var rangeA1 = d[META_INDEX.rangeA1];
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'Hoja no encontrada' };
  var range = sheet.getRange(rangeA1);
  var values = range.getValues();
  if (!values || values.length === 0) return { error: 'Rango vacío' };
  var headers = values[0].map(function(v) {
    return normalizeHeaderEntry(v === null ? '' : String(v));
  });
  return { headers: headers };
}
