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
    hasDescription: false
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
    normalized.label = String(labelValue === undefined || labelValue === null ? '' : labelValue).trim();
    normalized.description = hasDescription ? String(descriptionValue || '').trim() : '';
    normalized.hasDescription = hasDescription;
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
  headerMeta = headerMeta.map(function(entry) {
    var labelValue = entry && entry.label !== undefined && entry.label !== null ? String(entry.label) : '';
    var descriptionValue = entry && entry.description !== undefined && entry.description !== null ? String(entry.description) : '';
    var trimmedLabel = labelValue.trim();
    var trimmedDescription = descriptionValue.trim();
    var explicitHasDescription = entry && Object.prototype.hasOwnProperty.call(entry, 'hasDescription') ? !!entry.hasDescription : false;
    var hasDescription = explicitHasDescription || trimmedDescription !== '';
    return {
      label: trimmedLabel,
      description: hasDescription ? trimmedDescription : '',
      hasDescription: hasDescription
    };
  });
  var headerMetaForReturn = headerMeta.map(function(item) {
    return {
      label: item.label || '',
      description: item.hasDescription ? (item.description || '') : '',
      hasDescription: !!item.hasDescription
    };
  });
  var headerValues = headerMetaForReturn.map(function(item) {
    return item.label || '';
  });
  var previousTableData = null;
  var previousTableRows = 0;
  var previousTableCols = 0;
  if (doFormat) {
    if (id) {
      var entry = findMetaById(id);
      if (entry) {
        var prevRangeA1 = entry.data[META_INDEX.rangeA1];
        var prevSheetName = entry.data[META_INDEX.sheet];
        if (prevRangeA1) {
          var ss = SpreadsheetApp.getActive();
          var prevSheet = ss.getSheetByName(prevSheetName);
          if (prevSheet) {
            try {
              var prevRange = prevSheet.getRange(prevRangeA1);
              var sameRange = prevSheetName === sheet.getName() && prevRange.getA1Notation() === normalizedRange;
              if (!sameRange) {
                try {
                  previousTableData = {
                    values: prevRange.getValues(),
                    formulas: prevRange.getFormulas(),
                    horizontalAlignments: prevRange.getHorizontalAlignments(),
                    verticalAlignments: prevRange.getVerticalAlignments(),
                    fontWeights: prevRange.getFontWeights()
                  };
                  previousTableRows = prevRange.getNumRows();
                  previousTableCols = prevRange.getNumColumns();
                } catch (readErr) {
                  previousTableData = null;
                }
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
                clearRangeAndFormatting(prevRange);
              }
              deleteFilterViewsByTitle(prevSheet, 'TableCrafter_' + id);
            } catch (prevErr) {
              // Si el rango anterior no existe, continuar sin detener la ejecución.
            }
          }
        }
      }
    }
    if (previousTableData) {
      range.clearContent();
    }
    applyFormattingToRange(range, appliedStyle);
    if (shouldOverwriteHeaders) {
      headerRange.setValues([headerValues]);
    } else if (previousTableData && previousTableData.values && previousTableData.values.length > 0) {
      var prevHeaderValues = previousTableData.values[0] || [];
      var headerRowValues = [];
      for (var hc = 0; hc < cols; hc++) {
        headerRowValues.push(hc < prevHeaderValues.length ? prevHeaderValues[hc] : '');
      }
      headerRange.setValues([headerRowValues]);
      if (previousTableData.formulas && previousTableData.formulas.length > 0) {
        var headerFormulas = previousTableData.formulas[0] || [];
        for (var hf = 0; hf < cols; hf++) {
          var headerFormula = headerFormulas[hf];
          if (headerFormula && headerFormula !== '') {
            headerRange.getCell(1, hf + 1).setFormula(headerFormula);
          }
        }
      }
    }
    if (previousTableData && previousTableData.values && previousTableData.values.length > 0) {
      var prevValues = previousTableData.values;
      var totalTargetRows = rows - 1;
      if (totalTargetRows > 0) {
        var dataRange = range.offset(1, 0, totalTargetRows, cols);
        var valuesMatrix = [];
        for (var dr = 0; dr < totalTargetRows; dr++) {
          var sourceRow = (dr + 1 < prevValues.length) ? prevValues[dr + 1] : null;
          var newRow = [];
          for (var dc = 0; dc < cols; dc++) {
            if (sourceRow && dc < sourceRow.length && sourceRow[dc] !== undefined && sourceRow[dc] !== null) {
              newRow.push(sourceRow[dc]);
            } else {
              newRow.push('');
            }
          }
          valuesMatrix.push(newRow);
        }
        dataRange.setValues(valuesMatrix);
        if (previousTableData.formulas && previousTableData.formulas.length > 1) {
          var prevFormulas = previousTableData.formulas;
          var maxFormulaRows = Math.min(totalTargetRows, prevFormulas.length - 1);
          for (var fr = 0; fr < maxFormulaRows; fr++) {
            var formulaRow = prevFormulas[fr + 1] || [];
            for (var fc = 0; fc < cols; fc++) {
              var formula = formulaRow[fc];
              if (formula && formula !== '') {
                dataRange.getCell(fr + 1, fc + 1).setFormula(formula);
              }
            }
          }
        }
        if (previousTableData.horizontalAlignments && previousTableData.horizontalAlignments.length > 1) {
          var availableHorizontalRows = Math.min(totalTargetRows, previousTableData.horizontalAlignments.length - 1);
          if (availableHorizontalRows > 0) {
            var horizontalMatrix = [];
            for (var har = 0; har < availableHorizontalRows; har++) {
              var horizontalSource = previousTableData.horizontalAlignments[har + 1] || [];
              var horizontalRow = [];
              for (var hac = 0; hac < cols; hac++) {
                var horizontalValue = (horizontalSource && hac < horizontalSource.length) ? horizontalSource[hac] : null;
                horizontalRow.push(horizontalValue ? horizontalValue : null);
              }
              horizontalMatrix.push(horizontalRow);
            }
            dataRange.offset(0, 0, availableHorizontalRows, cols).setHorizontalAlignments(horizontalMatrix);
          }
        }
        if (previousTableData.verticalAlignments && previousTableData.verticalAlignments.length > 1) {
          var availableVerticalRows = Math.min(totalTargetRows, previousTableData.verticalAlignments.length - 1);
          if (availableVerticalRows > 0) {
            var verticalMatrix = [];
            for (var varr = 0; varr < availableVerticalRows; varr++) {
              var verticalSource = previousTableData.verticalAlignments[varr + 1] || [];
              var verticalRow = [];
              for (var vac = 0; vac < cols; vac++) {
                var verticalValue = (verticalSource && vac < verticalSource.length) ? verticalSource[vac] : null;
                verticalRow.push(verticalValue ? verticalValue : null);
              }
              verticalMatrix.push(verticalRow);
            }
            dataRange.offset(0, 0, availableVerticalRows, cols).setVerticalAlignments(verticalMatrix);
          }
        }
        if (previousTableData.fontWeights && previousTableData.fontWeights.length > 1) {
          var availableWeightRows = Math.min(totalTargetRows, previousTableData.fontWeights.length - 1);
          if (availableWeightRows > 0) {
            var weightMatrix = [];
            for (var wr = 0; wr < availableWeightRows; wr++) {
              var weightSource = previousTableData.fontWeights[wr + 1] || [];
              var weightRow = [];
              for (var wc = 0; wc < cols; wc++) {
                var weightValue = (weightSource && wc < weightSource.length) ? weightSource[wc] : '';
                weightRow.push(weightValue === 'bold' ? 'bold' : 'normal');
              }
              weightMatrix.push(weightRow);
            }
            dataRange.offset(0, 0, availableWeightRows, cols).setFontWeights(weightMatrix);
          }
        }
      }
    }
    if (id) {
      deleteFilterViewsByTitle(sheet, 'TableCrafter_' + id);
      createFilterViewForRange(range, 'TableCrafter_' + id);
    }
  } else if (shouldOverwriteHeaders) {
    headerRange.setValues([headerValues]);
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
 * @param {string} tableId ID de la tabla.
 * @param {string} question Pregunta del usuario.
 * @returns {Object} Objeto con answer o error.
 */
function askQuestion(tableId, question) {
  var apiKey = PropertiesService.getUserProperties().getProperty('TC_API_KEY');
  if (!apiKey) {
    return { error: 'No hay API Key configurada. Configure su clave en la sección de configuración.' };
  }
  var meta = findMetaById(tableId);
  if (!meta) {
    return { error: 'Tabla no encontrada.' };
  }
  var rangeA1 = meta.data[META_INDEX.rangeA1];
  var sheetName = meta.data[META_INDEX.sheet];
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    return { error: 'No se encontró la hoja con la tabla.' };
  }
  var range = sheet.getRange(rangeA1);
  var values = range.getDisplayValues();
  if (values.length === 0) {
    return { error: 'El rango está vacío.' };
  }
  var headers = values[0];
  // Preparar muestra de hasta 500 filas (excluyendo encabezado)
  var maxRows = Math.min(500, values.length - 1);
  var sampleRows = [];
  for (var i = 1; i <= maxRows; i++) {
    sampleRows.push(values[i].join(', '));
  }
  // Construir contexto
  var context = 'Encabezados: ' + headers.join(', ') + '\n';
  context += 'Muestra:\n' + sampleRows.join('\n');
  var messages = [
    { role: 'system', content: 'Eres un asistente experto en análisis de datos de Google Sheets. Responde de forma breve y clara en español.' },
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
    var json = JSON.parse(response.getContentText());
    if (json.error) {
      return { error: json.error.message || 'Error en la respuesta de la IA.' };
    }
    var answer = (json.choices && json.choices[0] && json.choices[0].message && json.choices[0].message.content) || '';
    return { answer: answer.trim() };
  } catch (err) {
    return { error: 'Error al consultar la IA: ' + err.message };
  }
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
