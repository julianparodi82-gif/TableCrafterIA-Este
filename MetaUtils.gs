/**
 * TableCrafter AI – Utilidades de monitoreo estructural.
 *
 * README DE INSTALACIÓN RÁPIDA
 * 1. Abre Extensiones → Apps Script dentro de la hoja.
 * 2. En Herramientas → Triggers, crea dos disparadores instalables:
 *    • onChange → evento "Al cambiar" (Change)
 *    • onSelectionChange → evento "Al cambiar de selección"
 * 3. Ambos triggers deben ejecutarse como el usuario actual con el alcance
 *    predeterminado del proyecto.
 */

var META_SHEET_NAME = '__TableCrafter_Meta';
var STRUCTURE_SELECTION_CACHE_KEY = 'TC_LAST_SELECTION';

function recordSelectionContext_(range) {
  if (!range) {
    return;
  }
  try {
    var sheet = range.getSheet();
    if (!sheet) {
      return;
    }
    var payload = {
      sheetId: sheet.getSheetId(),
      sheetName: sheet.getName(),
      row: range.getRow(),
      column: range.getColumn(),
      timestamp: Date.now()
    };
    var serialized = JSON.stringify(payload);
    try {
      var cache = CacheService.getDocumentCache();
      if (cache) {
        cache.put(STRUCTURE_SELECTION_CACHE_KEY, serialized, 21600);
      }
    } catch (cacheErr) {
      console.error('Error al guardar contexto en caché:', cacheErr);
    }
    try {
      PropertiesService.getDocumentProperties().setProperty(
        STRUCTURE_SELECTION_CACHE_KEY,
        serialized
      );
    } catch (propErr) {
      console.error('Error al persistir contexto de selección:', propErr);
    }
  } catch (err) {
    console.error('recordSelectionContext_ falló:', err);
  }
}

function readSelectionContext_() {
  var raw = null;
  try {
    var cache = CacheService.getDocumentCache();
    if (cache) {
      raw = cache.get(STRUCTURE_SELECTION_CACHE_KEY);
    }
  } catch (cacheErr) {
    console.error('No se pudo leer la caché de selección:', cacheErr);
  }
  if (!raw) {
    try {
      raw = PropertiesService.getDocumentProperties().getProperty(
        STRUCTURE_SELECTION_CACHE_KEY
      );
    } catch (propErr) {
      console.error('No se pudo leer las propiedades de selección:', propErr);
    }
  }
  if (!raw) {
    return null;
  }
  try {
    return JSON.parse(raw);
  } catch (err) {
    console.error('Contexto de selección inválido:', err);
    return null;
  }
}

function inferChange_(e) {
  if (!e || !e.changeType) {
    return null;
  }
  var typeMap = {
    INSERT_COLUMN: 'INSERT_COLUMN',
    REMOVE_COLUMN: 'REMOVE_COLUMN',
    INSERT_ROW: 'INSERT_ROW',
    REMOVE_ROW: 'REMOVE_ROW'
  };
  var normalizedType = typeMap[e.changeType];
  if (!normalizedType) {
    return null;
  }
  var ss = e.source || SpreadsheetApp.getActive();
  var sheet = null;
  if (e.range && typeof e.range.getSheet === 'function') {
    sheet = e.range.getSheet();
  }
  if (!sheet && e.sheetId && ss && typeof ss.getSheetById === 'function') {
    sheet = ss.getSheetById(e.sheetId);
  }
  if (!sheet && ss && typeof ss.getActiveSheet === 'function') {
    sheet = ss.getActiveSheet();
  }
  if (!sheet && ss) {
    var cached = readSelectionContext_();
    if (cached && cached.sheetId && typeof ss.getSheetById === 'function') {
      sheet = ss.getSheetById(cached.sheetId);
    }
    if (!sheet && cached && cached.sheetName && typeof ss.getSheetByName === 'function') {
      sheet = ss.getSheetByName(cached.sheetName);
    }
  }
  if (!sheet) {
    return null;
  }
  var index = null;
  var count = 1;
  var startRow = coerceNumber_(e && e.startRow, e && e.rowStart);
  var endRow = coerceNumber_(e && e.endRow, e && e.rowEnd);
  var startColumn = coerceNumber_(e && e.startColumn, e && e.columnStart);
  var endColumn = coerceNumber_(e && e.endColumn, e && e.columnEnd);
  if (normalizedType === 'INSERT_ROW' || normalizedType === 'REMOVE_ROW') {
    if (startRow !== null) {
      index = startRow;
      if (endRow !== null) {
        count = Math.max(1, endRow - startRow + 1);
      } else if (e && e.numberOfRows) {
        count = Math.max(1, parseInt(e.numberOfRows, 10));
      }
    } else if (e.range) {
      index = e.range.getRow();
      count = Math.max(1, e.range.getNumRows());
    }
  } else {
    if (startColumn !== null) {
      index = startColumn;
      if (endColumn !== null) {
        count = Math.max(1, endColumn - startColumn + 1);
      } else if (e && e.numberOfColumns) {
        count = Math.max(1, parseInt(e.numberOfColumns, 10));
      }
    } else if (e.range) {
      index = e.range.getColumn();
      count = Math.max(1, e.range.getNumColumns());
    }
  }
  if (index === null) {
    index = 1;
  }
  return {
    type: normalizedType,
    sheet: sheet,
    sheetName: sheet.getName(),
    sheetId: sheet.getSheetId(),
    index: Math.max(1, parseInt(index, 10)),
    count: Math.max(1, parseInt(count, 10)),
    maxRows: sheet.getMaxRows(),
    maxColumns: sheet.getMaxColumns()
  };
}

function loadTablesForSheet_(sheetName) {
  var normalizedSheet = sheetName === null || sheetName === undefined
    ? ''
    : String(sheetName).trim();
  if (!normalizedSheet) {
    return [];
  }
  var metaSheet = getMetaSheet();
  if (!metaSheet) {
    return [];
  }
  var lastRow = metaSheet.getLastRow();
  if (lastRow < 2) {
    return [];
  }
  var width = META_HEADERS.length;
  var values = metaSheet.getRange(2, 1, lastRow - 1, width).getValues();
  var tables = [];
  for (var i = 0; i < values.length; i++) {
    var row = values[i];
    var id = normalizeMetaId(row[META_INDEX.id]);
    if (!id) {
      continue;
    }
    var storedSheet = String(row[META_INDEX.sheet] || '').trim();
    var storedRangeRaw = String(row[META_INDEX.rangeA1] || '').trim();
    var rangeParts = splitRangeNotation(storedRangeRaw);
    if (!storedSheet && rangeParts.sheet) {
      storedSheet = rangeParts.sheet;
    }
    if (storedSheet !== normalizedSheet) {
      continue;
    }
    var pureRange = rangeParts.range;
    if (!pureRange) {
      continue;
    }
    var headerModels = normalizeHeaderArray(
      parseJsonValue(row[META_INDEX.headers], [])
    );
    tables.push({
      id: id,
      rowIndex: i + 2,
      name: row[META_INDEX.name],
      description: row[META_INDEX.description],
      rangeA1: pureRange,
      fullRangeA1: storedRangeRaw || buildFullRangeNotation(storedSheet, pureRange),
      sheetName: storedSheet,
      cols: Number(row[META_INDEX.cols]) || 0,
      rows: Number(row[META_INDEX.rows]) || 0,
      headersValue: row[META_INDEX.headers],
      headerModels: headerModels,
      styleValue: row[META_INDEX.style],
      formulaRefs: row[META_INDEX.formulaRefs]
    });
  }
  return tables;
}

function updateMetaRanges_(updates) {
  if (!updates || updates.length === 0) {
    return;
  }
  var metaSheet = getMetaSheet();
  if (!metaSheet) {
    return;
  }
  var lastRow = metaSheet.getLastRow();
  if (lastRow < 2) {
    return;
  }
  var width = META_HEADERS.length;
  var dataRange = metaSheet.getRange(2, 1, lastRow - 1, width);
  var values = dataRange.getValues();
  var timestamp = new Date().toISOString();
  var updateMap = {};
  var removalMap = {};
  for (var i = 0; i < updates.length; i++) {
    var item = updates[i];
    if (item.remove) {
      removalMap[item.rowIndex] = true;
    } else {
      updateMap[item.rowIndex] = item;
    }
  }
  for (var r = 0; r < values.length; r++) {
    var rowIndex = r + 2;
    if (removalMap[rowIndex]) {
      values[r] = null;
      continue;
    }
    var update = updateMap[rowIndex];
    if (update) {
      var row = values[r];
      row[META_INDEX.rangeA1] = update.fullRangeA1 || update.rangeA1 || '';
      row[META_INDEX.cols] = update.cols;
      row[META_INDEX.rows] = update.rows;
      row[META_INDEX.updatedAt] = timestamp;
      values[r] = row;
    }
  }
  var rowsToWrite = [];
  for (var j = 0; j < values.length; j++) {
    var rowValues = values[j];
    if (!rowValues) {
      continue;
    }
    var keep = false;
    for (var c = 0; c < rowValues.length; c++) {
      if (rowValues[c] !== '' && rowValues[c] !== null) {
        keep = true;
        break;
      }
    }
    if (keep) {
      rowsToWrite.push(rowValues);
    }
  }
  dataRange.clearContent();
  if (rowsToWrite.length > 0) {
    metaSheet.getRange(2, 1, rowsToWrite.length, width).setValues(rowsToWrite);
  }
}

function a1ToBounds_(rangeA1) {
  var parsed = parseRangeA1(rangeA1);
  if (!parsed) {
    return null;
  }
  return {
    r1: parsed.startRow,
    c1: parsed.startColumn,
    r2: parsed.endRow,
    c2: parsed.endColumn
  };
}

function boundsToA1_(r1, c1, r2, c2) {
  var rows = r2 - r1 + 1;
  var cols = c2 - c1 + 1;
  return buildA1Notation(c1, r1, cols, rows);
}

function adjustBoundsForChange_(bounds, change) {
  if (!bounds || !change) {
    return bounds;
  }
  var r1 = bounds.r1;
  var r2 = bounds.r2;
  var c1 = bounds.c1;
  var c2 = bounds.c2;
  var height = r2 - r1 + 1;
  var width = c2 - c1 + 1;
  var index = change.index;
  var count = change.count;
  if (change.type === 'INSERT_COLUMN') {
    if (index <= c1) {
      c1 += count;
      c2 += count;
    } else if (index > c2) {
      // Sin cambios.
    } else {
      c2 += count;
    }
  } else if (change.type === 'REMOVE_COLUMN') {
    var removalStartCol = index;
    var removalEndCol = index + count - 1;
    if (removalEndCol < c1) {
      c1 = Math.max(1, c1 - count);
      c2 = c1 + width - 1;
    } else if (removalStartCol > c2) {
      // Sin cambios.
    } else {
      var removedInsideCols = Math.min(c2, removalEndCol) - Math.max(c1, removalStartCol) + 1;
      if (removedInsideCols >= width) {
        return null;
      }
      var newWidth = width - removedInsideCols;
      if (removalStartCol <= c1) {
        c1 = removalStartCol;
      }
      c2 = c1 + newWidth - 1;
    }
  } else if (change.type === 'INSERT_ROW') {
    if (index <= r1) {
      r1 += count;
      r2 += count;
    } else if (index > r2) {
      // Sin cambios.
    } else {
      r2 += count;
    }
  } else if (change.type === 'REMOVE_ROW') {
    var removalStartRow = index;
    var removalEndRow = index + count - 1;
    if (removalEndRow < r1) {
      r1 = Math.max(1, r1 - count);
      r2 = r1 + height - 1;
    } else if (removalStartRow > r2) {
      // Sin cambios.
    } else {
      var removedInsideRows = Math.min(r2, removalEndRow) - Math.max(r1, removalStartRow) + 1;
      if (removedInsideRows >= height) {
        return null;
      }
      var newHeight = height - removedInsideRows;
      if (removalStartRow <= r1) {
        r1 = removalStartRow;
      }
      r2 = r1 + newHeight - 1;
    }
  }
  return { r1: r1, r2: r2, c1: c1, c2: c2 };
}

function ensureBoundsWithinSheet_(bounds, maxRows, maxColumns) {
  var constrained = {
    r1: Math.max(1, Math.min(bounds.r1, maxRows)),
    c1: Math.max(1, Math.min(bounds.c1, maxColumns))
  };
  constrained.r2 = Math.max(
    constrained.r1,
    Math.min(bounds.r2, maxRows)
  );
  constrained.c2 = Math.max(
    constrained.c1,
    Math.min(bounds.c2, maxColumns)
  );
  return constrained;
}

function applyHeaderProtection_(sheet, table, enabled) {
  if (!sheet || !table || !table.id) {
    return;
  }
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE) || [];
  var description = 'TableCrafter Header - ' + table.id;
  var existing = null;
  for (var i = 0; i < protections.length; i++) {
    if (protections[i].getDescription() === description) {
      existing = protections[i];
      break;
    }
  }
  if (!enabled) {
    if (existing) {
      try {
        existing.remove();
      } catch (err) {
        console.error('No se pudo eliminar la protección:', err);
      }
    }
    return;
  }
  var bounds = table.bounds || a1ToBounds_(table.rangeA1);
  if (!bounds) {
    return;
  }
  var width = bounds.c2 - bounds.c1 + 1;
  if (width <= 0) {
    return;
  }
  try {
    var range = sheet.getRange(bounds.r1, bounds.c1, 1, width);
    if (existing) {
      existing.setRange(range);
      return;
    }
    var protection = range.protect();
    protection.setDescription(description);
    try {
      var editors = protection.getEditors();
      if (editors && editors.length > 0) {
        protection.removeEditors(editors);
      }
    } catch (editorErr) {
      console.error('No se pudieron ajustar los editores de la protección:', editorErr);
    }
    if (protection.canDomainEdit()) {
      protection.setDomainEdit(false);
    }
  } catch (err2) {
    console.error('Error al aplicar protección de encabezado:', err2);
  }
}

function restoreHeadersIfNeeded_(sheet, table, changeContext) {
  if (!sheet || !table || !table.headerModels || table.headerModels.length === 0) {
    return;
  }
  if (changeContext && changeContext.type === 'REMOVE_COLUMN' && wasWholeColumnDeletion_(changeContext.event)) {
    return;
  }
  var bounds = table.bounds || a1ToBounds_(table.rangeA1);
  if (!bounds) {
    return;
  }
  var width = bounds.c2 - bounds.c1 + 1;
  if (width <= 0) {
    return;
  }
  var headerRange = sheet.getRange(bounds.r1, bounds.c1, 1, width);
  var currentValues = headerRange.getValues();
  var firstRow = currentValues && currentValues.length > 0 ? currentValues[0] : [];
  var needsRestore = false;
  for (var i = 0; i < width; i++) {
    if (!firstRow[i] || String(firstRow[i]).trim() === '') {
      needsRestore = true;
      break;
    }
  }
  if (!needsRestore) {
    return;
  }
  var labels = buildHeaderLabelRow_(table.headerModels, width);
  headerRange.setValues([labels]);
  applyHeaderProtection_(sheet, table, true);
}

function wasWholeColumnDeletion_(e) {
  return !!(e && e.changeType === 'REMOVE_COLUMN');
}

function wasWholeRowDeletion_(e) {
  return !!(e && e.changeType === 'REMOVE_ROW');
}

function buildHeaderLabelRow_(headerModels, width) {
  var labels = [];
  for (var i = 0; i < width; i++) {
    var entry = normalizeHeaderEntry(headerModels[i]);
    labels.push(entry.label || '');
  }
  return labels;
}

function coerceNumber_() {
  for (var i = 0; i < arguments.length; i++) {
    var value = arguments[i];
    if (value === null || value === undefined) {
      continue;
    }
    var num = Number(value);
    if (!isNaN(num) && isFinite(num)) {
      return Math.floor(num);
    }
  }
  return null;
}

function handleStructuralChange_(change, event) {
  if (!change || !change.sheetName) {
    return null;
  }
  var sheet = change.sheet || SpreadsheetApp.getActive().getSheetByName(change.sheetName);
  if (!sheet) {
    return {
      sheet: change.sheetName,
      type: change.type,
      tablesEvaluated: 0,
      tablesUpdated: [],
      tables: []
    };
  }
  var tables = loadTablesForSheet_(sheet.getName());
  if (!tables || tables.length === 0) {
    return {
      sheet: sheet.getName(),
      type: change.type,
      tablesEvaluated: 0,
      tablesUpdated: [],
      tables: []
    };
  }
  var updates = [];
  var summary = [];
  var removalEntries = [];
  var maxRows = change.maxRows || sheet.getMaxRows();
  var maxColumns = change.maxColumns || sheet.getMaxColumns();
  var removalStartRow = change.index;
  var removalEndRow = change.index + change.count - 1;
  for (var i = 0; i < tables.length; i++) {
    var table = tables[i];
    var bounds = a1ToBounds_(table.rangeA1);
    if (!bounds) {
      continue;
    }
    var originalBounds = { r1: bounds.r1, r2: bounds.r2, c1: bounds.c1, c2: bounds.c2 };
    var oldPureRange = boundsToA1_(originalBounds.r1, originalBounds.c1, originalBounds.r2, originalBounds.c2);
    var oldFullRange = table.fullRangeA1 || buildFullRangeNotation(table.sheetName, oldPureRange);
    var adjusted = adjustBoundsForChange_(bounds, change);
    if (adjusted === null) {
      updates.push({ id: table.id, rowIndex: table.rowIndex, remove: true });
      removalEntries.push({ table: table, bounds: originalBounds });
      summary.push({ id: table.id, oldRange: oldFullRange, newRange: null });
      continue;
    }
    adjusted = ensureBoundsWithinSheet_(adjusted, maxRows, maxColumns);
    var newPureRange = boundsToA1_(adjusted.r1, adjusted.c1, adjusted.r2, adjusted.c2);
    var newFullRange = buildFullRangeNotation(sheet.getName(), newPureRange);
    var newRows = adjusted.r2 - adjusted.r1 + 1;
    var newCols = adjusted.c2 - adjusted.c1 + 1;
    var changedRange = newFullRange !== oldFullRange || newRows !== table.rows || newCols !== table.cols;
    if (changedRange) {
      updates.push({
        id: table.id,
        rowIndex: table.rowIndex,
        rangeA1: newPureRange,
        fullRangeA1: newFullRange,
        rows: newRows,
        cols: newCols,
        sheetName: sheet.getName()
      });
    }
    table.bounds = adjusted;
    table.rangeA1 = newPureRange;
    table.fullRangeA1 = newFullRange;
    table.rows = newRows;
    table.cols = newCols;
    if (change.type === 'REMOVE_ROW') {
      if (removalStartRow <= originalBounds.r1 && removalEndRow >= originalBounds.r1) {
        table.headerRestoreNeeded = true;
      }
    }
    summary.push({
      id: table.id,
      oldRange: oldFullRange,
      newRange: changedRange ? newFullRange : oldFullRange
    });
  }
  if (updates.length > 0) {
    updateMetaRanges_(updates);
    for (var u = 0; u < updates.length; u++) {
      var update = updates[u];
      if (update.remove) {
        removeNamedRangeForTable(update.id);
      } else {
        syncNamedRangeForTable(update.id, change.sheetName, update.rangeA1);
      }
    }
  }
  SpreadsheetApp.flush();
  for (var r = 0; r < removalEntries.length; r++) {
    var removal = removalEntries[r];
    applyHeaderProtection_(sheet, {
      id: removal.table.id,
      rangeA1: boundsToA1_(removal.bounds.r1, removal.bounds.c1, removal.bounds.r2, removal.bounds.c2)
    }, false);
  }
  for (var t = 0; t < tables.length; t++) {
    var tableEntry = tables[t];
    if (!tableEntry.bounds) {
      continue;
    }
    applyHeaderProtection_(sheet, tableEntry, true);
    if (tableEntry.headerRestoreNeeded) {
      restoreHeadersIfNeeded_(sheet, tableEntry, { type: change.type, event: event });
    }
  }
  SpreadsheetApp.flush();
  var updated = [];
  for (var s = 0; s < summary.length; s++) {
    if (summary[s].oldRange !== summary[s].newRange) {
      updated.push(summary[s]);
    }
  }
  return {
    sheet: sheet.getName(),
    type: change.type,
    tablesEvaluated: tables.length,
    tablesUpdated: updated,
    tables: summary
  };
}
