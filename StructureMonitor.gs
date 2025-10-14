
/**
 * Gestiona los cambios estructurales de la hoja (inserciones y eliminaciones
 * de filas/columnas) para mantener sincronizados los metadatos de TableCrafter
 * y proteger los encabezados de las tablas guardadas.
 */
function onChange(e) {
  if (!e || !e.changeType) {
    return;
  }
  try {
    switch (e.changeType) {
      case 'INSERT_COLUMN':
        handleInsertedColumnsChange(e);
        break;
      case 'INSERT_ROW':
        handleInsertedRowsChange(e);
        break;
      case 'REMOVE_COLUMN':
        handleRemovedColumnsChange(e);
        break;
      case 'REMOVE_ROW':
        handleRemovedRowsChange(e);
        break;
      case 'REMOVE_RANGE':
        handleRemovedRangeChange(e);
        break;
      default:
        break;
    }
    reconcileTablesAfterEvent(e);
  } catch (err) {
    console.error('Error en onChange: ' + err.message);
  }
}

function onEdit(e) {
  try {
    protectHeadersOnEdit(e);
  } catch (err) {
    console.error('Error en onEdit: ' + err.message);
  }
  try {
    reconcileTablesAfterEvent(e);
  } catch (err2) {
    console.error('Error al reconciliar tras onEdit: ' + err2.message);
  }
}

function reconcileTablesAfterEvent(e) {
  var structural = isStructuralChangeType(e && e.changeType);
  if (structural) {
    reconcileAllSheets();
    return;
  }
  var sheets = determineSheetsForEvent(e);
  if (!sheets || sheets.length === 0) {
    reconcileAllSheets();
    return;
  }
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (!sheet) {
      continue;
    }
    var sheetName = sheet.getName();
    if (sheetName === '__TableCrafter_Meta') {
      continue;
    }
    var entries = getMetaEntriesForSheet(sheetName);
    if (!entries || entries.length === 0) {
      continue;
    }
    var bounds = extractEventBounds(e, sheet);
    var targetedEntries = bounds && bounds.length > 0 ? filterEntriesByBounds(entries, bounds) : [];
    if (!targetedEntries || targetedEntries.length === 0) {
      continue;
    }
    reconcileSheetTables(sheet, targetedEntries);
  }
}

function determineSheetsForEvent(e) {
  var sheet = getEventSheet(e);
  if (sheet) {
    return [sheet];
  }
  return getSheetsWithTables();
}

function isStructuralChangeType(changeType) {
  if (!changeType) {
    return false;
  }
  switch (changeType) {
    case 'INSERT_COLUMN':
    case 'REMOVE_COLUMN':
    case 'INSERT_ROW':
    case 'REMOVE_ROW':
    case 'REMOVE_RANGE':
    case 'INSERT_GRID':
    case 'REMOVE_GRID':
    case 'FORMAT':
    case 'OTHER':
      return true;
    default:
      return false;
  }
}

function extractEventBounds(e, sheet) {
  var bounds = [];
  if (!e) {
    return bounds;
  }
  var maxRows = sheet ? sheet.getMaxRows() : 50000;
  var maxColumns = sheet ? sheet.getMaxColumns() : 18278;

  if (e.range && typeof e.range.getRow === 'function') {
    try {
      var range = e.range;
      bounds.push({
        startRow: range.getRow(),
        endRow: range.getRow() + range.getNumRows() - 1,
        startColumn: range.getColumn(),
        endColumn: range.getColumn() + range.getNumColumns() - 1
      });
    } catch (err) {
      // Ignorar errores al leer el rango del evento.
    }
  }

  var startRow = coerceNumber(e.startRow, e.rowStart);
  var endRow = coerceNumber(e.endRow, e.rowEnd);
  var startColumn = coerceNumber(e.startColumn, e.columnStart);
  var endColumn = coerceNumber(e.endColumn, e.columnEnd);

  if (startRow !== null) {
    bounds.push({
      startRow: startRow,
      endRow: endRow !== null ? endRow : startRow,
      startColumn: 1,
      endColumn: maxColumns
    });
  }

  if (startColumn !== null) {
    bounds.push({
      startRow: 1,
      endRow: maxRows,
      startColumn: startColumn,
      endColumn: endColumn !== null ? endColumn : startColumn
    });
  }

  if (e.oldRange && typeof e.oldRange.getRow === 'function') {
    try {
      var oldRange = e.oldRange;
      bounds.push({
        startRow: oldRange.getRow(),
        endRow: oldRange.getRow() + oldRange.getNumRows() - 1,
        startColumn: oldRange.getColumn(),
        endColumn: oldRange.getColumn() + oldRange.getNumColumns() - 1
      });
    } catch (err2) {
      // Ignorar si no se puede leer oldRange.
    }
  }

  return bounds;
}

function coerceNumber() {
  for (var i = 0; i < arguments.length; i++) {
    var candidate = arguments[i];
    if (typeof candidate === 'number' && !isNaN(candidate)) {
      return candidate;
    }
  }
  return null;
}

function cloneHeadersForMutation(rawHeaders, targetWidth) {
  var normalized = normalizeHeaderArray(rawHeaders);
  var cloned = [];
  for (var i = 0; i < normalized.length; i++) {
    var item = normalizeHeaderEntry(normalized[i]);
    cloned.push({
      label: item.label || '',
      description: item.description || '',
      hasDescription: !!item.hasDescription,
      key: normalizeHeaderKey(item.key)
    });
  }
  if (typeof targetWidth === 'number' && targetWidth >= 0) {
    while (cloned.length < targetWidth) {
      cloned.push(normalizeHeaderEntry(''));
    }
    if (cloned.length > targetWidth) {
      cloned = cloned.slice(0, targetWidth);
    }
  }
  return cloned;
}

function finalizeHeaderArray(headers, targetWidth) {
  var adjusted = Array.isArray(headers) ? headers.slice() : [];
  if (typeof targetWidth === 'number' && targetWidth >= 0) {
    while (adjusted.length < targetWidth) {
      adjusted.push(normalizeHeaderEntry(''));
    }
    if (adjusted.length > targetWidth) {
      adjusted = adjusted.slice(0, targetWidth);
    }
  }
  return ensureHeaderKeys(adjusted);
}

function adjustColumnsForInsertion(parsed, headers, insertStart, insertEnd) {
  if (!parsed) {
    return { changed: false, startColumn: null, cols: 0, headers: headers };
  }
  var currentStart = parsed.startColumn;
  var currentWidth = parsed.cols;
  var workingHeaders = Array.isArray(headers) ? headers.slice() : [];
  var changed = false;
  if (insertStart > insertEnd) {
    var tmp = insertStart;
    insertStart = insertEnd;
    insertEnd = tmp;
  }
  for (var col = insertStart; col <= insertEnd; col++) {
    if (col <= currentStart) {
      currentStart += 1;
      changed = true;
    } else if (col <= currentStart + currentWidth) {
      var relativeIndex = col - currentStart;
      if (relativeIndex < 0) {
        relativeIndex = 0;
      }
      if (relativeIndex > workingHeaders.length) {
        relativeIndex = workingHeaders.length;
      }
      workingHeaders.splice(relativeIndex, 0, normalizeHeaderEntry(''));
      currentWidth += 1;
      changed = true;
    }
  }
  return {
    changed: changed,
    startColumn: currentStart,
    cols: currentWidth,
    headers: workingHeaders
  };
}

function adjustColumnsForRemoval(parsed, headers, removeStart, removeEnd) {
  if (!parsed) {
    return { changed: false, startColumn: null, cols: 0, headers: headers };
  }
  var currentStart = parsed.startColumn;
  var currentWidth = parsed.cols;
  var workingHeaders = Array.isArray(headers) ? headers.slice() : [];
  var changed = false;
  if (removeStart > removeEnd) {
    var tmp = removeStart;
    removeStart = removeEnd;
    removeEnd = tmp;
  }
  for (var col = removeStart; col <= removeEnd; col++) {
    if (col < currentStart) {
      currentStart = Math.max(1, currentStart - 1);
      changed = true;
    } else if (col >= currentStart && col < currentStart + currentWidth) {
      var relativeIndex = col - currentStart;
      if (relativeIndex < 0) {
        relativeIndex = 0;
      }
      if (relativeIndex >= workingHeaders.length) {
        relativeIndex = workingHeaders.length - 1;
      }
      if (relativeIndex >= 0 && workingHeaders.length > 0) {
        workingHeaders.splice(relativeIndex, 1);
      }
      if (currentWidth > 0) {
        currentWidth -= 1;
      }
      changed = true;
    }
  }
  if (currentWidth < 0) {
    currentWidth = 0;
  }
  return {
    changed: changed,
    startColumn: currentStart,
    cols: currentWidth,
    headers: workingHeaders
  };
}

function adjustRowsForInsertion(parsed, insertStart, insertEnd) {
  if (!parsed) {
    return { changed: false, startRow: null, rows: 0 };
  }
  var currentStart = parsed.startRow;
  var currentHeight = parsed.rows;
  var changed = false;
  if (insertStart > insertEnd) {
    var tmp = insertStart;
    insertStart = insertEnd;
    insertEnd = tmp;
  }
  for (var row = insertStart; row <= insertEnd; row++) {
    if (row <= currentStart) {
      currentStart += 1;
      changed = true;
    } else if (row <= currentStart + currentHeight) {
      currentHeight += 1;
      changed = true;
    }
  }
  return {
    changed: changed,
    startRow: currentStart,
    rows: currentHeight
  };
}

function adjustRowsForRemoval(parsed, removeStart, removeEnd) {
  if (!parsed) {
    return { changed: false, startRow: null, rows: 0, headerRemoved: false };
  }
  var currentStart = parsed.startRow;
  var currentHeight = parsed.rows;
  var headerRemoved = false;
  var changed = false;
  if (removeStart > removeEnd) {
    var tmp = removeStart;
    removeStart = removeEnd;
    removeEnd = tmp;
  }
  for (var row = removeStart; row <= removeEnd; row++) {
    if (row < currentStart) {
      currentStart = Math.max(1, currentStart - 1);
      changed = true;
    } else if (row >= currentStart && row < currentStart + currentHeight) {
      if (row === currentStart) {
        headerRemoved = true;
      }
      if (currentHeight > 0) {
        currentHeight -= 1;
      }
      changed = true;
    }
  }
  if (currentHeight < 0) {
    currentHeight = 0;
  }
  return {
    changed: changed,
    startRow: currentStart,
    rows: currentHeight,
    headerRemoved: headerRemoved
  };
}

function filterEntriesByBounds(entries, bounds) {
  if (!entries || entries.length === 0 || !bounds || bounds.length === 0) {
    return [];
  }
  var results = [];
  var seen = {};
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    if (!entry || !entry.data) {
      continue;
    }
    var parsed = parseRangeA1(entry.data[META_INDEX.rangeA1]);
    if (!parsed) {
      continue;
    }
    for (var b = 0; b < bounds.length; b++) {
      if (boundsIntersectEntry(bounds[b], parsed)) {
        if (!seen[entry.id]) {
          results.push(entry);
          seen[entry.id] = true;
        }
        break;
      }
    }
  }
  return results;
}

function boundsIntersectEntry(bounds, parsed) {
  if (!bounds || !parsed) {
    return false;
  }
  var entryStartRow = parsed.startRow;
  var entryEndRow = parsed.startRow + parsed.rows - 1;
  var entryStartColumn = parsed.startColumn;
  var entryEndColumn = parsed.startColumn + parsed.cols - 1;
  var intersectsRows = !(bounds.endRow < entryStartRow || bounds.startRow > entryEndRow);
  var intersectsColumns = !(bounds.endColumn < entryStartColumn || bounds.startColumn > entryEndColumn);
  return intersectsRows && intersectsColumns;
}

function getSheetsWithTables() {
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return [];
  }
  var ss = SpreadsheetApp.getActive();
  var seen = {};
  var sheets = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    var sheetName = row[META_INDEX.sheet];
    if (!sheetName || sheetName === '__TableCrafter_Meta' || seen[sheetName]) {
      continue;
    }
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      continue;
    }
    sheets.push(sheet);
    seen[sheetName] = true;
  }
  return sheets;
}

function getEventSheet(e) {
  if (!e) {
    return null;
  }
  if (e.sheet) {
    return e.sheet;
  }
  if (typeof e.sheetId === 'number' && e.source) {
    try {
      if (typeof e.source.getSheetById === 'function') {
        var sheetById = e.source.getSheetById(e.sheetId);
        if (sheetById) {
          return sheetById;
        }
      }
      if (typeof e.source.getSheets === 'function') {
        var sheets = e.source.getSheets();
        if (sheets && sheets.length) {
          for (var i = 0; i < sheets.length; i++) {
            var candidate = sheets[i];
            try {
              if (candidate && typeof candidate.getSheetId === 'function' && candidate.getSheetId() === e.sheetId) {
                return candidate;
              }
            } catch (err2) {
              // Ignorar y continuar buscando.
            }
          }
        }
      }
    } catch (err) {
      // Ignorar y continuar con los demás intentos.
    }
  }
  if (e.range && typeof e.range.getSheet === 'function') {
    try {
      return e.range.getSheet();
    } catch (err) {
      return null;
    }
  }
  if (e.source && typeof e.source.getActiveSheet === 'function') {
    try {
      return e.source.getActiveSheet();
    } catch (err2) {
      return null;
    }
  }
  return null;
}

function handleInsertedColumnsChange(e) {
  if (!e) {
    return;
  }
  var sheet = getEventSheet(e);
  if (!sheet) {
    return;
  }
  var sheetName = sheet.getName();
  if (sheetName === '__TableCrafter_Meta') {
    return;
  }
  var startColumn = typeof e.startColumn === 'number' ? e.startColumn : null;
  if (startColumn === null) {
    startColumn = typeof e.columnStart === 'number' ? e.columnStart : null;
  }
  var endColumn = typeof e.endColumn === 'number' ? e.endColumn : null;
  if (endColumn === null) {
    endColumn = typeof e.columnEnd === 'number' ? e.columnEnd : startColumn;
  }
  if (startColumn === null || endColumn === null) {
    return;
  }
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return;
  }
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updates = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    if (row[META_INDEX.sheet] !== sheetName) {
      continue;
    }
    var storedRangeA1 = row[META_INDEX.rangeA1];
    if (!storedRangeA1) {
      continue;
    }
    var parsed = parseRangeA1(storedRangeA1);
    if (!parsed) {
      continue;
    }

    var headers = cloneHeadersForMutation(parseJsonValue(row[META_INDEX.headers], []), parsed.cols);
    var result = adjustColumnsForInsertion(parsed, headers, startColumn, endColumn);
    if (!result.changed) {
      continue;
    }

    var newCols = result.cols;
    if (newCols <= 0) {
      continue;
    }
    var finalHeaders = finalizeHeaderArray(result.headers, newCols);
    var newRange = buildA1Notation(result.startColumn, parsed.startRow, newCols, parsed.rows);
    var updatedRow = row.slice();
    updatedRow[META_INDEX.rangeA1] = buildFullRangeNotation(sheetName, newRange);
    updatedRow[META_INDEX.cols] = newCols;
    updatedRow[META_INDEX.headers] = stringifyJsonValue(finalHeaders);
    updatedRow[META_INDEX.updatedAt] = now;
    updates.push({
      rowNumber: i + 1,
      values: updatedRow,
      id: entryId,
      newRange: newRange
    });
  }

  applyStructuralUpdates(updates, sheet);
}

function handleInsertedRowsChange(e) {
  if (!e) {
    return;
  }
  var sheet = getEventSheet(e);
  if (!sheet) {
    return;
  }
  var sheetName = sheet.getName();
  if (sheetName === '__TableCrafter_Meta') {
    return;
  }
  var startRow = typeof e.startRow === 'number' ? e.startRow : null;
  if (startRow === null) {
    startRow = typeof e.rowStart === 'number' ? e.rowStart : null;
  }
  var endRow = typeof e.endRow === 'number' ? e.endRow : null;
  if (endRow === null) {
    endRow = typeof e.rowEnd === 'number' ? e.rowEnd : startRow;
  }
  if (startRow === null || endRow === null) {
    return;
  }
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return;
  }
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updates = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    if (row[META_INDEX.sheet] !== sheetName) {
      continue;
    }
    var storedRangeA1 = row[META_INDEX.rangeA1];
    if (!storedRangeA1) {
      continue;
    }
    var parsed = parseRangeA1(storedRangeA1);
    if (!parsed) {
      continue;
    }

    var result = adjustRowsForInsertion(parsed, startRow, endRow);
    if (!result.changed) {
      continue;
    }

    var newRows = result.rows;
    if (newRows <= 0) {
      continue;
    }

    var newRange = buildA1Notation(parsed.startColumn, result.startRow, parsed.cols, newRows);
    var updatedRow = row.slice();
    updatedRow[META_INDEX.rangeA1] = buildFullRangeNotation(sheetName, newRange);
    updatedRow[META_INDEX.rows] = newRows;
    updatedRow[META_INDEX.updatedAt] = now;
    updates.push({
      rowNumber: i + 1,
      values: updatedRow,
      id: entryId,
      newRange: newRange
    });
  }

  applyStructuralUpdates(updates, sheet);
}

function handleRemovedColumnsChange(e) {
  if (!e) {
    return;
  }
  var sheet = getEventSheet(e);
  if (!sheet) {
    return;
  }
  var sheetName = sheet.getName();
  if (sheetName === '__TableCrafter_Meta') {
    return;
  }
  var columnStart = typeof e.columnStart === 'number' ? e.columnStart : null;
  if (columnStart === null) {
    columnStart = typeof e.startColumn === 'number' ? e.startColumn : null;
  }
  var columnEnd = typeof e.columnEnd === 'number' ? e.columnEnd : null;
  if (columnEnd === null) {
    columnEnd = typeof e.endColumn === 'number' ? e.endColumn : columnStart;
  }
  if (columnStart === null || columnEnd === null) {
    return;
  }
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return;
  }
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updates = [];
  var deletions = [];

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    if (row[META_INDEX.sheet] !== sheetName) {
      continue;
    }
    var storedRangeA1 = row[META_INDEX.rangeA1];
    if (!storedRangeA1) {
      continue;
    }
    var parsed = parseRangeA1(storedRangeA1);
    if (!parsed) {
      continue;
    }

    var headers = cloneHeadersForMutation(parseJsonValue(row[META_INDEX.headers], []), parsed.cols);
    var result = adjustColumnsForRemoval(parsed, headers, columnStart, columnEnd);
    if (!result.changed) {
      continue;
    }

    if (result.cols <= 0) {
      deletions.push(entryId);
      continue;
    }

    var finalHeaders = finalizeHeaderArray(result.headers, result.cols);
    var newRange = buildA1Notation(result.startColumn, parsed.startRow, result.cols, parsed.rows);
    var updatedRow = row.slice();
    updatedRow[META_INDEX.rangeA1] = buildFullRangeNotation(sheetName, newRange);
    updatedRow[META_INDEX.cols] = result.cols;
    updatedRow[META_INDEX.headers] = stringifyJsonValue(finalHeaders);
    updatedRow[META_INDEX.updatedAt] = now;
    updates.push({ rowNumber: i + 1, values: updatedRow, id: entryId, newRange: newRange });
  }

  applyStructuralUpdates(updates, sheet);
  handleTableDeletions(deletions);
}


function handleRemovedRowsChange(e) {
  if (!e) {
    return;
  }
  var sheet = getEventSheet(e);
  if (!sheet) {
    return;
  }
  var sheetName = sheet.getName();
  if (sheetName === '__TableCrafter_Meta') {
    return;
  }
  var rowStart = typeof e.rowStart === 'number' ? e.rowStart : null;
  if (rowStart === null) {
    rowStart = typeof e.startRow === 'number' ? e.startRow : null;
  }
  var rowEnd = typeof e.rowEnd === 'number' ? e.rowEnd : null;
  if (rowEnd === null) {
    rowEnd = typeof e.endRow === 'number' ? e.endRow : rowStart;
  }
  if (rowStart === null || rowEnd === null) {
    return;
  }
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return;
  }

  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updates = [];
  var deletions = [];
  var restoredHeaders = {};

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    if (row[META_INDEX.sheet] !== sheetName) {
      continue;
    }
    var storedRangeA1 = row[META_INDEX.rangeA1];
    if (!storedRangeA1) {
      continue;
    }
    var parsed = parseRangeA1(storedRangeA1);
    if (!parsed) {
      continue;
    }

    var result = adjustRowsForRemoval(parsed, rowStart, rowEnd);
    if (!result.changed) {
      continue;
    }

    var headerRemoved = result.headerRemoved;
    var currentRows = result.rows;
    if (headerRemoved) {
      currentRows = Math.max(1, currentRows + 1);
      if (!restoredHeaders[entryId]) {
        restoreTableHeaderRow(sheet, parsed.startColumn, result.startRow, parsed.cols, row);
        restoredHeaders[entryId] = true;
      }
    }

    if (!headerRemoved && currentRows <= 0) {
      deletions.push(entryId);
      continue;
    }

    if (currentRows <= 0) {
      currentRows = 1;
    }

    var newRange = buildA1Notation(parsed.startColumn, result.startRow, parsed.cols, currentRows);
    var updatedRow = row.slice();
    updatedRow[META_INDEX.rangeA1] = buildFullRangeNotation(sheetName, newRange);
    updatedRow[META_INDEX.rows] = currentRows;
    updatedRow[META_INDEX.updatedAt] = now;
    updates.push({ rowNumber: i + 1, values: updatedRow, id: entryId, newRange: newRange });
  }

  applyStructuralUpdates(updates, sheet);
  handleTableDeletions(deletions);
}

function handleRemovedRangeChange(e) {
  if (!e || !e.removedRange) {
    return;
  }
  var removedRange = e.removedRange;
  var sheet = removedRange.getSheet();
  if (!sheet) {
    return;
  }
  var sheetName = sheet.getName();
  if (sheetName === '__TableCrafter_Meta') {
    return;
  }
  var removedValues;
  var removedFormulas = null;
  var removedBackgrounds = null;
  var removedNumberFormats = null;
  var removedHorizontal = null;
  var removedVertical = null;
  var removedFontWeights = null;
  try {
    removedValues = removedRange.getValues();
  } catch (err) {
    removedValues = null;
  }
  try {
    removedFormulas = removedRange.getFormulas();
  } catch (err2) {
    removedFormulas = null;
  }
  try {
    removedBackgrounds = removedRange.getBackgrounds();
  } catch (err3) {
    removedBackgrounds = null;
  }
  try {
    removedNumberFormats = removedRange.getNumberFormats();
  } catch (err4) {
    removedNumberFormats = null;
  }
  try {
    removedHorizontal = removedRange.getHorizontalAlignments();
  } catch (err5) {
    removedHorizontal = null;
  }
  try {
    removedVertical = removedRange.getVerticalAlignments();
  } catch (err6) {
    removedVertical = null;
  }
  try {
    removedFontWeights = removedRange.getFontWeights();
  } catch (err7) {
    removedFontWeights = null;
  }
  if (!removedValues) {
    return;
  }
  var removedStartRow = removedRange.getRow();
  var removedStartCol = removedRange.getColumn();
  var removedEndRow = removedStartRow + removedRange.getNumRows() - 1;
  var removedEndCol = removedStartCol + removedRange.getNumColumns() - 1;

  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return;
  }
  var restored = false;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    if (row[META_INDEX.sheet] !== sheetName) {
      continue;
    }
    var rangeA1 = row[META_INDEX.rangeA1];
    if (!rangeA1) {
      continue;
    }
    var parsed = parseRangeA1(rangeA1);
    if (!parsed) {
      continue;
    }
    var intersectStartRow = Math.max(parsed.startRow, removedStartRow);
    var intersectEndRow = Math.min(parsed.endRow, removedEndRow);
    var intersectStartCol = Math.max(parsed.startColumn, removedStartCol);
    var intersectEndCol = Math.min(parsed.endColumn, removedEndCol);
    if (intersectStartRow > intersectEndRow || intersectStartCol > intersectEndCol) {
      continue;
    }
    var intersectRows = intersectEndRow - intersectStartRow + 1;
    var intersectCols = intersectEndCol - intersectStartCol + 1;
    var rowOffset = intersectStartRow - removedStartRow;
    var colOffset = intersectStartCol - removedStartCol;
    var valuesSlice = extractSubMatrix(removedValues, rowOffset, colOffset, intersectRows, intersectCols);
    var formulasSlice = removedFormulas ? extractSubMatrix(removedFormulas, rowOffset, colOffset, intersectRows, intersectCols) : null;
    var backgroundsSlice = removedBackgrounds ? extractSubMatrix(removedBackgrounds, rowOffset, colOffset, intersectRows, intersectCols) : null;
    var numberFormatsSlice = removedNumberFormats ? extractSubMatrix(removedNumberFormats, rowOffset, colOffset, intersectRows, intersectCols) : null;
    var horizontalSlice = removedHorizontal ? extractSubMatrix(removedHorizontal, rowOffset, colOffset, intersectRows, intersectCols) : null;
    var verticalSlice = removedVertical ? extractSubMatrix(removedVertical, rowOffset, colOffset, intersectRows, intersectCols) : null;
    var fontWeightSlice = removedFontWeights ? extractSubMatrix(removedFontWeights, rowOffset, colOffset, intersectRows, intersectCols) : null;

    var combinedMatrix = [];
    for (var r = 0; r < intersectRows; r++) {
      var valueRow = valuesSlice[r] || [];
      var formulaRow = (formulasSlice && formulasSlice[r]) || [];
      var combinedRow = [];
      for (var c = 0; c < intersectCols; c++) {
        var formulaCell = formulaRow[c];
        if (formulaCell) {
          combinedRow.push(formulaCell);
        } else {
          combinedRow.push(c < valueRow.length ? valueRow[c] : '');
        }
      }
      combinedMatrix.push(combinedRow);
    }

    try {
      var targetRange = sheet.getRange(intersectStartRow, intersectStartCol, intersectRows, intersectCols);
      targetRange.setValues(combinedMatrix);
      if (numberFormatsSlice) {
        targetRange.setNumberFormats(numberFormatsSlice);
      }
      if (backgroundsSlice) {
        targetRange.setBackgrounds(backgroundsSlice);
      }
      if (horizontalSlice) {
        targetRange.setHorizontalAlignments(horizontalSlice);
      }
      if (verticalSlice) {
        targetRange.setVerticalAlignments(verticalSlice);
      }
      if (fontWeightSlice) {
        targetRange.setFontWeights(fontWeightSlice);
      }
      restored = true;
    } catch (err3) {
      // Ignorar errores al restaurar porciones del rango.
    }
  }

  if (restored) {
    SpreadsheetApp.flush();
  }
}

function applyStructuralUpdates(updates, sheet) {
  if (!updates || updates.length === 0) {
    return;
  }
  var metaSheet = getMetaSheet();
  var sheetName = null;
  if (sheet && typeof sheet.getName === 'function') {
    try {
      sheetName = sheet.getName();
    } catch (errName) {
      sheetName = null;
    }
  }
  updates.forEach(function(update) {
    metaSheet.getRange(update.rowNumber, 1, 1, META_HEADERS.length).setValues([update.values]);
    var effectiveSheetName = sheetName;
    if (!effectiveSheetName && update.values && update.values.length > META_INDEX.sheet) {
      effectiveSheetName = update.values[META_INDEX.sheet];
    }
    if (sheet && typeof sheet.getRange === 'function' && sheetName) {
      try {
        var targetRange = sheet.getRange(update.newRange);
        deleteFilterViewsByTitle(sheet, 'TableCrafter_' + update.id);
        createFilterViewForRange(targetRange, 'TableCrafter_' + update.id);
        syncNamedRangeForTable(update.id, sheetName, update.newRange);
      } catch (err) {
        syncNamedRangeForTable(update.id, sheetName, update.newRange);
        // Ignorar errores al reconstruir la vista de filtro.
      }
    } else {
      syncNamedRangeForTable(update.id, effectiveSheetName, update.newRange);
    }
  });
  SpreadsheetApp.flush();
}

function handleTableDeletions(deletions) {
  if (!deletions || deletions.length === 0) {
    return;
  }
  deletions.forEach(function(tableId) {
    try {
      deleteSavedTable(tableId);
    } catch (err) {
      // Ignorar fallos al eliminar metadatos inexistentes.
    }
  });
}

function restoreTableHeaderRow(sheet, startColumn, targetStartRow, columnCount, metaRow) {
  if (!sheet || !startColumn || !columnCount) {
    return;
  }
  var startRow = targetStartRow < 1 ? 1 : targetStartRow;
  try {
    sheet.insertRowsBefore(startRow, 1);
  } catch (err) {
    try {
      sheet.insertRows(startRow, 1);
    } catch (err2) {
      return;
    }
  }
  var headerRange = sheet.getRange(startRow, startColumn, 1, columnCount);
  var labels = collectHeaderLabelsFromMeta(metaRow, columnCount);
  while (labels.length < columnCount) {
    labels.push('');
  }
  try {
    headerRange.setValues([labels]);
  } catch (err3) {
    // Ignorar errores al restablecer los valores del encabezado.
  }
  try {
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
  } catch (err4) {
    // Ignorar si no se puede aplicar la alineación.
  }
  var style = parseJsonValue(metaRow[META_INDEX.style], {}) || {};
  var headerColor = style && style.headerColor ? style.headerColor : '#CFE8FC';
  try {
    headerRange.setBackground(headerColor);
  } catch (err5) {
    // Ignorar si no se puede aplicar el color.
  }
  var bold = true;
  if (style && Object.prototype.hasOwnProperty.call(style, 'bold')) {
    bold = !!style.bold;
  }
  try {
    headerRange.setFontWeight(bold ? 'bold' : 'normal');
  } catch (err6) {
    // Ignorar si no se puede aplicar el peso de fuente.
  }
  var applyBorders = true;
  if (style && Object.prototype.hasOwnProperty.call(style, 'border')) {
    applyBorders = !!style.border;
  }
  try {
    headerRange.setBorder(applyBorders, applyBorders, applyBorders, applyBorders, applyBorders, applyBorders);
  } catch (err7) {
    // Ignorar si no se pueden aplicar bordes.
  }
  SpreadsheetApp.flush();
}

function collectHeaderLabelsFromMeta(metaRow, columnCount) {
  var labels = [];
  if (metaRow && metaRow.length > META_INDEX.headers) {
    var storedHeaders = normalizeHeaderArray(parseJsonValue(metaRow[META_INDEX.headers], []));
    for (var i = 0; i < columnCount; i++) {
      var entry = storedHeaders[i] || normalizeHeaderEntry('');
      labels.push(entry.label || '');
    }
  }
  return labels;
}

function reconcileAllSheets() {
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return;
  }
  var ss = SpreadsheetApp.getActive();
  var namedRangeIndex = buildNamedRangeIndex();
  var sheetIdIndex = buildSheetIdIndex(ss);
  var grouped = {};
  var renameBuckets = {};
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }

    var storedRange = row[META_INDEX.rangeA1] || '';
    var rangeParts = splitRangeNotation(storedRange);
    var storedSheetName = row[META_INDEX.sheet];
    var candidateSheetName = storedSheetName || rangeParts.sheet;
    var resolvedSheet = candidateSheetName ? ss.getSheetByName(candidateSheetName) : null;
    var namedInfo = getNamedRangeInfoFromIndex(entryId, namedRangeIndex);
    if (namedInfo) {
      var namedSheet = resolveSheetFromNamedInfo(namedInfo, ss, sheetIdIndex);
      if (!resolvedSheet && namedSheet) {
        resolvedSheet = namedSheet;
      } else if (resolvedSheet && namedSheet && typeof resolvedSheet.getSheetId === 'function' && typeof namedSheet.getSheetId === 'function') {
        if (resolvedSheet.getSheetId() !== namedSheet.getSheetId()) {
          resolvedSheet = namedSheet;
        }
      }
    }

    var effectiveSheetName = resolvedSheet ? resolvedSheet.getName() : candidateSheetName;
    if (!effectiveSheetName || effectiveSheetName === '__TableCrafter_Meta') {
      continue;
    }

    var pureRange = stripSheetFromRange(storedRange);
    var desiredStoredRange = pureRange ? buildFullRangeNotation(effectiveSheetName, pureRange) : storedRange;
    var needsSheetUpdate = resolvedSheet && storedSheetName !== effectiveSheetName;
    var needsRangeUpdate = resolvedSheet && pureRange && desiredStoredRange && desiredStoredRange !== storedRange;
    var updatedRow = row;

    if ((needsSheetUpdate || needsRangeUpdate) && pureRange) {
      updatedRow = row.slice();
      updatedRow[META_INDEX.sheet] = effectiveSheetName;
      updatedRow[META_INDEX.rangeA1] = desiredStoredRange;
      updatedRow[META_INDEX.updatedAt] = now;
      var bucketKey = resolvedSheet && typeof resolvedSheet.getSheetId === 'function' ? String(resolvedSheet.getSheetId()) : '__none__';
      if (!renameBuckets[bucketKey]) {
        renameBuckets[bucketKey] = { sheet: resolvedSheet, updates: [] };
      }
      renameBuckets[bucketKey].updates.push({
        rowNumber: i + 1,
        values: updatedRow,
        id: entryId,
        newRange: pureRange
      });
    }

    if (resolvedSheet) {
      var groupKey = resolvedSheet.getName();
      if (!grouped[groupKey]) {
        grouped[groupKey] = [];
      }
      grouped[groupKey].push({
        rowNumber: i + 1,
        data: updatedRow,
        id: entryId
      });
    }
  }

  for (var bucketKey in renameBuckets) {
    if (!Object.prototype.hasOwnProperty.call(renameBuckets, bucketKey)) {
      continue;
    }
    var bucket = renameBuckets[bucketKey];
    if (bucket && bucket.updates && bucket.updates.length > 0) {
      applyStructuralUpdates(bucket.updates, bucket.sheet);
    }
  }

  for (var key in grouped) {
    if (!Object.prototype.hasOwnProperty.call(grouped, key)) {
      continue;
    }
    var sheet = ss.getSheetByName(key);
    if (!sheet) {
      continue;
    }
    reconcileSheetTables(sheet, grouped[key]);
  }
}

function getMetaEntriesForSheet(sheetName) {
  if (!sheetName) {
    return [];
  }
  var metaSheet = getMetaSheet();
  var data = metaSheet.getDataRange().getValues();
  if (!data || data.length <= 1) {
    return [];
  }
  var entries = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (!row || row.length === 0) {
      continue;
    }
    if (row[META_INDEX.sheet] !== sheetName) {
      continue;
    }
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    entries.push({
      rowNumber: i + 1,
      data: row,
      id: entryId
    });
  }
  return entries;
}

function reconcileSheetTables(sheet, optEntries) {
  if (!sheet) {
    return;
  }
  var sheetName = sheet.getName();
  if (sheetName === '__TableCrafter_Meta') {
    return;
  }
  var entries = optEntries || getMetaEntriesForSheet(sheetName);
  if (!entries || entries.length === 0) {
    return;
  }
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var updates = [];
  var namedRangeIndex = buildNamedRangeIndex();
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    if (!entry || !entry.data) {
      continue;
    }
    var parsed = parseRangeA1(entry.data[META_INDEX.rangeA1]);
    if (!parsed) {
      continue;
    }
    var storedHeaders = normalizeHeaderArray(parseJsonValue(entry.data[META_INDEX.headers], []));
    storedHeaders = ensureHeaderKeys(storedHeaders);
    var columnCount = storedHeaders.length > 0 ? storedHeaders.length : parsed.cols;
    if (columnCount <= 0) {
      columnCount = parsed.cols;
    }
    var storedRowsHint = parsed.rows;
    var namedRangeInfo = getNamedRangeInfoFromIndex(entry.id, namedRangeIndex);
    var headerLocation = null;
    if (namedRangeInfo && namedRangeInfo.sheetId && sheet.getSheetId && namedRangeInfo.sheetId === sheet.getSheetId()) {
      if (namedRangeInfo.row && namedRangeInfo.column) {
        headerLocation = {
          row: namedRangeInfo.row,
          column: namedRangeInfo.column
        };
      }
      if (namedRangeInfo.cols > 0) {
        columnCount = Math.max(columnCount, namedRangeInfo.cols);
      }
      if (namedRangeInfo.rows > 0) {
        storedRowsHint = Math.max(storedRowsHint, namedRangeInfo.rows);
      }
    }
    if (!headerLocation) {
      headerLocation = locateHeaderPosition(sheet, storedHeaders, parsed);
    }
    if (!headerLocation) {
      ensureHeaderRowIntegrity(sheet, parsed.startRow, parsed.startColumn, columnCount, entry.data, storedHeaders);
      continue;
    }
    var finalizedHeaders = ensureHeaderKeys(finalizeHeaderArray(storedHeaders, columnCount));
    var finalizedHeadersJson = stringifyJsonValue(finalizedHeaders);
    var effectiveColumnCount = finalizedHeaders.length > 0 ? finalizedHeaders.length : columnCount;
    ensureHeaderRowIntegrity(sheet, headerLocation.row, headerLocation.column, effectiveColumnCount, entry.data, finalizedHeaders);
    var rowCount = determineTableRowCount(sheet, headerLocation.row, headerLocation.column, effectiveColumnCount, storedRowsHint);
    if (namedRangeInfo && namedRangeInfo.rows > 0) {
      rowCount = Math.max(rowCount, namedRangeInfo.rows);
    }
    if (rowCount < 1) {
      rowCount = 1;
    }
    var newRange = buildA1Notation(headerLocation.column, headerLocation.row, effectiveColumnCount, rowCount);
    var storedRangeNotation = buildFullRangeNotation(sheetName, newRange);
    var storedColsValue = entry.data[META_INDEX.cols];
    var storedRowsValue = entry.data[META_INDEX.rows];
    var needsHeaderUpdate = (entry.data[META_INDEX.headers] || '') !== finalizedHeadersJson;
    if (storedRangeNotation !== entry.data[META_INDEX.rangeA1] || storedColsValue !== effectiveColumnCount || storedRowsValue !== rowCount || needsHeaderUpdate) {
      var updatedRow = entry.data.slice();
      updatedRow[META_INDEX.rangeA1] = storedRangeNotation;
      updatedRow[META_INDEX.cols] = effectiveColumnCount;
      updatedRow[META_INDEX.rows] = rowCount;
      updatedRow[META_INDEX.updatedAt] = now;
      updatedRow[META_INDEX.headers] = finalizedHeadersJson;
      updates.push({
        rowNumber: entry.rowNumber,
        values: updatedRow,
        id: entry.id,
        newRange: newRange
      });
    }
  }
  if (updates.length > 0) {
    applyStructuralUpdates(updates, sheet);
  }
}

function buildNamedRangeIndex() {
  var index = {};
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    return index;
  }
  var namedRanges;
  try {
    namedRanges = ss.getNamedRanges();
  } catch (err) {
    return index;
  }
  if (!namedRanges || namedRanges.length === 0) {
    return index;
  }
  for (var i = 0; i < namedRanges.length; i++) {
    var namedRange = namedRanges[i];
    if (!namedRange) {
      continue;
    }
    try {
      var name = namedRange.getName();
      if (name) {
        index[name] = namedRange;
      }
    } catch (err2) {
      // Ignorar errores al leer el nombre.
    }
  }
  return index;
}

function buildSheetIdIndex(ss) {
  var index = {};
  if (!ss || typeof ss.getSheets !== 'function') {
    return index;
  }
  var sheets = ss.getSheets();
  if (!sheets || sheets.length === 0) {
    return index;
  }
  for (var i = 0; i < sheets.length; i++) {
    var sheet = sheets[i];
    if (!sheet || typeof sheet.getSheetId !== 'function') {
      continue;
    }
    try {
      var sheetId = sheet.getSheetId();
      index[sheetId] = sheet;
    } catch (err) {
      // Ignorar errores al leer el ID de la hoja.
    }
  }
  return index;
}

function resolveSheetFromNamedInfo(namedRangeInfo, ss, sheetIdIndex) {
  if (!namedRangeInfo) {
    return null;
  }
  if (namedRangeInfo.sheetName) {
    var byName = ss.getSheetByName(namedRangeInfo.sheetName);
    if (byName) {
      return byName;
    }
  }
  if (typeof namedRangeInfo.sheetId === 'number') {
    if (sheetIdIndex && sheetIdIndex[namedRangeInfo.sheetId]) {
      return sheetIdIndex[namedRangeInfo.sheetId];
    }
    if (ss && typeof ss.getSheets === 'function') {
      var sheets = ss.getSheets();
      for (var i = 0; i < sheets.length; i++) {
        var candidate = sheets[i];
        try {
          if (candidate && typeof candidate.getSheetId === 'function' && candidate.getSheetId() === namedRangeInfo.sheetId) {
            return candidate;
          }
        } catch (err) {
          // Ignorar y continuar.
        }
      }
    }
  }
  return null;
}

function getNamedRangeInfoFromIndex(tableId, namedRangeIndex) {
  if (!namedRangeIndex) {
    return null;
  }
  var name = buildTableNamedRangeName(tableId);
  if (!name) {
    return null;
  }
  var namedRange = namedRangeIndex[name];
  if (!namedRange) {
    return null;
  }
  var range;
  try {
    range = namedRange.getRange();
  } catch (err) {
    return null;
  }
  if (!range) {
    return null;
  }
  var info = {
    row: null,
    column: null,
    rows: null,
    cols: null,
    sheetId: null,
    sheetName: ''
  };
  try {
    info.row = range.getRow();
    info.column = range.getColumn();
    info.rows = range.getNumRows();
    info.cols = range.getNumColumns();
  } catch (err2) {
    // Ignorar errores al leer dimensiones.
  }
  try {
    var rangeSheet = range.getSheet();
    if (rangeSheet) {
      info.sheetId = typeof rangeSheet.getSheetId === 'function' ? rangeSheet.getSheetId() : null;
      info.sheetName = rangeSheet.getName ? rangeSheet.getName() : '';
    }
  } catch (err3) {
    // Ignorar si no se puede obtener la hoja.
  }
  return info;
}

function locateHeaderPosition(sheet, storedHeaders, parsedRange) {
  if (!sheet || !storedHeaders) {
    return null;
  }
  var columnCount = storedHeaders.length > 0 ? storedHeaders.length : parsedRange.cols;
  if (columnCount <= 0) {
    columnCount = parsedRange.cols;
  }
  if (parsedRange && parsedRange.startRow && parsedRange.startColumn) {
    try {
      var candidate = sheet.getRange(parsedRange.startRow, parsedRange.startColumn, 1, columnCount);
      var candidateValues = candidate.getDisplayValues()[0];
      if (headersMatchStored(candidateValues, storedHeaders)) {
        return {
          row: parsedRange.startRow,
          column: parsedRange.startColumn
        };
      }
    } catch (err) {
      // Ignorar errores al leer el rango almacenado.
    }
  }
  var desiredLabels = buildDesiredHeaderLabels(storedHeaders, columnCount);
  var anchors = [];
  for (var i = 0; i < desiredLabels.length; i++) {
    var label = desiredLabels[i];
    if (label && String(label).trim() !== '') {
      anchors.push({ index: i, label: label });
    }
  }
  for (var a = 0; a < anchors.length; a++) {
    var anchor = anchors[a];
    try {
      var finder = sheet.createTextFinder(anchor.label);
      if (!finder) {
        continue;
      }
      finder = finder.matchEntireCell(true);
      var matches = finder.findAll();
      if (!matches || matches.length === 0) {
        continue;
      }
      for (var m = 0; m < matches.length; m++) {
        var range = matches[m];
        if (!range) {
          continue;
        }
        var startColumn = range.getColumn() - anchor.index;
        if (startColumn < 1) {
          continue;
        }
        var startRow = range.getRow();
        try {
          var rowValues = sheet.getRange(startRow, startColumn, 1, columnCount).getDisplayValues()[0];
          if (headersMatchStored(rowValues, storedHeaders)) {
            return {
              row: startRow,
              column: startColumn
            };
          }
        } catch (err2) {
          // Ignorar candidatos inválidos.
        }
      }
    } catch (err3) {
      // Ignorar errores en la búsqueda por texto.
    }
  }
  var scanned = scanSheetForHeaderSequence(sheet, desiredLabels);
  if (scanned) {
    return scanned;
  }
  return null;
}

function scanSheetForHeaderSequence(sheet, desiredLabels) {
  if (!sheet || !desiredLabels || desiredLabels.length === 0) {
    return null;
  }
  var dataRange;
  try {
    dataRange = sheet.getDataRange();
  } catch (err) {
    return null;
  }
  if (!dataRange) {
    return null;
  }
  var numRows = dataRange.getNumRows();
  var numCols = dataRange.getNumColumns();
  if (numRows === 0 || numCols === 0 || numCols < desiredLabels.length) {
    return null;
  }
  var startRow = dataRange.getRow();
  var startColumn = dataRange.getColumn();
  var values;
  try {
    values = dataRange.getDisplayValues();
  } catch (err2) {
    values = [];
  }
  if (!values || values.length === 0) {
    return null;
  }
  for (var r = 0; r < values.length; r++) {
    var rowValues = values[r];
    if (!rowValues || rowValues.length === 0) {
      continue;
    }
    for (var c = 0; c <= rowValues.length - desiredLabels.length; c++) {
      var matches = true;
      for (var offset = 0; offset < desiredLabels.length; offset++) {
        var desired = desiredLabels[offset];
        var candidate = rowValues[c + offset];
        if (normalizeHeaderLabelValue(candidate) !== normalizeHeaderLabelValue(desired)) {
          matches = false;
          break;
        }
      }
      if (matches) {
        return {
          row: startRow + r,
          column: startColumn + c
        };
      }
    }
  }
  return null;
}

function buildDesiredHeaderLabels(storedHeaders, columnCount) {
  var labels = [];
  for (var i = 0; i < columnCount; i++) {
    var entry = storedHeaders[i] || normalizeHeaderEntry('');
    labels.push(entry.label || '');
  }
  return labels;
}

function headersMatchStored(rowValues, storedHeaders) {
  if (!rowValues || !storedHeaders) {
    return false;
  }
  var length = storedHeaders.length;
  for (var i = 0; i < length; i++) {
    var expected = storedHeaders[i] || normalizeHeaderEntry('');
    var expectedLabel = expected.label || '';
    var cellValue = i < rowValues.length ? rowValues[i] : '';
    if (normalizeHeaderLabelValue(cellValue) !== normalizeHeaderLabelValue(expectedLabel)) {
      return false;
    }
  }
  return true;
}

function ensureHeaderRowIntegrity(sheet, startRow, startColumn, columnCount, metaRow, storedHeaders) {
  if (!sheet || !startRow || !startColumn || !columnCount) {
    return;
  }
  var headerRange;
  try {
    headerRange = sheet.getRange(startRow, startColumn, 1, columnCount);
  } catch (err) {
    return;
  }
  var desiredLabels = buildDesiredHeaderLabels(storedHeaders, columnCount);
  var currentValues = [];
  try {
    currentValues = headerRange.getDisplayValues()[0];
  } catch (err2) {
    currentValues = [];
  }
  var needsValues = false;
  for (var i = 0; i < desiredLabels.length; i++) {
    var cellValue = i < currentValues.length ? currentValues[i] : '';
    if (normalizeHeaderLabelValue(cellValue) !== normalizeHeaderLabelValue(desiredLabels[i])) {
      needsValues = true;
      break;
    }
  }
  if (needsValues) {
    try {
      headerRange.setValues([desiredLabels]);
    } catch (err3) {
      // Ignorar si no se pueden restablecer los valores.
    }
  }
  try {
    headerRange.setHorizontalAlignment('center');
    headerRange.setVerticalAlignment('middle');
  } catch (err4) {
    // Ignorar errores al aplicar alineaciones.
  }
  var style = parseJsonValue(metaRow[META_INDEX.style], {}) || {};
  var headerColor = style && style.headerColor ? style.headerColor : '#CFE8FC';
  try {
    headerRange.setBackground(headerColor);
  } catch (err5) {
    // Ignorar si no se puede aplicar el color.
  }
  var bold = true;
  if (style && Object.prototype.hasOwnProperty.call(style, 'bold')) {
    bold = !!style.bold;
  }
  try {
    headerRange.setFontWeight(bold ? 'bold' : 'normal');
  } catch (err6) {
    // Ignorar si no se puede aplicar el peso.
  }
  var applyBorders = true;
  if (style && Object.prototype.hasOwnProperty.call(style, 'border')) {
    applyBorders = !!style.border;
  }
  try {
    headerRange.setBorder(applyBorders, applyBorders, applyBorders, applyBorders, applyBorders, applyBorders);
  } catch (err7) {
    // Ignorar si no se pueden aplicar los bordes.
  }
}

function determineTableRowCount(sheet, startRow, startColumn, columnCount, storedRows) {
  if (!sheet || !startRow || !startColumn || !columnCount) {
    return storedRows > 0 ? storedRows : 1;
  }

  var hintedRows = storedRows > 0 ? storedRows : 1;
  var padding = Math.max(5, Math.min(100, Math.ceil(hintedRows * 0.25)));
  var desiredHeight = hintedRows + padding;

  var maxRowsAvailable = null;
  if (typeof sheet.getMaxRows === 'function') {
    try {
      maxRowsAvailable = sheet.getMaxRows();
    } catch (errMaxRows) {
      maxRowsAvailable = null;
    }
  }
  if (maxRowsAvailable === null && typeof sheet.getLastRow === 'function') {
    try {
      maxRowsAvailable = sheet.getLastRow();
    } catch (errLastRow) {
      maxRowsAvailable = null;
    }
  }

  var availableHeight = desiredHeight;
  if (maxRowsAvailable !== null && maxRowsAvailable >= startRow) {
    availableHeight = Math.min(desiredHeight, Math.max(1, maxRowsAvailable - startRow + 1));
  }
  if (availableHeight < hintedRows) {
    availableHeight = hintedRows;
  }

  var range;
  try {
    range = sheet.getRange(startRow, startColumn, availableHeight, columnCount);
  } catch (errRange) {
    return hintedRows;
  }

  var values;
  try {
    values = range.getDisplayValues();
  } catch (errValues) {
    values = [];
  }

  var rowCount = hintedRows;
  for (var r = values.length - 1; r >= 0; r--) {
    var rowValues = values[r];
    var hasContent = false;
    for (var c = 0; c < rowValues.length; c++) {
      if (String(rowValues[c]).trim() !== '') {
        hasContent = true;
        break;
      }
    }
    if (hasContent) {
      rowCount = Math.max(r + 1, 1);
      break;
    }
  }

  if (storedRows > 0) {
    rowCount = Math.max(rowCount, storedRows);
  }

  return rowCount;
}

function protectHeadersOnEdit(e) {
  if (!e || !e.range) {
    return;
  }
  var range = e.range;
  var sheet;
  try {
    sheet = range.getSheet();
  } catch (err) {
    return;
  }
  if (!sheet || sheet.getName() === '__TableCrafter_Meta') {
    return;
  }
  var entries = getMetaEntriesForSheet(sheet.getName());
  if (!entries || entries.length === 0) {
    return;
  }
  var editStartRow = range.getRow();
  var editEndRow = editStartRow + range.getNumRows() - 1;
  var editStartCol = range.getColumn();
  var editEndCol = editStartCol + range.getNumColumns() - 1;
  for (var i = 0; i < entries.length; i++) {
    var entry = entries[i];
    var parsed = parseRangeA1(entry.data[META_INDEX.rangeA1]);
    if (!parsed) {
      continue;
    }
    var headerRow = parsed.startRow;
    if (headerRow < editStartRow || headerRow > editEndRow) {
      continue;
    }
    var headerStartCol = parsed.startColumn;
    var headerEndCol = headerStartCol + parsed.cols - 1;
    var intersectsCols = !(editStartCol > headerEndCol || editEndCol < headerStartCol);
    if (!intersectsCols) {
      continue;
    }
    var headers = normalizeHeaderArray(parseJsonValue(entry.data[META_INDEX.headers], []));
    headers = ensureHeaderKeys(headers);
    ensureHeaderRowIntegrity(sheet, headerRow, headerStartCol, headers.length > 0 ? headers.length : parsed.cols, entry.data, headers);
  }
}
