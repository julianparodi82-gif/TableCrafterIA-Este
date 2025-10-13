
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
  } catch (err) {
    console.error('Error en onChange: ' + err.message);
  }
}

function handleInsertedColumnsChange(e) {
  if (!e) {
    return;
  }
  var sheet = e.sheet || (e.range ? e.range.getSheet() : null);
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
  if (endColumn < startColumn) {
    var temp = startColumn;
    startColumn = endColumn;
    endColumn = temp;
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

    var currentStart = parsed.startColumn;
    var currentEnd = parsed.endColumn;
    var headerInsertions = [];
    for (var colIndex = startColumn; colIndex <= endColumn; colIndex++) {
      if (colIndex <= currentStart) {
        currentStart += 1;
        currentEnd += 1;
      } else if (colIndex <= currentEnd + 1) {
        headerInsertions.push(colIndex - currentStart);
        currentEnd += 1;
      }
    }

    if (currentStart === parsed.startColumn && currentEnd === parsed.endColumn) {
      continue;
    }

    var newCols = currentEnd - currentStart + 1;
    if (newCols <= 0) {
      continue;
    }
    var headers = normalizeHeaderArray(parseJsonValue(row[META_INDEX.headers], []));
    if (headers.length < parsed.cols) {
      while (headers.length < parsed.cols) {
        headers.push(normalizeHeaderEntry(''));
      }
    }
    headers = ensureHeaderKeys(headers);
    if (headerInsertions.length > 0) {
      headerInsertions.sort(function(a, b) {
        return a - b;
      });
      for (var h = 0; h < headerInsertions.length; h++) {
        var insertionIndex = headerInsertions[h] + h;
        if (insertionIndex < 0) {
          insertionIndex = 0;
        }
        if (insertionIndex > headers.length) {
          insertionIndex = headers.length;
        }
        headers.splice(insertionIndex, 0, normalizeHeaderEntry(''));
      }
    }
    if (headers.length > newCols) {
      headers = headers.slice(0, newCols);
    } else if (headers.length < newCols) {
      while (headers.length < newCols) {
        headers.push(normalizeHeaderEntry(''));
      }
    }
    headers = ensureHeaderKeys(headers);

    var newRange = buildA1Notation(currentStart, parsed.startRow, newCols, parsed.rows);
    var updatedRow = row.slice();
    updatedRow[META_INDEX.rangeA1] = newRange;
    updatedRow[META_INDEX.cols] = newCols;
    updatedRow[META_INDEX.headers] = stringifyJsonValue(headers);
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
  var sheet = e.sheet || (e.range ? e.range.getSheet() : null);
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
  if (endRow < startRow) {
    var temp = startRow;
    startRow = endRow;
    endRow = temp;
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

    var currentStart = parsed.startRow;
    var currentEnd = parsed.endRow;
    for (var rowIndex = startRow; rowIndex <= endRow; rowIndex++) {
      if (rowIndex <= currentStart) {
        currentStart += 1;
        currentEnd += 1;
      } else if (rowIndex <= currentEnd + 1) {
        currentEnd += 1;
      }
    }

    if (currentStart === parsed.startRow && currentEnd === parsed.endRow) {
      continue;
    }

    var newRows = currentEnd - currentStart + 1;
    if (newRows <= 0) {
      continue;
    }

    var newRange = buildA1Notation(parsed.startColumn, currentStart, parsed.cols, newRows);
    var updatedRow = row.slice();
    updatedRow[META_INDEX.rangeA1] = newRange;
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
  var sheet = e.sheet || (e.range ? e.range.getSheet() : null);
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
  if (columnEnd < columnStart) {
    var temp = columnStart;
    columnStart = columnEnd;
    columnEnd = temp;
  }
  var removedCount = columnEnd - columnStart + 1;
  if (removedCount <= 0) {
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
    var tableStartCol = parsed.startColumn;
    var tableEndCol = parsed.endColumn;
    var tableCols = parsed.cols;
    var tableStartRow = parsed.startRow;
    var tableRows = parsed.rows;
    if (columnStart > tableEndCol) {
      continue;
    }
    if (columnEnd < tableStartCol) {
      var shiftedStartCol = tableStartCol - removedCount;
      if (shiftedStartCol < 1) {
        shiftedStartCol = 1;
      }
      var shiftedRange = buildA1Notation(shiftedStartCol, tableStartRow, tableCols, tableRows);
      var updatedRow = row.slice();
      updatedRow[META_INDEX.rangeA1] = shiftedRange;
      updatedRow[META_INDEX.cols] = tableCols;
      updatedRow[META_INDEX.updatedAt] = now;
      updates.push({ rowNumber: i + 1, values: updatedRow, id: entryId, newRange: shiftedRange });
      continue;
    }

    var beforeStart = Math.max(0, Math.min(columnEnd, tableStartCol - 1) - columnStart + 1);
    var overlap = Math.max(0, Math.min(columnEnd, tableEndCol) - Math.max(columnStart, tableStartCol) + 1);
    var newStartCol = tableStartCol - beforeStart;
    if (newStartCol < 1) {
      newStartCol = 1;
    }
    var newCols = tableCols - overlap;
    if (newCols <= 0) {
      deletions.push(entryId);
      continue;
    }
    var headers = normalizeHeaderArray(parseJsonValue(row[META_INDEX.headers], []));
    if (overlap > 0) {
      var removeIndex = Math.max(columnStart, tableStartCol) - tableStartCol;
      if (removeIndex < 0) {
        removeIndex = 0;
      }
      var availableToRemove = Math.max(0, headers.length - removeIndex);
      var removeCount = Math.min(overlap, availableToRemove);
      if (removeCount > 0) {
        headers.splice(removeIndex, removeCount);
      }
    }
    while (headers.length > newCols) {
      headers.pop();
    }
    while (headers.length < newCols) {
      headers.push(normalizeHeaderEntry(''));
    }
    headers = ensureHeaderKeys(headers);
    var newRange = buildA1Notation(newStartCol, tableStartRow, newCols, tableRows);
    var updatedRowWithOverlap = row.slice();
    updatedRowWithOverlap[META_INDEX.rangeA1] = newRange;
    updatedRowWithOverlap[META_INDEX.cols] = newCols;
    updatedRowWithOverlap[META_INDEX.headers] = stringifyJsonValue(headers);
    updatedRowWithOverlap[META_INDEX.updatedAt] = now;
    updates.push({ rowNumber: i + 1, values: updatedRowWithOverlap, id: entryId, newRange: newRange });
  }

  applyStructuralUpdates(updates, sheet);
  handleTableDeletions(deletions);
}


function handleRemovedRowsChange(e) {
  if (!e) {
    return;
  }
  var sheet = e.sheet || (e.range ? e.range.getSheet() : null);
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
  if (rowEnd < rowStart) {
    var tmp = rowStart;
    rowStart = rowEnd;
    rowEnd = tmp;
  }
  var removedCount = rowEnd - rowStart + 1;
  if (removedCount <= 0) {
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

    var tableStartRow = parsed.startRow;
    var tableEndRow = parsed.endRow;
    var tableRows = parsed.rows;
    var tableStartCol = parsed.startColumn;
    var tableCols = parsed.cols;

    if (rowStart > tableEndRow) {
      continue;
    }

    if (rowEnd < tableStartRow) {
      var shiftedStartRow = tableStartRow - removedCount;
      if (shiftedStartRow < 1) {
        shiftedStartRow = 1;
      }
      var shiftedRange = buildA1Notation(tableStartCol, shiftedStartRow, tableCols, tableRows);
      var updatedRow = row.slice();
      updatedRow[META_INDEX.rangeA1] = shiftedRange;
      updatedRow[META_INDEX.rows] = tableRows;
      updatedRow[META_INDEX.updatedAt] = now;
      updates.push({ rowNumber: i + 1, values: updatedRow, id: entryId, newRange: shiftedRange });
      continue;
    }

    var beforeStart = Math.max(0, Math.min(rowEnd, tableStartRow - 1) - rowStart + 1);
    var overlap = Math.max(0, Math.min(rowEnd, tableEndRow) - Math.max(rowStart, tableStartRow) + 1);
    var headerRemoved = rowStart <= tableStartRow && rowEnd >= tableStartRow;

    var newStartRow = tableStartRow - beforeStart;
    if (newStartRow < 1) {
      newStartRow = 1;
    }

    var effectiveOverlap = overlap;
    if (headerRemoved) {
      effectiveOverlap = Math.max(0, overlap - 1);
    }

    var newRows = tableRows - effectiveOverlap;
    if (headerRemoved) {
      newRows = Math.max(1, newRows);
      if (!restoredHeaders[entryId]) {
        restoreTableHeaderRow(sheet, tableStartCol, newStartRow, tableCols, row);
        restoredHeaders[entryId] = true;
      }
    }

    if (!headerRemoved && newRows <= 0) {
      deletions.push(entryId);
      continue;
    }

    if (headerRemoved && newRows <= 0) {
      newRows = 1;
    }

    var newRange = buildA1Notation(tableStartCol, newStartRow, tableCols, newRows);
    var updatedRowWithOverlap = row.slice();
    updatedRowWithOverlap[META_INDEX.rangeA1] = newRange;
    updatedRowWithOverlap[META_INDEX.rows] = newRows;
    updatedRowWithOverlap[META_INDEX.updatedAt] = now;
    updates.push({ rowNumber: i + 1, values: updatedRowWithOverlap, id: entryId, newRange: newRange });
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
  updates.forEach(function(update) {
    metaSheet.getRange(update.rowNumber, 1, 1, META_HEADERS.length).setValues([update.values]);
    try {
      var targetRange = sheet.getRange(update.newRange);
      deleteFilterViewsByTitle(sheet, 'TableCrafter_' + update.id);
      createFilterViewForRange(targetRange, 'TableCrafter_' + update.id);
    } catch (err) {
      // Ignorar errores al reconstruir la vista de filtro.
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
