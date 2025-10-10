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
      headers.push(firstRow[i] === null ? '' : String(firstRow[i]));
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

/**
 * Devuelve el listado de tablas guardadas desde la hoja oculta
 * __TableCrafter_Meta.  Cada entrada contiene id, nombre y rango A1.
 *
 * @returns {Array} Lista de objetos con los metadatos de las tablas.
 */
function listSavedTables() {
  var meta = getMetaSheet();
  var data = meta.getDataRange().getValues();
  var result = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    if (row[0]) {
      result.push({
        id: row[0],
        name: row[1],
        rangeA1: row[2],
        description: row[3],
        sheetName: row[4],
        cols: row[5],
        rows: row[6],
        createdAt: row[7],
        updatedAt: row[8]
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
  var rangeA1 = data[2];
  var sheetName = data[4];
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
 * @returns {Object} Resultado con id y mensaje, o error.
 */
function applyTableFormatting(tableId, rangeA1, name, description, headers, style) {
  // Obtener rango activo
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getActiveSheet();
  var range = sheet.getRange(rangeA1);
  var rows = range.getNumRows();
  var cols = range.getNumColumns();
  // Validar límites
  if (cols > 16 || rows > 1000) {
    SpreadsheetApp.getUi().alert('El rango seleccionado (' + cols + ' columnas y ' + rows + ' filas) supera los límites permitidos (máx. 16 columnas y 1000 filas).');
    return { error: 'exceeded' };
  }
  var id = tableId && tableId.trim() !== '' ? tableId : Utilities.getUuid();
  // Si existe, limpiar rango anterior
  if (tableId && tableId.trim() !== '') {
    var entry = findMetaById(tableId);
    if (entry) {
      var prevRangeA1 = entry.data[2];
      var prevSheetName = entry.data[4];
      if (prevRangeA1) {
        var prevSheet = ss.getSheetByName(prevSheetName);
        if (prevSheet) {
          var prevRange = prevSheet.getRange(prevRangeA1);
          clearRangeAndFormatting(prevRange);
          // Eliminar vista de filtro anterior
          var prevTitle = 'TableCrafter_' + tableId;
          deleteFilterViewsByTitle(prevSheet, prevTitle);
        }
      }
    }
  }
  // Aplicar encabezados si han sido modificados
  if (headers && headers.length > 0) {
    var headerRange = range.offset(0, 0, 1, cols);
    headerRange.setValues([headers]);
  }
  // Aplicar formato
  applyFormattingToRange(range, style);
  // Crear vista de filtro dedicada
  var filterTitle = 'TableCrafter_' + id;
  deleteFilterViewsByTitle(sheet, filterTitle); // eliminar si existe antes de crear
  createFilterViewForRange(range, filterTitle);
  // Guardar/actualizar metadatos
  saveOrUpdateMeta(id, name, description, rangeA1, sheet.getName(), cols, rows);
  return { id: id, message: 'Tabla aplicada' };
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
  // Aplicar banding con colores personalizados
  var banded = range.applyRowBanding();
  banded.setHeaderColor(style.headerColor);
  banded.setFirstBandColor(style.altColor1);
  banded.setSecondBandColor(style.altColor2);
  // Encabezado en negrita o normal
  var headerRange = range.offset(0, 0, 1, cols);
  headerRange.setFontWeight(style.bold ? 'bold' : 'normal');
  // Bordes finos negros
  if (style.border) {
    range.setBorder(true, true, true, true, true, true, '#000000', SpreadsheetApp.BorderStyle.SOLID);
  } else {
    range.setBorder(false, false, false, false, false, false);
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
    // Encabezados
    var headers = ['id', 'name', 'rangeA1', 'description', 'sheet', 'cols', 'rows', 'createdAt', 'updatedAt'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }
  return sheet;
}

/**
 * Busca una entrada de metadatos por ID.
 *
 * @param {string} id ID de la tabla.
 * @returns {Object|null} Objeto con {row, data} o null si no existe.
 */
function findMetaById(id) {
  var sheet = getMetaSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === id) {
      return { row: i + 1, data: data[i] };
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
function saveOrUpdateMeta(id, name, description, rangeA1, sheetName, cols, rows) {
  var sheet = getMetaSheet();
  var meta = findMetaById(id);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  if (meta) {
    // Actualizar
    var row = meta.row;
    sheet.getRange(row, 1, 1, 9).setValues([[id, name, rangeA1, description, sheetName, cols, rows, meta.data[7] || now, now]]);
  } else {
    // Crear nueva
    sheet.appendRow([id, name, rangeA1, description, sheetName, cols, rows, now, now]);
  }
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
  var rangeA1 = meta.data[2];
  var sheetName = meta.data[4];
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
  var meta = findMetaById(tableId);
  if (!meta) return null;
  var d = meta.data;
  return {
    id: d[0],
    name: d[1],
    rangeA1: d[2],
    description: d[3],
    sheetName: d[4],
    cols: d[5],
    rows: d[6],
    createdAt: d[7],
    updatedAt: d[8]
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
  var sheetName = d[4];
  var rangeA1 = d[2];
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'Hoja no encontrada' };
  var range = sheet.getRange(rangeA1);
  var values = range.getValues();
  if (!values || values.length === 0) return { error: 'Rango vacío' };
  var headers = values[0].map(function(v) { return v === null ? '' : String(v); });
  return { headers: headers };
}