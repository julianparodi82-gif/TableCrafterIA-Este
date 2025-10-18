/**
 * TableCrafterAI
 *
 * Este archivo contiene la lógica principal del complemento TableCrafterAI para
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
var SIDEBAR_THEME_PROPERTY_KEY = 'tablecrafter.sidebarThemeMode';
var SIDEBAR_THEME_WORD = 'word';
var SIDEBAR_THEME_TABLE = 'table';
var WELCOME_MESSAGE_PROPERTY_KEY = 'tablecrafter.hideWelcomeMessage';
var WELCOME_MESSAGE_PROPERTY_KEY_TABLE = 'tablecrafter.hideWelcomeMessage.table';
var WELCOME_MESSAGE_PROPERTY_KEY_WORD = 'tablecrafter.hideWelcomeMessage.word';
var DEFAULT_AI_FILL_DATA_TYPE = 'datos';

function onOpen() {
  var themeMode = getSidebarThemeMode();
  refreshAddonMenuForTheme(themeMode);
}

/**
 * Muestra la barra lateral principal del complemento.
 */
function showSidebar() {
  var template = HtmlService.createTemplateFromFile('UI');
  var themeMode = getSidebarThemeMode();
  template.initialThemeMode = themeMode;
  var sidebarTitle = themeMode === SIDEBAR_THEME_WORD ? 'WordCrafterAI' : 'TableCrafterAI';
  var html = template.evaluate().setTitle(sidebarTitle).setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
  refreshAddonMenuForTheme(themeMode);
}

/**
 * Muestra el panel de configuración de la API.
 */
function showConfig() {
  var html = HtmlService.createHtmlOutputFromFile('Config')
    .setWidth(400)
    .setHeight(220)
    .setTitle('Configuración de TableCrafterAI');
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración de TableCrafterAI');
}

/**
 * Muestra la página de ayuda.  Incluye enlaces a privacidad y términos.
 */
function showHelp() {
  var html = HtmlService.createHtmlOutputFromFile('Help')
    .setWidth(600)
    .setHeight(500)
    .setTitle('Ayuda de TableCrafterAI');
  SpreadsheetApp.getUi().showModalDialog(html, 'Ayuda de TableCrafterAI');
}

/**
 * Permite incluir archivos HTML parciales en las plantillas.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getSidebarThemeMode() {
  var userProperties = PropertiesService.getUserProperties();
  var stored = userProperties.getProperty(SIDEBAR_THEME_PROPERTY_KEY);
  if (stored === SIDEBAR_THEME_WORD || stored === SIDEBAR_THEME_TABLE) {
    return stored;
  }
  return SIDEBAR_THEME_TABLE;
}

function refreshAddonMenuForTheme(themeMode) {
  var ui = SpreadsheetApp.getUi();
  if (ui && typeof ui.removeMenu === 'function') {
    try {
      ui.removeMenu('TableCrafter AI');
    } catch (error) {}
    try {
      ui.removeMenu('TableCrafterAI');
    } catch (error) {}
    try {
      ui.removeMenu('WordCrafter AI');
    } catch (error) {}
    try {
      ui.removeMenu('WordCrafterAI');
    } catch (error) {}
  }
  ui.createMenu('TableCrafterAI')
    .addItem('Abrir TableCrafterAI', 'showTablecrafterSidebar')
    .addItem('Abrir WordCrafterAI', 'showWordcrafterSidebar')
    .addItem('Configurar API', 'showConfig')
    .addItem('Ayuda', 'showHelp')
    .addToUi();
}

function setSidebarThemeMode(mode) {
  var normalized = mode === SIDEBAR_THEME_WORD ? SIDEBAR_THEME_WORD : SIDEBAR_THEME_TABLE;
  PropertiesService.getUserProperties().setProperty(SIDEBAR_THEME_PROPERTY_KEY, normalized);
  refreshAddonMenuForTheme(normalized);
  return normalized;
}

function showTablecrafterSidebar() {
  setSidebarThemeMode(SIDEBAR_THEME_TABLE);
  showSidebar();
}

function showWordcrafterSidebar() {
  setSidebarThemeMode(SIDEBAR_THEME_WORD);
  showSidebar();
}

function getWelcomeMessagePreference() {
  var userProperties = PropertiesService.getUserProperties();
  var tableValue = userProperties.getProperty(WELCOME_MESSAGE_PROPERTY_KEY_TABLE);
  var wordValue = userProperties.getProperty(WELCOME_MESSAGE_PROPERTY_KEY_WORD);
  var legacyValue = userProperties.getProperty(WELCOME_MESSAGE_PROPERTY_KEY);
  var hasExplicitTable = tableValue !== null && tableValue !== undefined;
  var hasExplicitWord = wordValue !== null && wordValue !== undefined;
  var hideTable = hasExplicitTable ? tableValue === 'true' : false;
  var hideWord = hasExplicitWord ? wordValue === 'true' : false;
  if (!hasExplicitTable && !hasExplicitWord && legacyValue !== null && legacyValue !== undefined) {
    var legacyHidden = legacyValue === 'true';
    hideTable = legacyHidden;
    hideWord = legacyHidden;
  }
  return { hideTable: hideTable, hideWord: hideWord };
}

function setWelcomeMessagePreference(showMessage, theme) {
  var userProperties = PropertiesService.getUserProperties();
  var themesToUpdate;
  if (theme === SIDEBAR_THEME_WORD) {
    themesToUpdate = [SIDEBAR_THEME_WORD];
  } else if (theme === SIDEBAR_THEME_TABLE) {
    themesToUpdate = [SIDEBAR_THEME_TABLE];
  } else if (theme === 'all') {
    themesToUpdate = [SIDEBAR_THEME_TABLE, SIDEBAR_THEME_WORD];
  } else if (theme === undefined || theme === null) {
    themesToUpdate = [SIDEBAR_THEME_TABLE, SIDEBAR_THEME_WORD];
  } else {
    themesToUpdate = [SIDEBAR_THEME_TABLE];
  }
  for (var i = 0; i < themesToUpdate.length; i += 1) {
    var targetTheme = themesToUpdate[i];
    var propertyKey =
      targetTheme === SIDEBAR_THEME_WORD ? WELCOME_MESSAGE_PROPERTY_KEY_WORD : WELCOME_MESSAGE_PROPERTY_KEY_TABLE;
    if (showMessage) {
      userProperties.deleteProperty(propertyKey);
    } else {
      userProperties.setProperty(propertyKey, 'true');
    }
  }
  userProperties.deleteProperty(WELCOME_MESSAGE_PROPERTY_KEY);
  return getWelcomeMessagePreference();
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

var META_HEADERS = [
  'id',
  'name',
  'rangeA1',
  'description',
  'sheet',
  'cols',
  'rows',
  'createdAt',
  'updatedAt',
  'style',
  'headers',
  'formulaRefs',
  'type',
  'reportConfig'
];
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
  headers: 10,
  formulaRefs: 11,
  recordType: 12,
  reportConfig: 13
};

var REPORT_FEATURE_MAX_QUANTITY = 3;
var REPORT_FEATURES_WITHOUT_QUANTITY = {};
var REPORT_FEATURE_FIXED_QUANTITY = {};
var REPORT_CHART_TYPE_DEFAULT = 'auto';
var REPORT_CHART_TYPE_ALLOWED = {
  auto: true,
  line: true,
  bar: true,
  column: true,
  pie: true,
  area: true
};
var REPORT_LIST_ORDER_DEFAULT = 'ai';
var REPORT_LIST_ORDER_ALLOWED = {
  ai: true,
  alphabetical: true,
  reversealphabetical: true,
  priority: true,
  original: true
};
var REPORT_LIST_ENUMERATE_DEFAULT = 'no';
var REPORT_LIST_ENUMERATE_ALLOWED = {
  yes: true,
  no: true
};
var REPORT_LIST_ENUM_STYLE_DEFAULT = 'auto';
var REPORT_LIST_ENUM_STYLE_ALLOWED = {
  auto: true,
  numbers: true,
  bullets: true,
  dashes: true,
  checks: true,
  letters: true,
  roman: true
};
var REPORT_FEATURE_SEQUENCE = ['chart', 'table', 'list', 'summary'];
var REPORT_OUTPUT_LANGUAGE_DEFAULT = 'auto';
var REPORT_OUTPUT_LANGUAGE_ALLOWED = {
  auto: true,
  es: true,
  en: true,
  pt: true,
  fr: true,
  other: true
};
var REPORT_TONE_DEFAULT = 'neutral';
var REPORT_TONE_ALLOWED = {
  neutral: true,
  executive: true,
  friendly: true,
  analytical: true,
  custom: true
};
var REPORT_SUMMARY_LENGTH_DEFAULT = 'standard';
var REPORT_SUMMARY_LENGTH_ALLOWED = {
  concise: true,
  standard: true,
  detailed: true,
  executive: true
};
var REPORT_PERSONA_DEFAULT = 'auto';
var REPORT_PERSONA_ALLOWED = {
  auto: true,
  executive: true,
  'data-analyst': true,
  finance: true,
  marketing: true,
  operations: true,
  hr: true,
  it: true,
  sales: true,
  custom: true
};

var ALL_TABLES_OPTION_VALUE = '__ALL__';

var ASK_TONE_INSTRUCTIONS = {
  neutral: 'Responde con un tono neutral y profesional, evitando juicios subjetivos.',
  executive: 'Mantén un tono ejecutivo con foco en decisiones y conclusiones accionables.',
  friendly: 'Utiliza un tono cercano y motivador que refuerce la colaboración del equipo.',
  analytical: 'Redacta con un tono analítico y detallado destacando datos y métricas relevantes.'
};

var ASK_PERSONA_INSTRUCTIONS = {
  auto: '',
  executive: 'Enmarca la explicación para la dirección general, resaltando impacto y riesgos estratégicos.',
  'data-analyst': 'Incluye detalles técnicos, tendencias y cálculos útiles para un analista de datos.',
  finance: 'Prioriza indicadores financieros, márgenes, costos y retornos esperados.',
  marketing: 'Resalta métricas de marketing, comportamiento de audiencias y posicionamiento de marca.',
  operations: 'Destaca eficiencia operativa, tiempos de entrega y procesos clave.',
  hr: 'Enfoca la información en talento, clima laboral y desarrollo del equipo.',
  it: 'Incluye consideraciones técnicas, integraciones y requisitos de sistemas.',
  sales: 'Orienta la respuesta a objetivos comerciales, conversiones y próximos pasos de ventas.',
  custom: ''
};

function normalizeMetaId(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function normalizeMetaRecordType(value) {
  if (value === null || value === undefined) {
    return 'table';
  }
  var text = String(value).trim().toLowerCase();
  if (!text) {
    return 'table';
  }
  if (text === 'reportfavorite' || text === 'report_favorite' || text === 'report-favorite') {
    return 'reportFavorite';
  }
  if (text === 'table') {
    return 'table';
  }
  return 'table';
}

function parseBooleanValue(value, defaultValue) {
  if (value === null || value === undefined || value === '') {
    return defaultValue;
  }
  if (typeof value === 'boolean') {
    return value;
  }
  var normalized = String(value).trim().toLowerCase();
  if (normalized === 'true' || normalized === '1' || normalized === 'si' || normalized === 'sí' || normalized === 'yes') {
    return true;
  }
  if (normalized === 'false' || normalized === '0' || normalized === 'no') {
    return false;
  }
  return defaultValue;
}

function sanitizeReportFileName(value) {
  if (value === null || value === undefined) {
    return '';
  }
  var text = String(value).trim();
  if (!text) {
    return '';
  }
  if (typeof text.normalize === 'function') {
    text = text.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  }
  text = text.replace(/\s+/g, '');
  text = text.replace(/[^A-Za-z0-9_-]/g, '');
  return text;
}

function buildTableNamedRangeName(id) {
  var normalized = normalizeMetaId(id);
  if (!normalized) {
    return '';
  }
  var sanitized = normalized.replace(/[^A-Za-z0-9_]/g, '_');
  if (!/^[A-Za-z_]/.test(sanitized)) {
    sanitized = 'TC_' + sanitized;
  }
  return 'TableCrafterRange_' + sanitized;
}

function findNamedRangeByName(name, optSpreadsheet) {
  if (!name) {
    return null;
  }
  var ss = optSpreadsheet || SpreadsheetApp.getActive();
  if (!ss) {
    return null;
  }
  var namedRanges;
  try {
    namedRanges = ss.getNamedRanges();
  } catch (err) {
    return null;
  }
  if (!namedRanges || namedRanges.length === 0) {
    return null;
  }
  for (var i = 0; i < namedRanges.length; i++) {
    var candidate = namedRanges[i];
    if (!candidate) {
      continue;
    }
    try {
      if (candidate.getName && candidate.getName() === name) {
        return candidate;
      }
    } catch (err2) {
      // Ignorar y continuar buscando.
    }
  }
  return null;
}

function syncNamedRangeForTable(id, sheetName, rangeA1) {
  var name = buildTableNamedRangeName(id);
  if (!name) {
    return;
  }
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    return;
  }
  var namedRange = findNamedRangeByName(name, ss);
  if (!sheetName || !rangeA1) {
    if (namedRange && namedRange.remove) {
      try {
        namedRange.remove();
      } catch (err) {
        // Ignorar fallos al eliminar el rango con nombre.
      }
    }
    return;
  }
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    if (namedRange && namedRange.remove) {
      try {
        namedRange.remove();
      } catch (err2) {
        // Ignorar fallos al eliminar el rango con nombre.
      }
    }
    return;
  }
  var range;
  try {
    range = getRangeWithinSheet(sheet, rangeA1);
  } catch (err3) {
    range = null;
  }
  if (!range) {
    if (namedRange && namedRange.remove) {
      try {
        namedRange.remove();
      } catch (err4) {
        // Ignorar fallos al eliminar el rango con nombre.
      }
    }
    return;
  }
  if (namedRange) {
    try {
      namedRange.setRange(range);
    } catch (err5) {
      // Si no se puede actualizar, intentar recrearlo.
      try {
        namedRange.remove();
      } catch (err6) {
        // Ignorar y continuar creando uno nuevo.
      }
      ss.setNamedRange(name, range);
    }
  } else {
    ss.setNamedRange(name, range);
  }
}

function removeNamedRangeForTable(id) {
  var name = buildTableNamedRangeName(id);
  if (!name) {
    return;
  }
  var ss = SpreadsheetApp.getActive();
  if (!ss) {
    return;
  }
  var namedRange = findNamedRangeByName(name, ss);
  if (namedRange && namedRange.remove) {
    try {
      namedRange.remove();
    } catch (err) {
      // Ignorar fallos al eliminar.
    }
  }
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
    key: '',
    aiFillEnabled: false,
    aiFillPrompt: '',
    aiFillDataType: DEFAULT_AI_FILL_DATA_TYPE
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
    var hasExplicitAiFlag = Object.prototype.hasOwnProperty.call(entry, 'aiFillEnabled');
    var aiEnabledValue = hasExplicitAiFlag ? entry.aiFillEnabled : false;
    var aiPromptValue = Object.prototype.hasOwnProperty.call(entry, 'aiFillPrompt') ? entry.aiFillPrompt : '';
    var aiDataTypeValue = Object.prototype.hasOwnProperty.call(entry, 'aiFillDataType') ? entry.aiFillDataType : '';
    var trimmedPrompt = String(aiPromptValue === undefined || aiPromptValue === null ? '' : aiPromptValue).trim();
    var trimmedDataType = String(aiDataTypeValue === undefined || aiDataTypeValue === null ? '' : aiDataTypeValue).trim();
    var aiEnabled = hasExplicitAiFlag ? !!aiEnabledValue : trimmedPrompt !== '';
    normalized.label = String(labelValue === undefined || labelValue === null ? '' : labelValue).trim();
    normalized.description = hasDescription ? String(descriptionValue || '').trim() : '';
    normalized.hasDescription = hasDescription;
    normalized.key = normalizeHeaderKey(keyValue);
    normalized.aiFillEnabled = aiEnabled;
    normalized.aiFillPrompt = trimmedPrompt;
    normalized.aiFillDataType = trimmedDataType || DEFAULT_AI_FILL_DATA_TYPE;
    return normalized;
  }
  if (entry !== undefined && entry !== null) {
    normalized.label = String(entry).trim();
  }
  return normalized;
}

function columnLetterToNumber(letter) {
  if (!letter) {
    return null;
  }
  var value = 0;
  var upper = String(letter).replace(/\$/g, '').toUpperCase();
  for (var i = 0; i < upper.length; i++) {
    var code = upper.charCodeAt(i);
    if (code < 65 || code > 90) {
      return null;
    }
    value = value * 26 + (code - 64);
  }
  return value;
}

function columnNumberToLetter(number) {
  var num = parseInt(number, 10);
  if (!num || num < 1) {
    return null;
  }
  var result = '';
  while (num > 0) {
    var mod = (num - 1) % 26;
    result = String.fromCharCode(65 + mod) + result;
    num = Math.floor((num - 1) / 26);
  }
  return result;
}

function parseRangeA1(rangeA1) {
  if (!rangeA1) {
    return null;
  }
  var trimmed = String(rangeA1).trim();
  if (trimmed === '') {
    return null;
  }
  var exclIndex = trimmed.lastIndexOf('!');
  if (exclIndex !== -1) {
    trimmed = trimmed.substring(exclIndex + 1);
  }
  var parts = trimmed.split(':');
  var startRef = parts[0];
  var endRef = parts.length > 1 ? parts[1] : parts[0];
  var startMatch = startRef.match(/\$?([A-Za-z]+)\$?(\d+)/);
  if (!startMatch) {
    return null;
  }
  var endMatch = endRef.match(/\$?([A-Za-z]+)\$?(\d+)/);
  if (!endMatch) {
    endMatch = startMatch;
  }
  var startColumn = columnLetterToNumber(startMatch[1]);
  var startRow = parseInt(startMatch[2], 10);
  var endColumn = columnLetterToNumber(endMatch[1]);
  var endRow = parseInt(endMatch[2], 10);
  if (!startColumn || !startRow || !endColumn || !endRow) {
    return null;
  }
  var normalizedStartColumn = Math.min(startColumn, endColumn);
  var normalizedEndColumn = Math.max(startColumn, endColumn);
  var normalizedStartRow = Math.min(startRow, endRow);
  var normalizedEndRow = Math.max(startRow, endRow);
  return {
    startColumn: normalizedStartColumn,
    startRow: normalizedStartRow,
    endColumn: normalizedEndColumn,
    endRow: normalizedEndRow,
    cols: normalizedEndColumn - normalizedStartColumn + 1,
    rows: normalizedEndRow - normalizedStartRow + 1
  };
}

function buildA1Notation(startColumn, startRow, cols, rows) {
  var col = parseInt(startColumn, 10);
  var row = parseInt(startRow, 10);
  var totalCols = parseInt(cols, 10);
  var totalRows = parseInt(rows, 10);
  if (!col || !row || !totalCols || !totalRows) {
    return '';
  }
  var endColumn = col + totalCols - 1;
  var endRow = row + totalRows - 1;
  var startLetter = columnNumberToLetter(col);
  var endLetter = columnNumberToLetter(endColumn);
  if (!startLetter || !endLetter) {
    return '';
  }
  var startRef = startLetter + row;
  if (totalCols === 1 && totalRows === 1) {
    return startRef;
  }
  if (totalCols === 1) {
    return startRef + ':' + startLetter + endRow;
  }
  if (totalRows === 1) {
    return startRef + ':' + endLetter + row;
  }
  return startRef + ':' + endLetter + endRow;
}

function splitRangeNotation(rangeA1) {
  var result = { sheet: '', range: '' };
  if (rangeA1 === null || rangeA1 === undefined) {
    return result;
  }
  var text = String(rangeA1).trim();
  if (text === '') {
    return result;
  }
  var exclIndex = text.lastIndexOf('!');
  if (exclIndex === -1) {
    result.range = text;
    return result;
  }
  var sheetPart = text.substring(0, exclIndex);
  var rangePart = text.substring(exclIndex + 1);
  if (sheetPart.length >= 2 && sheetPart.charAt(0) === "'" && sheetPart.charAt(sheetPart.length - 1) === "'") {
    sheetPart = sheetPart.substring(1, sheetPart.length - 1).replace(/''/g, "'");
  }
  result.sheet = sheetPart;
  result.range = rangePart;
  return result;
}

function stripSheetFromRange(rangeA1) {
  return splitRangeNotation(rangeA1).range;
}

function extractSheetNameFromRange(rangeA1) {
  return splitRangeNotation(rangeA1).sheet;
}

function buildFullRangeNotation(sheetName, rangeA1) {
  var pureRange = stripSheetFromRange(rangeA1);
  if (!pureRange) {
    return '';
  }
  var normalizedSheet = sheetName === null || sheetName === undefined ? '' : String(sheetName).trim();
  if (!normalizedSheet) {
    return pureRange;
  }
  var escaped = normalizedSheet.replace(/'/g, "''");
  return "'" + escaped + "'!" + pureRange;
}

function getRangeWithinSheet(sheet, rangeA1) {
  if (!sheet || typeof sheet.getRange !== 'function') {
    throw new Error('Hoja no encontrada.');
  }
  var pureRange = stripSheetFromRange(rangeA1);
  if (!pureRange) {
    throw new Error('Rango no especificado.');
  }
  return sheet.getRange(pureRange);
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
      if (normalizeMetaRecordType(row[META_INDEX.recordType]) !== 'table') {
        continue;
      }
      var style = parseJsonValue(row[META_INDEX.style], null);
      var headers = normalizeHeaderArray(parseJsonValue(row[META_INDEX.headers], null));
      var rangeParts = splitRangeNotation(row[META_INDEX.rangeA1]);
      var entrySheetName = row[META_INDEX.sheet] || rangeParts.sheet;
      var pureRange = rangeParts.range || '';
      var fullRange = buildFullRangeNotation(entrySheetName, pureRange);
      result.push({
        id: entryId,
        name: row[META_INDEX.name],
        rangeA1: pureRange,
        fullRangeA1: fullRange,
        description: row[META_INDEX.description],
        sheetName: entrySheetName,
        cols: row[META_INDEX.cols],
        rows: row[META_INDEX.rows],
        createdAt: row[META_INDEX.createdAt],
        updatedAt: row[META_INDEX.updatedAt],
        style: style,
        headers: headers,
        updateFormulaReferences: parseBooleanValue(row[META_INDEX.formulaRefs], true)
      });
    }
  }
  return result;
}

function listSavedActions() {
  SpreadsheetApp.flush();
  var meta = getMetaSheet();
  var data = meta.getDataRange().getValues();
  var actions = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    var type = normalizeMetaRecordType(row[META_INDEX.recordType]);
    if (type !== 'table' && type !== 'reportFavorite') {
      continue;
    }
    actions.push({
      id: entryId,
      type: type,
      name: row[META_INDEX.name] || '',
      description: row[META_INDEX.description] || '',
      createdAt: row[META_INDEX.createdAt] || '',
      updatedAt: row[META_INDEX.updatedAt] || ''
    });
  }
  actions.sort(function(a, b) {
    var nameA = (a && a.name ? String(a.name) : '').toLowerCase();
    var nameB = (b && b.name ? String(b.name) : '').toLowerCase();
    if (nameA < nameB) {
      return -1;
    }
    if (nameA > nameB) {
      return 1;
    }
    return 0;
  });
  return actions;
}

function buildReportFavoriteResponse(entry) {
  var data = entry && entry.data ? entry.data : [];
  var config = parseJsonValue(data[META_INDEX.reportConfig], {});
  if (!config || typeof config !== 'object') {
    config = {};
  }
  var features = Array.isArray(config.features) ? config.features : [];
  var featureDetails = Array.isArray(config.featureDetails) ? config.featureDetails : [];
  var tableList = Array.isArray(config.tables) ? config.tables : [];
  var normalizedTableSet = {};
  tableList.forEach(function(tableId) {
    var normalizedId = normalizeMetaId(tableId);
    if (normalizedId) {
      normalizedTableSet[normalizedId] = true;
    }
  });
  var detailMap = {};
  var normalizedDetails = [];
  var storedFileName = config.fileName || '';
  var normalizedFileName = sanitizeReportFileName(storedFileName);
  featureDetails.forEach(function(detail) {
    if (!detail) {
      return;
    }
    var featureId = normalizeMetaId(detail.feature);
    if (!featureId) {
      return;
    }
    if (detailMap[featureId]) {
      return;
    }
    var quantityValue = parseInt(detail.quantity, 10);
    if (isNaN(quantityValue) || quantityValue < 1) {
      quantityValue = 1;
    }
    if (quantityValue > 3) {
      quantityValue = 3;
    }
    var instanceDescriptions = ['', '', ''];
    if (Array.isArray(detail.instanceDescriptions)) {
      detail.instanceDescriptions.slice(0, 3).forEach(function(entry, index) {
        instanceDescriptions[index] = entry !== null && entry !== undefined ? String(entry) : '';
      });
    }
    var instanceTitles = normalizeFeatureTitleArray(detail.instanceTitles);
    var instanceChartTypes = normalizeFeatureChartTypeArray(detail.instanceChartTypes);
    var instanceListOrders = normalizeFeatureListOrderArray(detail.instanceListOrders);
    var instanceListEnumerate = normalizeFeatureListEnumerateArray(detail.instanceListEnumerate);
    var instanceListEnumStyles = normalizeFeatureListEnumStyleArray(detail.instanceListEnumStyles);
    var summaryLengthValue = normalizeSummaryLengthValue(detail.summaryLength);
    var instanceTablesEnabled = normalizeFeatureTableToggleArray(detail.instanceTablesEnabled);
    var instanceTables = normalizeFeatureTableSelectionArray(detail.instanceTables, normalizedTableSet);
    normalizedDetails.push({
      feature: featureId,
      quantity: quantityValue,
      instanceDescriptions: instanceDescriptions,
      instanceTitles: instanceTitles,
      instanceChartTypes: instanceChartTypes,
      instanceListOrders: instanceListOrders,
      instanceListEnumerate: instanceListEnumerate,
      instanceListEnumStyles: instanceListEnumStyles,
      summaryLength: summaryLengthValue,
      instanceTablesEnabled: instanceTablesEnabled,
      instanceTables: instanceTables
    });
    detailMap[featureId] = true;
  });
  features.forEach(function(feature) {
    var featureId = normalizeMetaId(feature);
    if (!featureId || detailMap[featureId]) {
      return;
    }
    normalizedDetails.push({
      feature: featureId,
      quantity: getReportFeatureDefaultQuantity(featureId),
      instanceDescriptions: ['', '', ''],
      instanceTitles: createEmptyFeatureTitleArray(),
      instanceChartTypes: normalizeFeatureChartTypeArray([]),
      instanceListOrders: normalizeFeatureListOrderArray([]),
      instanceListEnumerate: normalizeFeatureListEnumerateArray([]),
      instanceListEnumStyles: normalizeFeatureListEnumStyleArray([]),
      summaryLength: REPORT_SUMMARY_LENGTH_DEFAULT,
      instanceTablesEnabled: createEmptyFeatureTableToggleArray(),
      instanceTables: createEmptyFeatureTableSelectionArray()
    });
    detailMap[featureId] = true;
  });
  var quantityMap = buildFeatureQuantityMap(normalizedDetails);
  var normalizedOrder = normalizeFeatureOrderEntries(config.featureOrder, quantityMap);
  return {
    id: normalizeMetaId(data[META_INDEX.id]),
    type: 'reportFavorite',
    name: data[META_INDEX.name] || '',
    description: data[META_INDEX.description] || '',
    createdAt: data[META_INDEX.createdAt] || '',
    updatedAt: data[META_INDEX.updatedAt] || '',
    format: config.format || '',
    fileName: normalizedFileName,
    fileExtension: config.fileExtension || '',
    appendDateToFile: parseBooleanValue(config.appendDateToFile, false),
    tables: Array.isArray(config.tables) ? config.tables : [],
    channels: Array.isArray(config.channels) ? config.channels : [],
    descriptionEnabled: parseBooleanValue(config.descriptionEnabled, false),
    descriptionText: config.description || '',
    features: features,
    featureDetails: normalizedDetails,
    featureOrder: normalizedOrder,
    emails: Array.isArray(config.emails) ? config.emails : [],
    phones: Array.isArray(config.phones) ? config.phones : [],
    customization: normalizeReportCustomization(config.customization)
  };
}

function getSavedActionDetails(actionId) {
  var id = normalizeMetaId(actionId);
  if (!id) {
    return { error: 'Acción no encontrada.' };
  }
  var entry = findMetaById(id);
  if (!entry || !entry.data) {
    return { error: 'Acción no encontrada.' };
  }
  var type = normalizeMetaRecordType(entry.data[META_INDEX.recordType]);
  if (type === 'table') {
    var row = entry.data;
    var rangeParts = splitRangeNotation(row[META_INDEX.rangeA1]);
    var sheetName = row[META_INDEX.sheet] || rangeParts.sheet;
    var pureRange = rangeParts.range || '';
    return {
      id: normalizeMetaId(row[META_INDEX.id]),
      type: 'table',
      name: row[META_INDEX.name] || '',
      description: row[META_INDEX.description] || '',
      sheetName: sheetName || '',
      rangeA1: pureRange,
      fullRangeA1: buildFullRangeNotation(sheetName, pureRange),
      cols: row[META_INDEX.cols] || '',
      rows: row[META_INDEX.rows] || '',
      createdAt: row[META_INDEX.createdAt] || '',
      updatedAt: row[META_INDEX.updatedAt] || ''
    };
  }
  if (type === 'reportFavorite') {
    return buildReportFavoriteResponse(entry);
  }
  return { error: 'Acción no soportada.' };
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
  if (normalizeMetaRecordType(entry.data[META_INDEX.recordType]) !== 'table') {
    return { error: 'Tabla no encontrada.' };
  }
  var rowData = entry.data || [];
  var rangeParts = splitRangeNotation(rowData[META_INDEX.rangeA1]);
  var sheetName = rowData[META_INDEX.sheet] || rangeParts.sheet;
  var rangeA1 = rangeParts.range;
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
    range = getRangeWithinSheet(sheet, rowData[META_INDEX.rangeA1] || rangeA1);
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
    rangeA1: buildFullRangeNotation(sheetName, rangeA1),
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
  if (normalizeMetaRecordType(entry.data[META_INDEX.recordType]) !== 'table') {
    return { error: 'Tabla no encontrada' };
  }
  var row = entry.row;
  var data = entry.data;
  var rangeParts = splitRangeNotation(data[META_INDEX.rangeA1]);
  var rangeA1 = rangeParts.range;
  var sheetName = data[META_INDEX.sheet] || rangeParts.sheet;
  if (rangeA1) {
    try {
      var ss = SpreadsheetApp.getActive();
      var sheet = ss.getSheetByName(sheetName);
      var range = getRangeWithinSheet(sheet, data[META_INDEX.rangeA1] || rangeA1);
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
  if (normalizeMetaRecordType(entry.data[META_INDEX.recordType]) !== 'table') {
    return { error: 'Tabla no encontrada.' };
  }

  var data = entry.data;
  var rangeParts = splitRangeNotation(data[META_INDEX.rangeA1]);
  var rangeA1 = rangeParts.range;
  var sheetName = data[META_INDEX.sheet] || rangeParts.sheet;
  var warnings = [];

  if (rangeA1 && sheetName) {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      try {
        var range = getRangeWithinSheet(sheet, data[META_INDEX.rangeA1] || rangeA1);
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
  removeNamedRangeForTable(id);
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
  var formulaPreference = shouldUpdateFormulaReferences;
  var hasExplicitFormulaPreference = false;
  if (options && Object.prototype.hasOwnProperty.call(options, 'formulaPreference')) {
    formulaPreference = !!options.formulaPreference;
    hasExplicitFormulaPreference = true;
  } else if (options && Object.prototype.hasOwnProperty.call(options, 'updateFormulaReferences')) {
    formulaPreference = !!options.updateFormulaReferences;
    hasExplicitFormulaPreference = true;
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
    var hasExplicitAiFlag = entry && Object.prototype.hasOwnProperty.call(entry, 'aiFillEnabled');
    var aiEnabledValue = hasExplicitAiFlag ? entry.aiFillEnabled : false;
    var aiPromptValue = entry && entry.aiFillPrompt !== undefined && entry.aiFillPrompt !== null ? String(entry.aiFillPrompt) : '';
    var aiDataTypeValue = entry && entry.aiFillDataType !== undefined && entry.aiFillDataType !== null ? String(entry.aiFillDataType) : '';
    var trimmedPrompt = aiPromptValue.trim();
    var trimmedDataType = aiDataTypeValue.trim() || DEFAULT_AI_FILL_DATA_TYPE;
    var aiEnabled = hasExplicitAiFlag ? !!aiEnabledValue : trimmedPrompt !== '';
    return {
      label: trimmedLabel,
      description: hasDescription ? trimmedDescription : '',
      hasDescription: hasDescription,
      key: normalizeHeaderKey(keyValue),
      aiFillEnabled: aiEnabled,
      aiFillPrompt: trimmedPrompt,
      aiFillDataType: trimmedDataType
    };
  });
  var headerMetaForReturn = headerMeta.map(function(item) {
    return {
      label: item.label || '',
      description: item.hasDescription ? (item.description || '') : '',
      hasDescription: !!item.hasDescription,
      key: item.key || '',
      aiFillEnabled: !!item.aiFillEnabled,
      aiFillPrompt: item.aiFillPrompt ? String(item.aiFillPrompt).trim() : '',
      aiFillDataType: item.aiFillDataType ? String(item.aiFillDataType).trim() : DEFAULT_AI_FILL_DATA_TYPE
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
        var storedFormulaPreference = parseBooleanValue(entry.data[META_INDEX.formulaRefs], true);
        if (!hasExplicitFormulaPreference) {
          formulaPreference = storedFormulaPreference;
          shouldUpdateFormulaReferences = storedFormulaPreference;
        }
        storedHeaderMeta = normalizeHeaderArray(parseJsonValue(entry.data[META_INDEX.headers], null));
        var prevRangeStored = entry.data[META_INDEX.rangeA1];
        var prevRangeParts = splitRangeNotation(prevRangeStored);
        var prevRangeA1 = prevRangeParts.range;
        var prevSheetName = entry.data[META_INDEX.sheet] || prevRangeParts.sheet;
        if (prevRangeA1 && prevSheetName) {
          var ss = SpreadsheetApp.getActive();
          var prevSheet = ss.getSheetByName(prevSheetName);
          if (prevSheet) {
            var allowOverlapWithHeaderRow = false;
            try {
              var prevRange = getRangeWithinSheet(prevSheet, prevRangeStored || prevRangeA1);
              var sameSheet = prevSheet.getSheetId() === sheet.getSheetId();
              var sameRange = sameSheet && prevRange.getA1Notation() === normalizedRange;
              if (!sameRange && sameSheet && rangesIntersect(prevRange, range)) {
                if (prevRange.getRow() === range.getRow()) {
                  allowOverlapWithHeaderRow = true;
                } else {
                  return {
                    error: 'No se puede superponer el nuevo rango con el anterior.'
                  };
                }
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
                previousRangeDetails = {
                  range: prevRange,
                  sheet: prevSheet,
                  allowOverlap: allowOverlapWithHeaderRow
                };
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
      if (previousRangeDetails.allowOverlap) {
        clearRangePortionsOutsideTarget(previousRangeDetails.range, range);
      } else {
        clearRangeAndFormatting(previousRangeDetails.range);
      }
    } catch (cleanupErr) {
      // Ignorar errores al limpiar el rango anterior para no interrumpir la operación principal.
    }
  }

  if (doSave) {
    saveOrUpdateMeta(
      id,
      name,
      description,
      normalizedRange,
      sheet.getName(),
      cols,
      rows,
      appliedStyle,
      headerMetaForReturn,
      formulaPreference
    );
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
      headers: headerMetaForReturn,
      updateFormulaReferences: formulaPreference
    },
    message: message
  };
  if (doSave) {
    result.isNew = isNew;
  }
  return result;
}

function extractSubMatrix(matrix, rowOffset, colOffset, numRows, numCols) {
  var result = [];
  if (!Array.isArray(matrix)) {
    return result;
  }
  for (var r = 0; r < numRows; r++) {
    var sourceRow = matrix[rowOffset + r] || [];
    var rowResult = [];
    for (var c = 0; c < numCols; c++) {
      rowResult.push(sourceRow[colOffset + c]);
    }
    result.push(rowResult);
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

function clearRangePortionsOutsideTarget(previousRange, targetRange) {
  if (!previousRange || !targetRange) {
    return;
  }
  var previousSheet = previousRange.getSheet();
  var targetSheet = targetRange.getSheet();
  if (!previousSheet || !targetSheet || previousSheet.getSheetId() !== targetSheet.getSheetId()) {
    clearRangeAndFormatting(previousRange);
    return;
  }
  if (!rangesIntersect(previousRange, targetRange)) {
    clearRangeAndFormatting(previousRange);
    return;
  }
  var prevStartRow = previousRange.getRow();
  var prevRows = previousRange.getNumRows();
  var prevCols = previousRange.getNumColumns();
  var prevStartCol = previousRange.getColumn();
  var prevEndRow = prevStartRow + prevRows - 1;
  var prevEndCol = prevStartCol + prevCols - 1;

  var targetStartRow = targetRange.getRow();
  var targetRows = targetRange.getNumRows();
  var targetCols = targetRange.getNumColumns();
  var targetStartCol = targetRange.getColumn();
  var targetEndRow = targetStartRow + targetRows - 1;
  var targetEndCol = targetStartCol + targetCols - 1;

  if (targetStartRow > prevStartRow) {
    var topRows = targetStartRow - prevStartRow;
    if (topRows > 0) {
      clearRangeAndFormatting(previousRange.offset(0, 0, topRows, prevCols));
    }
  }

  if (targetEndRow < prevEndRow) {
    var bottomRows = prevEndRow - targetEndRow;
    if (bottomRows > 0) {
      clearRangeAndFormatting(previousRange.offset(targetEndRow - prevStartRow + 1, 0, bottomRows, prevCols));
    }
  }

  if (targetStartCol > prevStartCol) {
    var leftCols = targetStartCol - prevStartCol;
    if (leftCols > 0) {
      clearRangeAndFormatting(previousRange.offset(0, 0, prevRows, leftCols));
    }
  }

  if (targetEndCol < prevEndCol) {
    var rightCols = prevEndCol - targetEndCol;
    if (rightCols > 0) {
      clearRangeAndFormatting(previousRange.offset(0, targetEndCol - prevStartCol + 1, prevRows, rightCols));
    }
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
  var typeFilter = [];
  if (Object.prototype.hasOwnProperty.call(opts, 'type')) {
    var rawTypes = Array.isArray(opts.type) ? opts.type : [opts.type];
    rawTypes.forEach(function(value) {
      var normalizedType = normalizeMetaRecordType(value);
      if (normalizedType && typeFilter.indexOf(normalizedType) === -1) {
        typeFilter.push(normalizedType);
      }
    });
  } else {
    typeFilter.push('table');
  }
  var sheet = getMetaSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var entryId = normalizeMetaId(row[META_INDEX.id]);
    if (!entryId) {
      continue;
    }
    var recordType = normalizeMetaRecordType(row[META_INDEX.recordType]);
    if (typeFilter.length > 0 && typeFilter.indexOf(recordType) === -1) {
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
function saveOrUpdateMeta(id, name, description, rangeA1, sheetName, cols, rows, style, headers, formulaPreference) {
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
  var normalizedFormulaPreference = typeof formulaPreference === 'boolean' ? formulaPreference : true;
  var formulaPreferenceValue = normalizedFormulaPreference ? 'TRUE' : 'FALSE';
  var normalizedRangeOnly = stripSheetFromRange(rangeA1);
  var storedRange = buildFullRangeNotation(sheetName, normalizedRangeOnly);
  if (meta) {
    // Actualizar
    var row = meta.row;
    var createdAt = meta.data[META_INDEX.createdAt] || now;
    var storedConfig = meta.data.length > META_INDEX.reportConfig ? meta.data[META_INDEX.reportConfig] : '';
    sheet
      .getRange(row, 1, 1, META_HEADERS.length)
      .setValues([
        [
          normalizedId,
          name,
          storedRange,
          description,
          sheetName,
          cols,
          rows,
          createdAt,
          now,
          styleValue,
          headersValue,
          formulaPreferenceValue,
          'table',
          storedConfig
        ]
      ]);
  } else {
    // Crear nueva
    sheet.appendRow([
      normalizedId,
      name,
      storedRange,
      description,
      sheetName,
      cols,
      rows,
      now,
      now,
      styleValue,
      headersValue,
      formulaPreferenceValue,
      'table',
      ''
    ]);
  }

  syncNamedRangeForTable(normalizedId, sheetName, storedRange || normalizedRangeOnly);
  SpreadsheetApp.flush();
}

function getReportFeatureFixedQuantity(featureId) {
  if (!featureId) {
    return 0;
  }
  var fixed = REPORT_FEATURE_FIXED_QUANTITY[featureId];
  if (!fixed) {
    return 0;
  }
  if (fixed < 1) {
    return 1;
  }
  if (fixed > REPORT_FEATURE_MAX_QUANTITY) {
    return REPORT_FEATURE_MAX_QUANTITY;
  }
  return fixed;
}

function getReportFeatureDefaultQuantity(featureId) {
  var fixed = getReportFeatureFixedQuantity(featureId);
  if (fixed > 0) {
    return fixed;
  }
  return 1;
}

function reportFeatureAllowsQuantity(featureId) {
  if (!featureId) {
    return true;
  }
  return !REPORT_FEATURES_WITHOUT_QUANTITY[featureId];
}

function clampReportFeatureQuantityValue(featureId, value) {
  if (!reportFeatureAllowsQuantity(featureId)) {
    return getReportFeatureDefaultQuantity(featureId);
  }
  var parsed = parseInt(value, 10);
  if (isNaN(parsed) || parsed < 1) {
    return 1;
  }
  if (parsed > REPORT_FEATURE_MAX_QUANTITY) {
    return REPORT_FEATURE_MAX_QUANTITY;
  }
  return parsed;
}

function normalizeChartTypeValue(value) {
  if (value === null || value === undefined) {
    return REPORT_CHART_TYPE_DEFAULT;
  }
  var text = String(value).trim().toLowerCase();
  if (!text) {
    return REPORT_CHART_TYPE_DEFAULT;
  }
  if (REPORT_CHART_TYPE_ALLOWED[text]) {
    return text;
  }
  return REPORT_CHART_TYPE_DEFAULT;
}

function normalizeFeatureChartTypeArray(source) {
  var base = [
    REPORT_CHART_TYPE_DEFAULT,
    REPORT_CHART_TYPE_DEFAULT,
    REPORT_CHART_TYPE_DEFAULT
  ];
  if (Array.isArray(source)) {
    source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
      base[index] = normalizeChartTypeValue(entry);
    });
  }
  return base;
}

function createEmptyFeatureTitleArray() {
  return ['', '', ''];
}

function normalizeFeatureTitleArray(source) {
  var base = createEmptyFeatureTitleArray();
  if (Array.isArray(source)) {
    source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
      base[index] = entry !== null && entry !== undefined ? String(entry).trim() : '';
    });
  }
  return base;
}

function createEmptyFeatureListOrderArray() {
  return [
    REPORT_LIST_ORDER_DEFAULT,
    REPORT_LIST_ORDER_DEFAULT,
    REPORT_LIST_ORDER_DEFAULT
  ];
}

function createEmptyFeatureListEnumerateArray() {
  return [
    REPORT_LIST_ENUMERATE_DEFAULT,
    REPORT_LIST_ENUMERATE_DEFAULT,
    REPORT_LIST_ENUMERATE_DEFAULT
  ];
}

function createEmptyFeatureListEnumStyleArray() {
  return [
    REPORT_LIST_ENUM_STYLE_DEFAULT,
    REPORT_LIST_ENUM_STYLE_DEFAULT,
    REPORT_LIST_ENUM_STYLE_DEFAULT
  ];
}

function normalizeListOrderValue(value) {
  if (value === null || value === undefined) {
    return REPORT_LIST_ORDER_DEFAULT;
  }
  var normalized = String(value).trim().toLowerCase();
  if (!normalized) {
    return REPORT_LIST_ORDER_DEFAULT;
  }
  if (REPORT_LIST_ORDER_ALLOWED[normalized]) {
    return normalized;
  }
  return REPORT_LIST_ORDER_DEFAULT;
}

function normalizeFeatureListOrderArray(source) {
  var base = createEmptyFeatureListOrderArray();
  if (Array.isArray(source)) {
    source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
      base[index] = normalizeListOrderValue(entry);
    });
  }
  return base;
}

function normalizeListEnumerateValue(value) {
  if (value === null || value === undefined) {
    return REPORT_LIST_ENUMERATE_DEFAULT;
  }
  var normalized = String(value).trim().toLowerCase();
  if (!normalized) {
    return REPORT_LIST_ENUMERATE_DEFAULT;
  }
  if (REPORT_LIST_ENUMERATE_ALLOWED[normalized]) {
    return normalized;
  }
  return REPORT_LIST_ENUMERATE_DEFAULT;
}

function normalizeFeatureListEnumerateArray(source) {
  var base = createEmptyFeatureListEnumerateArray();
  if (Array.isArray(source)) {
    source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
      base[index] = normalizeListEnumerateValue(entry);
    });
  }
  return base;
}

function normalizeListEnumStyleValue(value) {
  if (value === null || value === undefined) {
    return REPORT_LIST_ENUM_STYLE_DEFAULT;
  }
  var normalized = String(value).trim().toLowerCase();
  if (!normalized) {
    return REPORT_LIST_ENUM_STYLE_DEFAULT;
  }
  if (REPORT_LIST_ENUM_STYLE_ALLOWED[normalized]) {
    return normalized;
  }
  return REPORT_LIST_ENUM_STYLE_DEFAULT;
}

function normalizeFeatureListEnumStyleArray(source) {
  var base = createEmptyFeatureListEnumStyleArray();
  if (Array.isArray(source)) {
    source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
      base[index] = normalizeListEnumStyleValue(entry);
    });
  }
  return base;
}

function createEmptyFeatureTableToggleArray() {
  return [false, false, false];
}

function normalizeFeatureTableToggleArray(source) {
  var base = createEmptyFeatureTableToggleArray();
  if (Array.isArray(source)) {
    source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
      base[index] = !!entry;
    });
  }
  return base;
}

function createEmptyFeatureTableSelectionArray() {
  return [[], [], []];
}

function normalizeFeatureTableSelectionArray(source, allowedSet) {
  var base = createEmptyFeatureTableSelectionArray();
  if (!Array.isArray(source)) {
    return base;
  }
  var allowAll = !allowedSet;
  source.slice(0, REPORT_FEATURE_MAX_QUANTITY).forEach(function(entry, index) {
    if (!Array.isArray(entry)) {
      base[index] = [];
      return;
    }
    var seen = {};
    var values = [];
    entry.forEach(function(value) {
      if (value === null || value === undefined) {
        return;
      }
      var text = String(value).trim();
      if (!text || text === ALL_TABLES_OPTION_VALUE || seen[text]) {
        return;
      }
      if (!allowAll && !allowedSet[text]) {
        return;
      }
      seen[text] = true;
      values.push(text);
    });
    base[index] = values;
  });
  return base;
}

function buildFeatureQuantityMap(details) {
  var map = {};
  if (!Array.isArray(details)) {
    return map;
  }
  details.forEach(function(detail) {
    if (!detail) {
      return;
    }
    var featureId = normalizeMetaId(detail.feature);
    if (!featureId) {
      return;
    }
    map[featureId] = clampReportFeatureQuantityValue(featureId, detail.quantity);
  });
  return map;
}

function buildActiveFeatureInstanceKeysFromMap(quantityMap) {
  var keys = [];
  if (!quantityMap) {
    return keys;
  }
  REPORT_FEATURE_SEQUENCE.forEach(function(featureId) {
    if (!featureId || !Object.prototype.hasOwnProperty.call(quantityMap, featureId)) {
      return;
    }
    var quantity = clampReportFeatureQuantityValue(featureId, quantityMap[featureId]);
    for (var idx = 0; idx < quantity; idx++) {
      keys.push(featureId + ':' + idx);
    }
  });
  return keys;
}

function parseFeatureOrderEntry(entry) {
  if (entry === null || entry === undefined) {
    return null;
  }
  if (typeof entry === 'string') {
    var parts = String(entry).split(':');
    if (parts.length !== 2) {
      return null;
    }
    var feature = normalizeMetaId(parts[0]);
    var index = parseInt(parts[1], 10);
    if (!feature || isNaN(index) || index < 0 || index >= REPORT_FEATURE_MAX_QUANTITY) {
      return null;
    }
    return { feature: feature, index: index };
  }
  var featureId = normalizeMetaId(entry.feature);
  if (!featureId) {
    return null;
  }
  var parsedIndex = parseInt(entry.index, 10);
  if (isNaN(parsedIndex) || parsedIndex < 0 || parsedIndex >= REPORT_FEATURE_MAX_QUANTITY) {
    return null;
  }
  return { feature: featureId, index: parsedIndex };
}

function normalizeFeatureOrderEntries(source, quantityMap) {
  var activeKeys = buildActiveFeatureInstanceKeysFromMap(quantityMap);
  var activeSet = activeKeys.reduce(function(map, key) {
    map[key] = true;
    return map;
  }, {});
  var normalizedKeys = [];
  if (Array.isArray(source)) {
    source.forEach(function(entry) {
      var parsed = parseFeatureOrderEntry(entry);
      if (!parsed) {
        return;
      }
      var key = parsed.feature + ':' + parsed.index;
      if (!activeSet[key] || normalizedKeys.indexOf(key) !== -1) {
        return;
      }
      normalizedKeys.push(key);
    });
  }
  activeKeys.forEach(function(key) {
    if (normalizedKeys.indexOf(key) === -1) {
      normalizedKeys.push(key);
    }
  });
  return normalizedKeys.map(function(key) {
    var parts = key.split(':');
    return { feature: parts[0], index: parseInt(parts[1], 10) };
  });
}

function normalizeSummaryLengthValue(value) {
  if (value === null || value === undefined) {
    return REPORT_SUMMARY_LENGTH_DEFAULT;
  }
  var text = String(value).trim().toLowerCase();
  if (!text) {
    return REPORT_SUMMARY_LENGTH_DEFAULT;
  }
  return REPORT_SUMMARY_LENGTH_ALLOWED[text] ? text : REPORT_SUMMARY_LENGTH_DEFAULT;
}

function normalizeOutputLanguageValue(value) {
  if (value === null || value === undefined) {
    return REPORT_OUTPUT_LANGUAGE_DEFAULT;
  }
  var text = String(value).trim().toLowerCase();
  if (!text) {
    return REPORT_OUTPUT_LANGUAGE_DEFAULT;
  }
  return REPORT_OUTPUT_LANGUAGE_ALLOWED[text]
    ? text
    : REPORT_OUTPUT_LANGUAGE_DEFAULT;
}

function normalizeToneValue(value) {
  if (value === null || value === undefined) {
    return REPORT_TONE_DEFAULT;
  }
  var text = String(value).trim().toLowerCase();
  if (!text) {
    return REPORT_TONE_DEFAULT;
  }
  return REPORT_TONE_ALLOWED[text] ? text : REPORT_TONE_DEFAULT;
}

function normalizeCustomTextValue(value) {
  if (value === null || value === undefined) {
    return '';
  }
  return String(value).trim();
}

function normalizePersonaValue(value) {
  if (value === null || value === undefined) {
    return REPORT_PERSONA_DEFAULT;
  }
  var text = String(value).trim().toLowerCase();
  if (!text) {
    return REPORT_PERSONA_DEFAULT;
  }
  return REPORT_PERSONA_ALLOWED[text] ? text : REPORT_PERSONA_DEFAULT;
}

function normalizeReportCustomization(source) {
  var base = {
    outputLanguage: {
      enabled: false,
      value: REPORT_OUTPUT_LANGUAGE_DEFAULT,
      customValue: ''
    },
    tone: {
      enabled: false,
      value: REPORT_TONE_DEFAULT,
      customValue: ''
    },
    persona: {
      enabled: false,
      value: REPORT_PERSONA_DEFAULT,
      customNote: ''
    }
  };
  if (!source || typeof source !== 'object') {
    return base;
  }
  if (source.outputLanguage && typeof source.outputLanguage === 'object') {
    base.outputLanguage.enabled = !!source.outputLanguage.enabled;
    base.outputLanguage.value = normalizeOutputLanguageValue(source.outputLanguage.value);
    base.outputLanguage.customValue = normalizeCustomTextValue(
      source.outputLanguage.customValue
    );
  }
  if (source.tone && typeof source.tone === 'object') {
    base.tone.enabled = !!source.tone.enabled;
    base.tone.value = normalizeToneValue(source.tone.value);
    base.tone.customValue = normalizeCustomTextValue(source.tone.customValue);
  }
  if (source.persona && typeof source.persona === 'object') {
    base.persona.enabled = !!source.persona.enabled;
    base.persona.value = normalizePersonaValue(source.persona.value);
    base.persona.customNote = normalizeCustomTextValue(source.persona.customNote);
  }
  if (!base.outputLanguage.enabled) {
    base.outputLanguage.value = REPORT_OUTPUT_LANGUAGE_DEFAULT;
    base.outputLanguage.customValue = '';
  }
  if (!base.tone.enabled) {
    base.tone.value = REPORT_TONE_DEFAULT;
    base.tone.customValue = '';
  }
  if (!base.persona.enabled) {
    base.persona.value = REPORT_PERSONA_DEFAULT;
    base.persona.customNote = '';
  }
  if (base.outputLanguage.value !== 'other') {
    base.outputLanguage.customValue = '';
  }
  if (base.tone.value !== 'custom') {
    base.tone.customValue = '';
  }
  if (base.persona.value !== 'custom') {
    base.persona.customNote = '';
  }
  return base;
}

function saveReportFavorite(payload) {
  if (!payload || typeof payload !== 'object') {
    throw new Error('No se recibieron datos del reporte.');
  }
  var missing = [];
  var name = payload.name ? String(payload.name).trim() : '';
  if (!name) {
    missing.push('nombre del reporte favorito');
  }
  var fileName = sanitizeReportFileName(payload.fileName);
  if (!fileName) {
    missing.push('nombre del archivo');
  }
  var format = normalizeMetaId(payload.format);
  if (!format) {
    missing.push('formato del reporte');
  }
  var features = Array.isArray(payload.features) ? payload.features : [];
  var seenFeatures = {};
  var normalizedFeatures = [];
  features.forEach(function(value) {
    var normalized = normalizeMetaId(value);
    if (normalized && !seenFeatures[normalized]) {
      seenFeatures[normalized] = true;
      normalizedFeatures.push(normalized);
    }
  });
  if (normalizedFeatures.length === 0) {
    missing.push('al menos una característica configurada');
  }
  var featureLabelMap = {
    chart: 'Gráfico',
    table: 'Tabla',
    list: 'Lista',
    summary: 'Resumen'
  };
  var tables = Array.isArray(payload.tables) ? payload.tables : [];
  var seenTables = {};
  var normalizedTables = [];
  tables.forEach(function(value) {
    var normalized = normalizeMetaId(value);
    if (normalized && !seenTables[normalized]) {
      seenTables[normalized] = true;
      normalizedTables.push(normalized);
    }
  });
  var featureDetails = Array.isArray(payload.featureDetails) ? payload.featureDetails : [];
  var normalizedFeatureDetails = [];
  var seenFeatureDetails = {};
  var missingFeatureDescriptions = [];
  var missingFeatureAssociations = [];
  var normalizedTableSet = {};
  normalizedTables.forEach(function(tableId) {
    normalizedTableSet[tableId] = true;
  });
  featureDetails.forEach(function(entry) {
    if (!entry) {
      return;
    }
    var featureId = normalizeMetaId(entry.feature);
    if (!featureId) {
      return;
    }
    if (seenFeatureDetails[featureId]) {
      return;
    }
    if (!seenFeatures[featureId] && normalizedFeatures.indexOf(featureId) === -1) {
      return;
    }
    var quantityValue = clampReportFeatureQuantityValue(featureId, entry.quantity);
    var rawInstances = Array.isArray(entry.instanceDescriptions) ? entry.instanceDescriptions : [];
    var normalizedInstances = ['', '', ''];
    rawInstances.slice(0, 3).forEach(function(value, index) {
      normalizedInstances[index] = value !== null && value !== undefined ? String(value).trim() : '';
    });
    var rawTitles = Array.isArray(entry.instanceTitles) ? entry.instanceTitles : [];
    var normalizedTitles = normalizeFeatureTitleArray(rawTitles);
    var rawChartTypes = Array.isArray(entry.instanceChartTypes) ? entry.instanceChartTypes : [];
    var normalizedChartTypes = normalizeFeatureChartTypeArray(rawChartTypes);
    var rawListOrders = Array.isArray(entry.instanceListOrders) ? entry.instanceListOrders : [];
    var normalizedListOrders = normalizeFeatureListOrderArray(rawListOrders);
    var rawListEnumerate = Array.isArray(entry.instanceListEnumerate) ? entry.instanceListEnumerate : [];
    var normalizedListEnumerate = normalizeFeatureListEnumerateArray(rawListEnumerate);
    var rawListEnumStyles = Array.isArray(entry.instanceListEnumStyles) ? entry.instanceListEnumStyles : [];
    var normalizedListEnumStyles = normalizeFeatureListEnumStyleArray(rawListEnumStyles);
    var summaryLengthValue = normalizeSummaryLengthValue(entry.summaryLength);
    var normalizedTableToggles = normalizeFeatureTableToggleArray(entry.instanceTablesEnabled);
    var normalizedTableSelections = normalizeFeatureTableSelectionArray(
      entry.instanceTables,
      normalizedTableSet
    );
    for (var idx = 0; idx < quantityValue; idx++) {
      if (!normalizedInstances[idx]) {
        var descLabel = 'Descripción ' + (idx + 1) + ' para ' + (featureLabelMap[featureId] || featureId);
        if (missingFeatureDescriptions.indexOf(descLabel) === -1) {
          missingFeatureDescriptions.push(descLabel);
        }
      }
      if (normalizedTableToggles[idx] && normalizedTableSelections[idx].length === 0) {
        var assocLabel = 'Tablas asociadas para ' + (featureLabelMap[featureId] || featureId) + ' #' + (idx + 1);
        if (missingFeatureAssociations.indexOf(assocLabel) === -1) {
          missingFeatureAssociations.push(assocLabel);
        }
      }
    }
    normalizedFeatureDetails.push({
      feature: featureId,
      quantity: quantityValue,
      instanceDescriptions: normalizedInstances,
      instanceTitles: normalizedTitles,
      instanceChartTypes: normalizedChartTypes,
      instanceListOrders: normalizedListOrders,
      instanceListEnumerate: normalizedListEnumerate,
      instanceListEnumStyles: normalizedListEnumStyles,
      summaryLength: summaryLengthValue,
      instanceTablesEnabled: normalizedTableToggles,
      instanceTables: normalizedTableSelections
    });
    seenFeatureDetails[featureId] = true;
  });
  normalizedFeatures.forEach(function(featureId) {
    if (!seenFeatureDetails[featureId]) {
      var missingLabel = featureLabelMap[featureId] || featureId;
      var message = 'Descripciones para ' + missingLabel;
      if (missingFeatureDescriptions.indexOf(message) === -1) {
        missingFeatureDescriptions.push(message);
      }
      normalizedFeatureDetails.push({
        feature: featureId,
        quantity: getReportFeatureDefaultQuantity(featureId),
        instanceDescriptions: ['', '', ''],
        instanceTitles: createEmptyFeatureTitleArray(),
        instanceChartTypes: normalizeFeatureChartTypeArray([]),
        instanceListOrders: normalizeFeatureListOrderArray([]),
        instanceListEnumerate: normalizeFeatureListEnumerateArray([]),
        instanceListEnumStyles: normalizeFeatureListEnumStyleArray([]),
        summaryLength: REPORT_SUMMARY_LENGTH_DEFAULT,
        instanceTablesEnabled: createEmptyFeatureTableToggleArray(),
        instanceTables: createEmptyFeatureTableSelectionArray()
      });
      seenFeatureDetails[featureId] = true;
    }
  });
  var channels = Array.isArray(payload.channels) ? payload.channels : [];
  var seenChannels = {};
  var normalizedChannels = [];
  channels.forEach(function(value) {
    var normalized = normalizeMetaId(value);
    if (normalized && !seenChannels[normalized]) {
      seenChannels[normalized] = true;
      normalizedChannels.push(normalized);
    }
  });
  if (normalizedChannels.length === 0) {
    missing.push('al menos un canal para compartir');
  }
  var customization = normalizeReportCustomization(payload.customization);
  if (
    customization.outputLanguage.enabled &&
    customization.outputLanguage.value === 'other' &&
    !customization.outputLanguage.customValue
  ) {
    missing.push('el idioma personalizado del reporte');
  }
  if (customization.tone.enabled && customization.tone.value === 'custom' && !customization.tone.customValue) {
    missing.push('el tono personalizado del reporte');
  }
  if (missingFeatureDescriptions.length > 0) {
    missingFeatureDescriptions.forEach(function(message) {
      if (missing.indexOf(message) === -1) {
        missing.push(message);
      }
    });
  }
  if (missingFeatureAssociations.length > 0) {
    missingFeatureAssociations.forEach(function(message) {
      if (missing.indexOf(message) === -1) {
        missing.push(message);
      }
    });
  }
  if (missing.length > 0) {
    throw new Error('Faltan datos para guardar el reporte favorito:\n- ' + missing.join('\n- '));
  }
  if (name) {
    var duplicateFavorite = findMetaByName(name, { type: 'reportFavorite' });
    if (duplicateFavorite) {
      throw new Error('Ya existe un reporte favorito con ese nombre. Elige otro diferente.');
    }
  }
  if (normalizedTables.length === 0) {
    throw new Error('Debe asociar al menos una tabla al reporte favorito.');
  }
  var emails = Array.isArray(payload.emails) ? payload.emails : [];
  var seenEmails = {};
  var normalizedEmails = [];
  emails.forEach(function(value) {
    var text = value === null || value === undefined ? '' : String(value).trim();
    if (text && !seenEmails[text]) {
      seenEmails[text] = true;
      normalizedEmails.push(text);
    }
  });
  var phones = Array.isArray(payload.phones) ? payload.phones : [];
  var seenPhones = {};
  var normalizedPhones = [];
  phones.forEach(function(value) {
    var text = value === null || value === undefined ? '' : String(value).trim();
    if (text && !seenPhones[text]) {
      seenPhones[text] = true;
      normalizedPhones.push(text);
    }
  });
  var quantityMapForOrder = buildFeatureQuantityMap(normalizedFeatureDetails);
  var normalizedFeatureOrder = normalizeFeatureOrderEntries(payload.featureOrder, quantityMapForOrder);
  var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
  var config = {
    format: format,
    fileName: fileName,
    fileExtension: payload.fileExtension ? String(payload.fileExtension).trim() : '',
    appendDateToFile: !!payload.appendDateToFile,
    tables: normalizedTables,
    channels: normalizedChannels,
    descriptionEnabled: !!payload.descriptionEnabled,
    description: payload.description ? String(payload.description) : '',
    features: normalizedFeatures,
    featureDetails: normalizedFeatureDetails,
    featureOrder: normalizedFeatureOrder,
    featureDescriptionsEnabled: false,
    featureDescriptions: '',
    emails: normalizedEmails,
    phones: normalizedPhones,
    customization: customization,
    savedAt: now
  };
  var favoriteId = 'reportFavorite:' + Utilities.getUuid();
  var sheet = getMetaSheet();
  sheet.appendRow([
    favoriteId,
    name,
    '',
    config.descriptionEnabled ? config.description : '',
    '',
    '',
    '',
    now,
    now,
    '',
    '',
    '',
    'reportFavorite',
    stringifyJsonValue(config)
  ]);
  SpreadsheetApp.flush();
  return { ok: true, id: favoriteId };
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
function askQuestion(tableSelection, question, customization) {
  var apiKey = PropertiesService.getUserProperties().getProperty('TC_API_KEY');
  if (!apiKey) {
    return { error: 'No hay API Key configurada. Configure su clave en la sección de configuración.' };
  }

  var customizationData = customization && typeof customization === 'object' ? customization : {};
  var toneValue = normalizeToneValue(customizationData.tone);
  var personaValue = normalizePersonaValue(customizationData.persona);
  var toneCustom = normalizeCustomTextValue(customizationData.toneCustom);
  var personaCustom = normalizeCustomTextValue(customizationData.personaCustom);
  if (toneValue !== 'custom') {
    toneCustom = '';
  }
  if (personaValue !== 'custom') {
    personaCustom = '';
  }

  var normalizedIds = normalizeTableSelection(tableSelection);
  if (normalizedIds.length === 0) {
    return { error: 'Seleccione al menos una tabla válida para consultar.' };
  }

  var normalizedQuestion = '';
  if (question !== null && question !== undefined) {
    normalizedQuestion = String(question).trim();
  }
  if (!normalizedQuestion) {
    return { error: 'Escriba una pregunta.' };
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
    var rangeParts = splitRangeNotation(data[META_INDEX.rangeA1]);
    var rangeA1 = rangeParts.range;
    var sheetName = data[META_INDEX.sheet] || rangeParts.sheet;
    if (!rangeA1 || !sheetName) {
      continue;
    }

    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      continue;
    }

    var range = getRangeWithinSheet(sheet, data[META_INDEX.rangeA1] || rangeA1);
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
  var baseInstruction =
    'Eres un asistente experto en análisis de datos de Google Sheets. Responde de forma breve y clara en español.';
  var systemParts = [baseInstruction];
  var toneInstruction = '';
  if (toneValue === 'custom' && toneCustom) {
    toneInstruction = 'Adapta el tono de la respuesta según esta indicación: ' + toneCustom + '.';
  } else if (ASK_TONE_INSTRUCTIONS[toneValue]) {
    toneInstruction = ASK_TONE_INSTRUCTIONS[toneValue];
  }
  if (toneInstruction) {
    systemParts.push(toneInstruction);
  }
  var personaInstruction = '';
  if (personaValue === 'custom' && personaCustom) {
    personaInstruction = 'Enfoca la respuesta para este perfil profesional: ' + personaCustom + '.';
  } else if (ASK_PERSONA_INSTRUCTIONS[personaValue]) {
    personaInstruction = ASK_PERSONA_INSTRUCTIONS[personaValue];
  }
  if (personaInstruction) {
    systemParts.push(personaInstruction);
  }
  var systemContent = systemParts.join(' ');
  var messages = [
    {
      role: 'system',
      content: systemContent
    },
    { role: 'user', content: context + '\n\nPregunta: ' + normalizedQuestion }
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
  if (normalizeMetaRecordType(meta.data[META_INDEX.recordType]) !== 'table') {
    return null;
  }
  var d = meta.data;
  var style = parseJsonValue(d[META_INDEX.style], null);
  var headers = normalizeHeaderArray(parseJsonValue(d[META_INDEX.headers], null));
  var rangeParts = splitRangeNotation(d[META_INDEX.rangeA1]);
  var sheetName = d[META_INDEX.sheet] || rangeParts.sheet;
  var pureRange = rangeParts.range || '';
  return {
    id: normalizeMetaId(d[META_INDEX.id]),
    name: d[META_INDEX.name],
    rangeA1: pureRange,
    fullRangeA1: buildFullRangeNotation(sheetName, pureRange),
    description: d[META_INDEX.description],
    sheetName: sheetName,
    cols: d[META_INDEX.cols],
    rows: d[META_INDEX.rows],
    createdAt: d[META_INDEX.createdAt],
    updatedAt: d[META_INDEX.updatedAt],
    style: style,
    headers: headers,
    updateFormulaReferences: parseBooleanValue(d[META_INDEX.formulaRefs], true)
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
  if (normalizeMetaRecordType(meta.data[META_INDEX.recordType]) !== 'table') {
    return { error: 'Tabla no encontrada' };
  }
  var d = meta.data;
  var storedHeaders = normalizeHeaderArray(parseJsonValue(d[META_INDEX.headers], null));
  if (storedHeaders && storedHeaders.length) {
    return { headers: storedHeaders };
  }
  var rangeParts = splitRangeNotation(d[META_INDEX.rangeA1]);
  var sheetName = d[META_INDEX.sheet] || rangeParts.sheet;
  var rangeA1 = rangeParts.range;
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { error: 'Hoja no encontrada' };
  var range = getRangeWithinSheet(sheet, d[META_INDEX.rangeA1] || rangeA1);
  var values = range.getValues();
  if (!values || values.length === 0) return { error: 'Rango vacío' };
  var headers = values[0].map(function(v) {
    return normalizeHeaderEntry(v === null ? '' : String(v));
  });
  return { headers: headers };
}
