var SheetService = (function() {
  var cache = {};

  /**
   * Retrieves a sheet by name, with caching.
   * @param {string} sheetName
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   */
  function getSheet(sheetName) {
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('getSheet: sheetName must be a non-empty string');
    }
    if (cache[sheetName]) {
      return cache[sheetName];
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error('getSheet: No active spreadsheet found');
    }
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('getSheet: Sheet "' + sheetName + '" not found');
    }
    cache[sheetName] = sheet;
    return sheet;
  }

  /**
   * Validates that a sheet exists by name.
   * @param {string} sheetName
   * @returns {GoogleAppsScript.Spreadsheet.Sheet}
   * @throws {Error} if the sheet does not exist
   * @private
   */
  function validateSheetExists(sheetName) {
    return getSheet(sheetName);
  }

  /**
   * Reads values from a specified range.
   * @param {string} rangeA1
   * @param {string=} sheetName
   * @returns {Array<Array<any>>}
   */
  function readRange(rangeA1, sheetName) {
    if (!rangeA1 || typeof rangeA1 !== 'string') {
      throw new Error('readRange: rangeA1 must be a non-empty string');
    }
    sheetName = sheetName || constants.SHEET_NAMES.data;
    try {
      var sheet = getSheet(sheetName);
      return sheet.getRange(rangeA1).getValues();
    } catch (e) {
      LoggerService.logError({ action: 'readRange', sheet: sheetName, range: rangeA1, error: e.message });
      throw e;
    }
  }

  /**
   * Writes values to a specified range, overwriting existing contents.
   * @param {string} rangeA1
   * @param {Array<Array<any>>} values
   * @param {string=} sheetName
   */
  function writeRange(rangeA1, values, sheetName) {
    if (!rangeA1 || typeof rangeA1 !== 'string') {
      throw new Error('writeRange: rangeA1 must be a non-empty string');
    }
    if (!Array.isArray(values) || !Array.isArray(values[0])) {
      throw new Error('writeRange: values must be a 2D array');
    }
    sheetName = sheetName || constants.SHEET_NAMES.data;
    try {
      var sheet = getSheet(sheetName);
      var range = sheet.getRange(rangeA1);
      var numRows = range.getNumRows();
      var numCols = range.getNumColumns();
      if (values.length !== numRows || values[0].length !== numCols) {
        throw new Error('writeRange: dimension mismatch between range and values');
      }
      range.setValues(values);
      LoggerService.logSuccess({ action: 'writeRange', sheet: sheetName, range: rangeA1, rows: numRows, cols: numCols });
    } catch (e) {
      LoggerService.logError({ action: 'writeRange', sheet: sheetName, range: rangeA1, error: e.message });
      throw e;
    }
  }

  /**
   * Appends rows of data below the last non-empty row.
   * @param {Array<Array<any>>} rows
   * @param {string=} sheetName
   */
  function appendRows(rows, sheetName) {
    if (!Array.isArray(rows) || rows.length === 0) {
      throw new Error('appendRows: rows must be a non-empty 2D array');
    }
    if (!Array.isArray(rows[0])) {
      throw new Error('appendRows: rows must be a 2D array');
    }
    sheetName = sheetName || constants.SHEET_NAMES.data;
    try {
      var sheet = getSheet(sheetName);
      var lastRow = sheet.getLastRow();
      var startRow = lastRow + 1;
      var numCols = rows[0].length;
      sheet.getRange(startRow, 1, rows.length, numCols).setValues(rows);
      LoggerService.logSuccess({ action: 'appendRows', sheet: sheetName, rowsAdded: rows.length });
    } catch (e) {
      LoggerService.logError({ action: 'appendRows', sheet: sheetName, error: e.message });
      throw e;
    }
  }

  /**
   * Clears contents of a specified range.
   * @param {string} rangeA1
   * @param {string=} sheetName
   */
  function clearRange(rangeA1, sheetName) {
    if (!rangeA1 || typeof rangeA1 !== 'string') {
      throw new Error('clearRange: rangeA1 must be a non-empty string');
    }
    sheetName = sheetName || constants.SHEET_NAMES.data;
    try {
      var sheet = getSheet(sheetName);
      sheet.getRange(rangeA1).clearContents();
      LoggerService.logSuccess({ action: 'clearRange', sheet: sheetName, range: rangeA1 });
    } catch (e) {
      LoggerService.logError({ action: 'clearRange', sheet: sheetName, range: rangeA1, error: e.message });
      throw e;
    }
  }

  /**
   * Clears all data from the default data sheet.
   * @public
   * @returns {void}
   */
  function clearData() {
    var sheetName = constants.SHEET_NAMES.data;
    try {
      var sheet = getSheet(sheetName);
      sheet.clearContents();
      LoggerService.logSuccess({ action: 'clearData', sheet: sheetName });
    } catch (e) {
      LoggerService.logError({ action: 'clearData', sheet: sheetName, error: e.message });
      throw e;
    }
  }

  /**
   * Clears existing data and writes new rows to the default data sheet.
   * @public
   * @param {Array<Array<any>>} rows The data rows to write.
   * @returns {void}
   */
  function writeData(rows) {
    if (!Array.isArray(rows)) {
      throw new Error('writeData: rows must be a 2D array');
    }
    var sheetName = constants.SHEET_NAMES.data;
    try {
      clearData();
      appendRows(rows, sheetName);
      LoggerService.logSuccess({ action: 'writeData', sheet: sheetName, rowsWritten: rows.length });
    } catch (e) {
      LoggerService.logError({ action: 'writeData', sheet: sheetName, error: e.message });
      throw e;
    }
  }

  return {
    clearData: clearData,
    writeData: writeData
  };
})();