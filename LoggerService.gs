var LoggerService = (function() {
  'use strict';
  var LOG_LEVELS = { DEBUG: 0, INFO: 1, WARN: 2, ERROR: 3 };
  var DEFAULT_LEVEL = 'INFO';
  var config = {
    level: PropertiesService.getScriptProperties().getProperty('LOG_LEVEL') || DEFAULT_LEVEL,
    sheetName: constants.SHEET_NAMES.log
  };

  function getLevelIndex(level) {
    return LOG_LEVELS.hasOwnProperty(level) ? LOG_LEVELS[level] : LOG_LEVELS[DEFAULT_LEVEL];
  }

  function shouldLog(level) {
    return getLevelIndex(level) >= getLevelIndex(config.level);
  }

  function getLogSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(config.sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(config.sheetName);
      sheet.appendRow(['Timestamp', 'Level', 'Message', 'Details']);
    }
    return sheet;
  }

  function formatDetails(details) {
    try {
      return details != null ? JSON.stringify(details) : '';
    } catch (e) {
      return String(details);
    }
  }

  function log(level, message, details) {
    try {
      if (!shouldLog(level)) return;
      var sheet = getLogSheet();
      sheet.appendRow([new Date(), level, message, formatDetails(details)]);
    } catch (e) {
      console.error('LoggerService error logging:', e);
      // Fallback behavior: logging failed. Could retry or write to alternative storage here.
    }
  }

  /**
   * Logs an error-level message or Error object to the configured log sheet.
   *
   * @param {string|object} msgOrDetails The error message string, or an object with a message property.
   * @param {object} [details] Optional additional details to include in the log entry.
   */
  function logError(msgOrDetails, details) {
    if (typeof msgOrDetails === 'string') {
      log('ERROR', msgOrDetails, details);
    } else {
      var msg = msgOrDetails && msgOrDetails.message ? msgOrDetails.message : 'Error';
      log('ERROR', msg, details || msgOrDetails);
    }
  }

  /**
   * Logs a success event at INFO level to the configured log sheet.
   *
   * @param {object} details Additional details about the successful operation.
   */
  function logSuccess(details) {
    log('INFO', 'Operation succeeded', details);
  }

  return {
    logError: logError,
    logSuccess: logSuccess
  };
})();