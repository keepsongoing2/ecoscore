/**
 * @OnlyCurrentDoc
 */

/**
 * Adds the "EcoScore" custom menu to the active spreadsheet.
 */
function onOpen(e) {
  try {
    var ui = SpreadsheetApp.getUi();
    var ecoMenu = ui.createMenu('EcoScore');
    ecoMenu.addItem('Sync Data', 'syncData');
    // Register stable menu command ID for ecoSyncCmd
    EventEmitter.registerCommand('ecoSyncCmd', 'syncData');
    ecoMenu.addToUi();
  } catch (err) {
    console.error('onOpen error:', err);
    throw err;
  }
}

/**
 * Orchestrates the data synchronization workflow.
 */
function syncData() {
  // Emit syncStarted event
  eventEmitter.emit('syncStarted');
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var title = 'EcoScore';
  ss.toast('Starting data synchronization...', title, 5);
  try {
    var startTime = new Date();
    var data = fetchData();
    if (!Array.isArray(data)) {
      throw new Error('fetchData must return an array');
    }
    clearData();
    writeData(data);
    var duration = (new Date() - startTime) / 1000;
    PropertiesService.getScriptProperties().setProperty('LAST_SYNC', new Date().toISOString());
    logSuccess({ rows: data.length, duration: duration });
    // Emit syncSucceeded event
    eventEmitter.emit('syncSucceeded');
    ss.toast('Data synchronized successfully!', title, 5);
  } catch (err) {
    logError({ message: err.message, stack: err.stack });
    // Emit syncFailed event
    eventEmitter.emit('syncFailed', err);
    ss.toast('Sync failed: ' + err.message, title, 5);
    console.error('syncData error:', err);
  }
}

/**
 * Handles GET HTTP requests. Placeholder for UI rendering.
 */
function doGet(e) {
  try {
    return HtmlService.createHtmlOutput(
      '<p>EcoScore Add-on</p>' +
      '<div id="data-preview" aria-label="Data Preview"></div>' +
      '<div id="log-view" aria-label="Log View"></div>'
    );
  } catch (err) {
    console.error('doGet error:', err);
    return HtmlService.createHtmlOutput('Error: ' + err.message);
  }
}

/**
 * Handles POST HTTP requests. Echoes received JSON payload.
 *
 * @param {object} e Event parameter.
 */
function doPost(e) {
  try {
    var content = e.postData && e.postData.contents;
    if (!content) {
      throw new Error('Missing postData contents');
    }
    var payload = JSON.parse(content);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', data: payload }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    console.error('doPost error:', err);
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Retrieves a value from the script cache.
 *
 * @param {string} key The cache key.
 * @return {*} The parsed cached value, or null.
 */
function getCachedData(key) {
  var cache = CacheService.getScriptCache();
  var value = cache.get(key);
  return value ? JSON.parse(value) : null;
}

/**
 * Stores a value in the script cache.
 *
 * @param {string} key The cache key.
 * @param {*} value The value to cache.
 * @param {number} expirationSeconds Time to live in seconds.
 */
function setCachedData(key, value, expirationSeconds) {
  var cache = CacheService.getScriptCache();
  cache.put(key, JSON.stringify(value), expirationSeconds);
}