function fetchData() {
  if (!SpreadsheetApp.getActiveSpreadsheet()) {
    throw new Error('FetchService.fetchData must be run in a container-bound context');
  }
  try {
    var response = makeRequest_();
    var data = parseResponse_(response);
    validateSchema_(data);
    return data;
  } catch (e) {
    logError({ step: 'fetchData', message: e.message, stack: e.stack });
    throw new Error('FetchService.fetchData failed: ' + e.message);
  }
}

function makeRequest_() {
  if (!/^https?:\/\/.+/.test(API_URL)) {
    throw new Error('Invalid API_URL: ' + API_URL);
  }
  var attempts = 3;
  var delay = 1000;
  for (var attempt = 1; attempt <= attempts; attempt++) {
    try {
      var options = {
        method: 'get',
        headers: buildHeaders_(),
        muteHttpExceptions: true,
        timeout: TIMEOUT_MS
      };
      var response = UrlFetchApp.fetch(API_URL, options);
      var code = response.getResponseCode();
      if ((code >= 500 && code < 600) || code === 429) {
        throw new Error('HTTP Error ' + code);
      }
      return response;
    } catch (e) {
      if (attempt === attempts) {
        throw e;
      }
      Utilities.sleep(delay);
      delay *= 2;
    }
  }
}

function buildHeaders_() {
  return {
    Authorization: 'Bearer ' + getAccessToken_(),
    'Content-Type': 'application/json'
  };
}

function getAccessToken_() {
  return ScriptApp.getOAuthToken();
}

function parseResponse_(response) {
  var code = response.getResponseCode();
  var text = response.getContentText();
  if (code >= 200 && code < 300) {
    try {
      return JSON.parse(text);
    } catch (e) {
      throw new Error('Invalid JSON response: ' + e.message);
    }
  }
  throw new Error('HTTP Error ' + code + ': ' + text);
}

function validateSchema_(data) {
  for (var key in EXPECTED_SCHEMA) {
    if (!data.hasOwnProperty(key)) {
      throw new Error('Missing expected key in response: ' + key);
    }
    var expectedType = EXPECTED_SCHEMA[key];
    var value = data[key];
    var actualType = Array.isArray(value) ? 'array' : typeof value;
    if (actualType !== expectedType) {
      throw new Error('Type mismatch for key "' + key + '": expected ' + expectedType + ' but got ' + actualType);
    }
  }
}