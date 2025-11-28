// TestSuite.gs
var TestSuite = (function(){
  // Import LoggerService
  var LoggerService = LoggerService;
  // Mockable external dependencies
  var mocks = { UrlFetchApp: UrlFetchApp, SpreadsheetApp: SpreadsheetApp };

  /**
   * Inject mocks for external services like UrlFetchApp and SpreadsheetApp.
   * @param {Object} newMocks
   */
  function setMocks(newMocks) {
    mocks = Object.assign({}, mocks, newMocks);
  }

  var tests = [];
  var beforeAllHooks = [];
  var beforeEachHooks = [];
  var afterEachHooks = [];
  var afterAllHooks = [];
  var config = { tags: [], retryLimit: 0 };
  var results = [];

  var Assert = {
    equal: function(actual, expected, msg) {
      if (actual !== expected) throw new Error(msg || 'Expected ' + expected + ' but got ' + actual);
    },
    notEqual: function(actual, unexpected, msg) {
      if (actual === unexpected) throw new Error(msg || 'Did not expect ' + unexpected);
    },
    isTrue: function(val, msg) {
      if (val !== true) throw new Error(msg || 'Expected true but got ' + val);
    },
    isFalse: function(val, msg) {
      if (val !== false) throw new Error(msg || 'Expected false but got ' + val);
    },
    throws: function(fn, msg) {
      var threw = false;
      try {
        fn();
      } catch (e) {
        threw = true;
      }
      if (!threw) throw new Error(msg || 'Expected function to throw');
    },
    matches: function(actual, regex, msg) {
      if (!regex.test(actual)) throw new Error(msg || 'Expected ' + actual + ' to match ' + regex);
    }
  };

  function addTestCase(name, fn, tags) {
    tests.push({ name: name, fn: fn, tags: tags || [] });
  }

  function beforeAll(fn) {
    beforeAllHooks.push(fn);
  }

  function beforeEach(fn) {
    beforeEachHooks.push(fn);
  }

  function afterEach(fn) {
    afterEachHooks.push(fn);
  }

  function afterAll(fn) {
    afterAllHooks.push(fn);
  }

  function runAllTests(options) {
    options = options || {};
    results = [];
    var total = 0;
    var passed = 0;
    var failed = 0;
    config.tags = options.tags || [];
    config.retryLimit = options.retryLimit || 0;

    beforeAllHooks.forEach(function(h) { h(); });

    tests.filter(function(test) {
      if (config.tags.length === 0) return true;
      return test.tags.some(function(tag) { return config.tags.indexOf(tag) > -1; });
    }).forEach(function(test) {
      total++;
      beforeEachHooks.forEach(function(h) { h(); });
      var start = new Date();
      try {
        var attempts = 0;
        while (true) {
          try {
            test.fn();
            recordResult(test.name, 'PASS', start);
            passed++;
            break;
          } catch (e) {
            attempts++;
            if (attempts > config.retryLimit) {
              throw e;
            }
          }
        }
      } catch (e) {
        recordResult(test.name, 'FAIL', start, e);
        failed++;
      } finally {
        afterEachHooks.forEach(function(h) { h(); });
      }
    });

    afterAllHooks.forEach(function(h) { h(); });

    reportSummary(total, passed, failed);
  }

  function recordResult(name, status, startTime, error) {
    var duration = new Date() - startTime;
    var entry = {
      testName: name,
      status: status,
      durationMs: duration,
      timestamp: new Date()
    };
    if (error) {
      entry.errorMessage = error.message;
      entry.stack = error.stack;
    }
    results.push(entry);
    if (status === 'PASS') {
      LoggerService.logSuccess({ testName: entry.testName, status: entry.status, duration: entry.durationMs });
    } else {
      LoggerService.logError({ testName: entry.testName, status: entry.status, message: entry.errorMessage });
    }
  }

  function reportSummary(total, passed, failed) {
    var msg = 'Tests complete: ' + passed + '/' + total + ' passed, ' + failed + ' failed.';
    mocks.SpreadsheetApp.getActive().toast(msg, 'TestSuite Results', 5);
  }

  // Contract integrity check for API_URL
  beforeAll(function() {
    Assert.matches(CONSTANTS.API.API_URL, /^https?:\/\/.+/, 'API_URL must be valid');
  });

  // Example lifecycle hooks
  beforeAll(function() {
    // Setup before all tests
  });

  beforeEach(function() {
    // Setup before each test
  });

  afterEach(function() {
    // Cleanup after each test
  });

  afterAll(function() {
    // Cleanup after all tests
  });

  // Example test cases
  addTestCase('FetchService returns object', function() {
    var data = fetchData();
    Assert.isTrue(typeof data === 'object', 'fetchData should return an object');
  }, ['fetch']);

  addTestCase('Schema validation passes for fetched data', function() {
    var data = fetchData();
    var schema = EXPECTED_SCHEMA;
    for (var key in schema) {
      Assert.isTrue(data.hasOwnProperty(key), 'Missing key: ' + key);
      Assert.equal(typeof data[key], schema[key], 'Type mismatch for ' + key);
    }
  }, ['fetch']);

  addTestCase('SheetService clear and write data', function() {
    var sheetName = SHEET_NAMES.data;
    var sheet = mocks.SpreadsheetApp.getActive().getSheetByName(sheetName);
    sheet.clear();
    sheet.appendRow(['temp']);
    clearData();
    Assert.equal(sheet.getLastRow(), 0, 'Sheet should be empty after clear');
    var rows = [['a','b'],['c','d']];
    writeData(rows);
    Assert.equal(sheet.getLastRow(), 2, 'Sheet should have two rows after write');
  }, ['sheet']);

  return {
    addTestCase: addTestCase,
    beforeAll: beforeAll,
    beforeEach: beforeEach,
    afterEach: afterEach,
    afterAll: afterAll,
    runAllTests: runAllTests,
    setMocks: setMocks,
    assert: Assert
  };
})();

function runAllTests() {
  TestSuite.runAllTests();
}