/**
 * @file Constants.gs
 * @description Immutable configuration constants for EcoScore project.
 * @version 1.0.0
 * @changelog
 *   - 1.0.0: Initial release of configuration constants.
 */

/**
 * Centralized, immutable configuration constants.
 * @namespace CONSTANTS
 */
const CONSTANTS = Object.freeze({
  /** @section API Configuration
   * Base endpoint and settings for UrlFetchApp calls.
   */
  API: Object.freeze({
    /** Base URL for EcoScore data fetch API.
     * @type {string}
     * @example
     *   UrlFetchApp.fetch(CONSTANTS.API.API_URL, { timeout: CONSTANTS.API.TIMEOUT_MS });
     */
    API_URL: 'https://api.ecoscore.example.com/v1/scores',
    /** Request timeout in milliseconds.
     * @type {number}
     * @default 30000
     */
    TIMEOUT_MS: 30000,
    /** Default HTTP headers for API requests.
     * @type {Object<string, string>}
     */
    HEADERS: Object.freeze({
      'Accept': 'application/json'
    })
  }),

  /** @section Sheet Names
   * Worksheet names used by the SheetService.
   * @type {{data: string, log: string}}
   */
  SHEET_NAMES: Object.freeze({
    /** Name of the data sheet.
     * @type {string}
     */
    data: 'Data',
    /** Name of the log sheet.
     * @type {string}
     */
    log: 'Log'
  }),

  /** @section Expected JSON Schema
   * Defines expected keys and their JavaScript types.
   * @type {Object<string, string>}
   */
  EXPECTED_SCHEMA: Object.freeze({
    id: 'number',
    name: 'string',
    score: 'number',
    timestamp: 'string'
  }),
});

/**
 * Validates core constant formats and presence.
 * Throws an error if any validation fails.
 * @private
 */
(function validateConstants() {
  if (!/^https?:\/\/\S+/.test(CONSTANTS.API.API_URL)) {
    throw new Error(`Invalid API_URL: ${CONSTANTS.API.API_URL}`);
  }
  if (typeof CONSTANTS.API.TIMEOUT_MS !== 'number' || CONSTANTS.API.TIMEOUT_MS <= 0) {
    throw new Error(`Invalid TIMEOUT_MS: ${CONSTANTS.API.TIMEOUT_MS}`);
  }
  ['data', 'log'].forEach(key => {
    if (!CONSTANTS.SHEET_NAMES[key] || typeof CONSTANTS.SHEET_NAMES[key] !== 'string') {
      throw new Error(`Invalid sheet name for ${key}`);
    }
  });
  for (const [field, type] of Object.entries(CONSTANTS.EXPECTED_SCHEMA)) {
    if (typeof field !== 'string' || typeof type !== 'string') {
      throw new Error(`Invalid schema entry: ${field}: ${type}`);
    }
  }
})();

/**
 * Unit tests for constant values.
 * Logs assertions to the execution log.
 */
function testConstants() {
  console.assert(typeof CONSTANTS.API.API_URL === 'string', 'API_URL must be a string');
  console.assert(CONSTANTS.API.API_URL.startsWith('http'), 'API_URL must start with http');
  console.assert(typeof CONSTANTS.API.TIMEOUT_MS === 'number', 'TIMEOUT_MS must be a number');
  console.assert(CONSTANTS.API.TIMEOUT_MS > 0, 'TIMEOUT_MS must be > 0');
  console.assert(typeof CONSTANTS.SHEET_NAMES.data === 'string', 'SHEET_NAMES.data must be a string');
  console.assert(typeof CONSTANTS.SHEET_NAMES.log === 'string', 'SHEET_NAMES.log must be a string');
  console.assert(typeof CONSTANTS.EXPECTED_SCHEMA === 'object', 'EXPECTED_SCHEMA must be an object');
}