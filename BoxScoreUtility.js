// ===== BOX SCORE UTILITY MODULE =====
// Shared helper functions for Box Score automation
// Used by all other modules

// ===== SCRIPT PROPERTIES MANAGEMENT =====

/**
 * Get a Script Property key with sheet ID suffix
 * @param {string} baseName - Base property name from config
 * @param {string} sheetId - Sheet ID to append
 * @return {string} Full property key
 */
function getPropertyKey(baseName, sheetId) {
  return baseName + "_" + sheetId;
}

/**
 * Get game state (shadow storage of at-bat grid)
 * @param {string} sheetId - Sheet ID
 * @return {Object} Game state object
 */
function getGameState(sheetId) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_GAME_STATE, sheetId);
  var json = properties.getProperty(key);
  return json ? JSON.parse(json) : {};
}

/**
 * Save game state
 * @param {string} sheetId - Sheet ID
 * @param {Object} gameState - Game state object
 */
function saveGameState(sheetId, gameState) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_GAME_STATE, sheetId);
  properties.setProperty(key, JSON.stringify(gameState));
}

/**
 * Get pitcher stats
 * @param {string} sheetId - Sheet ID
 * @return {Object} Pitcher stats object
 */
function getPitcherStats(sheetId) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_PITCHER_STATS, sheetId);
  var json = properties.getProperty(key);
  return json ? JSON.parse(json) : {};
}

/**
 * Save pitcher stats
 * @param {string} sheetId - Sheet ID
 * @param {Object} pitcherStats - Pitcher stats object
 */
function savePitcherStats(sheetId, pitcherStats) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_PITCHER_STATS, sheetId);
  properties.setProperty(key, JSON.stringify(pitcherStats));
}

/**
 * Get batter stats
 * @param {string} sheetId - Sheet ID
 * @return {Object} Batter stats object
 */
function getBatterStats(sheetId) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_BATTER_STATS, sheetId);
  var json = properties.getProperty(key);
  return json ? JSON.parse(json) : {};
}

/**
 * Save batter stats
 * @param {string} sheetId - Sheet ID
 * @param {Object} batterStats - Batter stats object
 */
function saveBatterStats(sheetId, batterStats) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_BATTER_STATS, sheetId);
  properties.setProperty(key, JSON.stringify(batterStats));
}

/**
 * Get active pitchers
 * @param {string} sheetId - Sheet ID
 * @return {Object} Active pitchers {away: string, home: string}
 */
function getActivePitchers(sheetId) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_ACTIVE_PITCHERS, sheetId);
  var json = properties.getProperty(key);
  return json ? JSON.parse(json) : {away: "", home: ""};
}

/**
 * Save active pitchers
 * @param {string} sheetId - Sheet ID
 * @param {Object} activePitchers - Active pitchers object
 */
function saveActivePitchers(sheetId, activePitchers) {
  var properties = PropertiesService.getScriptProperties();
  var key = getPropertyKey(BOX_SCORE_CONFIG.PROPERTY_ACTIVE_PITCHERS, sheetId);
  properties.setProperty(key, JSON.stringify(activePitchers));
}

/**
 * Clear all properties for a sheet (used by reset)
 * Includes defensive stat associations to prevent storage leak
 * @param {string} sheetId - Sheet ID
 */
function clearAllProperties(sheetId) {
  var properties = PropertiesService.getScriptProperties();
  
  // Clear main stat storage
  var keys = [
    BOX_SCORE_CONFIG.PROPERTY_GAME_STATE,
    BOX_SCORE_CONFIG.PROPERTY_PITCHER_STATS,
    BOX_SCORE_CONFIG.PROPERTY_BATTER_STATS,
    BOX_SCORE_CONFIG.PROPERTY_ACTIVE_PITCHERS
  ];
  
  for (var i = 0; i < keys.length; i++) {
    var key = getPropertyKey(keys[i], sheetId);
    properties.deleteProperty(key);
  }
  
  // Clear defensive stat associations (NP and E)
  // These have keys like "NP_123456_7", "E_123456_10"
  clearDefensiveAssociations(sheetId);
  
  logInfo("Utility", "Cleared all properties for sheet: " + sheetId);
}

/**
 * Clear all defensive stat associations for a sheet
 * Prevents Properties Service storage leak
 * @param {string} sheetId - Sheet ID
 */
function clearDefensiveAssociations(sheetId) {
  var properties = PropertiesService.getScriptProperties();
  var allProperties = properties.getProperties();
  
  var npPrefix = BOX_SCORE_CONFIG.PROPERTY_PREFIX_NP + '_' + sheetId + '_';
  var ePrefix = BOX_SCORE_CONFIG.PROPERTY_PREFIX_E + '_' + sheetId + '_';
  
  var deletedCount = 0;
  
  // Find and delete all defensive associations for this sheet
  for (var key in allProperties) {
    if (key.indexOf(npPrefix) === 0 || key.indexOf(ePrefix) === 0) {
      properties.deleteProperty(key);
      deletedCount++;
    }
  }
  
  if (deletedCount > 0) {
    logInfo("Utility", "Cleared " + deletedCount + " defensive associations for sheet: " + sheetId);
  }
}

// ===== SHEET OPERATIONS =====

/**
 * Clear pitcher and defensive stats in sheet (skip protected rows)
 * @param {Sheet} sheet - The game sheet
 */
function clearPitcherStatsInSheet(sheet) {
  var awayRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  var pitcherCols = BOX_SCORE_CONFIG.PITCHER_STATS_COLUMNS;
  var fieldingCols = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS;
  
  // Get column range for all stats (pitcher + fielding)
  var firstCol = Math.min(
    pitcherCols.BF, pitcherCols.IP, pitcherCols.H, pitcherCols.HR, 
    pitcherCols.R, pitcherCols.BB, pitcherCols.K,
    fieldingCols.NP, fieldingCols.E, fieldingCols.SB
  );
  var lastCol = Math.max(
    pitcherCols.BF, pitcherCols.IP, pitcherCols.H, pitcherCols.HR, 
    pitcherCols.R, pitcherCols.BB, pitcherCols.K,
    fieldingCols.NP, fieldingCols.E, fieldingCols.SB
  );
  var numCols = lastCol - firstCol + 1;
  
  // Clear away pitcher/defensive stats
  for (var row = awayRange.startRow; row <= awayRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      sheet.getRange(row, firstCol, 1, numCols).setValue(0);
    }
  }
  
  // Clear home pitcher/defensive stats
  for (var row = homeRange.startRow; row <= homeRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      sheet.getRange(row, firstCol, 1, numCols).setValue(0);
    }
  }
}

/**
 * Clear hitting and stolen base stats in sheet (skip protected rows)
 * @param {Sheet} sheet - The game sheet
 */
function clearHittingStatsInSheet(sheet) {
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;
  var hittingCols = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS;
  var sbCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.SB;
  
  // Clear away hitting stats
  for (var row = hittingRange.awayStartRow; row <= hittingRange.awayEndRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      // Clear all hitting stat columns
      sheet.getRange(row, hittingCols.AB).setValue(0);
      sheet.getRange(row, hittingCols.H).setValue(0);
      sheet.getRange(row, hittingCols.HR).setValue(0);
      sheet.getRange(row, hittingCols.RBI).setValue(0);
      sheet.getRange(row, hittingCols.BB).setValue(0);
      sheet.getRange(row, hittingCols.K).setValue(0);
      sheet.getRange(row, hittingCols.ROB).setValue(0);
      sheet.getRange(row, hittingCols.DP).setValue(0);
      sheet.getRange(row, hittingCols.TB).setValue(0);
    }
  }
  
  // Clear home hitting stats
  for (var row = hittingRange.homeStartRow; row <= hittingRange.homeEndRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      // Clear all hitting stat columns
      sheet.getRange(row, hittingCols.AB).setValue(0);
      sheet.getRange(row, hittingCols.H).setValue(0);
      sheet.getRange(row, hittingCols.HR).setValue(0);
      sheet.getRange(row, hittingCols.RBI).setValue(0);
      sheet.getRange(row, hittingCols.BB).setValue(0);
      sheet.getRange(row, hittingCols.K).setValue(0);
      sheet.getRange(row, hittingCols.ROB).setValue(0);
      sheet.getRange(row, hittingCols.DP).setValue(0);
      sheet.getRange(row, hittingCols.TB).setValue(0);
    }
  }
  
  // Clear SB from fielding section (rows 7-15, 18-26)
  var awayFieldingRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeFieldingRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  
  for (var row = awayFieldingRange.startRow; row <= awayFieldingRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      sheet.getRange(row, sbCol).setValue(0);
    }
  }
  
  for (var row = homeFieldingRange.startRow; row <= homeFieldingRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      sheet.getRange(row, sbCol).setValue(0);
    }
  }
}

/**
 * Clear at-bat grid (optional - removes all entered at-bats)
 * @param {Sheet} sheet - The game sheet
 */
function clearAtBatGrid(sheet) {
  var awayRange = BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;
  
  // Clear away at-bats
  var awayRows = awayRange.endRow - awayRange.startRow + 1;
  var awayCols = awayRange.endCol - awayRange.startCol + 1;
  sheet.getRange(awayRange.startRow, awayRange.startCol, awayRows, awayCols).clearContent();
  
  // Clear home at-bats
  var homeRows = homeRange.endRow - homeRange.startRow + 1;
  var homeCols = homeRange.endCol - homeRange.startCol + 1;
  sheet.getRange(homeRange.startRow, homeRange.startCol, homeRows, homeCols).clearContent();
}

/**
 * Clear pitcher dropdowns
 * @param {Sheet} sheet - The game sheet
 */
function clearPitcherDropdowns(sheet) {
  sheet.getRange(BOX_SCORE_CONFIG.AWAY_PITCHER_CELL).clearContent();
  sheet.getRange(BOX_SCORE_CONFIG.HOME_PITCHER_CELL).clearContent();
}

// ===== HELPER FUNCTIONS =====

/**
 * Check if cell is in at-bat range
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @return {boolean} True if in at-bat range
 */
function isAtBatCell(row, col) {
  var awayRange = BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;
  
  var isAwayAtBat = (row >= awayRange.startRow && row <= awayRange.endRow &&
                     col >= awayRange.startCol && col <= awayRange.endCol);
  
  var isHomeAtBat = (row >= homeRange.startRow && row <= homeRange.endRow &&
                     col >= homeRange.startCol && col <= homeRange.endCol);
  
  return isAwayAtBat || isHomeAtBat;
}

/**
 * Determine batting team from row
 * @param {number} row - Row number
 * @return {string} "away" or "home" or null
 */
function getBattingTeam(row) {
  var awayRange = BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;
  
  if (row >= awayRange.startRow && row <= awayRange.endRow) {
    return "away";
  }
  if (row >= homeRange.startRow && row <= homeRange.endRow) {
    return "home";
  }
  return null;
}

/**
 * Get batter row from at-bat cell
 * @param {number} row - At-bat cell row
 * @return {number} Hitting stats row number, or null
 */
function getBatterRowFromAtBatCell(row) {
  var awayRange = BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;
  
  // Away batters: rows 7-15 → batter positions 1-9 → hitting rows 30-38
  if (row >= awayRange.startRow && row <= awayRange.endRow) {
    var batterPosition = row - awayRange.startRow;
    return hittingRange.awayStartRow + batterPosition;
  }
  
  // Home batters: rows 18-26 → batter positions 1-9 → hitting rows 41-49
  if (row >= homeRange.startRow && row <= homeRange.endRow) {
    var batterPosition = row - homeRange.startRow;
    return hittingRange.homeStartRow + batterPosition;
  }
  
  return null;
}

/**
 * Get player name from batter row
 * @param {Sheet} sheet - The game sheet
 * @param {number} row - Batter row number
 * @return {string} Player name
 */
function getPlayerNameFromBatterRow(sheet, row) {
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;
  var name = sheet.getRange(row, hittingRange.nameCol).getValue();
  return String(name).trim();
}

/**
 * Initialize pitcher stats for a pitcher
 * @param {string} sheetId - Sheet ID
 * @param {string} pitcherName - Pitcher name
 */
function initializePitcherIfNeeded(sheetId, pitcherName) {
  if (!pitcherName || pitcherName === "") return;
  
  var pitcherStats = getPitcherStats(sheetId);
  
  if (!pitcherStats[pitcherName]) {
    pitcherStats[pitcherName] = {
      BF: 0,
      outs: 0,
      H: 0,
      HR: 0,
      R: 0,
      BB: 0,
      K: 0
    };
    savePitcherStats(sheetId, pitcherStats);
  }
}

/**
 * Initialize batter stats for a batter
 * @param {string} sheetId - Sheet ID
 * @param {string} batterName - Batter name
 */
function initializeBatterIfNeeded(sheetId, batterName) {
  if (!batterName || batterName === "") return;
  
  var batterStats = getBatterStats(sheetId);
  
  if (!batterStats[batterName]) {
    batterStats[batterName] = {
      AB: 0,
      H: 0,
      HR: 0,
      RBI: 0,
      BB: 0,
      K: 0,
      ROB: 0,
      DP: 0,
      TB: 0
    };
    saveBatterStats(sheetId, batterStats);
  }
}

// ===== LOGGING FUNCTIONS =====

/**
 * Log info message
 * @param {string} module - Module name
 * @param {string} message - Message
 */
function logInfo(module, message) {
  if (typeof Logger !== 'undefined') {
    Logger.log("INFO [" + module + "]: " + message);
  }
}

/**
 * Log warning message
 * @param {string} module - Module name
 * @param {string} message - Message
 * @param {string} entity - Affected entity
 */
function logWarning(module, message, entity) {
  if (typeof Logger !== 'undefined') {
    Logger.log("WARNING [" + module + "]: " + message + " (Entity: " + entity + ")");
  }
}

/**
 * Log error message
 * @param {string} module - Module name
 * @param {string} message - Error message
 * @param {string} entity - Affected entity
 */
function logError(module, message, entity) {
  if (typeof Logger !== 'undefined') {
    Logger.log("ERROR [" + module + "]: " + message + " (Entity: " + entity + ")");
  }
}

// ===== POSITION TRACKING UTILITIES =====

/**
 * Parse position string and return current (rightmost) position
 * Examples:
 *   "SS" → "SS"
 *   "2B / P" → "P"
 *   "RF / P / SS" → "SS"
 * @param {string} positionString - Position value from Column A
 * @return {string} Current position
 */
function getCurrentPosition(positionString) {
  if (!positionString) return '';
  
  var value = positionString.toString().trim();
  
  if (value.indexOf('/') !== -1) {
    var positions = value.split('/').map(function(p) { return p.trim(); });
    return positions[positions.length - 1]; // Rightmost = current
  }
  
  return value; // No history yet
}

/**
 * Append new position to position history string
 * Examples:
 *   ("SS", "P") → "SS / P"
 *   ("SS / P", "RF") → "SS / P / RF"
 * @param {string} currentValue - Current position value
 * @param {string} newPosition - New position to append
 * @return {string} Updated position string
 */
function appendPosition(currentValue, newPosition) {
  if (!currentValue || currentValue.trim() === '') {
    return newPosition;
  }
  
  var current = currentValue.toString().trim();
  
  // Check if already at this position (avoid "P / P")
  if (getCurrentPosition(current) === newPosition) {
    return current;
  }
  
  // Append with delimiter
  return current + ' / ' + newPosition;
}

/**
 * Get full position history as array
 * Examples:
 *   "SS" → ["SS"]
 *   "2B / P / SS" → ["2B", "P", "SS"]
 * @param {string} positionString - Position value from Column A
 * @return {Array} Array of positions
 */
function getPositionHistory(positionString) {
  if (!positionString) return [];
  
  var value = positionString.toString().trim();
  
  if (value.indexOf('/') !== -1) {
    return value.split('/').map(function(p) { return p.trim(); });
  }
  
  return [value];
}

/**
 * Find player row by name in fielding roster
 * Searches both away (7-15) and home (18-26) rosters
 * @param {Sheet} sheet - The game sheet
 * @param {string} playerName - Player name to find
 * @return {number} Row number or -1 if not found
 */
function findPlayerRowByName(sheet, playerName) {
  if (!playerName) return -1;
  
  var nameCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.nameCol;
  
  // Search away roster
  var awayRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var awayNames = sheet.getRange(
    awayRange.startRow, 
    nameCol, 
    awayRange.numPlayers, 
    1
  ).getValues();
  
  for (var i = 0; i < awayNames.length; i++) {
    if (awayNames[i][0] === playerName) {
      return awayRange.startRow + i;
    }
  }
  
  // Search home roster
  var homeRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  var homeNames = sheet.getRange(
    homeRange.startRow, 
    nameCol, 
    homeRange.numPlayers, 
    1
  ).getValues();
  
  for (var i = 0; i < homeNames.length; i++) {
    if (homeNames[i][0] === playerName) {
      return homeRange.startRow + i;
    }
  }
  
  return -1; // Not found
}

/**
 * Find player row by current position in fielding roster
 * @param {Sheet} sheet - The game sheet
 * @param {string} targetPosition - Position to find (e.g., "SS", "P")
 * @param {string} team - "away" or "home" (which roster to search)
 * @return {number} Row number or -1 if not found
 */
function findPlayerRowByPosition(sheet, targetPosition, team) {
  if (!targetPosition) return -1;
  
  var positionCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
  var range = (team === "away") ? 
    BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE : 
    BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  
  var positions = sheet.getRange(
    range.startRow, 
    positionCol, 
    range.numPlayers, 
    1
  ).getValues();
  
  for (var i = 0; i < positions.length; i++) {
    var currentPos = getCurrentPosition(positions[i][0]);
    if (currentPos.toUpperCase() === targetPosition.toUpperCase()) {
      return range.startRow + i;
    }
  }
  
  return -1; // Not found
}