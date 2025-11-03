// ===== SCORE UTILITY MODULE =====
// Shared helper functions for score automation
// Used by all other modules

// ===== SHEET OPERATIONS =====

/**
 * Clear pitcher and defensive stats in sheet (skip protected rows)
 * v3: Optimized with batch operations
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

  // Build array of zeros for batch write
  var zeroRow = [];
  for (var i = 0; i < numCols; i++) {
    zeroRow.push(0);
  }

  // Clear away pitcher/defensive stats (batch operation)
  var awayRows = [];
  for (var row = awayRange.startRow; row <= awayRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      awayRows.push(zeroRow.slice());
    }
  }
  if (awayRows.length > 0) {
    sheet.getRange(awayRange.startRow, firstCol, awayRows.length, numCols).setValues(awayRows);
  }

  // Clear home pitcher/defensive stats (batch operation)
  var homeRows = [];
  for (var row = homeRange.startRow; row <= homeRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      homeRows.push(zeroRow.slice());
    }
  }
  if (homeRows.length > 0) {
    sheet.getRange(homeRange.startRow, firstCol, homeRows.length, numCols).setValues(homeRows);
  }
}

/**
 * Clear hitting and stolen base stats in sheet (skip protected rows)
 * v3: Optimized with batch operations
 * @param {Sheet} sheet - The game sheet
 */
function clearHittingStatsInSheet(sheet) {
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;
  var hittingCols = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS;
  var sbCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.SB;

  // Build zero row for hitting stats (9 columns)
  var zeroHittingRow = [0, 0, 0, 0, 0, 0, 0, 0, 0];

  // Clear away hitting stats (batch operation)
  var awayRows = [];
  for (var row = hittingRange.awayStartRow; row <= hittingRange.awayEndRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      awayRows.push(zeroHittingRow.slice());
    }
  }
  if (awayRows.length > 0) {
    sheet.getRange(hittingRange.awayStartRow, hittingCols.AB, awayRows.length, hittingRange.numStatCols).setValues(awayRows);
  }

  // Clear home hitting stats (batch operation)
  var homeRows = [];
  for (var row = hittingRange.homeStartRow; row <= hittingRange.homeEndRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      homeRows.push(zeroHittingRow.slice());
    }
  }
  if (homeRows.length > 0) {
    sheet.getRange(hittingRange.homeStartRow, hittingCols.AB, homeRows.length, hittingRange.numStatCols).setValues(homeRows);
  }

  // Clear SB from fielding section - batch operations
  var awayFieldingRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeFieldingRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;

  var awaySBRows = [];
  for (var row = awayFieldingRange.startRow; row <= awayFieldingRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      awaySBRows.push([0]);
    }
  }
  if (awaySBRows.length > 0) {
    sheet.getRange(awayFieldingRange.startRow, sbCol, awaySBRows.length, 1).setValues(awaySBRows);
  }

  var homeSBRows = [];
  for (var row = homeFieldingRange.startRow; row <= homeFieldingRange.endRow; row++) {
    if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) === -1) {
      homeSBRows.push([0]);
    }
  }
  if (homeSBRows.length > 0) {
    sheet.getRange(homeFieldingRange.startRow, sbCol, homeSBRows.length, 1).setValues(homeSBRows);
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