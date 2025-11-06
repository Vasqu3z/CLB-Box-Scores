// ===== SCORE TRIGGERS MODULE =====
// Purpose: Orchestrates automation via onEdit trigger and menu-driven bulk processor.
// Dependencies: ScoreConfig.js, ScoreNotation.js, ScoreUtility.js
// Entry Point(s): onEdit, processGameStatsBulk

/**
 * Main onEdit trigger - entry point for all automation
 * Simplified design - no LockService needed since we only handle position swaps
 * @param {Event} e - Edit event object
 */
function onEdit(e) {
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var sheetName = sheet.getName();

  // Only run on game sheets (sheets starting with configured prefix)
  if (!sheetName.startsWith(BOX_SCORE_CONFIG.GAME_SHEET_PREFIX)) return;

  var cell = e.range.getA1Notation();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var newValue = e.value || "";

  try {
    // Process edit (no lock needed - simple operations only)
    processEdit(sheet, cell, row, col, newValue, e.oldValue, e.range);

  } catch (error) {
    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("ERROR [Triggers]: " + error.toString() + " (Entity: " + sheetName + "!" + cell + ")");
    }

    // Show user-friendly error
    var ui = SpreadsheetApp.getUi();
    ui.alert(
      'Box Score Automation Error',
      'An error occurred while processing your entry.\n\n' +
      'Cell: ' + cell + '\n' +
      'Error: ' + error.toString() + '\n\n' +
      'Please check the Apps Script logs for details.',
      ui.ButtonSet.OK
    );
  }
}

/**
 * Process edit (extracted from onEdit for lock management)
 * @param {Sheet} sheet - The game sheet
 * @param {string} cell - Cell address
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} newValue - New cell value
 * @param {string} oldValue - Old cell value (from event)
 * @param {Range} range - The range object from event
 */
function processEdit(sheet, cell, row, col, newValue, oldValue, range) {
  // ============================================
  // Handle pitcher dropdown changes with position swaps
  // ============================================
  if (cell === BOX_SCORE_CONFIG.AWAY_PITCHER_CELL ||
      cell === BOX_SCORE_CONFIG.HOME_PITCHER_CELL) {

    // Get old value for position swap
    oldValue = oldValue || "";

    // Handle position swap when pitcher changes
    if (oldValue && newValue && oldValue !== newValue) {
      handlePositionSwap(sheet, oldValue, newValue);

      // Auto-insert PC[X] notation when pitcher changes
      if (BOX_SCORE_CONFIG.AUTO_INSERT_PITCHER_CHANGE) {
        autoInsertPitcherChange(sheet, cell, oldValue, newValue);
      }
    }
    return;
  }

  // ============================================
  // Real-time scoring: Auto-process stats after each at-bat
  // ============================================
  if (isAtBatCell(row, col)) {
    // If auto-processing is enabled, trigger bulk processor for real-time scoring
    if (BOX_SCORE_CONFIG.AUTO_PROCESS_ON_AT_BAT) {
      // Use background processing to avoid blocking the user
      processGameStatsBulkBackground(sheet);
    }
    return;
  }
}

// ============================================
// Handle Position Swaps
// ============================================

/**
 * Count how many relief pitchers have been used for a specific team (RP1, RP2, etc.)
 * @param {Sheet} sheet - The game sheet
 * @param {string} team - "away" or "home"
 * @return {number} Highest RP number found for that team (0 if none)
 */
function countReliefPitchers(sheet, team) {
  var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
  var range = (team === 'away') ?
    BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE :
    BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;

  var maxRP = 0;

  // Check only the specified team
  var positions = sheet.getRange(range.startRow, posCol, range.numPlayers, 1).getValues();
  for (var i = 0; i < positions.length; i++) {
    var history = getPositionHistory(positions[i][0]);
    for (var j = 0; j < history.length; j++) {
      var match = history[j].match(/^RP(\d+)$/);
      if (match) {
        maxRP = Math.max(maxRP, parseInt(match[1]));
      }
    }
  }

  return maxRP;
}

/**
 * Handle position swap when pitcher changes
 * @param {Sheet} sheet - The game sheet
 * @param {string} oldPitcher - Previous pitcher name
 * @param {string} newPitcher - New pitcher name
 */
function handlePositionSwap(sheet, oldPitcher, newPitcher) {
  if (!oldPitcher || !newPitcher || oldPitcher === newPitcher) {
    return;
  }
  
  // Find both players in roster
  var newPitcherRow = findPlayerRowByName(sheet, newPitcher);
  var oldPitcherRow = findPlayerRowByName(sheet, oldPitcher);
  
  // Edge case: New pitcher not found
  if (newPitcherRow === -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '⚠️ ' + newPitcher + ' not found in roster',
      'Position Swap',
      5
    );
    return;
  }
  
  // Edge case: Old pitcher not found (shouldn't happen in CLB)
  if (oldPitcherRow === -1) {
    // First pitcher of the game - use SP
    var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
    var newPitcherPosition = sheet.getRange(newPitcherRow, posCol).getValue();
    var updatedPosition = appendPosition(newPitcherPosition, 'SP');
    sheet.getRange(newPitcherRow, posCol).setValue(updatedPosition);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      newPitcher + ' moved to SP',
      'Position Swap',
      3
    );
    return;
  }

  // Get current positions (batch read for performance)
  // Read a range covering both players' rows, then index into the array
  var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
  var minRow = Math.min(newPitcherRow, oldPitcherRow);
  var maxRow = Math.max(newPitcherRow, oldPitcherRow);
  var rowCount = maxRow - minRow + 1;
  var positions = sheet.getRange(minRow, posCol, rowCount, 1).getValues();

  var newPitcherPositionCell = positions[newPitcherRow - minRow][0];
  var oldPitcherPositionCell = positions[oldPitcherRow - minRow][0];

  var newPitcherCurrentPos = getCurrentPosition(newPitcherPositionCell);
  var oldPitcherCurrentPos = getCurrentPosition(oldPitcherPositionCell);

  // Determine which team the new pitcher is on
  var awayRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  var newPitcherTeam;
  if (newPitcherRow >= awayRange.startRow && newPitcherRow <= awayRange.endRow) {
    newPitcherTeam = 'away';
  } else if (newPitcherRow >= homeRange.startRow && newPitcherRow <= homeRange.endRow) {
    newPitcherTeam = 'home';
  }

  // Determine pitcher notation for old and new pitchers
  var oldPitcherHistory = getPositionHistory(oldPitcherPositionCell);
  var oldPitcherNotation;

  // If old pitcher is currently at "P" with no history, they're the starting pitcher
  if (oldPitcherCurrentPos === 'P' && oldPitcherHistory.length === 1) {
    oldPitcherNotation = 'SP';
  } else {
    // Old pitcher keeps their current position notation (already SP/RP#)
    oldPitcherNotation = oldPitcherCurrentPos;
  }

  // Count existing relief pitchers for this team to assign next RP number
  var reliefCount = countReliefPitchers(sheet, newPitcherTeam);
  var newPitcherNotation = 'RP' + (reliefCount + 1);

  // Check if new pitcher already pitched (re-entry warning)
  var newPitcherHistory = getPositionHistory(newPitcherPositionCell);
  var hasPitched = false;
  for (var i = 0; i < newPitcherHistory.length; i++) {
    if (newPitcherHistory[i] === 'P' || newPitcherHistory[i] === 'SP' || newPitcherHistory[i].indexOf('RP') === 0) {
      hasPitched = true;
      break;
    }
  }

  if (hasPitched) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '⚠️ ' + newPitcher + ' already pitched this game (was at ' + newPitcherHistory[i] + '). Allowing swap...',
      'Pitcher Re-Entry',
      5
    );
  }

  // Perform position swap (batch write)
  var newPitcherUpdated = appendPosition(newPitcherPositionCell, newPitcherNotation);
  var oldPitcherUpdated = appendPosition(oldPitcherPositionCell, newPitcherCurrentPos);

  positions[newPitcherRow - minRow][0] = newPitcherUpdated;
  positions[oldPitcherRow - minRow][0] = oldPitcherUpdated;
  sheet.getRange(minRow, posCol, rowCount, 1).setValues(positions);

  // Success toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    newPitcher + ' moved to ' + newPitcherNotation + ', ' + oldPitcher + ' moved to ' + newPitcherCurrentPos,
    'Position Swap',
    3
  );
}

/**
 * Install onEdit trigger (run this once manually if needed)
 * Note: Simple onEdit triggers install automatically, but this is here for reference
 */
function installTriggers() {
  var triggers = ScriptApp.getProjectTriggers();

  // Check if onEdit trigger already exists
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'onEdit') {
      if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
        Logger.log("INFO [Triggers]: onEdit trigger already installed");
      }
      return;
    }
  }

  // Simple triggers (like onEdit) don't need manual installation
  // They work automatically when the function is named "onEdit"
  if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
    Logger.log("INFO [Triggers]: onEdit trigger uses simple trigger (automatic)");
  }
}

// ============================================
// AUTO PITCHER CHANGE TRACKING
// ============================================

/**
 * Auto-insert PC[X] notation when pitcher changes
 * Appends PC notation to the last at-bat cell
 * @param {Sheet} sheet - The game sheet
 * @param {string} pitcherCell - Pitcher dropdown cell (D3 or D4)
 * @param {string} oldPitcher - Previous pitcher name
 * @param {string} newPitcher - New pitcher name
 */
function autoInsertPitcherChange(sheet, pitcherCell, oldPitcher, newPitcher) {
  try {
    // Determine which team is batting (opposite of pitching team)
    var battingTeam = (pitcherCell === BOX_SCORE_CONFIG.AWAY_PITCHER_CELL) ? 'home' : 'away';

    // Find last at-bat cell to append PC notation
    var result = findLastAtBatCell(sheet, battingTeam);
    if (!result) {
      if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
        Logger.log("INFO [AutoPC]: No at-bat entries found for pitcher change - game just started?");
      }
      return;
    }

    // Calculate inherited runners from current inning state
    var inheritedRunners = calculateInheritedRunners(sheet, battingTeam, result.col);

    // Append PC[X] notation to last at-bat cell
    var currentValue = sheet.getRange(result.row, result.col).getValue();
    var pcNotation = currentValue + " PC" + inheritedRunners;
    sheet.getRange(result.row, result.col).setValue(pcNotation);

    // Show toast notification
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Appended PC' + inheritedRunners + ' to last at-bat (' + inheritedRunners + ' inherited runners)',
      'Pitcher Change',
      5
    );

    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("INFO [AutoPC]: Appended PC" + inheritedRunners + " to " + result.row + "," + result.col + " (was: " + currentValue + ")");
    }

  } catch (error) {
    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("ERROR [AutoPC]: Failed to auto-insert PC notation: " + error.toString() + " (Entity: " + sheet.getName() + ")");
    }
  }
}

/**
 * Find the last filled at-bat cell for a team (for appending PC notation)
 * @param {Sheet} sheet - The game sheet
 * @param {string} team - "away" or "home"
 * @return {Object} {row, col} or null if no filled cells found
 */
function findLastAtBatCell(sheet, team) {
  var range = (team === 'away') ?
    BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE :
    BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;

  // Read all at-bat cells
  var numRows = range.endRow - range.startRow + 1;
  var numCols = range.endCol - range.startCol + 1;
  var values = sheet.getRange(range.startRow, range.startCol, numRows, numCols).getValues();

  // Scan from right to left (most recent inning), top to bottom
  // Find the last filled cell
  for (var c = numCols - 1; c >= 0; c--) {
    for (var r = numRows - 1; r >= 0; r--) {
      if (values[r][c] && values[r][c] !== "") {
        return {row: range.startRow + r, col: range.startCol + c};
      }
    }
  }

  // No filled cells found
  return null;
}

/**
 * Find the next available at-bat cell for a team
 * @param {Sheet} sheet - The game sheet
 * @param {string} team - "away" or "home"
 * @return {Object} {row, col} or null if no cell found
 */
function findNextAtBatCell(sheet, team) {
  var range = (team === 'away') ?
    BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE :
    BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;

  // Read all at-bat cells
  var numRows = range.endRow - range.startRow + 1;
  var numCols = range.endCol - range.startCol + 1;
  var values = sheet.getRange(range.startRow, range.startCol, numRows, numCols).getValues();

  // Scan column by column (inning by inning), then row by row (batter by batter)
  // Find the first empty cell in the rightmost active column
  var lastActiveCol = -1;
  var firstEmptyInLastCol = null;

  for (var c = 0; c < numCols; c++) {
    var hasData = false;
    for (var r = 0; r < numRows; r++) {
      if (values[r][c] && values[r][c] !== "") {
        hasData = true;
        lastActiveCol = c;
      }
    }
  }

  // If no active column found, use first column
  if (lastActiveCol === -1) {
    return {row: range.startRow, col: range.startCol};
  }

  // Find first empty cell in the active column
  for (var r = 0; r < numRows; r++) {
    if (!values[r][lastActiveCol] || values[r][lastActiveCol] === "") {
      return {row: range.startRow + r, col: range.startCol + lastActiveCol};
    }
  }

  // Active column is full, move to next column if available
  if (lastActiveCol + 1 < numCols) {
    return {row: range.startRow, col: range.startCol + lastActiveCol + 1};
  }

  // No available cells
  return null;
}

/**
 * Calculate inherited runners from current inning state
 * Logic: Track runners who reached base but haven't scored or made outs
 * @param {Sheet} sheet - The game sheet
 * @param {string} battingTeam - "away" or "home"
 * @param {number} currentCol - Current inning column (absolute column number)
 * @return {number} Number of inherited runners (0-3)
 */
function calculateInheritedRunners(sheet, battingTeam, currentCol) {
  var range = (battingTeam === 'away') ?
    BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE :
    BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;

  // Read current column to check if it has data
  var numRows = range.endRow - range.startRow + 1;
  var currentColValues = sheet.getRange(range.startRow, currentCol, numRows, 1).getValues();

  // Check if current column is empty (new inning scenario)
  var hasData = false;
  for (var i = 0; i < currentColValues.length; i++) {
    if (currentColValues[i][0] && currentColValues[i][0] !== "") {
      hasData = true;
      break;
    }
  }

  // If current column is empty and we're not in the first column, check previous column
  // This handles the case where pitcher changes at start of new inning
  var columnToRead = currentCol;
  if (!hasData && currentCol > range.startCol) {
    columnToRead = currentCol - 1;
  }

  // If still no data (game just started), return 0
  if (!hasData && currentCol === range.startCol) {
    return 0;
  }

  // Read the correct column
  var values = sheet.getRange(range.startRow, columnToRead, numRows, 1).getValues();

  var runnersOnBase = 0;
  var outsRecorded = 0;

  for (var r = 0; r < values.length; r++) {
    var value = values[r][0];
    if (!value || value === "") continue;

    // Skip PC[X] notations (they're pitcher changes, not at-bats)
    if (String(value).toUpperCase().indexOf("PC") === 0) continue;

    var stats = parseNotation(value);

    // Track runners reaching base (hits, walks, errors, fielder's choice)
    // Check all notations: H, BB, E (legacy), E[1-9] (new), FC (no out)
    if (stats.H > 0 || stats.BB > 0 || stats.E || stats.isError || stats.FC) {
      runnersOnBase++;
    }

    // Track runners scoring (subtract RBIs from runners on base)
    if (stats.R > 0) {
      runnersOnBase = Math.max(0, runnersOnBase - stats.R);
    }

    // Track outs
    outsRecorded += stats.outs;
  }

  // If 3 outs recorded, inherited runners should be 0 (inning ended)
  if (outsRecorded >= 3) {
    return 0;
  }

  // Cap at 3 runners (bases loaded)
  return Math.min(3, runnersOnBase);
}

// ============================================
// BULK PROCESSOR (ABSOLUTE STATE ENGINE)
// ============================================

/**
 * Normalize starting pitchers: convert "P" to "SP" for pitchers with no history
 * This ensures starting pitchers are properly marked before processing
 * Uses batch operations for performance
 * @param {Sheet} sheet - The game sheet
 */
function normalizeStartingPitchers(sheet) {
  var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
  var awayRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;

  // Read both teams' positions in batch
  var awayPositions = sheet.getRange(awayRange.startRow, posCol, awayRange.numPlayers, 1).getValues();
  var homePositions = sheet.getRange(homeRange.startRow, posCol, homeRange.numPlayers, 1).getValues();

  var needsUpdate = false;

  // Convert P → SP in the arrays
  for (var i = 0; i < awayPositions.length; i++) {
    if (awayPositions[i][0] === 'P') {
      awayPositions[i][0] = 'SP';
      needsUpdate = true;
    }
  }

  for (var i = 0; i < homePositions.length; i++) {
    if (homePositions[i][0] === 'P') {
      homePositions[i][0] = 'SP';
      needsUpdate = true;
    }
  }

  // Write back in batch only if changes were made
  if (needsUpdate) {
    sheet.getRange(awayRange.startRow, posCol, awayRange.numPlayers, 1).setValues(awayPositions);
    sheet.getRange(homeRange.startRow, posCol, homeRange.numPlayers, 1).setValues(homePositions);
  }
}

/**
 * Background processor for real-time scoring (no UI alerts)
 * Called automatically after each at-bat entry when AUTO_PROCESS_ON_AT_BAT is true
 * @param {Sheet} sheet - The game sheet
 */
function processGameStatsBulkBackground(sheet) {
  try {
    // Normalize starting pitchers (convert P → SP)
    normalizeStartingPitchers(sheet);

    // Step 1: Clear all old stat data
    clearPitcherStatsInSheet(sheet);
    clearHittingStatsInSheet(sheet);

    // Step 2: Build roster map (player name -> row/position)
    var rosterMap = buildRosterMap(sheet);

    // Step 3: Read at-bat grids
    var awayAtBats = readAtBatGrid(sheet, true);  // true = away team
    var homeAtBats = readAtBatGrid(sheet, false); // false = home team

    // Step 4: Initialize stat storage
    var playerStats = {}; // {playerName: {pitching: {...}, hitting: {...}, fielding: {...}}}

    // Step 5: Process away team at-bats (home pitcher pitching)
    var awayState = processTeamAtBats(sheet, awayAtBats, 'away', rosterMap, playerStats);

    // Step 6: Process home team at-bats (away pitcher pitching)
    var homeState = processTeamAtBats(sheet, homeAtBats, 'home', rosterMap, playerStats);

    // Step 7: Write all stats to sheet in batch
    writeStatsToSheet(sheet, playerStats, rosterMap);

    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("INFO [Processor]: Background processing completed (real-time mode)");
    }

  } catch (error) {
    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("ERROR [Processor]: Background processing failed: " + error.toString() + " (Entity: " + sheet.getName() + ")");
    }
    // Don't show UI alert - just log the error
  }
}

/**
 * Process all game stats from at-bat grid
 * This is the main menu-driven bulk processor that calculates all stats from scratch
 */
function processGameStatsBulk() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var ui = SpreadsheetApp.getUi();

  // Show progress message
  var startTime = new Date().getTime();
  ui.alert(
    'Processing Game Stats',
    'Reading at-bat grids and calculating stats...\n\nThis may take a few seconds.',
    ui.ButtonSet.OK
  );

  try {
    // Normalize starting pitchers (convert P → SP)
    normalizeStartingPitchers(sheet);

    // Step 1: Clear all old stat data
    clearPitcherStatsInSheet(sheet);
    clearHittingStatsInSheet(sheet);

    // Step 2: Build roster map (player name -> row/position)
    var rosterMap = buildRosterMap(sheet);

    // Step 3: Read at-bat grids
    var awayAtBats = readAtBatGrid(sheet, true);  // true = away team
    var homeAtBats = readAtBatGrid(sheet, false); // false = home team

    // Step 4: Initialize stat storage
    var playerStats = {}; // {playerName: {pitching: {...}, hitting: {...}, fielding: {...}}}

    // Step 5: Process away team at-bats (home pitcher pitching)
    var awayState = processTeamAtBats(sheet, awayAtBats, 'away', rosterMap, playerStats);

    // Step 6: Process home team at-bats (away pitcher pitching)
    var homeState = processTeamAtBats(sheet, homeAtBats, 'home', rosterMap, playerStats);

    // Step 7: Write all stats to sheet in batch
    writeStatsToSheet(sheet, playerStats, rosterMap);

    // Show completion message
    var endTime = new Date().getTime();
    var duration = ((endTime - startTime) / 1000).toFixed(1);

    ui.alert(
      'Processing Complete',
      'Game stats have been calculated and updated.\n\n' +
      'Processing time: ' + duration + ' seconds',
      ui.ButtonSet.OK
    );

    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("INFO [Processor]: Bulk processing completed in " + duration + "s");
    }

  } catch (error) {
    ui.alert(
      'Processing Error',
      'An error occurred while processing game stats:\n\n' +
      error.toString() + '\n\n' +
      'Please check the Apps Script logs for details.',
      ui.ButtonSet.OK
    );
    if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
      Logger.log("ERROR [Processor]: " + error.toString() + " (Entity: " + sheet.getName() + ")");
    }
  }
}

/**
 * Build roster map for quick player lookup
 * @param {Sheet} sheet - The game sheet
 * @return {Object} Map of player name to {row, position, team}
 */
function buildRosterMap(sheet) {
  var map = {};

  var awayRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;

  // Read away team roster
  var awayNames = sheet.getRange(awayRange.startRow, awayRange.nameCol, awayRange.numPlayers, 1).getValues();
  var awayPositions = sheet.getRange(awayRange.startRow, awayRange.positionCol, awayRange.numPlayers, 1).getValues();

  for (var i = 0; i < awayNames.length; i++) {
    var name = String(awayNames[i][0]).trim();
    if (name) {
      map[name] = {
        row: awayRange.startRow + i,
        position: getCurrentPosition(awayPositions[i][0]),
        team: 'away',
        batterIndex: i  // 0-8 for lineup position
      };
    }
  }

  // Read home team roster
  var homeNames = sheet.getRange(homeRange.startRow, homeRange.nameCol, homeRange.numPlayers, 1).getValues();
  var homePositions = sheet.getRange(homeRange.startRow, homeRange.positionCol, homeRange.numPlayers, 1).getValues();

  for (var i = 0; i < homeNames.length; i++) {
    var name = String(homeNames[i][0]).trim();
    if (name) {
      map[name] = {
        row: homeRange.startRow + i,
        position: getCurrentPosition(homePositions[i][0]),
        team: 'home',
        batterIndex: i  // 0-8 for lineup position
      };
    }
  }

  return map;
}

/**
 * Build pitcher timeline from position history (SP, RP1, RP2, etc.)
 * Uses batch read for performance
 * @param {Sheet} sheet - The game sheet
 * @param {string} team - "away" or "home" (the fielding team)
 * @param {Object} rosterMap - Player roster map
 * @return {Array} Array of pitcher names in order: [SP, RP1, RP2, ...]
 */
function buildPitcherTimeline(sheet, team, rosterMap) {
  var timeline = [];
  var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;

  // Determine which range to read based on team
  var range = (team === 'away') ?
    BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE :
    BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;

  // Read all position cells for this team in batch
  var positionCells = sheet.getRange(range.startRow, posCol, range.numPlayers, 1).getValues();
  var nameCol = range.nameCol;
  var names = sheet.getRange(range.startRow, nameCol, range.numPlayers, 1).getValues();

  // Scan through positions to build timeline
  for (var i = 0; i < positionCells.length; i++) {
    var positionCell = positionCells[i][0];
    var name = String(names[i][0]).trim();

    if (!name || !positionCell) continue;

    var history = getPositionHistory(positionCell);

    // Check for SP (starting pitcher)
    for (var j = 0; j < history.length; j++) {
      if (history[j] === 'SP') {
        timeline[0] = name;
      }

      // Check for RP# (relief pitchers)
      var rpMatch = history[j].match(/^RP(\d+)$/);
      if (rpMatch) {
        var rpNum = parseInt(rpMatch[1]);
        timeline[rpNum] = name;
      }
    }
  }

  return timeline;
}

/**
 * Read at-bat grid for a team
 * @param {Sheet} sheet - The game sheet
 * @param {boolean} isAway - True for away team, false for home team
 * @return {Array} 2D array of at-bat values
 */
function readAtBatGrid(sheet, isAway) {
  var range = isAway ? BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE : BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;
  var numRows = range.endRow - range.startRow + 1;
  var numCols = range.endCol - range.startCol + 1;

  return sheet.getRange(range.startRow, range.startCol, numRows, numCols).getValues();
}

/**
 * Process at-bats for one team
 * @param {Sheet} sheet - The game sheet
 * @param {Array} atBatGrid - 2D array of at-bat values
 * @param {string} battingTeam - "away" or "home"
 * @param {Object} rosterMap - Player roster map
 * @param {Object} playerStats - Stats storage object (modified in place)
 * @return {Object} Final state {activePitcher, inheritedRunners}
 */
function processTeamAtBats(sheet, atBatGrid, battingTeam, rosterMap, playerStats) {
  var fieldingTeam = (battingTeam === 'away') ? 'home' : 'away';

  // Build pitcher timeline from position history (SP, RP1, RP2, ...)
  var pitcherTimeline = buildPitcherTimeline(sheet, fieldingTeam, rosterMap);

  // Start with first pitcher (SP)
  // Note: normalizeStartingPitchers() ensures P → SP conversion before processing
  var pitcherIndex = 0;
  var activePitcher = pitcherTimeline[0] || null;
  var previousPitcher = null;
  var inheritedRunners = 0;

  // Process each batter (row) and inning (column)
  for (var col = 0; col < atBatGrid[0].length; col++) {
    for (var row = 0; row < atBatGrid.length; row++) {
      var value = atBatGrid[row][col];
      if (!value || value === "") continue;

      // Parse notation
      var stats = parseNotation(value);

      // Check if this is a standalone PC notation (no at-bat stats)
      var isStandalonePitcherChange = stats.isPitcherChange && stats.BF === 0;

      // Handle standalone pitcher change (no at-bat in this cell)
      if (isStandalonePitcherChange) {
        // Store current pitcher as previous (for inherited runs)
        previousPitcher = activePitcher;

        // Switch to next pitcher in timeline
        pitcherIndex++;
        activePitcher = pitcherTimeline[pitcherIndex];

        inheritedRunners = stats.inheritedRunners;
        continue;
      }

      // If we reach here, there's an at-bat to process (may or may not have PC appended)

      // Get batter name
      var batterNames = Object.keys(rosterMap).filter(function(name) {
        return rosterMap[name].team === battingTeam && rosterMap[name].batterIndex === row;
      });

      if (batterNames.length === 0) continue;
      var batterName = batterNames[0];

      // Initialize player stats if needed
      if (!playerStats[batterName]) {
        playerStats[batterName] = {};
      }
      if (!playerStats[batterName].hitting) {
        playerStats[batterName].hitting = {AB: 0, H: 0, HR: 0, RBI: 0, BB: 0, K: 0, ROB: 0, DP: 0, TB: 0};
      }
      if (!playerStats[batterName].fielding) {
        playerStats[batterName].fielding = {NP: 0, E: 0, SB: 0};
      }

      if (activePitcher) {
        if (!playerStats[activePitcher]) {
          playerStats[activePitcher] = {};
        }
        if (!playerStats[activePitcher].pitching) {
          playerStats[activePitcher].pitching = {BF: 0, outs: 0, H: 0, HR: 0, R: 0, BB: 0, K: 0};
        }
      }

      // Apply hitting stats
      playerStats[batterName].hitting.AB += stats.AB;
      playerStats[batterName].hitting.H += stats.H;
      playerStats[batterName].hitting.HR += stats.HR;
      playerStats[batterName].hitting.RBI += stats.R;  // R from notation = RBI for batter
      playerStats[batterName].hitting.BB += stats.BB;
      playerStats[batterName].hitting.K += stats.K;
      playerStats[batterName].hitting.DP += stats.DP ? 1 : 0;
      playerStats[batterName].hitting.TB += stats.TB;

      // Apply pitching stats (to active pitcher)
      if (activePitcher && playerStats[activePitcher].pitching) {
        playerStats[activePitcher].pitching.BF += stats.BF;
        playerStats[activePitcher].pitching.outs += stats.outs;
        playerStats[activePitcher].pitching.H += stats.H;
        playerStats[activePitcher].pitching.HR += stats.HR;
        playerStats[activePitcher].pitching.BB += stats.BB;
        playerStats[activePitcher].pitching.K += stats.K;

        // Handle inherited runs
        if (stats.R > 0) {
          if (inheritedRunners > 0) {
            // Assign runs to previous pitcher
            var runsToInherit = Math.min(stats.R, inheritedRunners);
            if (previousPitcher && playerStats[previousPitcher] && playerStats[previousPitcher].pitching) {
              playerStats[previousPitcher].pitching.R += runsToInherit;
            }
            playerStats[activePitcher].pitching.R += (stats.R - runsToInherit);
            inheritedRunners -= runsToInherit;
          } else {
            playerStats[activePitcher].pitching.R += stats.R;
          }
        }
      }

      // Handle fielding stats (NP, E)
      if (stats.isNicePlay && stats.fielderPosition) {
        var fielder = findPlayerByPosition(rosterMap, fieldingTeam, stats.fielderPosition);
        if (fielder) {
          if (!playerStats[fielder]) {
            playerStats[fielder] = {};
          }
          if (!playerStats[fielder].fielding) {
            playerStats[fielder].fielding = {NP: 0, E: 0, SB: 0};
          }
          playerStats[fielder].fielding.NP += 1;
          // Add ROB to batter
          playerStats[batterName].hitting.ROB += 1;
        }
      }

      if (stats.isError && stats.fielderPosition) {
        var fielder = findPlayerByPosition(rosterMap, fieldingTeam, stats.fielderPosition);
        if (fielder) {
          if (!playerStats[fielder]) {
            playerStats[fielder] = {};
          }
          if (!playerStats[fielder].fielding) {
            playerStats[fielder].fielding = {NP: 0, E: 0, SB: 0};
          }
          playerStats[fielder].fielding.E += 1;
        }
      }

      // Handle stolen bases
      if (stats.SB) {
        playerStats[batterName].fielding.SB += 1;
      }

      // Handle pitcher change AFTER processing at-bat stats
      // Order matters: "K PC2" means the DEPARTING pitcher gets credit for the strikeout,
      // then we switch to the relief pitcher for subsequent at-bats
      if (stats.isPitcherChange) {
        // Store current pitcher as previous (for inherited runs)
        previousPitcher = activePitcher;

        // Switch to next pitcher in timeline
        pitcherIndex++;
        activePitcher = pitcherTimeline[pitcherIndex];

        inheritedRunners = stats.inheritedRunners;
      }
    }

    // Clear inherited runners at end of each inning
    inheritedRunners = 0;
  }

  return {
    activePitcher: activePitcher,
    inheritedRunners: inheritedRunners
  };
}

/**
 * Find player name by fielding position
 * @param {Object} rosterMap - Player roster map
 * @param {string} team - "away" or "home"
 * @param {number} position - Position number (1=P, 2=C, 3=1B, 4=2B, 5=3B, 6=SS, 7=LF, 8=CF, 9=RF)
 * @return {string} Player name or null
 */
function findPlayerByPosition(rosterMap, team, position) {
  var positionMap = {
    1: 'P', 2: 'C', 3: '1B', 4: '2B', 5: '3B', 6: 'SS', 7: 'LF', 8: 'CF', 9: 'RF'
  };

  var targetPos = positionMap[position];
  if (!targetPos) return null;

  for (var name in rosterMap) {
    if (rosterMap[name].team === team && rosterMap[name].position === targetPos) {
      return name;
    }
  }

  return null;
}

/**
 * Write all stats to sheet in batch
 * @param {Sheet} sheet - The game sheet
 * @param {Object} playerStats - Stats storage object
 * @param {Object} rosterMap - Player roster map
 */
function writeStatsToSheet(sheet, playerStats, rosterMap) {
  var pCols = BOX_SCORE_CONFIG.PITCHER_STATS_COLUMNS;
  var fCols = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS;
  var hCols = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS;
  var awayPitcherRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homePitcherRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;

  // 1. Create empty 2D arrays to hold all stats
  var numPitcherCols = 7; // IP to K
  var numFieldingCols = 3; // NP to SB
  var numHittingCols = 9;  // AB to TB

  var awayPitchingBatch = createEmptyBatch(awayPitcherRange.numPlayers, numPitcherCols);
  var homePitchingBatch = createEmptyBatch(homePitcherRange.numPlayers, numPitcherCols);

  var awayFieldingBatch = createEmptyBatch(awayPitcherRange.numPlayers, numFieldingCols);
  var homeFieldingBatch = createEmptyBatch(homePitcherRange.numPlayers, numFieldingCols);

  var awayHittingBatch = createEmptyBatch(hittingRange.numPlayers, numHittingCols);
  var homeHittingBatch = createEmptyBatch(hittingRange.numPlayers, numHittingCols);

  // 2. Loop through playerStats ONCE to populate batch arrays
  for (var name in playerStats) {
    if (!rosterMap[name]) continue;

    var map = rosterMap[name];
    var stats = playerStats[name];
    var batchIndex = map.batterIndex; // 0-8

    // Pitching & Fielding Stats (go in the same roster block)
    if (map.team === 'away') {
      if (stats.pitching) {
        var p = stats.pitching;
        awayPitchingBatch[batchIndex] = [calculateIP(p.outs), p.BF, p.H, p.HR, p.R, p.BB, p.K];
      }
      if (stats.fielding) {
        var f = stats.fielding;
        awayFieldingBatch[batchIndex] = [f.NP, f.E, f.SB];
      }
    } else { // Home Team
      if (stats.pitching) {
        var p = stats.pitching;
        homePitchingBatch[batchIndex] = [calculateIP(p.outs), p.BF, p.H, p.HR, p.R, p.BB, p.K];
      }
      if (stats.fielding) {
        var f = stats.fielding;
        homeFieldingBatch[batchIndex] = [f.NP, f.E, f.SB];
      }
    }

    // Hitting Stats (go in a separate roster block)
    if (stats.hitting) {
      var h = stats.hitting;
      var hittingArray = [h.AB, h.H, h.HR, h.RBI, h.BB, h.K, h.ROB, h.DP, h.TB];
      if (map.team === 'away') {
        awayHittingBatch[batchIndex] = hittingArray;
      } else {
        homeHittingBatch[batchIndex] = hittingArray;
      }
    }
  }

  // 3. Write all stats in 6 batch operations (AFTER the loop)
  sheet.getRange(awayPitcherRange.startRow, pCols.IP, awayPitcherRange.numPlayers, numPitcherCols).setValues(awayPitchingBatch);
  sheet.getRange(homePitcherRange.startRow, pCols.IP, homePitcherRange.numPlayers, numPitcherCols).setValues(homePitchingBatch);

  sheet.getRange(awayPitcherRange.startRow, fCols.NP, awayPitcherRange.numPlayers, numFieldingCols).setValues(awayFieldingBatch);
  sheet.getRange(homePitcherRange.startRow, fCols.NP, homePitcherRange.numPlayers, numFieldingCols).setValues(homeFieldingBatch);

  sheet.getRange(hittingRange.awayStartRow, hCols.AB, hittingRange.numPlayers, numHittingCols).setValues(awayHittingBatch);
  sheet.getRange(hittingRange.homeStartRow, hCols.AB, hittingRange.numPlayers, numHittingCols).setValues(homeHittingBatch);
}