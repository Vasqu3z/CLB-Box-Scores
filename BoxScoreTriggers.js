// ===== BOX SCORE TRIGGERS MODULE =====
// Orchestrates all automation via onEdit trigger
// Handles pitcher changes and at-bat entries
// NOW WITH LOCKSERVICE FOR RAPID EDIT SAFETY

/**
 * Main onEdit trigger - entry point for all automation
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
  
  // ===== ACQUIRE LOCK FOR SEQUENTIAL PROCESSING =====
  var lock = LockService.getScriptLock();
  var lockAcquired = false;
  
  try {
    // Wait for lock (timeout from config)
    lockAcquired = lock.tryLock(BOX_SCORE_CONFIG.LOCK_TIMEOUT_MS);
    
    if (!lockAcquired) {
      // Lock timeout - very rare, only if script hung
      var ui = SpreadsheetApp.getUi();
      ui.alert(
        'Box Score Busy',
        'Another edit is processing. Please wait and try again.',
        ui.ButtonSet.OK
      );
      logWarning("Triggers", "Lock timeout for cell: " + cell, sheetName);
      return;
    }
    
    // ===== PROCESS EDIT (WITHIN LOCK) =====
    processEdit(sheet, cell, row, col, newValue, e.oldValue, e.range);
    
  } catch (error) {
    logError("Triggers", error.toString(), sheetName + "!" + cell);
    
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
    
  } finally {
    // ===== ALWAYS RELEASE LOCK =====
    if (lockAcquired) {
      lock.releaseLock();
    }
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
    }
    return;
  }
  
  // At-bat entries are no longer processed in onEdit
  // They will be processed by the bulk processor menu function
  if (isAtBatCell(row, col)) {
    return;
  }
}

// ============================================
// Handle Position Swaps
// ============================================

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
    // Just move new pitcher to P
    var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
    var newPitcherPosition = sheet.getRange(newPitcherRow, posCol).getValue();
    var updatedPosition = appendPosition(newPitcherPosition, 'P');
    sheet.getRange(newPitcherRow, posCol).setValue(updatedPosition);
    
    SpreadsheetApp.getActiveSpreadsheet().toast(
      newPitcher + ' moved to P',
      'Position Swap',
      3
    );
    return;
  }
  
  // Get current positions
  var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
  var newPitcherPositionCell = sheet.getRange(newPitcherRow, posCol).getValue();
  var oldPitcherPositionCell = sheet.getRange(oldPitcherRow, posCol).getValue();
  
  var newPitcherCurrentPos = getCurrentPosition(newPitcherPositionCell);
  var oldPitcherCurrentPos = getCurrentPosition(oldPitcherPositionCell);
  
  // Check if new pitcher already pitched (re-entry warning)
  var newPitcherHistory = getPositionHistory(newPitcherPositionCell);
  if (newPitcherHistory.indexOf('P') !== -1) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      '⚠️ ' + newPitcher + ' already pitched this game (was at P). Allowing swap...',
      'Pitcher Re-Entry',
      5
    );
  }
  
  // Perform position swap
  var newPitcherUpdated = appendPosition(newPitcherPositionCell, 'P');
  var oldPitcherUpdated = appendPosition(oldPitcherPositionCell, newPitcherCurrentPos);
  
  sheet.getRange(newPitcherRow, posCol).setValue(newPitcherUpdated);
  sheet.getRange(oldPitcherRow, posCol).setValue(oldPitcherUpdated);
  
  // Success toast
  SpreadsheetApp.getActiveSpreadsheet().toast(
    newPitcher + ' moved to P, ' + oldPitcher + ' moved to ' + newPitcherCurrentPos,
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
      logInfo("Triggers", "onEdit trigger already installed");
      return;
    }
  }

  // Simple triggers (like onEdit) don't need manual installation
  // They work automatically when the function is named "onEdit"
  logInfo("Triggers", "onEdit trigger uses simple trigger (automatic)");
}

// ============================================
// v3: BULK PROCESSOR (ABSOLUTE STATE ENGINE)
// ============================================

/**
 * v3: Process all game stats from at-bat grid
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

    logInfo("Processor", "Bulk processing completed in " + duration + "s");

  } catch (error) {
    ui.alert(
      'Processing Error',
      'An error occurred while processing game stats:\n\n' +
      error.toString() + '\n\n' +
      'Please check the Apps Script logs for details.',
      ui.ButtonSet.OK
    );
    logError("Processor", error.toString(), sheet.getName());
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

  // Get active pitcher from dropdown
  var pitcherCell = (battingTeam === 'away') ?
    BOX_SCORE_CONFIG.AWAY_PITCHER_CELL :
    BOX_SCORE_CONFIG.HOME_PITCHER_CELL;

  // For away batters, home pitcher is in D4. For home batters, away pitcher is in D3
  var activePitcherCell = (battingTeam === 'away') ?
    BOX_SCORE_CONFIG.HOME_PITCHER_CELL :
    BOX_SCORE_CONFIG.AWAY_PITCHER_CELL;

  var activePitcher = sheet.getRange(activePitcherCell).getValue();
  var previousPitcher = null;
  var inheritedRunners = 0;

  // Process each batter (row) and inning (column)
  for (var col = 0; col < atBatGrid[0].length; col++) {
    for (var row = 0; row < atBatGrid.length; row++) {
      var value = atBatGrid[row][col];
      if (!value || value === "") continue;

      // Parse notation
      var stats = parseNotation(value);

      // Handle pitcher change
      if (stats.isPitcherChange) {
        previousPitcher = activePitcher;
        // Read new pitcher from dropdown (would need to be updated by user)
        // For now, we just track that a change occurred
        inheritedRunners = stats.inheritedRunners;
        continue;
      }

      // Get batter name
      var batterNames = Object.keys(rosterMap).filter(function(name) {
        return rosterMap[name].team === battingTeam && rosterMap[name].batterIndex === row;
      });

      if (batterNames.length === 0) continue;
      var batterName = batterNames[0];

      // Initialize player stats if needed
      if (!playerStats[batterName]) {
        playerStats[batterName] = {
          hitting: {AB: 0, H: 0, HR: 0, RBI: 0, BB: 0, K: 0, ROB: 0, DP: 0, TB: 0},
          fielding: {NP: 0, E: 0, SB: 0}
        };
      }

      if (activePitcher && !playerStats[activePitcher]) {
        playerStats[activePitcher] = {
          pitching: {BF: 0, outs: 0, H: 0, HR: 0, R: 0, BB: 0, K: 0}
        };
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
        if (fielder && playerStats[fielder]) {
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
        if (fielder && playerStats[fielder]) {
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
  var pitcherCols = BOX_SCORE_CONFIG.PITCHER_STATS_COLUMNS;
  var fieldingCols = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS;
  var hittingCols = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS;
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;

  // Write pitcher and fielding stats
  for (var name in playerStats) {
    if (!rosterMap[name]) continue;

    var row = rosterMap[name].row;
    var team = rosterMap[name].team;

    // Write pitching stats if player has them
    if (playerStats[name].pitching) {
      var p = playerStats[name].pitching;
      var ip = calculateIP(p.outs);

      var pitchingArray = [[ip, p.BF, p.H, p.HR, p.R, p.BB, p.K]];
      sheet.getRange(row, pitcherCols.IP, 1, 7).setValues(pitchingArray);
    }

    // Write fielding stats if player has them
    if (playerStats[name].fielding) {
      var f = playerStats[name].fielding;
      sheet.getRange(row, fieldingCols.NP).setValue(f.NP);
      sheet.getRange(row, fieldingCols.E).setValue(f.E);
      sheet.getRange(row, fieldingCols.SB).setValue(f.SB);
    }

    // Write hitting stats
    if (playerStats[name].hitting) {
      var h = playerStats[name].hitting;
      var batterIndex = rosterMap[name].batterIndex;
      var hittingRow = (team === 'away') ?
        hittingRange.awayStartRow + batterIndex :
        hittingRange.homeStartRow + batterIndex;

      var hittingArray = [[h.AB, h.H, h.HR, h.RBI, h.BB, h.K, h.ROB, h.DP, h.TB]];
      sheet.getRange(hittingRow, hittingCols.AB, 1, 9).setValues(hittingArray);
    }
  }
}