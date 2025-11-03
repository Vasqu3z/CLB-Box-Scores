// ===== SCORE MENU MODULE =====
// User interface, menu system, and stat viewers

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  addBoxScoreMenu_v3();
}

/**
 * Add Box Score v3 menu to UI
 */
function addBoxScoreMenu_v3() {
  var ui = SpreadsheetApp.getUi();

  ui.createMenu('📊 Box Score Tools')
    .addItem('🚀 Process Game Stats', 'processGameStatsBulk')
    .addSeparator()
    .addItem('⚾ View Pitcher Stats', 'showPitcherStats')
    .addItem('🏏 View Hitting Stats', 'showBatterStats')
    .addSeparator()
    .addItem('🗑️ Reset Game Stats', 'resetCurrentGame')
    .addToUi();
}

// ===== STAT VIEWERS =====

/**
 * Show pitcher stats viewer - Condensed format, separated by team, in pitching order
 * v3: Reads stats directly from the sheet instead of PropertiesService
 */
function showPitcherStats() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get rosters IN ORDER using config
  var awayRange = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE;
  var homeRange = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  var pitcherCols = BOX_SCORE_CONFIG.PITCHER_STATS_COLUMNS;

  var awayRoster = sheet.getRange(awayRange.startRow, awayRange.nameCol, awayRange.numPlayers, 1).getValues();
  var homeRoster = sheet.getRange(homeRange.startRow, homeRange.nameCol, homeRange.numPlayers, 1).getValues();

  // Read pitcher stats from sheet
  var numPitcherCols = Object.keys(pitcherCols).length;
  var awayStats = sheet.getRange(awayRange.startRow, pitcherCols.IP, awayRange.numPlayers, numPitcherCols).getValues();
  var homeStats = sheet.getRange(homeRange.startRow, pitcherCols.IP, homeRange.numPlayers, numPitcherCols).getValues();
  
  var message = "╔═══════════════════════════════════╗\n";
  message += "║       PITCHING STATS SUMMARY         ║\n";
  message += "╚═══════════════════════════════════╝\n\n";

  // Away team
  message += "──── AWAY TEAM ────\n";
  var awayCount = 0;
  for (var i = 0; i < awayRoster.length; i++) {
    var name = awayRoster[i][0];
    var stats = awayStats[i]; // [IP, BF, H, HR, R, BB, K]
    if (name && stats[0] > 0) { // Check if pitcher has pitched (IP > 0)
      var paddedName = (name + "            ").substring(0, 12);
      message += paddedName + ": " +
                 stats[0].toFixed(2) + " IP, " +
                 stats[2] + " H, " +
                 stats[4] + " R, " +
                 stats[5] + " BB, " +
                 stats[6] + " K\n";
      awayCount++;
    }
  }
  if (awayCount === 0) {
    message += "(No pitching stats yet)\n";
  }

  message += "\n──── HOME TEAM ────\n";
  var homeCount = 0;
  for (var i = 0; i < homeRoster.length; i++) {
    var name = homeRoster[i][0];
    var stats = homeStats[i]; // [IP, BF, H, HR, R, BB, K]
    if (name && stats[0] > 0) { // Check if pitcher has pitched (IP > 0)
      var paddedName = (name + "            ").substring(0, 12);
      message += paddedName + ": " +
                 stats[0].toFixed(2) + " IP, " +
                 stats[2] + " H, " +
                 stats[4] + " R, " +
                 stats[5] + " BB, " +
                 stats[6] + " K\n";
      homeCount++;
    }
  }
  if (homeCount === 0) {
    message += "(No pitching stats yet)\n";
  }

  message += "\n" + "─".repeat(40) + "\n";
  message += "Format: IP, H, R, BB, K\n";
  message += "Order: Roster order (for W/L/SV tracking)";

  var ui = SpreadsheetApp.getUi();
  ui.alert('Pitcher Stats', message, ui.ButtonSet.OK);
}

/**
 * Show batter stats viewer - Baseball statline format, separated by team, in batting order
 * v3: Reads stats directly from the sheet instead of PropertiesService
 */
function showBatterStats() {
  var sheet = SpreadsheetApp.getActiveSheet();

  // Get rosters IN BATTING ORDER using config
  var hittingRange = BOX_SCORE_CONFIG.HITTING_RANGE;
  var hittingCols = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS;

  var awayRoster = sheet.getRange(hittingRange.awayStartRow, hittingRange.nameCol, hittingRange.numPlayers, 1).getValues();
  var homeRoster = sheet.getRange(hittingRange.homeStartRow, hittingRange.nameCol, hittingRange.numPlayers, 1).getValues();

  // Read hitting stats from sheet
  var awayStats = sheet.getRange(hittingRange.awayStartRow, hittingCols.AB, hittingRange.numPlayers, hittingRange.numStatCols).getValues();
  var homeStats = sheet.getRange(hittingRange.homeStartRow, hittingCols.AB, hittingRange.numPlayers, hittingRange.numStatCols).getValues();
  
  var message = "╔═══════════════════════════════════╗\n";
  message += "║        BATTING STATS SUMMARY         ║\n";
  message += "╚═══════════════════════════════════╝\n\n";

  // Away team
  message += "──── AWAY TEAM ────\n";
  for (var i = 0; i < awayRoster.length; i++) {
    var name = awayRoster[i][0];
    var stats = awayStats[i]; // [AB, H, HR, RBI, BB, K, ROB, DP, TB]
    if (name && stats[0] > 0) { // Check if batter has AB > 0
      var paddedName = (name + "          ").substring(0, 10);
      var line = (i + 1) + ". " + paddedName + ": " + stats[1] + "-" + stats[0];

      // HR
      if (stats[2] > 0) {
        line += ", " + stats[2] + "HR";
      }

      // Other XBH (approximate from TB)
      var otherXBH = stats[8] - stats[1] - (stats[2] * 3);
      if (otherXBH > 0) {
        line += ", " + otherXBH + "XBH";
      }

      // RBI
      if (stats[3] > 0) {
        line += ", " + stats[3] + "RBI";
      }

      // Hits Stolen (ROB)
      if (stats[6] > 0) {
        line += ", " + stats[6] + " Stolen";
      }

      // BB and K
      if (stats[4] > 0 || stats[5] > 0) {
        line += " (" + stats[4] + "BB, " + stats[5] + "K)";
      }

      message += line + "\n";
    }
  }

  message += "\n──── HOME TEAM ────\n";
  for (var i = 0; i < homeRoster.length; i++) {
    var name = homeRoster[i][0];
    var stats = homeStats[i]; // [AB, H, HR, RBI, BB, K, ROB, DP, TB]
    if (name && stats[0] > 0) { // Check if batter has AB > 0
      var paddedName = (name + "          ").substring(0, 10);
      var line = (i + 1) + ". " + paddedName + ": " + stats[1] + "-" + stats[0];

      // HR
      if (stats[2] > 0) {
        line += ", " + stats[2] + "HR";
      }

      // Other XBH (approximate from TB)
      var otherXBH = stats[8] - stats[1] - (stats[2] * 3);
      if (otherXBH > 0) {
        line += ", " + otherXBH + "XBH";
      }

      // RBI
      if (stats[3] > 0) {
        line += ", " + stats[3] + "RBI";
      }

      // Hits Stolen (ROB)
      if (stats[6] > 0) {
        line += ", " + stats[6] + " Stolen";
      }

      // BB and K
      if (stats[4] > 0 || stats[5] > 0) {
        line += " (" + stats[4] + "BB, " + stats[5] + "K)";
      }

      message += line + "\n";
    }
  }

  message += "\n" + "─".repeat(40) + "\n";
  message += "Format: H-AB, XBH, RBI, Hits Stolen, BB, K\n";
  message += "Order: Batting order (1-9)";

  var ui = SpreadsheetApp.getUi();
  ui.alert('Batter Stats', message, ui.ButtonSet.OK);
}

// ===== RESET GAME =====

/**
 * Reset current game - clear all stats
 * v3: Simplified since PropertiesService is no longer used
 */
function resetCurrentGame() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    'Reset Game Stats',
    'This will clear all tracked stats for this game sheet.\n\n' +
    'The following will be reset:\n' +
    '• Pitcher stats (columns I-O)\n' +
    '• Defensive stats (columns P-R)\n' +
    '• Hitting stats (columns C-K)\n' +
    '• Pitcher dropdowns (D3, D4)\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) return;
  
  // Ask if they want to clear at-bats too
  var clearAtBats = ui.alert(
    'Clear At-Bats?',
    'Do you also want to clear all entered at-bat results?\n\n' +
    '(This will erase the game grid C7:H15 and C18:H26)',
    ui.ButtonSet.YES_NO
  );
  
  var sheet = SpreadsheetApp.getActiveSheet();

  try {
    clearPitcherStatsInSheet(sheet);
    clearHittingStatsInSheet(sheet);
    clearPitcherDropdowns(sheet);
    
    if (clearAtBats === ui.Button.YES) {
      clearAtBatGrid(sheet);
    }
    
    ui.alert(
      'Game Reset Complete', 
      'All stats have been cleared.\n\n' +
      (clearAtBats === ui.Button.YES ? 'At-bat grid has been cleared.\n\n' : '') +
      'Ready for a new game!',
      ui.ButtonSet.OK
    );
    
    logInfo("Menu", "Game reset completed for sheet: " + sheet.getName());
  } catch (error) {
    ui.alert('Reset Error', 'Error: ' + error.toString(), ui.ButtonSet.OK);
    logError("Menu", "Reset failed: " + error.toString(), sheet.getName());
  }
}

