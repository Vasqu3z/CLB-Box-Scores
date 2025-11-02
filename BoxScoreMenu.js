// ===== BOX SCORE MENU MODULE =====
// User interface, menu system, and utility functions

/**
 * Create custom menu when spreadsheet opens
 */
function onOpen() {
  addBoxScoreMenu();
}

/**
 * Add Box Score menu to UI
 */
function addBoxScoreMenu() {
  var ui = SpreadsheetApp.getUi();
  
  ui.createMenu('📊 Box Score Tools')
    .addItem('⚾ View Pitcher Stats', 'showPitcherStats')
    .addItem('🏏 View Hitting Stats', 'showBatterStats')
    .addSeparator()
    .addItem('🔄 Reprocess Selected At-Bat', 'reprocessSelectedAtBat')
    .addItem('🔁 Reprocess Pitcher (All At-Bats)', 'reprocessPitcherAtBats')
    .addSeparator()
    .addItem('📜 View Position Changes', 'showPositionHistory')
    .addSeparator()
    .addItem('🗑️ Reset Game Stats', 'resetCurrentGame')
    .addToUi();
}

// ===== STAT VIEWERS =====

/**
 * Show pitcher stats viewer - Condensed format, separated by team, in pitching order
 */
function showPitcherStats() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetId = sheet.getSheetId().toString();
  var pitcherStats = getPitcherStats(sheetId);
  
  if (Object.keys(pitcherStats).length === 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Pitcher Stats', 'No stats tracked yet.', ui.ButtonSet.OK);
    return;
  }
  
  // Get rosters IN ORDER (column B has names)
  var awayRoster = sheet.getRange(7, 2, 9, 1).getValues(); // B7:B15
  var homeRoster = sheet.getRange(18, 2, 9, 1).getValues(); // B18:B26
  
  var message = "╔═══════════════════════════════════╗\n";
  message += "║       PITCHING STATS SUMMARY         ║\n";
  message += "╚═══════════════════════════════════╝\n\n";
  
  // Away team
  message += "──── AWAY TEAM ────\n";
  var awayCount = 0;
  for (var i = 0; i < awayRoster.length; i++) {
    var name = awayRoster[i][0];
    if (name && pitcherStats[name]) {
      message += formatPitcherStatline(name, pitcherStats[name]) + "\n";
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
    if (name && pitcherStats[name]) {
      message += formatPitcherStatline(name, pitcherStats[name]) + "\n";
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
 * Format pitcher statline: IP, H, R, BB, K (condensed)
 */
function formatPitcherStatline(name, stats) {
  var ip = calculateIP(stats.outs);
  
  // Use fixed-width formatting for alignment
  var ipStr = ip.toFixed(2);
  
  // Pad name to 12 characters for alignment
  var paddedName = (name + "            ").substring(0, 12);
  
  return paddedName + ": " + 
         ipStr + " IP, " + 
         stats.H + " H, " + 
         stats.R + " R, " + 
         stats.BB + " BB, " + 
         stats.K + " K";
}

/**
 * Show batter stats viewer - Baseball statline format, separated by team, in batting order
 */
function showBatterStats() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetId = sheet.getSheetId().toString();
  var batterStats = getBatterStats(sheetId);
  
  if (Object.keys(batterStats).length === 0) {
    var ui = SpreadsheetApp.getUi();
    ui.alert('Batter Stats', 'No stats tracked yet.', ui.ButtonSet.OK);
    return;
  }
  
  // Get rosters IN BATTING ORDER (column B has names)
  var awayRoster = sheet.getRange(30, 2, 9, 1).getValues(); // B30:B38
  var homeRoster = sheet.getRange(41, 2, 9, 1).getValues(); // B41:B49
  
  var message = "╔═══════════════════════════════════╗\n";
  message += "║        BATTING STATS SUMMARY         ║\n";
  message += "╚═══════════════════════════════════╝\n\n";
  
  // Away team
  message += "──── AWAY TEAM ────\n";
  for (var i = 0; i < awayRoster.length; i++) {
    var name = awayRoster[i][0];
    if (name && batterStats[name]) {
      message += (i + 1) + ". " + formatBatterStatline(name, batterStats[name]) + "\n";
    }
  }
  
  message += "\n──── HOME TEAM ────\n";
  for (var i = 0; i < homeRoster.length; i++) {
    var name = homeRoster[i][0];
    if (name && batterStats[name]) {
      message += (i + 1) + ". " + formatBatterStatline(name, batterStats[name]) + "\n";
    }
  }
  
  message += "\n" + "─".repeat(40) + "\n";
  message += "Format: H-AB, XBH, RBI, Hits Stolen, BB, K\n";
  message += "Order: Batting order (1-9)";
  
  var ui = SpreadsheetApp.getUi();
  ui.alert('Batter Stats', message, ui.ButtonSet.OK);
}

/**
 * Format batter statline: H-AB, XBH, RBI, ROB, BB, K
 */
function formatBatterStatline(name, stats) {
  // Pad name to 10 characters for alignment
  var paddedName = (name + "          ").substring(0, 10);
  
  var line = paddedName + ": " + stats.H + "-" + stats.AB;
  
  // Extra base hits (HR only for now)
  if (stats.HR > 0) {
    line += ", " + stats.HR + "HR";
  }
  
  // Other XBH (approximate from TB)
  var otherXBH = stats.TB - stats.H - (stats.HR * 3);
  if (otherXBH > 0) {
    line += ", " + otherXBH + "XBH";
  }
  
  // RBI
  if (stats.RBI > 0) {
    line += ", " + stats.RBI + "RBI";
  }
  
  // Hits Stolen (ROB)
  if (stats.ROB > 0) {
    line += ", " + stats.ROB + " Stolen";
  }
  
  // BB and K (always show for context)
  if (stats.BB > 0 || stats.K > 0) {
    line += " (" + stats.BB + "BB, " + stats.K + "K)";
  }
  
  return line;
}

// ===== REPROCESS FUNCTIONS (IMPROVED) =====

/**
 * Reprocess the currently selected at-bat cell
 * (Renamed from reprocessLastAtBat for clarity)
 */
function reprocessSelectedAtBat() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var selectedRange = sheet.getActiveRange();
  
  if (!selectedRange) {
    ui.alert(
      'No Cell Selected',
      'Please click on an at-bat cell first, then try again.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  var row = selectedRange.getRow();
  var col = selectedRange.getColumn();
  
  // Verify it's an at-bat cell
  if (!isAtBatCell(row, col)) {
    ui.alert(
      'Invalid Cell',
      'Please select an at-bat cell from the game grid:\n\n' +
      '• Away team: C7:H15\n' +
      '• Home team: C18:H26',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Show confirmation with cell details
  var cell = selectedRange.getA1Notation();
  var currentValue = selectedRange.getValue() || "(empty)";
  
  var response = ui.alert(
    'Reprocess At-Bat',
    'Reprocess cell ' + cell + '?\n\nCurrent value: "' + currentValue + '"\n\n' +
    'This will recalculate stats for this at-bat.',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  // Clear shadow storage for this cell to force reprocess
  var sheetId = sheet.getSheetId().toString();
  var gameState = getGameState(sheetId);
  delete gameState[cell];
  saveGameState(sheetId, gameState);
  
  // Trigger reprocess by simulating edit
  var notation = selectedRange.getValue();
  handleAtBatEntry(sheet, cell, row, col, notation);
  
  ui.toast('Cell ' + cell + ' reprocessed successfully', 'Reprocess Complete', 3);
  
  logInfo("Menu", "Reprocessed at-bat: " + cell);
}

/**
 * Reprocess all at-bats for a specific pitcher
 * Useful when pitcher stats look wrong but you don't know which at-bat is the problem
 */
function reprocessPitcherAtBats() {
  var ui = SpreadsheetApp.getUi();
  var sheet = SpreadsheetApp.getActiveSheet();
  var sheetId = sheet.getSheetId().toString();
  
  // Get pitcher stats to build list
  var pitcherStats = getPitcherStats(sheetId);
  
  if (Object.keys(pitcherStats).length === 0) {
    ui.alert(
      'No Pitchers Found',
      'No pitcher stats have been tracked yet.\n\n' +
      'Enter some at-bats first, then try again.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Build pitcher list (numbered, in roster order)
  var awayRoster = sheet.getRange(7, 2, 9, 1).getValues(); // B7:B15
  var homeRoster = sheet.getRange(18, 2, 9, 1).getValues(); // B18:B26
  
  var pitcherList = [];
  var pitcherDisplay = 'AWAY TEAM:\n';
  var counter = 1;
  
  // Add away pitchers who have stats
  for (var i = 0; i < awayRoster.length; i++) {
    var name = awayRoster[i][0];
    if (name && pitcherStats[name]) {
      pitcherList.push(name);
      pitcherDisplay += counter + '. ' + name + '\n';
      counter++;
    }
  }
  
  pitcherDisplay += '\nHOME TEAM:\n';
  
  // Add home pitchers who have stats
  for (var i = 0; i < homeRoster.length; i++) {
    var name = homeRoster[i][0];
    if (name && pitcherStats[name]) {
      pitcherList.push(name);
      pitcherDisplay += counter + '. ' + name + '\n';
      counter++;
    }
  }
  
  // Prompt for pitcher number
  var response = ui.prompt(
    'Reprocess Pitcher',
    'Enter pitcher number:\n\n' + pitcherDisplay,
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) return;
  
  var input = response.getResponseText().trim();
  var pitcherIndex = parseInt(input) - 1;
  
  // Validate input
  if (isNaN(pitcherIndex) || pitcherIndex < 0 || pitcherIndex >= pitcherList.length) {
    ui.alert(
      'Invalid Number',
      'Please enter a number from 1 to ' + pitcherList.length,
      ui.ButtonSet.OK
    );
    return;
  }
  
  var pitcherName = pitcherList[pitcherIndex];
  
  // Validate pitcher exists (should always be true based on list)
  if (!pitcherStats[pitcherName]) {
    ui.alert(
      'Error',
      'Something went wrong. Please try again.',
      ui.ButtonSet.OK
    );
    return;
  }
  
  // Confirm before processing
  var confirmResponse = ui.alert(
    'Confirm Reprocess',
    'Reprocess all at-bats for ' + pitcherName + '?\n\n' +
    'Current stats:\n' +
    '• IP: ' + calculateIP(pitcherStats[pitcherName].outs).toFixed(2) + '\n' +
    '• H: ' + pitcherStats[pitcherName].H + '\n' +
    '• R: ' + pitcherStats[pitcherName].R + '\n' +
    '• K: ' + pitcherStats[pitcherName].K + '\n\n' +
    'This will recalculate from scratch.',
    ui.ButtonSet.YES_NO
  );
  
  if (confirmResponse !== ui.Button.YES) return;
  
  // Find all at-bats where this pitcher was active
  var activePitchers = getActivePitchers(sheetId);
  var gameState = getGameState(sheetId);
  
  var atBatsToReprocess = [];
  
  // Check away at-bats (if pitcher is home team pitcher)
  var awayRange = BOX_SCORE_CONFIG.AWAY_ATBAT_RANGE;
  for (var row = awayRange.startRow; row <= awayRange.endRow; row++) {
    for (var col = awayRange.startCol; col <= awayRange.endCol; col++) {
      var cell = columnToLetter(col) + row;
      var notation = sheet.getRange(cell).getValue();
      
      if (notation) {
        // Check if this pitcher was active when this at-bat happened
        // We'll reprocess all at-bats and let the script figure it out
        atBatsToReprocess.push({cell: cell, row: row, col: col, notation: notation});
      }
    }
  }
  
  // Check home at-bats (if pitcher is away team pitcher)
  var homeRange = BOX_SCORE_CONFIG.HOME_ATBAT_RANGE;
  for (var row = homeRange.startRow; row <= homeRange.endRow; row++) {
    for (var col = homeRange.startCol; col <= homeRange.endCol; col++) {
      var cell = columnToLetter(col) + row;
      var notation = sheet.getRange(cell).getValue();
      
      if (notation) {
        atBatsToReprocess.push({cell: cell, row: row, col: col, notation: notation});
      }
    }
  }
  
  // Clear pitcher stats
  delete pitcherStats[pitcherName];
  savePitcherStats(sheetId, pitcherStats);
  
  // Clear shadow storage for all at-bats
  for (var i = 0; i < atBatsToReprocess.length; i++) {
    delete gameState[atBatsToReprocess[i].cell];
  }
  saveGameState(sheetId, gameState);
  
  // Reprocess all at-bats in order
  var count = 0;
  for (var i = 0; i < atBatsToReprocess.length; i++) {
    var ab = atBatsToReprocess[i];
    handleAtBatEntry(sheet, ab.cell, ab.row, ab.col, ab.notation);
    count++;
  }
  
  // Show results
  var updatedStats = getPitcherStats(sheetId);
  var newPitcherStats = updatedStats[pitcherName] || {outs: 0, H: 0, R: 0, K: 0};
  
  ui.alert(
    'Reprocess Complete',
    'Reprocessed ' + count + ' at-bats for ' + pitcherName + '\n\n' +
    'Updated stats:\n' +
    '• IP: ' + calculateIP(newPitcherStats.outs).toFixed(2) + '\n' +
    '• H: ' + newPitcherStats.H + '\n' +
    '• R: ' + newPitcherStats.R + '\n' +
    '• K: ' + newPitcherStats.K,
    ui.ButtonSet.OK
  );
  
  logInfo("Menu", "Reprocessed " + count + " at-bats for pitcher: " + pitcherName);
}

// ===== RESET GAME =====

/**
 * Reset current game - clear all stats and storage
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
    '• Pitcher dropdowns (C3, C4)\n' +
    '• All invisible tracking data\n\n' +
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
  var sheetId = sheet.getSheetId().toString();
  
  try {
    clearAllProperties(sheetId);
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

// ===== POSITION HISTORY VIEWER =====

/**
 * Show position change history for all players
 */
function showPositionHistory() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var posCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.positionCol;
  var nameCol = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.nameCol;
  
  var history = 'POSITION CHANGE HISTORY:\n\n';
  
  // Away team
  var awayStart = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.startRow;
  var awayEnd = BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.endRow;
  var awayPositions = sheet.getRange(awayStart, posCol, awayEnd - awayStart + 1, 1).getValues();
  var awayNames = sheet.getRange(awayStart, nameCol, awayEnd - awayStart + 1, 1).getValues();
  
  history += 'AWAY TEAM:\n';
  for (var i = 0; i < awayPositions.length; i++) {
    if (awayPositions[i][0] && awayNames[i][0]) {
      var posHistory = getPositionHistory(awayPositions[i][0]);
      if (posHistory.length > 1) {
        history += awayNames[i][0] + ': ' + posHistory.join(' → ') + '\n';
      }
    }
  }
  
  // Home team
  var homeStart = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE.startRow;
  var homeEnd = BOX_SCORE_CONFIG.HOME_PITCHER_RANGE.endRow;
  var homePositions = sheet.getRange(homeStart, posCol, homeEnd - homeStart + 1, 1).getValues();
  var homeNames = sheet.getRange(homeStart, nameCol, homeEnd - homeStart + 1, 1).getValues();
  
  history += '\nHOME TEAM:\n';
  for (var i = 0; i < homePositions.length; i++) {
    if (homePositions[i][0] && homeNames[i][0]) {
      var posHistory = getPositionHistory(homePositions[i][0]);
      if (posHistory.length > 1) {
        history += homeNames[i][0] + ': ' + posHistory.join(' → ') + '\n';
      }
    }
  }
  
  var ui = SpreadsheetApp.getUi();
  ui.alert('Position History', history, ui.ButtonSet.OK);
}