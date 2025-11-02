// ===== BOX SCORE HITTING MODULE =====
// Handles batter stat tracking with delta calculation
// Attributes stats based on which row the at-bat cell is in

/**
 * Apply delta to batter stats
 * @param {Sheet} sheet - The game sheet
 * @param {number} atBatRow - At-bat cell row
 * @param {Object} delta - Stats delta to apply
 */
function applyDeltaToBatter(sheet, atBatRow, delta) {
  var sheetId = sheet.getSheetId().toString();
  var batterRow = getBatterRowFromAtBatCell(atBatRow);
  
  if (!batterRow) {
    logWarning("Hitting", "Could not determine batter row from at-bat row: " + atBatRow, sheet.getName());
    return;
  }
  
  var batterName = getPlayerNameFromBatterRow(sheet, batterRow);
  
  if (!batterName || batterName === "") {
    logWarning("Hitting", "No batter name found at row: " + batterRow, sheet.getName());
    return;
  }
  
  var batterStats = getBatterStats(sheetId);
  
  // Initialize if needed
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
  }
  
  // Apply delta to current values
  batterStats[batterName].AB = Math.max(0, (batterStats[batterName].AB || 0) + delta.AB);
  batterStats[batterName].H = Math.max(0, (batterStats[batterName].H || 0) + delta.H);
  batterStats[batterName].HR = Math.max(0, (batterStats[batterName].HR || 0) + delta.HR);
  batterStats[batterName].RBI = Math.max(0, (batterStats[batterName].RBI || 0) + delta.RBI);
  batterStats[batterName].BB = Math.max(0, (batterStats[batterName].BB || 0) + delta.BB);
  batterStats[batterName].K = Math.max(0, (batterStats[batterName].K || 0) + delta.K);
  batterStats[batterName].ROB = Math.max(0, (batterStats[batterName].ROB || 0) + delta.ROB);
  batterStats[batterName].DP = Math.max(0, (batterStats[batterName].DP || 0) + delta.DP);
  batterStats[batterName].TB = Math.max(0, (batterStats[batterName].TB || 0) + delta.TB);
  
  // Save back to properties
  saveBatterStats(sheetId, batterStats);
  
  // Update sheet
  updateBatterRow(sheet, batterRow, batterStats[batterName]);
  
  logInfo("Hitting", "Updated " + batterName + ": AB=" + batterStats[batterName].AB + 
          ", H=" + batterStats[batterName].H + ", HR=" + batterStats[batterName].HR + 
          ", RBI=" + batterStats[batterName].RBI);
}

/**
 * Update batter row in sheet using batch write
 * @param {Sheet} sheet - The game sheet
 * @param {number} batterRow - Batter row number
 * @param {Object} stats - Batter stats
 */
function updateBatterRow(sheet, batterRow, stats) {
  // Skip protected rows
  if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(batterRow) !== -1) {
    logWarning("Hitting", "Skipped protected row: " + batterRow, "");
    return;
  }
  
  var cols = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS;
  
  // Build batch write array: [[AB, H, HR, RBI, BB, K, ROB, DP, TB]]
  var statsArray = [[
    stats.AB,
    stats.H,
    stats.HR,
    stats.RBI,
    stats.BB,
    stats.K,
    stats.ROB,
    stats.DP,
    stats.TB
  ]];
  
  // Single batch write (columns C through K)
  sheet.getRange(batterRow, cols.AB, 1, 9).setValues(statsArray);
}

/**
 * Calculate hitting delta from parsed notation
 * @param {Object} oldStats - Old parsed stats
 * @param {Object} newStats - New parsed stats
 * @return {Object} Hitting delta
 */
function calculateHittingDelta(oldStats, newStats) {
  return {
    AB: newStats.AB - oldStats.AB,
    H: newStats.H - oldStats.H,
    HR: newStats.HR - oldStats.HR,
    RBI: newStats.R - oldStats.R,  // Note: R from notation = RBI for batter
    BB: newStats.BB - oldStats.BB,
    K: newStats.K - oldStats.K,
    ROB: 0,  // ROB is handled by defensive module
    DP: (newStats.DP ? 1 : 0) - (oldStats.DP ? 1 : 0),
    TB: newStats.TB - oldStats.TB
  };
}

/**
 * Add ROB (Run On Base - hit into nice play) to batter
 * Called by defensive module when NP occurs
 * @param {Sheet} sheet - The game sheet
 * @param {number} atBatRow - At-bat cell row
 */
function addROBToBatter(sheet, atBatRow) {
  var sheetId = sheet.getSheetId().toString();
  var batterRow = getBatterRowFromAtBatCell(atBatRow);
  
  if (!batterRow) {
    logWarning("Hitting", "Could not determine batter row for ROB: " + atBatRow, sheet.getName());
    return;
  }
  
  var batterName = getPlayerNameFromBatterRow(sheet, batterRow);
  
  if (!batterName || batterName === "") {
    logWarning("Hitting", "No batter name found for ROB at row: " + batterRow, sheet.getName());
    return;
  }
  
  var batterStats = getBatterStats(sheetId);
  
  // Initialize if needed
  if (!batterStats[batterName]) {
    initializeBatterIfNeeded(sheetId, batterName);
    batterStats = getBatterStats(sheetId);
  }
  
  // Add 1 ROB
  batterStats[batterName].ROB = (batterStats[batterName].ROB || 0) + 1;
  
  // Save and update
  saveBatterStats(sheetId, batterStats);
  updateBatterRow(sheet, batterRow, batterStats[batterName]);
  
  logInfo("Hitting", "Added ROB to " + batterName);
}

/**
 * Add stolen base to batter
 * NOTE: SB goes in column R of the FIELDING section (rows 7-15, 18-26)
 * NOT the hitting section (rows 30-38, 41-49)
 * @param {Sheet} sheet - The game sheet
 * @param {number} atBatRow - At-bat cell row
 */
function addStolenBaseToBatter(sheet, atBatRow) {
  // Determine which team is batting
  var batterTeam = getBattingTeam(atBatRow);
  
  if (!batterTeam) {
    logWarning("Hitting", "Could not determine batting team for SB: " + atBatRow, sheet.getName());
    return;
  }
  
  // Get batter row to find player name
  var batterRow = getBatterRowFromAtBatCell(atBatRow);
  
  if (!batterRow) {
    logWarning("Hitting", "Could not determine batter row for SB: " + atBatRow, sheet.getName());
    return;
  }
  
  var batterName = getPlayerNameFromBatterRow(sheet, batterRow);
  
  if (!batterName || batterName === "") {
    logWarning("Hitting", "No batter name found for SB at row: " + batterRow, sheet.getName());
    return;
  }
  
  // Find player in the FIELDING section (rows 7-15 or 18-26)
  var fieldingRange = (batterTeam === "away") ? 
    BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE : 
    BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  
  // Read player names from fielding section
  var playersData = sheet.getRange(
    fieldingRange.startRow,
    fieldingRange.nameCol,
    fieldingRange.numPlayers,
    1
  ).getValues();
  
  // Find matching player
  for (var i = 0; i < playersData.length; i++) {
    if (String(playersData[i][0]).trim() === batterName) {
      var playerFieldingRow = fieldingRange.startRow + i;
      
      // Write SB to column R in fielding section
      var sbCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.SB;
      var currentSB = sheet.getRange(playerFieldingRow, sbCol).getValue() || 0;
      sheet.getRange(playerFieldingRow, sbCol).setValue(currentSB + 1);
      
      logInfo("Hitting", "Added SB to " + batterName + " at fielding row " + playerFieldingRow + ", col " + sbCol);
      return;
    }
  }
  
  logWarning("Hitting", "Could not find " + batterName + " in fielding roster for SB", sheet.getName());
}