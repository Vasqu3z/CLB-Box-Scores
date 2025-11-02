// ===== BOX SCORE PITCHING MODULE =====
// Handles pitcher stat tracking with delta calculation
// Attributes stats to active pitcher based on C3/C4 dropdowns

/**
 * Handle pitcher dropdown change
 * @param {Sheet} sheet - The game sheet
 * @param {string} cell - Cell address (C3 or C4)
 * @param {string} newPitcher - New pitcher name
 */
function handlePitcherChange(sheet, cell, newPitcher) {
  var sheetId = sheet.getSheetId().toString();
  var activePitchers = getActivePitchers(sheetId);
  
  // Update active pitcher
  if (cell === BOX_SCORE_CONFIG.AWAY_PITCHER_CELL) {
    activePitchers.away = newPitcher.trim();
  } else {
    activePitchers.home = newPitcher.trim();
  }
  
  saveActivePitchers(sheetId, activePitchers);
  initializePitcherIfNeeded(sheetId, newPitcher.trim());
  
  logInfo("Pitching", "Pitcher changed: " + cell + " = " + newPitcher);
}

/**
 * Apply delta to pitcher stats
 * @param {Sheet} sheet - The game sheet
 * @param {string} pitcherName - Pitcher name
 * @param {Object} delta - Stats delta to apply
 */
function applyDeltaToPitcher(sheet, pitcherName, delta) {
  var sheetId = sheet.getSheetId().toString();
  var pitcherStats = getPitcherStats(sheetId);
  
  // Initialize if needed
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
  }
  
  // Apply delta to current values
  pitcherStats[pitcherName].BF = Math.max(0, (pitcherStats[pitcherName].BF || 0) + delta.BF);
  pitcherStats[pitcherName].outs = Math.max(0, (pitcherStats[pitcherName].outs || 0) + delta.outs);
  pitcherStats[pitcherName].H = Math.max(0, (pitcherStats[pitcherName].H || 0) + delta.H);
  pitcherStats[pitcherName].HR = Math.max(0, (pitcherStats[pitcherName].HR || 0) + delta.HR);
  pitcherStats[pitcherName].R = Math.max(0, (pitcherStats[pitcherName].R || 0) + delta.R);
  pitcherStats[pitcherName].BB = Math.max(0, (pitcherStats[pitcherName].BB || 0) + delta.BB);
  pitcherStats[pitcherName].K = Math.max(0, (pitcherStats[pitcherName].K || 0) + delta.K);
  
  // Calculate IP
  var ip = calculateIP(pitcherStats[pitcherName].outs);
  
  // Save back to properties
  savePitcherStats(sheetId, pitcherStats);
  
  // Update sheet
  updatePitcherRow(sheet, pitcherName, pitcherStats[pitcherName], ip);
  
  logInfo("Pitching", "Updated " + pitcherName + ": BF=" + pitcherStats[pitcherName].BF + 
          ", IP=" + ip.toFixed(2) + ", H=" + pitcherStats[pitcherName].H + 
          ", K=" + pitcherStats[pitcherName].K);
}

/**
 * Update pitcher row in sheet using batch write
 * @param {Sheet} sheet - The game sheet
 * @param {string} pitcherName - Pitcher name
 * @param {Object} stats - Pitcher stats
 * @param {number} ip - Innings pitched
 */
function updatePitcherRow(sheet, pitcherName, stats, ip) {
  var ranges = [
    BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE,
    BOX_SCORE_CONFIG.HOME_PITCHER_RANGE
  ];
  
  // Search for pitcher in both teams
  for (var r = 0; r < ranges.length; r++) {
    var range = ranges[r];
    
    // Read player names
    var pitcherData = sheet.getRange(
      range.startRow,
      range.nameCol,
      range.endRow - range.startRow + 1,
      1
    ).getValues();
    
    // Find matching pitcher
    for (var i = 0; i < pitcherData.length; i++) {
      if (String(pitcherData[i][0]).trim() === pitcherName) {
        var row = range.startRow + i;
        
        // Skip protected rows
        if (BOX_SCORE_CONFIG.PROTECTED_ROWS.indexOf(row) !== -1) {
          logWarning("Pitching", "Skipped protected row: " + row, pitcherName);
          continue;
        }
        
        var cols = BOX_SCORE_CONFIG.PITCHER_STATS_COLUMNS;
        
        // Build batch write array: [[IP, BF, H, HR, R, BB, K]]
        // Order matches columns I through O
        var statsArray = [[
          ip,
          stats.BF,
          stats.H,
          stats.HR,
          stats.R,
          stats.BB,
          stats.K
        ]];
        
        // Single batch write (columns I through O)
        sheet.getRange(row, cols.IP, 1, 7).setValues(statsArray);
        
        return;  // Found and updated, exit
      }
    }
  }
  
  logWarning("Pitching", "Pitcher not found in roster: " + pitcherName, sheet.getName());
}

/**
 * Get active pitcher for a batting team
 * @param {string} sheetId - Sheet ID
 * @param {string} batterTeam - "away" or "home"
 * @return {string} Active pitcher name
 */
function getActivePitcherForBatter(sheetId, batterTeam) {
  var activePitchers = getActivePitchers(sheetId);
  
  // Away batters face home pitcher, home batters face away pitcher
  if (batterTeam === "away") {
    return activePitchers.home;
  } else {
    return activePitchers.away;
  }
}