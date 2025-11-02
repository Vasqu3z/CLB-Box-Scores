// ===== BOX SCORE DEFENSIVE STATS MODULE =====
// Handles defensive stats (NP, E) with interactive prompts
// Also handles stolen bases (SB) and ROB tracking

// ============================================
// DEFENSIVE STATS WITH DELTA
// ============================================

/**
 * Apply defensive stats with delta support (handles additions and removals)
 * @param {Sheet} sheet - The game sheet
 * @param {string} batterTeam - "away" or "home"
 * @param {number} atBatRow - At-bat cell row
 * @param {Object} oldStats - Old parsed stats 
 * @param {Object} newStats - New parsed stats
 */
function applyDefensiveStats(sheet, batterTeam, atBatRow, oldStats, newStats) {
  // Calculate deltas
  var oldNP = oldStats.NP ? 1 : 0;
  var newNP = newStats.NP ? 1 : 0;
  var oldE = oldStats.E ? 1 : 0;
  var newE = newStats.E ? 1 : 0;
  var oldSB = oldStats.SB ? 1 : 0;
  var newSB = newStats.SB ? 1 : 0;
  
  var deltaNP = newNP - oldNP;
  var deltaE = newE - oldE;
  var deltaSB = newSB - oldSB;
  
  var fieldingTeam = (batterTeam === "away") ? "home" : "away";
  
  // ===== NICE PLAY DELTA =====
  if (deltaNP > 0) {
    // Adding NP - prompt for fielder
    var npRow = promptForDefensivePlayer(sheet, fieldingTeam, "NP");
    
    if (npRow) {
      var npCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.NP;
      var currentNP = sheet.getRange(npRow, npCol).getValue() || 0;
      sheet.getRange(npRow, npCol).setValue(currentNP + 1);
      
      var playerName = sheet.getRange(npRow, BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.nameCol).getValue();
      
      // Store association for removal
      storeDefensiveAssociation(sheet, atBatRow, 'NP', playerName);
      
      logInfo("Defensive", "NP added to " + playerName + " at row " + npRow);
      
      // Add ROB to batter (hitting section)
      addROBToBatter(sheet, atBatRow);
    }
  } else if (deltaNP < 0) {
    // Removing NP - only process if there's a stored association
    var fielderName = getDefensiveAssociation(sheet, atBatRow, 'NP');
    if (fielderName) {
      removeNicePlayFromFielder(sheet, atBatRow);
    } else {
      // No stored association - NP was from before tracking started
      logInfo("Defensive", "Skipping NP removal - no association found (likely old data)");
    }
  }
  
  // ===== ERROR DELTA =====
  if (deltaE > 0) {
    // Adding E - prompt for fielder
    var eRow = promptForDefensivePlayer(sheet, fieldingTeam, "E");
    
    if (eRow) {
      var eCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.E;
      var currentE = sheet.getRange(eRow, eCol).getValue() || 0;
      sheet.getRange(eRow, eCol).setValue(currentE + 1);
      
      var playerName = sheet.getRange(eRow, BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE.nameCol).getValue();
      
      // Store association for removal (FIXED: correct parameter order)
      storeDefensiveAssociation(sheet, atBatRow, 'E', playerName);
      
      logInfo("Defensive", "E added to " + playerName + " at row " + eRow);
    }
  } else if (deltaE < 0) {
    // Removing E - only process if there's a stored association
    var fielderName = getDefensiveAssociation(sheet, atBatRow, 'E');
    if (fielderName) {
      removeErrorFromFielder(sheet, atBatRow);
    } else {
      // No stored association - E was from before tracking started
      logInfo("Defensive", "Skipping E removal - no association found (likely old data)");
    }
  }
  
  // ===== STOLEN BASE DELTA =====
  if (deltaSB > 0) {
    addStolenBaseToBatter(sheet, atBatRow);
  }
  // Note: No SB removal logic needed (SB doesn't get removed in practice)
}

/**
 * Prompt user to select defensive player BY POSITION
 * @param {Sheet} sheet - The game sheet
 * @param {string} fieldingTeam - "away" or "home"
 * @param {string} statType - "NP" or "E"
 * @return {number} Row number of selected player, or null if cancelled
 */
function promptForDefensivePlayer(sheet, fieldingTeam, statType) {
  var statName = (statType === "NP") ? "Nice Play" : "Error";
  var teamName = (fieldingTeam === "away") ? "AWAY" : "HOME";
  
  // Get fielding team roster
  var rosterRange = (fieldingTeam === "away") ?
    BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE :
    BOX_SCORE_CONFIG.HOME_PITCHER_RANGE;
  
  var startRow = rosterRange.startRow;
  var endRow = rosterRange.endRow;
  var posCol = rosterRange.positionCol;
  var nameCol = rosterRange.nameCol;
  
  // Build roster display with CURRENT positions
  var positions = sheet.getRange(startRow, posCol, endRow - startRow + 1, 1).getValues();
  var names = sheet.getRange(startRow, nameCol, endRow - startRow + 1, 1).getValues();
  
  var rosterDisplay = '';
  for (var i = 0; i < positions.length; i++) {
    if (positions[i][0] && names[i][0]) {
      var currentPos = getCurrentPosition(positions[i][0]);
      rosterDisplay += currentPos + ': ' + names[i][0] + '\n';
    }
  }
  
  // Show prompt
  var ui = SpreadsheetApp.getUi();
  var response = ui.prompt(
    'Select Fielder - ' + teamName,
    statName + ' - Enter position:\n\n' + rosterDisplay + '\nExample: SS, P, 1B, CF',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  var inputPosition = response.getResponseText().trim().toUpperCase();
  
  // Find player by current position
  var playerRow = findPlayerRowByPosition(sheet, inputPosition, fieldingTeam);
  
  if (playerRow === -1) {
    ui.alert('Position "' + inputPosition + '" not found. Please try again.');
    return promptForDefensivePlayer(sheet, fieldingTeam, statType); // Retry
  }
  
  return playerRow;
}

// ============================================
// STORAGE HELPERS
// ============================================

/**
 * Store defensive stat association (which fielder got the NP/E)
 */
function storeDefensiveAssociation(sheet, atBatRow, statType, fielderName) {
  var props = PropertiesService.getScriptProperties();
  var prefix = (statType === 'NP') ? 
    BOX_SCORE_CONFIG.PROPERTY_PREFIX_NP : 
    BOX_SCORE_CONFIG.PROPERTY_PREFIX_E;
  var key = prefix + '_' + sheet.getSheetId() + '_' + atBatRow;
  props.setProperty(key, fielderName);
}

/**
 * Get stored defensive stat association
 */
function getDefensiveAssociation(sheet, atBatRow, statType) {
  var props = PropertiesService.getScriptProperties();
  var prefix = (statType === 'NP') ? 
    BOX_SCORE_CONFIG.PROPERTY_PREFIX_NP : 
    BOX_SCORE_CONFIG.PROPERTY_PREFIX_E;
  var key = prefix + '_' + sheet.getSheetId() + '_' + atBatRow;
  return props.getProperty(key);
}

/**
 * Clear defensive stat association
 */
function clearDefensiveAssociation(sheet, atBatRow, statType) {
  var props = PropertiesService.getScriptProperties();
  var prefix = (statType === 'NP') ? 
    BOX_SCORE_CONFIG.PROPERTY_PREFIX_NP : 
    BOX_SCORE_CONFIG.PROPERTY_PREFIX_E;
  var key = prefix + '_' + sheet.getSheetId() + '_' + atBatRow;
  props.deleteProperty(key);
}

// ============================================
// REMOVAL FUNCTIONS
// ============================================

/**
 * Remove Nice Play from FIELDER when NP removed from at-bat
 * Also removes ROB from BATTER
 */
function removeNicePlayFromFielder(sheet, atBatRow) {
  // Get which fielder was credited with the NP
  var fielderName = getDefensiveAssociation(sheet, atBatRow, 'NP');
  
  if (!fielderName) {
    // No stored fielder - this is OK (might be first edit after game start)
    logInfo("Defensive", "No NP fielder association found for row " + atBatRow + " (likely no NP to remove)");
    return;
  }
  
  // Find fielder in roster
  var fielderRow = findPlayerRowByName(sheet, fielderName);
  if (fielderRow === -1) {
    logWarning("Defensive", "Fielder not found: " + fielderName + " (clearing stale association)", "");
    clearDefensiveAssociation(sheet, atBatRow, 'NP');
    return;
  }
  
  // Remove NP from FIELDER (fielding section, column P)
  var npCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.NP;
  var currentNP = sheet.getRange(fielderRow, npCol).getValue() || 0;
  
  // Safety check: only decrement if NP > 0
  if (currentNP > 0) {
    var newNP = Math.max(0, currentNP - 1);
    sheet.getRange(fielderRow, npCol).setValue(newNP);
    logInfo("Defensive", "Removed NP from fielder " + fielderName + " at row " + fielderRow + " (" + currentNP + " → " + newNP + ")");
  } else {
    logWarning("Defensive", "Fielder " + fielderName + " already has 0 nice plays (skipping removal)", "");
  }
  
  // Remove ROB from BATTER (hitting section, column I)
  removeROBFromBatter(sheet, atBatRow);
  
  // Clear stored association
  clearDefensiveAssociation(sheet, atBatRow, 'NP');
}

/**
 * Remove Error from FIELDER when E removed from at-bat
 */
function removeErrorFromFielder(sheet, atBatRow) {
  // Get which fielder was credited with the E
  var fielderName = getDefensiveAssociation(sheet, atBatRow, 'E');
  
  if (!fielderName) {
    // No stored fielder - this is OK (might be first edit after game start)
    logInfo("Defensive", "No E fielder association found for row " + atBatRow + " (likely no E to remove)");
    return;
  }
  
  // Find fielder in roster
  var fielderRow = findPlayerRowByName(sheet, fielderName);
  if (fielderRow === -1) {
    logWarning("Defensive", "Fielder not found: " + fielderName + " (clearing stale association)", "");
    clearDefensiveAssociation(sheet, atBatRow, 'E');
    return;
  }
  
  // Remove E from FIELDER (fielding section, column Q)
  var eCol = BOX_SCORE_CONFIG.FIELDING_STATS_COLUMNS.E;
  var currentE = sheet.getRange(fielderRow, eCol).getValue() || 0;
  
  // Safety check: only decrement if E > 0
  if (currentE > 0) {
    var newE = Math.max(0, currentE - 1);
    sheet.getRange(fielderRow, eCol).setValue(newE);
    logInfo("Defensive", "Removed E from fielder " + fielderName + " at row " + fielderRow + " (" + currentE + " → " + newE + ")");
  } else {
    logWarning("Defensive", "Fielder " + fielderName + " already has 0 errors (skipping removal)", "");
  }
  
  // Clear stored association
  clearDefensiveAssociation(sheet, atBatRow, 'E');
}

/**
 * Remove ROB from batter (when NP is removed)
 */
function removeROBFromBatter(sheet, atBatRow) {
  // Get hitting section row from at-bat row
  var hittingRow = getBatterRowFromAtBatCell(atBatRow);
  if (!hittingRow) {
    logWarning("Defensive", "Could not find hitting row for at-bat row " + atBatRow, "");
    return;
  }
  
  // Remove ROB (hitting section, column I)
  var robCol = BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS.ROB;
  var currentROB = sheet.getRange(hittingRow, robCol).getValue() || 0;
  var newROB = Math.max(0, currentROB - 1);
  sheet.getRange(hittingRow, robCol).setValue(newROB);
  
  logInfo("Defensive", "Removed ROB from batter at hitting row " + hittingRow);
}