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
  
  // Only run on game sheets (sheets starting with #)
  if (!sheetName.startsWith("#")) return;
  
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
    
    // Handle position swap FIRST (before pitcher change tracking)
    if (oldValue && newValue && oldValue !== newValue) {
      handlePositionSwap(sheet, oldValue, newValue);
    }
    
    // Then handle pitcher change tracking (existing logic)
    handlePitcherChange(sheet, cell, newValue);
    return;
  }
  
  // Handle at-bat entries
  if (isAtBatCell(row, col)) {
    // Check if this is a range paste (multiple cells at once)
    var numRows = range.getNumRows();
    var numCols = range.getNumColumns();
    var totalCells = numRows * numCols;
    
    if (numRows > 1 || numCols > 1) {
      // Count how many cells are actually in at-bat range
      var atBatCount = 0;
      for (var r = row; r < row + numRows; r++) {
        for (var c = col; c < col + numCols; c++) {
          if (isAtBatCell(r, c)) {
            atBatCount++;
          }
        }
      }
      
      // Warn based on actual at-bat cells (not total cells)
      var ui = SpreadsheetApp.getUi();
      var thresholds = BOX_SCORE_CONFIG.PASTE_THRESHOLDS;
      
      if (atBatCount > thresholds.DANGER) {
        // High risk of timeout
        var response = ui.alert(
          'Large Paste Warning',
          'You are pasting ' + atBatCount + ' at-bat cells.\n\n' +
          '⚠️ This will likely TIMEOUT before completing.\n\n' +
          'Recommendation:\n' +
          '• Paste smaller sections (≤20 cells at a time)\n' +
          '• Or paste all, then use "Reprocess Selected At-Bat"\n' +
          '  for cells that didn\'t process\n\n' +
          'Continue anyway?',
          ui.ButtonSet.YES_NO
        );
        
        if (response !== ui.Button.YES) {
          return; // User cancelled
        }
      } else if (atBatCount > thresholds.CAUTION) {
        // Medium risk - inform but don't block
        var response = ui.alert(
          'Medium Paste Detected',
          'You are pasting ' + atBatCount + ' at-bat cells.\n\n' +
          'This may take 15-20 seconds to process.\n' +
          'Some cells may timeout.\n\n' +
          'Continue?',
          ui.ButtonSet.YES_NO
        );
        
        if (response !== ui.Button.YES) {
          return; // User cancelled
        }
      }
      
      // Multi-cell paste - process each cell in the range
      handleRangePaste(sheet, range);
    } else {
      // Single cell edit - normal processing
      handleAtBatEntry(sheet, cell, row, col, newValue);
    }
    return;
  }
}

// ============================================
// Handle Position Swaps
// ============================================

/**
 * Handle range paste (copy/paste multiple cells)
 * @param {Sheet} sheet - The game sheet
 * @param {Range} range - Pasted range
 */
function handleRangePaste(sheet, range) {
  var startRow = range.getRow();
  var startCol = range.getColumn();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();
  var totalCells = numRows * numCols;
  
  // Show progress for large pastes (from config)
  var toastThreshold = BOX_SCORE_CONFIG.PASTE_THRESHOLDS.TOAST;
  if (totalCells > toastThreshold) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Processing ' + totalCells + ' cells...',
      'Range Paste',
      3
    );
  }
  
  // Get all values in the range
  var values = range.getValues();
  
  var processedCount = 0;
  
  // Process each cell in the pasted range
  for (var r = 0; r < numRows; r++) {
    for (var c = 0; c < numCols; c++) {
      var row = startRow + r;
      var col = startCol + c;
      
      // Only process if it's an at-bat cell
      if (isAtBatCell(row, col)) {
        var cell = columnToLetter(col) + row;
        var newValue = values[r][c] || "";
        
        handleAtBatEntry(sheet, cell, row, col, newValue);
        processedCount++;
      }
    }
  }
  
  // Show completion for large pastes
  if (totalCells > toastThreshold) {
    SpreadsheetApp.getActiveSpreadsheet().toast(
      'Processed ' + processedCount + ' at-bats',
      'Complete',
      2
    );
  }
  
  logInfo("Triggers", "Processed range paste: " + range.getA1Notation() + " (" + processedCount + " cells)");
}

/**
 * Helper: Convert column number to letter (1→A, 2→B, etc.)
 */
function columnToLetter(column) {
  var letter = '';
  while (column > 0) {
    var temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
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
 * Handle at-bat entry - orchestrates all stat updates
 * @param {Sheet} sheet - The game sheet
 * @param {string} cell - Cell address (e.g., "C7")
 * @param {number} row - Row number
 * @param {number} col - Column number
 * @param {string} newValue - New cell value
 */
function handleAtBatEntry(sheet, cell, row, col, newValue) {
  var sheetId = sheet.getSheetId().toString();
  
  // ===== STEP 1: Get shadow storage =====
  var gameState = getGameState(sheetId);
  var oldValue = gameState[cell] || "";
  
  // ===== STEP 2: Get active pitchers =====
  var activePitchers = getActivePitchers(sheetId);
  var batterTeam = getBattingTeam(row);
  
  if (!batterTeam) {
    logWarning("Triggers", "Could not determine batting team", cell);
    return;
  }
  
  var pitcher = getActivePitcherForBatter(sheetId, batterTeam);
  
  if (!pitcher || pitcher === "") {
    // No pitcher selected yet - just update shadow storage and return
    gameState[cell] = newValue;
    saveGameState(sheetId, gameState);
    logInfo("Triggers", "At-bat entered but no active pitcher: " + cell);
    return;
  }
  
  // ===== STEP 3: Parse notation =====
  var oldStats = parseNotation(oldValue);
  var newStats = parseNotation(newValue);
  
  // ===== STEP 4: Calculate deltas =====
  var pitchingDelta = {
    BF: newStats.BF - oldStats.BF,
    outs: newStats.outs - oldStats.outs,
    H: newStats.H - oldStats.H,
    HR: newStats.HR - oldStats.HR,
    R: newStats.R - oldStats.R,
    BB: newStats.BB - oldStats.BB,
    K: newStats.K - oldStats.K
  };
  
  var hittingDelta = calculateHittingDelta(oldStats, newStats);
  
  // ===== STEP 5: Apply pitching stats =====
  applyDeltaToPitcher(sheet, pitcher, pitchingDelta);
  
  // ===== STEP 6: Apply hitting stats =====
  applyDeltaToBatter(sheet, row, hittingDelta);
  
  // ===== STEP 7: Apply defensive stats (NP, E, SB) =====
  // Check both old and new for defensive codes (to handle additions and removals)
  if (oldStats.NP || newStats.NP || oldStats.E || newStats.E || oldStats.SB || newStats.SB) {
    applyDefensiveStats(sheet, batterTeam, row, oldStats, newStats);
  }
  
  // ===== STEP 8: Update shadow storage =====
  gameState[cell] = newValue;
  saveGameState(sheetId, gameState);
  
  logInfo("Triggers", "At-bat processed: " + cell + " = '" + newValue + "' (Pitcher: " + pitcher + ")");
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