### The "Gold Standard" Principles

All refactoring must adhere to these four principles:

1.  **P1 (Performance): Read Once, Write Once.** All I/O (Input/Output) with `SpreadsheetApp` must be batched. There will be **no** `.getValue()`, `.setValue()`, or `.setFormula()` calls inside loops.
2.  **P2 (Configurability): No Magic Numbers.** All sheet names, row/column/range values, and logical thresholds must be defined in and read from the `ScoreConfig.js` file. This includes **0-based relative indices** for processing arrays from batch reads.
3.  **P3 (Data Flow):** (Already Compliant) The "Absolute State" engine in `processGameStatsBulk` is a model of P3 compliance.
4.  **P4 (Commenting): Professional & Structural.** Comments must be clean, standardized, and structural. They must provide JSDoc context and high-level section guidance, not explain obvious code.

-----

### Part 1: Global Commenting Standard (P4)

Apply this standard to *every* `.js` file in the suite.

1.  **File Headers:** Every `.js` file must begin with this 4-line header, replacing placeholder text.

    ```javascript
    // ===== {MODULE_NAME_IN_CAPS} =====
    // Purpose: {Brief, one-line description of the file's responsibility.}
    // Dependencies: {Key config files or modules, e.g., ScoreConfig.js, ScoreNotation.js}
    // Entry Point(s): {Primary function(s) called from other modules, e.g., onEdit, processGameStatsBulk}
    ```

2.  **Section Headers:** Use `// ===== {SECTION_NAME_IN_CAPS} =====` to break up files logically (e.g., `// ===== STAT VIEWERS =====`, `// ===== HELPER FUNCTIONS =====`).

3.  **Function Headers (JSDoc):** Ensure all functions, except for trivial private helpers, have a JSDoc header (the suite is already very good at this).

    ```javascript
    /**
     * {Brief description of what the function does.}
     * @param {Sheet} sheet - The game sheet.
     * @param {Object} playerStats - Stats storage object.
     * @param {Object} rosterMap - Player roster map.
     */
    function writeStatsToSheet(sheet, playerStats, rosterMap) { ... }
    ```

4.  **Logging Standardization (P4):** The suite has two logging systems (`Logger.log()` vs. the functions in `ScoreUtility.js`). Standardize on a single, config-driven model.

      * **Action (in `ScoreConfig.js`):** Add a new `DEBUG` object inside the `BOX_SCORE_CONFIG` object.
        ```javascript
        var BOX_SCORE_CONFIG = {
          GAME_SHEET_PREFIX: "#",
          // ...
          AUTO_PROCESS_ON_AT_BAT: true,
          
          // ADD THIS
          DEBUG: {
            ENABLE_LOGGING: true
          },
          
          // ===== PITCHER TRACKING =====
          // ...
        ```
      * **Action (in `ScoreUtility.js`):**
        1.  Delete the functions `logInfo`, `logWarning`, and `logError`.
      * **Action (in all other files, e.g., `ScoreTriggers.js`, `ScoreMenu.js`):**
        1.  Replace all calls to the deleted `log...` functions with the new standard.
      * **Example (Before):** `logError("Triggers", error.toString(), sheetName + "!" + cell);`
      * **Example (After):**
        ```javascript
        if (BOX_SCORE_CONFIG.DEBUG.ENABLE_LOGGING) {
          Logger.log("ERROR [Triggers]: " + error.toString() + " (Entity: " + sheetName + "!" + cell + ")");
        }
        ```

-----

### Part 2: Configuration File Refactor (P2)

**File:** `ScoreConfig.js`

  * **Principle Violation:** P2 (Missing relative indices). The config correctly defines 1-based *absolute* columns (e.g., `IP: 9`), but it lacks 0-based *relative* indices for processing arrays read from those columns.

  * **Task:** Add relative 0-based index maps to `BOX_SCORE_CONFIG`.

  * **Action:** Add the following three new objects inside the `BOX_SCORE_CONFIG` object.

    ```javascript
    // ADD THIS (near PITCHER_STATS_COLUMNS)
    // 0-based relative indices for processing pitcher stat arrays
    // (Array starts at IP, so IP is index 0)
    PITCHER_STATS_INDICES: {
      IP: 0,    // (Column I = 9) - 9 = 0
      BF: 1,    // (Column J = 10) - 9 = 1
      H: 2,
      HR: 3,
      R: 4,
      BB: 5,
      K: 6
    },

    // ADD THIS (near FIELDING_STATS_COLUMNS)
    // 0-based relative indices for processing fielding stat arrays
    // (Array starts at NP, so NP is index 0)
    FIELDING_STATS_INDICES: {
      NP: 0,    // (Column P = 16) - 16 = 0
      E: 1,
      SB: 2
    },

    // ADD THIS (near HITTING_STATS_COLUMNS)
    // 0-based relative indices for processing hitting stat arrays
    // (Array starts at AB, so AB is index 0)
    HITTING_STATS_INDICES: {
      AB: 0,    // (Column C = 3) - 3 = 0
      H: 1,     // (Column D = 4) - 3 = 1
      HR: 2,
      RBI: 3,
      BB: 4,
      K: 5,
      ROB: 6,
      DP: 7,
      TB: 8
    },
    ```

-----

### Part 3: Module-Specific Refactoring (P1 & P2)

#### `ScoreTriggers.js`

  * **Principle Violation:** P1 (Critical N+1 Bottleneck).
  * **Task:** Refactor `writeStatsToSheet` to use batch writes, eliminating all `.setValues()` calls from inside the loop.
  * **Action:** Rewrite the `writeStatsToSheet` function as follows:
    ```javascript
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
    ```

#### `ScoreUtility.js`

  * **Principle Violation:** P1 (Helper function missing), P4 (Redundant logging).
  * **Task 1 (P4):** Delete the functions `logInfo`, `logWarning`, and `logError` (as defined in Part 1).
  * **Task 2 (P1):** Add the `createEmptyBatch` helper function required by the `writeStatsToSheet` refactor.
  * **Action:** Add the following function to `ScoreUtility.js`.
    ```javascript
    /**
     * Creates a 2D array pre-filled with zeros for batch operations.
     * @param {number} numRows - The number of rows.
     * @param {number} numCols - The number of columns.
     * @return {Array<Array<number>>} A 2D array.
     */
    function createEmptyBatch(numRows, numCols) {
      var batch = [];
      for (var r = 0; r < numRows; r++) {
        var row = [];
        for (var c = 0; c < numCols; c++) {
          row.push(0); // Default to 0
        }
        batch.push(row);
      }
      return batch;
    }
    ```

#### `ScoreMenu.js`

  * **Principle Violation:** P2 (Brittle, hardcoded array indices).
  * **Task:** Refactor `showPitcherStats` and `showBatterStats` to use the new 0-based `..._INDICES` maps from the config.
  * **Action (for `showPitcherStats`):**
      * **Before:**
        ```javascript
        var stats = awayStats[i]; // [IP, BF, H, HR, R, BB, K]
        if (name && stats[0] > 0) { // Check if pitcher has pitched (IP > 0)
          message += paddedName + ": " +
                     stats[0].toFixed(2) + " IP, " +
                     stats[2] + " H, " +
                     stats[4] + " R, " +
                     stats[5] + " BB, " +
                     stats[6] + " K\n";
        ```
      * **After:**
        ```javascript
        var pIdx = BOX_SCORE_CONFIG.PITCHER_STATS_INDICES; // Get 0-based index map
        // ...
        var stats = awayStats[i];
        if (name && stats[pIdx.IP] > 0) { // Check IP by its mapped index
          message += paddedName + ": " +
                     stats[pIdx.IP].toFixed(2) + " IP, " +
                     stats[pIdx.H] + " H, " +
                     stats[pIdx.R] + " R, " +
                     stats[pIdx.BB] + " BB, " +
                     stats[pIdx.K] + " K\n";
        ```
  * **Action (for `showBatterStats`):**
      * **Before:**
        ```javascript
        var stats = awayStats[i]; // [AB, H, HR, RBI, BB, K, ROB, DP, TB]
        if (name && stats[0] > 0) { // Check if batter has AB > 0
          var line = (i + 1) + ". " + paddedName + ": " + stats[1] + "-" + stats[0];
          // ...
          if (stats[2] > 0) { line += ", " + stats[2] + "HR"; }
          // ...
          var otherXBH = stats[8] - stats[1] - (stats[2] * 3);
        ```
      * **After:**
        ```javascript
        var hIdx = BOX_SCORE_CONFIG.HITTING_STATS_INDICES; // Get 0-based index map
        // ...
        var stats = awayStats[i];
        if (name && stats[hIdx.AB] > 0) { // Check AB by its mapped index
          var line = (i + 1) + ". " + paddedName + ": " + stats[hIdx.H] + "-" + stats[hIdx.AB];
          // ...
          if (stats[hIdx.HR] > 0) { line += ", " + stats[hIdx.HR] + "HR"; }
          // ...
          var otherXBH = stats[hIdx.TB] - stats[hIdx.H] - (stats[hIdx.HR] * 3);
        ```

This completes the refactor for the "Box Score" suite, bringing it into full compliance by fixing the critical P1 write-loop and standardizing its P2 configuration.