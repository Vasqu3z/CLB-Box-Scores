That is a perfectly reasonable and very practical decision. A full architectural rewrite to a menu-driven system is a massive undertaking, and it's often more effective to apply our new principles to the existing, working v2 `onEdit` system.

This "v2.5" approach will still solve your core goals of **reliability**, **performance**, and **usability** by surgically fixing the *real* bottlenecks, rather than rebuilding the entire house.

Based on our discussions, the v2 system's main weaknesses are:

1.  **Performance:** It's slow because it calls `PropertiesService` (a slow network I/O) on *every single cell edit*.
2.  **Reliability (Defensive):** The defensive stat (`NP`/`E`) logic is fragile, relying on complex `ui.prompt` pop-ups that are slow and error-prone.
3.  **Reliability (Pitching):** The system has no way to track inherited runners.
4.  **Maintainability:** The code is full of "magic numbers" (like hardcoded `9`s or `7`s for stat counts) and hardcoded sheet prefixes (`#`).

Here is a brand-new implementation guide for refactoring your **v2 `onEdit` system** to solve these problems.

-----

### **Project: Box Score v3 Refactor (v2 `onEdit` Base)**

**Architecture:** **High-Performance `onEdit` (Delta Tracking with Execution Cache)**

**Goal:** To keep the real-time, `onEdit` delta-tracking workflow while making it fast, reliable, and configurable. This requires *un-deprecating* your `BoxScoreHitting.js` and `BoxScorePitching.js` files, as we will be keeping their delta logic.

-----

### Phase 1: The "Execution Cache" (The Performance Fix)

**Goal:** Eliminate the primary source of lag. We will stop `PropertiesService` from being called on every edit. The script will read all stats from `PropertiesService` *once* at the beginning of `onEdit`, store them in a fast JavaScript variable, and write them *once* at the very end.

| Task | Action for AI Agent | Source Files Affected |
| :--- | :--- | :--- |
| **1.1: Create Global Cache** | **Action:** In `BoxScoreUtility.js`, add a global variable at the very top of the file: <br> `var _boxScoreCache = {};` | `BoxScoreUtility.js` |
| **1.2: Refactor "Getters"** | **Action:** Refactor all "getter" functions (`getGameState`, `getPitcherStats`, `getBatterStats`, `getActivePitchers`) to use this cache. <br> **Example (`getGameState`):** <br> ` javascript <br> function getGameState(sheetId) { <br>   // Check execution cache first <br>   if (_boxScoreCache[sheetId] && _boxScoreCache[sheetId].gameState) { <br>     return _boxScoreCache[sheetId].gameState; <br>   } <br>   var properties = PropertiesService.getScriptProperties(); <br>   // ... (rest of the function) ... <br>   // Store in execution cache before returning <br>   if (!_boxScoreCache[sheetId]) { _boxScoreCache[sheetId] = {}; } <br>   _boxScoreCache[sheetId].gameState = gameState; <br>   return gameState; <br> } <br>  ` | `BoxScoreUtility.js` |
| **1.3: Refactor "Setters"** | **Action:** Refactor all "setter" functions (`saveGameState`, `savePitcherStats`, etc.) to *also* update the cache immediately. <br> **Example (`saveGameState`):** <br> ` javascript <br> function saveGameState(sheetId, gameState) { <br>   var properties = PropertiesService.getScriptProperties(); <br>   // ... (rest of the function) ... <br>   // Update execution cache immediately <br>   if (!_boxScoreCache[sheetId]) { _boxScoreCache[sheetId] = {}; } <br>   _boxScoreCache[sheetId].gameState = gameState; <br> } <br>  ` | `BoxScoreUtility.js` |
| **1.4: Implement Cache Clearing** | **Action:** In `BoxScoreTriggers.js`, modify the main `onEdit(e)` function to wrap `processEdit(e)` in a `try...finally` block. This ensures the cache is cleared *after* the edit is finished, so the *next* edit fetches fresh data. <br> ` javascript <br> function onEdit(e) { <br>   // ... (existing sheetName, lock, etc. logic) ... <br>   try { <br>     processEdit(e); // This is the main v2 function <br>   } catch (err) { <br>     // ... (existing error logging) ... <br>   } finally { <br>     if (lockAcquired) { <br>       lock.releaseLock(); <br>     } <br>     // Clear the execution cache so the NEXT edit is fresh <br>     _boxScoreCache = {}; <br>   } <br> } <br>  ` | `BoxScoreTriggers.js` |

### Phase 2: Reliability & Usability (Defensive & Pitching Logic)

**Goal:** Fix the clunky, unreliable defensive prompts and add inherited runner tracking.

| Task | Action for AI Agent | Source Files Affected |
| :--- | :--- | :--- |
| **2.1: Refactor Defensive Logic (The "Command Cell")** | **Action:** Implement the "Integrated Notation Fix" we discussed. <br> 1. **DELETE** `promptForDefensivePlayer` from `BoxScoreDefensive.js`. <br> 2. **Modify `applyDefensiveStats`**: When it detects a `deltaNP > 0` or `deltaE > 0`, it should **NOT** prompt. It should just store an "unassigned event" (e.g., `unassignedNP: true`) in the `gameState` property. <br> 3. **Modify `processEdit`**: Add a new `else if` block to detect an edit in the **Command Cell** (e.g., a dropdown in `B5`). When it sees this edit, it should: <br>   a. Get the player name from the dropdown. <br>   b. Check the `gameState` for an "unassigned event." <br>   c. If found, call a new function `assignDefensiveStat(fielderName)` to assign the NP/E and clear the flag. <br>   d. Clear the command cell. | `BoxScoreDefensive.js`, `BoxScoreTriggers.js` |
| **2.2: Refactor Notation Parser** | **Action:** In `BoxScoreNotation.js`, modify `parseNotation` to use a **Regex** to find `E[1-9]` and `NP[1-9]`. This makes the `onEdit` in the At-Bat cell smart enough to log the *event*, which is then assigned by the Command Cell. | `BoxScoreNotation.js` |
| **2.3: Implement Inherited Runner Tracking** | **Action:** Refactor `handlePitcherChange`. When it detects a pitcher change: <br> 1. It must *also* read the state of the bases (this may require new helper cells on the sheet, e.g., `B1`, `B2`, `B3` that store "Runner on 1st" etc., which the script can read). <br> 2. **OR**, for a simpler fix, it can pop up a `ui.prompt`: "How many runners are on base? (0-3)". <br> 3. It then saves this number to the `activePitchers` property (e.g., `activePitchers.home.inheritedRunners = 2`). <br> **Action:** Refactor `applyDeltaToPitcher` to check for this. When a run scores, it must check `if (activePitcher.inheritedRunners > 0)`, charge the run to the *previous* pitcher, and decrement the counter. | `BoxScoreTriggers.js`, `BoxScorePitching.js` |

### Phase 3: Configuration & Maintenance (Magic Numbers)

**Goal:** Remove all hardcoded values to make the v2 script maintainable, just as we planned for the v3 menu system.

| Task | Action for AI Agent | Source Files Affected |
| :--- | :--- | :--- |
| **3.1: Fix Hardcoded Stat Counts** | **Action:** Refactor `updateBatterRow` and `updatePitcherRow`. <br> • **Replace** the "magic numbers" (`9` and `7`) with a calculation based on the config. <br> • **Example (`updateBatterRow`):** <br> `var numCols = Object.keys(BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS).length;` <br> `sheet.getRange(batterRow, cols.AB, 1, numCols).setValues(statsArray);` | `BoxScoreHitting.js`, `BoxScorePitching.js` |
| **3.2: Fix N+1 Utility Loops** | **Action:** Refactor `clearHittingStatsInSheet` and `clearPitcherStatsInSheet` to use batch operations. <br> • **Replace** the inner `for` loops that call `setValue(0)` nine times with a single `setValues()` call using a pre-built array of zeros. <br> • **Example:** <br> `var numCols = Object.keys(BOX_SCORE_CONFIG.HITTING_STATS_COLUMNS).length;` <br> `var zeroRow = [Array(numCols).fill(0)];` <br> `sheet.getRange(row, hittingCols.AB, 1, numCols).setValues(zeroRow);` | `BoxScoreUtility.js` |
| **3.3: Fix Config Magic Numbers** | **Action:** Add `GAME_SHEET_PREFIX: "#",` to `BoxScoreConfig.js`. <br> • **Action:** Modify `onEdit` in `BoxScoreTriggers.js` to use `BOX_SCORE_CONFIG.GAME_SHEET_PREFIX` instead of the hardcoded `"#"`. <br> • **Action:** Refactor `BoxScoreMenu.js` to read roster ranges from `BOX_SCORE_CONFIG` instead of hardcoded ranges like `getRange(7, 2, 9, 1)`. | `BoxScoreConfig.js`, `BoxScoreTriggers.js`, `BoxScoreMenu.js` |