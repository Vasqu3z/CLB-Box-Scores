Here is the comprehensive v3 implementation guide for the **CLB Box Score** script suite, incorporating all our architectural decisions.

This refactor will transition the suite from a fragile, real-time `onEdit` system to a robust, high-performance **Menu-Driven Bulk Processor**. This change solves all identified issues with reliability, performance, and complex notation by centralizing logic into a single, manually-triggered function that calculates stats from a "single source of truth": the raw At-Bat grid.

-----

### **Project: CLB Box Score v3 Refactor**

**Architecture:** **Menu-Driven Bulk Processor (Absolute State Logic)**

**Goal:** To refactor the Box Score automation suite to be 100% reliable, configurable, and high-performance. This will be achieved by:

1.  **Eliminating `onEdit` for Scoring:** Migrating all stat calculations from the `onEdit` trigger to a single, manually-triggered menu function.
2.  **Removing Delta Logic:** Deleting all complex "delta" tracking (`applyDeltaToBatter`, `applyDeltaToPitcher`) and `PropertiesService` stat caching, which are sources of data corruption and lag.
3.  **Implementing "Absolute State":** The new processor will read the raw At-Bat grid as the single source of truth and calculate *total* stats from scratch every time it's run.
4.  **Adopting New Notations:** Implementing the robust `PC[X]` and `E[1-9]` / `NP[1-9]` notations to handle complex baseball logic (inherited runners, fielder assignment) seamlessly within the bulk processor.

-----

### Phase 1: Configuration & Code Demolition

**Goal:** Strip all deprecated `onEdit` and delta-tracking logic from the suite to prepare for the new "Absolute State" processor.

| Task | Action for AI Agent | Source Files Affected |
| :--- | :--- | :--- |
| **1.1: Prune `BoxScoreConfig.js`** | **DELETE** all configuration variables related to the old, trigger-based system. They are no longer needed. <br> • **Delete:** `LOCK_TIMEOUT_MS`. <br> • **Delete:** `PASTE_THRESHOLDS`. <br> • **Delete:** All `PROPERTY_` keys (`PROPERTY_GAME_STATE`, `PROPERTY_PITCHER_STATS`, `PROPERTY_BATTER_STATS`, `PROPERTY_ACTIVE_PITCHERS`, `PROPERTY_PREFIX_NP`, `PROPERTY_PREFIX_E`). | `BoxScoreConfig.js` |
| **1.2: Add Config Prefix** | **ADD** the new configuration variable `GAME_SHEET_PREFIX: "#",` to `BoxScoreConfig.js` to replace the hardcoded prefix in the `onEdit` trigger. | `BoxScoreConfig.js` |
| **1.3: Delete Delta Logic** | **DELETE** all functions that calculate "deltas" (changes) or apply small, incremental updates. This is the core of the old `onEdit` system. <br> • **Delete:** `handleAtBatEntry`. <br> • **Delete:** `handleRangePaste`. <br> • **Delete:** `applyDeltaToPitcher`. <br> • **Delete:** `applyDeltaToBatter`. <br> • **Delete:** `calculateHittingDelta`. | `BoxScoreTriggers.js`, `BoxScorePitching.js`, `BoxScoreHitting.js` |
| **1.4: Delete `PropertiesService` Logic** | **DELETE** all functions that read from or write to `PropertiesService` for stat tracking. The "Absolute State" processor will calculate stats live from the grid. <br> • **Delete:** `getGameState`, `saveGameState`. <br> • **Delete:** `getPitcherStats`, `savePitcherStats`. <br> • **Delete:** `getBatterStats`, `saveBatterStats`. <br> • **Delete:** `getActivePitchers`, `saveActivePitchers`. <br> • **Delete:** `clearAllProperties`, `clearDefensiveAssociations`. | `BoxScoreUtility.js`, `BoxScoreDefensive.js` |
| **1.5: Delete Defensive Prompt Logic**| **DELETE** all functions related to the old, fragile defensive stat system. <br> • **Delete:** `applyDefensiveStats`. <br> • **Delete:** `promptForDefensivePlayer`. <br> • **Delete:** All storage helpers (`storeDefensiveAssociation`, `getDefensiveAssociation`, `clearDefensiveAssociation`). | `BoxScoreDefensive.js` |
| **1.6: Delete Unused Hitting/Pitching Helpers** | **DELETE** `addROBToBatter`, `addStolenBaseToBatter`, and `getActivePitcherForBatter` as their logic will be centralized into the new bulk processor. | `BoxScoreHitting.js`, `BoxScorePitching.js` |

### Phase 2: Build the Menu-Driven Bulk Processor

**Goal:** Create the new "Absolute State" engine that reads the grid, calculates all stats, and performs complex attribution in one pass.

| Task | Action for AI Agent | Source Files Affected |
| :--- | :--- | :--- |
| **2.1: Create New Menu Trigger** | **Action:** In `BoxScoreMenu.js`, gut the old menu and replace it with the new v3 menu. <br> • **`onOpen()`** should now call `addBoxScoreMenu_v3()`. <br> • **`addBoxScoreMenu_v3()`** will create one primary item: `ui.createMenu('📊 Box Score Tools').addItem('🚀 Process Game Stats', 'processGameStatsBulk')` (plus `Reset Game Stats` and the `View...` functions). | `BoxScoreMenu.js` |
| **2.2: Create the Bulk Processor Controller** | **Action:** Create the new controller function: `processGameStatsBulk()`. <br> • This function will be the main orchestrator. It gets the active sheet. <br> • It must **clear all old stat data** *before* processing by calling `clearPitcherStatsInSheet()` and `clearHittingStatsInSheet()`. <br> • It will then call the core processing logic (Task 2.4). | `BoxScoreTriggers.js` (or a new `BoxScoreProcessor.js` file) |
| **2.3: Rewrite `parseNotation`** | **Action:** The `parseNotation` function must be **completely rewritten** to support the new, robust notations and return a rich object. <br> • It must parse **`PC[X]`** (e.g., `PC0`, `PC1`, `PC2`, `PC3`) using Regex and return `{ isPitcherChange: true, inheritedRunners: X }`. <br> • It must parse **`E[1-9]` / `NP[1-9]`** (e.g., `1B E6`) using Regex and return `{ isError: true, fielderPosition: 6 }` or `{ isNicePlay: true, fielderPosition: 5 }`. (Note: `1`=P, `2`=C, `3`=1B, `4`=2B, `5`=3B, `6`=SS, `7`=LF, `8`=CF, `9`=RF). <br> • It must correctly parse all standard notations (`1B`, `HR`, `K`, `BB`, `RBI`, `SB`, etc.). | `BoxScoreNotation.js` |
| **2.4: Implement "Absolute State" Processor Logic** | **Action:** The `processGameStatsBulk()` function must: <br> 1. **Initialize** empty in-memory objects (e.g., `playerStats = {}`, `teamStats = {}`). <br> 2. **Build Roster Map:** Read the Pitcher/Fielder rosters *once* to create a map: `{'Player Name': {row: 7, pos: 'P'}, ...}`. <br> 3. **Read** the *entire* Away At-Bat grid and Home At-Bat grid *once*. <br> 4. **Iterate** through every cell in the grids, managing an `activeState = { activePitcherName: "Name", inheritedRunners: 0 }`. <br> 5. **Call `parseNotation()`** on each cell and **accumulate** stats into the `playerStats` object. | `BoxScoreTriggers.js` (or new Processor file) |
| **2.5: Implement Pitcher Change Logic** | **Action:** Inside the processor loop, when `parseNotation` returns `{ isPitcherChange: true }`: <br> 1. Read the `C3`/`C4` dropdown cell to get the new pitcher's name. <br> 2. Set `activeState.activePitcherName` to this new name. <br> 3. Set `activeState.inheritedRunners = X` (from the `PC[X]` notation). | `BoxScoreTriggers.js` (or new Processor file) |
| **2.6: Implement Inherited Run Logic** | **Action:** Inside the processor loop, when a run scores (e.g., `RBI > 0`): <br> 1. Check `if (activeState.inheritedRunners > 0)`. <br> 2. If **true**, assign the Run to the **previous** pitcher's in-memory stats and *decrement* `activeState.inheritedRunners`. <br> 3. If **false**, assign the Run to the `activeState.activePitcherName`. | `BoxScoreTriggers.js` (or new Processor file) |
| **2.7: Implement Fielder Assignment** | **Action:** Inside the processor loop, when `parseNotation` returns `{ isError: true, fielderPosition: 6 }`: <br> 1. Call a **new helper** `findPlayerNameByPosition(rosterMap, 6)` to get the fielder's name (e.g., "J. Jeter"). <br> 2. Increment the `E` (Error) count in the `playerStats["J. Jeter"]` object. | `BoxScoreTriggers.js`, `BoxScoreUtility.js` (new helper) |
| **2.8: Implement Automated Cleanup Logic** | **Action:** The processor loop must detect the end of a half-inning (e.g., moving from the Away grid to the Home grid, or finishing the Home grid). <br> • At this structural break, it **must** set `activeState.inheritedRunners = 0`. This clears the counter automatically, as requested. | `BoxScoreTriggers.js` (or new Processor file) |
| **2.9: Implement Final Batch Write** | **Action:** After the loop is complete, format the in-memory `playerStats` object into 2D arrays (one for Hitting, one for Pitching/Fielding) that match the roster order. <br> • **Action:** Use **batched `setValues()`** to write all stats to the Pitching (I-O) and Hitting (C-K) sections **once**. This replaces the old `updateBatterRow` and `updatePitcherRow` functions. | `BoxScoreTriggers.js` (or new Processor file) |

### Phase 3: Refactor Utilities & `onEdit` Trigger

**Goal:** Clean up all remaining utilities and correctly "neuter" the `onEdit` trigger to only handle non-scoring events.

| Task | Action for AI Agent | Source Files Affected |
| :--- | :--- | :--- |
| **3.1: "Neuter" the `onEdit` Trigger** | **Action:** The `onEdit(e)` function **must** be refactored. <br> 1. It must **`return` immediately** if the edit `col` and `row` are within the At-Bat grid ranges. This stops it from processing scoring. <br> 2. It should **only** process edits to the pitcher dropdowns (`AWAY_PITCHER_CELL`, `HOME_PITCHER_CELL`) to run `handlePositionSwap`. The `LockService` should be retained for this one process. <br> 3. **Action:** Update the first line to use the new config variable: `if (!sheetName.startsWith(BOX_SCORE_CONFIG.GAME_SHEET_PREFIX)) return;` | `BoxScoreTriggers.js` |
| **3.2: Refactor N+1 Loops in Utilities** | **Action:** Refactor `clearHittingStatsInSheet`. **DELETE** the entire `for...loop` that calls `setValue(0)` nine times per row. <br> • **REPLACE** with logic that clears the *entire range* (respecting `PROTECTED_ROWS`) using `clearContent()` or `setValues([Array(numCols).fill(0)])`. <br> • **Action:** Refactor `clearPitcherStatsInSheet` to do the same for the Pitching/Fielding stat columns. | `BoxScoreUtility.js` |
| **3.3: Refactor Menu Magic Numbers** | **Action:** In `BoxScoreMenu.js`, refactor `showPitcherStats` and `showBatterStats`. <br> • **REPLACE** all hardcoded ranges like `sheet.getRange(7, 2, 9, 1)` with config-driven ranges using `BOX_SCORE_CONFIG.AWAY_PITCHER_RANGE`, `HITTING_RANGE`, etc.. <br> • **Note:** These functions will now be for *viewing* the stats written by the bulk processor, not the (deleted) `PropertiesService` cache. They must be refactored to read directly from the sheet cells. | `BoxScoreMenu.js` |