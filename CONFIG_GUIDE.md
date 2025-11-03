# CLB Box Score Configuration Guide

Backend configuration reference for `ScoreConfig.js`

## Overview

The `ScoreConfig.js` file contains all layout and behavioral settings for the CLB Box Score system. This centralizes configuration so that sheet structure changes only require updates in one place.

**⚠️ Important:** Changes to this file must match the actual Google Sheet layout. Incorrect values will cause automation failures.

---

## Configuration Sections

### 1. Game Sheet Identification

```javascript
GAME_SHEET_PREFIX: "#"
```

**Purpose:** Identifies which sheets are game sheets (vs. stats sheets or other tabs)

**Usage:** The automation only runs on sheets whose names start with this prefix

**Examples:**
- `"#"` → Sheets like "#Game1", "#2024-05-15" will be processed
- `"Game-"` → Sheets like "Game-2024-05-15" will be processed

**When to change:** If you want to use a different naming convention for game sheets

---

### 2. Experimental Features

#### AUTO_INSERT_PITCHER_CHANGE

```javascript
AUTO_INSERT_PITCHER_CHANGE: true
```

**Purpose:** Automatically inserts `PC#` notation when pitcher dropdown changes

**How it works:**
1. User changes pitcher dropdown (D3 or D4)
2. System calculates inherited runners from current inning
3. Appends `PC#` to the last at-bat cell
4. Updates positions (SP/RP# notation)

**When to disable:**
- If users prefer manual `PC#` entry
- If auto-calculation is inaccurate for your league's rules
- During troubleshooting

---

#### AUTO_PROCESS_ON_AT_BAT

```javascript
AUTO_PROCESS_ON_AT_BAT: true
```

**Purpose:** Real-time scoring mode - processes stats after every at-bat entry

**Impact:**
- **Enabled:** Stats update immediately (like v2), slightly slower data entry
- **Disabled:** Stats only update when clicking "Process Game Stats" menu, faster data entry

**Performance considerations:**
- Runs bulk processor on every keystroke
- Uses batch operations for performance
- Should feel responsive even in late innings

**When to disable:**
- Batch entry of historical games
- Performance issues on slow connections
- Prefer manual control over when stats calculate

---

### 3. Pitcher Tracking

```javascript
AWAY_PITCHER_CELL: "D3"
HOME_PITCHER_CELL: "D4"
```

**Purpose:** Defines where pitcher dropdown cells are located

**Requirements:**
- Must be dropdown/data validation cells
- Should contain player names matching roster
- Used by position swap automation

**When to change:** If you move pitcher dropdown cells

---

#### AWAY_PITCHER_RANGE / HOME_PITCHER_RANGE

```javascript
AWAY_PITCHER_RANGE: {
  startRow: 7,
  endRow: 15,
  nameCol: 2,        // Column B
  positionCol: 5,    // Column E
  numPlayers: 9
}
```

**Purpose:** Defines roster layout for each team

**Fields:**
- `startRow` / `endRow`: Row range containing player roster
- `nameCol`: Column containing player names (for lookups)
- `positionCol`: Column showing position history (SP→CF, etc.)
- `numPlayers`: Total roster size (used for batch operations)

**When to change:**
- Roster moved to different rows
- Changed column layout
- Roster size increased/decreased

---

### 4. Pitcher Stats Layout

```javascript
PITCHER_STATS_COLUMNS: {
  IP: 9,    // Column I - Innings Pitched
  BF: 10,   // Column J - Batters Faced
  H: 11,    // Column K - Hits Allowed
  HR: 12,   // Column L - Home Runs Allowed
  R: 13,    // Column M - Runs Allowed
  BB: 14,   // Column N - Walks Allowed
  K: 15     // Column O - Strikeouts
}
```

**Purpose:** Maps stat types to column numbers

**Format:** Column numbers (1-based, A=1, B=2, etc.)

**When to change:**
- Moved stat columns to different location
- Reordered stats
- Added/removed stat columns

**⚠️ Note:** If you add/remove columns, update all related functions

---

### 5. Fielding Stats Layout

```javascript
FIELDING_STATS_COLUMNS: {
  NP: 16,   // Column P - Nice Plays
  E: 17,    // Column Q - Errors
  SB: 18    // Column R - Stolen Bases Allowed
}
```

**Purpose:** Maps defensive stats to columns

**Usage:** Written alongside pitching stats in same row

---

### 6. Hitting Stats Layout

```javascript
HITTING_RANGE: {
  awayStartRow: 7,
  homeStartRow: 18,
  nameCol: 2,        // Column B
  numPlayers: 9,
  numStatCols: 9     // AB, H, HR, RBI, BB, K, ROB, DP, TB
}

HITTING_STATS_COLUMNS: {
  AB: 3,    // Column C - At Bats
  H: 4,     // Column D - Hits
  HR: 5,    // Column E - Home Runs
  RBI: 6,   // Column F - Runs Batted In
  BB: 7,    // Column G - Walks
  K: 8,     // Column H - Strikeouts
  ROB: 9,   // Column I - Robbed (nice plays against)
  DP: 10,   // Column J - Double Plays
  TB: 11    // Column K - Total Bases
}
```

**Purpose:** Defines hitting stats location (separate from roster grid)

**Layout:**
- Away team: rows 7-15
- Home team: rows 18-26
- Each row = one player's hitting stats

---

### 7. At-Bat Grid Layout

```javascript
AWAY_ATBAT_RANGE: {
  startRow: 7,
  endRow: 15,
  startCol: 3,       // Column C
  endCol: 8          // Column H
}

HOME_ATBAT_RANGE: {
  startRow: 18,
  endRow: 26,
  startCol: 3,       // Column C
  endCol: 8          // Column H
}
```

**Purpose:** Defines where at-bat notation is entered

**Layout:**
- Rows = batting order (1-9)
- Columns = innings (typically 6 innings)
- Each cell contains at-bat notation (e.g., "1B", "K PC2")

**When to change:**
- Expanded/reduced number of innings
- Moved at-bat grid location
- Changed roster size

---

### 8. Protected Rows

```javascript
PROTECTED_ROWS: [6, 16, 17]
```

**Purpose:** Rows that should never be cleared during reset operations

**Typical values:**
- Header rows
- Team separator rows
- Totals/summary rows

**When to change:**
- Added header rows
- Changed sheet layout structure

---

## Common Configuration Scenarios

### Scenario 1: Adding an Extra Inning

**Problem:** Game went 7 innings, grid only has 6 columns

**Solution:**
1. Insert column in sheet (after column H)
2. Update config:
```javascript
AWAY_ATBAT_RANGE: {
  endCol: 9  // Was 8, now 9 (Column I)
}
HOME_ATBAT_RANGE: {
  endCol: 9  // Was 8, now 9 (Column I)
}
```

---

### Scenario 2: Moving Stat Columns

**Problem:** Want to rearrange stat columns

**Solution:**
1. Move columns in sheet
2. Update ALL related column numbers in config:
```javascript
PITCHER_STATS_COLUMNS: {
  IP: 10,   // Was 9, moved right
  BF: 11,   // Was 10, moved right
  // ... update all
}
```

**⚠️ Important:** Update EVERY column reference, not just the ones that moved

---

### Scenario 3: Larger Roster (10 players)

**Problem:** League uses 10-player rosters

**Solution:**
1. Add row in sheet for 10th player
2. Update config:
```javascript
AWAY_PITCHER_RANGE: {
  endRow: 16,       // Was 15
  numPlayers: 10    // Was 9
}
// Also update hitting range, at-bat range, etc.
```

---

### Scenario 4: Disabling Real-Time Scoring

**Problem:** Data entry feels sluggish

**Solution:**
```javascript
AUTO_PROCESS_ON_AT_BAT: false
```

**Workflow changes:**
- Stats won't update automatically
- Must click "Process Game Stats" menu to see updated stats
- Faster data entry, less responsive feedback

---

## Validation Checklist

After changing configuration, verify:

- [ ] **Row numbers** match actual sheet rows (1-based)
- [ ] **Column numbers** match actual sheet columns (A=1, B=2, etc.)
- [ ] **Ranges** don't overlap incorrectly
- [ ] **numPlayers** matches actual roster size
- [ ] **startRow/endRow** pairs are consistent
- [ ] **Protected rows** still make sense

**Test by:**
1. Enter a test at-bat
2. Verify stats appear in correct columns
3. Change pitcher dropdown
4. Verify position swap works
5. Click "View Pitcher Stats" to confirm data is read correctly

---

## Troubleshooting

### Stats appearing in wrong columns
→ Check `PITCHER_STATS_COLUMNS` and `HITTING_STATS_COLUMNS` mappings

### Pitcher change not inserting PC notation
→ Verify `AWAY_PITCHER_CELL` and `HOME_PITCHER_CELL` values
→ Check `AUTO_INSERT_PITCHER_CHANGE` is `true`

### Stats not updating in real-time
→ Check `AUTO_PROCESS_ON_AT_BAT` is `true`
→ Verify script has proper permissions

### Reset clearing too many/too few rows
→ Update `PROTECTED_ROWS` array
→ Check `AWAY_PITCHER_RANGE` and `HOME_PITCHER_RANGE` boundaries

### Script not running on game sheets
→ Verify sheet name starts with `GAME_SHEET_PREFIX`

---

## Advanced: Adding New Stats

To add a new stat column:

1. **Add column** in sheet
2. **Update column config:**
```javascript
PITCHER_STATS_COLUMNS: {
  // ... existing stats
  NEWSTAT: 16  // Column P
}
```

3. **Update parser** (`ScoreNotation.js`):
```javascript
stats.NEWSTAT = 0;  // Add to stats object
// Add parsing logic
```

4. **Update processor** (`ScoreTriggers.js`):
```javascript
playerStats[name].pitching.NEWSTAT += stats.NEWSTAT;
```

5. **Update writer** (`ScoreTriggers.js`):
```javascript
var pitchingArray = [[ip, p.BF, p.H, p.HR, p.R, p.BB, p.K, p.NEWSTAT]];
sheet.getRange(row, pitcherCols.IP, 1, 8).setValues(pitchingArray);
```

**⚠️ Complexity:** Requires code changes in multiple files. Only attempt if comfortable with JavaScript.

---

## Variable Naming Convention

All config variables use `SCREAMING_SNAKE_CASE` to distinguish them from regular code variables.

**Examples:**
- `AWAY_PITCHER_CELL` ✅
- `awayPitcherCell` ❌ (regular variable)

**Usage in code:**
```javascript
var cell = sheet.getRange(BOX_SCORE_CONFIG.AWAY_PITCHER_CELL);
```

Always reference via `BOX_SCORE_CONFIG.VARIABLE_NAME` for clarity and consistency.

---

## Version History

**v3.0:**
- Centralized all configuration
- Removed hardcoded values
- Added experimental feature flags
- Introduced SP/RP# pitcher timeline system

**v2.x:**
- Scattered configuration across files
- Many hardcoded values
- No feature toggles

---

## Support

For configuration assistance:
1. Check this guide first
2. Verify sheet layout matches config values
3. Test with a simple game scenario
4. Review Apps Script execution logs

For code-level questions, consult:
- `ScoreTriggers.js` - Main automation logic
- `ScoreNotation.js` - Parsing rules
- `ScoreUtility.js` - Helper functions
