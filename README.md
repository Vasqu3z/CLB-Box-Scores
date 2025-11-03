# CLB Box Score Automation

Automated box score tracking system for the Casual League Baseball league, built on Google Apps Script.

## Features

### Real-Time Scoring
- **Automatic stat calculation** after each at-bat entry
- **Live updates** to pitcher and hitting statistics
- **Instant feedback** while scoring games

### Smart Pitcher Tracking
- **Automatic pitcher change detection** when you swap the pitcher dropdown
- **Inherited runner calculation** - automatically determines how many runners are on base
- **Timeline preservation** - SP (Starting Pitcher) and RP1, RP2, RP3... tracking
- **Position swap automation** - old pitcher and new pitcher automatically swap positions

### Comprehensive Notation System
Enter at-bat results using simple shorthand notation that gets automatically parsed into full statistics.

---

## At-Bat Notation Guide

### Basic Outcomes

#### Hits
- `1B` - Single
- `2B` - Double
- `3B` - Triple
- `HR` - Home Run

#### Outs
- `OUT` - Generic out (flyout, groundout, etc.)
- `K` - Strikeout
- `DP` - Double play (records 2 outs)
- `TP` - Triple play (records 3 outs)

#### Walks & Sacrifices
- `BB` - Walk (base on balls)
- `SF` - Sacrifice fly
- `SH` - Sacrifice hit/bunt

#### Fielder's Choice
- `FC` - Fielder's choice (batter reaches, no out)
- `FC OUT` - Fielder's choice with an out

#### Baserunning
- `SB` - Stolen base
- `CS` - Caught stealing (adds an out)

### RBIs (Runs Batted In)

Add RBI notation to any at-bat:
- `RBI` or `1RBI` - 1 run scored
- `2RBI` - 2 runs scored
- `3RBI` - 3 runs scored
- `4RBI` - 4 runs scored (grand slam)

**Examples:**
- `HR 4RBI` - Grand slam home run
- `2B 2RBI` - 2-run double
- `SF RBI` - Sacrifice fly with 1 RBI

### Defensive Plays

#### Errors
Assign errors to specific fielders by position number:
- `E[1-9]` - Error by fielder at position #

**Position Numbers:**
1. Pitcher (P)
2. Catcher (C)
3. First Base (1B)
4. Second Base (2B)
5. Third Base (3B)
6. Shortstop (SS)
7. Left Field (LF)
8. Center Field (CF)
9. Right Field (RF)

**Examples:**
- `1B E6` - Single, error by shortstop
- `OUT E4` - Out, but second baseman made an error

#### Nice Plays
Recognize outstanding defensive plays:
- `NP[1-9]` - Nice play by fielder at position #
- Adds +1 NP to the fielder's stats
- Adds +1 "Robbed" (ROB) to the batter's stats

**Examples:**
- `OUT NP6` - Batter was robbed by shortstop
- `OUT NP8` - Center fielder made a great catch

### Pitcher Changes

Pitcher changes are tracked **automatically** when you change the pitcher dropdown. The system will:
1. Calculate how many runners are on base
2. Insert `PC#` notation (where # = inherited runners)
3. Update positions: old pitcher → `SP` or `RP#`, new pitcher gets next RP number

**Notation Format:**
- `PC0` - Pitcher change, no inherited runners
- `PC1` - Pitcher change, 1 inherited runner
- `PC2` - Pitcher change, 2 inherited runners
- `PC3` - Pitcher change, bases loaded

**Common Patterns:**
- `K PC2` - Pitcher struck out batter, then was taken out (2 runners on base)
- `OUT OUT OUT PC0` - Pitcher finished inning, changed between innings

---

## Notation Examples

### Complex At-Bats

Combine notations to describe exactly what happened:

```
HR 4RBI          → Grand slam (4 runs)
2B 2RBI          → 2-run double
1B E6            → Single + error by shortstop
OUT NP8          → Flyout, great catch by CF
SF RBI           → Sacrifice fly, 1 RBI
DP               → Double play (2 outs)
K PC2            → Strikeout, pitcher removed (2 inherited runners)
```

### Sample Inning

```
Inning 1:
Batter 1: 1B
Batter 2: OUT
Batter 3: 2B RBI
Batter 4: K PC1
Batter 5: (New pitcher) HR 2RBI
Batter 6: OUT
Batter 7: OUT
```

---

## Using the System

### Starting a Game

1. **Enter rosters** in columns A-E (names and starting positions)
2. **Set starting pitchers** in cells D3 (Away) and D4 (Home)
3. **Start scoring** - enter at-bat notation in the grid (columns C-H)

### During the Game

**Entering At-Bats:**
- Type notation in at-bat cells (grid starts at row 7)
- Stats update automatically after each entry
- Watch pitcher/hitting stats populate in real-time

**Changing Pitchers:**
1. Select new pitcher from dropdown (D3 or D4)
2. System automatically:
   - Calculates inherited runners
   - Inserts `PC#` notation
   - Swaps pitcher positions

**Viewing Stats:**
- **Menu → View Pitcher Stats** - See all pitching lines
- **Menu → View Hitting Stats** - See all batting stats
- **Live stats** appear in columns I-R as you score

### After the Game

**Manual Processing:**
If auto-processing is disabled, click **Menu → Process Game Stats** to calculate all statistics.

**Resetting:**
**Menu → Reset Game Stats** - Clears all stats and optionally clears at-bat grid for a new game.

---

## Position Notation

The position column (E) shows each player's positional history using the `→` arrow format:

**Examples:**
- `SP` - Starting Pitcher (first pitcher)
- `SP→CF` - Started as pitcher, moved to center field
- `CF→RP1` - Started in center field, came in as 1st relief pitcher
- `2B→RP2→LF` - Played 2B, then relief pitched, then moved to LF

**Pitcher Timeline:**
- `SP` = Starting Pitcher
- `RP1` = 1st Relief Pitcher
- `RP2` = 2nd Relief Pitcher
- `RP3` = 3rd Relief Pitcher
- etc.

---

## Configuration

Settings can be adjusted in `ScoreConfig.js`:

### Experimental Features

**AUTO_INSERT_PITCHER_CHANGE** (default: `true`)
- Automatically inserts `PC#` when pitcher changes
- Set to `false` to manually enter pitcher changes

**AUTO_PROCESS_ON_AT_BAT** (default: `true`)
- Real-time scoring - stats update after each at-bat
- Set to `false` for manual processing (faster data entry, no live updates)

---

## Troubleshooting

### Stats Not Updating
- **Check** `AUTO_PROCESS_ON_AT_BAT` setting in config
- **Try** manually processing: Menu → Process Game Stats

### Pitcher Change Issues
- **Ensure** pitcher dropdown is updated first (D3 or D4)
- **Check** that `AUTO_INSERT_PITCHER_CHANGE` is enabled
- **Verify** position column shows `SP` or `RP#` markers

### Error Messages
- Check Apps Script logs (Extensions → Apps Script → Executions)
- Verify notation format matches examples above
- Ensure ScoreNotation.js file is uploaded to project

---

## Technical Details

### Statistics Tracked

**Pitching:**
- IP (Innings Pitched)
- BF (Batters Faced)
- H (Hits Allowed)
- HR (Home Runs Allowed)
- R (Runs Allowed)
- BB (Walks)
- K (Strikeouts)

**Hitting:**
- AB (At Bats)
- H (Hits)
- HR (Home Runs)
- RBI (Runs Batted In)
- BB (Walks)
- K (Strikeouts)
- ROB (Robbed - nice plays against)
- DP (Double Plays)
- TB (Total Bases)

**Fielding:**
- NP (Nice Plays)
- E (Errors)
- SB (Stolen Bases Allowed - for pitchers/catchers)

---

## Credits

Built for Casual League Baseball (CLB)
Developed by: Anthony Vasquez
Version: 3.0

For issues or feature requests, contact your league administrator.
