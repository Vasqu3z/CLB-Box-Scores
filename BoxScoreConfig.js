// ===== BOX SCORE CONFIGURATION =====
// Configuration for CLB Box Score Automation
// 
// This configuration defines the exact layout of CLB box score sheets.
// Any changes to sheet structure must be reflected here.
//
// ONLY used in Box Score Spreadsheet (not Stats Processing)
//
// RULE: Never hardcode column numbers or thresholds in code - always reference this config

var BOX_SCORE_CONFIG = {
  
  // ===== PITCHER TRACKING =====
  // Dropdown cells for active pitchers
  AWAY_PITCHER_CELL: "D3",
  HOME_PITCHER_CELL: "D4",
  
  // ===== AT-BAT ENTRY RANGES =====
  // Where users enter at-bat results during live games
  AWAY_ATBAT_RANGE: {
    startRow: 7,      // First batter row
    endRow: 15,       // Last batter row (9 batters)
    startCol: 3,      // Column C (inning 1)
    endCol: 8         // Column H (inning 6)
  },
  HOME_ATBAT_RANGE: {
    startRow: 18,     // First batter row
    endRow: 26,       // Last batter row (9 batters)
    startCol: 3,      // Column C (inning 1)
    endCol: 8         // Column H (inning 6)
  },
  
  // ===== PITCHER STATS LAYOUT =====
  // Columns I-O contain pitching statistics
  // Order: IP, BF, H, HR, R, BB, K
  PITCHER_STATS_COLUMNS: {
    IP: 9,    // Column I - Innings Pitched
    BF: 10,   // Column J - Batters Faced
    H: 11,    // Column K - Hits Allowed
    HR: 12,   // Column L - Home Runs Allowed
    R: 13,    // Column M - Runs Allowed
    BB: 14,   // Column N - Walks Allowed
    K: 15     // Column O - Strikeouts
  },
  
  // ===== FIELDING STATS LAYOUT =====
  // Columns P-R contain defensive statistics
  FIELDING_STATS_COLUMNS: {
    NP: 16,   // Column P - Nice Plays
    E: 17,    // Column Q - Errors
    SB: 18    // Column R - Stolen Bases (for batters)
  },
  
  // ===== HITTING STATS LAYOUT =====
  // Columns C-K contain batting statistics
  // Order: AB, H, HR, RBI, BB, K, ROB, DP, TB
  HITTING_STATS_COLUMNS: {
    AB: 3,    // Column C - At Bats
    H: 4,     // Column D - Hits
    HR: 5,    // Column E - Home Runs
    RBI: 6,   // Column F - Runs Batted In
    BB: 7,    // Column G - Walks
    K: 8,     // Column H - Strikeouts
    ROB: 9,   // Column I - Reached On Base (hits stolen via nice plays)
    DP: 10,   // Column J - Double Plays
    TB: 11    // Column K - Total Bases
  },
  
  // ===== PITCHER/FIELDER ROSTER RANGES =====
  // Where player names and positions are listed (also where pitching/fielding stats go)
  AWAY_PITCHER_RANGE: {
    startRow: 7,      // First data row (row 6 is header)
    endRow: 15,       // Last data row (row 16 is team totals)
    numPlayers: 9,    // Number of players
    nameCol: 2,       // Column B - Player names
    positionCol: 1,   // Column A - Positions (P, C, 1B, etc.)
    statsStartCol: 9  // Column I - Where stats begin
  },
  HOME_PITCHER_RANGE: {
    startRow: 18,     // First data row (row 17 is header)
    endRow: 26,       // Last data row (row 27 is team totals)
    numPlayers: 9,    // Number of players
    nameCol: 2,       // Column B - Player names
    positionCol: 1,   // Column A - Positions (P, C, 1B, etc.)
    statsStartCol: 9  // Column I - Where stats begin
  },
  
  // ===== HITTING ROSTER RANGES =====
  // Where batting statistics are displayed (separate from pitching/fielding)
  HITTING_RANGE: {
    awayStartRow: 30,   // First away batter (row 29 is header)
    awayEndRow: 38,     // Last away batter (row 39 is team totals)
    homeStartRow: 41,   // First home batter (row 40 is header)
    homeEndRow: 49,     // Last home batter (row 50 is team totals)
    numPlayers: 9,      // Number of batters per team
    nameCol: 2,         // Column B - Player names
    statsStartCol: 3,   // Column C - Where hitting stats begin
    numStatCols: 9      // Number of stat columns (C through K)
  },
  
  // ===== PROTECTED ROWS (NEVER MODIFY THESE) =====
  // These rows contain headers, formulas, or team totals
  // Scripts should NEVER write to these rows
  PROTECTED_ROWS: [
    6,    // Away team pitching/fielding header
    16,   // Away team pitching/fielding totals
    17,   // Home team pitching/fielding header
    27,   // Home team pitching/fielding totals
    29,   // Away team hitting header
    39,   // Away team hitting totals
    40,   // Home team hitting header
    50    // Home team hitting totals
  ],
  
  // ===== PERFORMANCE & BEHAVIOR SETTINGS =====
  // Thresholds and timeouts for script behavior
  
  LOCK_TIMEOUT_MS: 30000,  // 30 seconds - How long to wait for lock acquisition
  
  PASTE_THRESHOLDS: {
    DANGER: 27,    // Show strong warning (likely to timeout)
    CAUTION: 18,   // Show medium warning (may timeout)
    TOAST: 9      // Show progress toast for pastes larger than this
  },
  
  // ===== SCRIPT PROPERTIES KEYS =====
  // Keys for storing game state in Script Properties
  // Each game sheet has its own set of properties (suffixed with _{sheetId})
  PROPERTY_GAME_STATE: "boxScoreGameState",           // Shadow storage of at-bat grid
  PROPERTY_PITCHER_STATS: "boxScorePitcherStats",     // Cumulative pitcher stats
  PROPERTY_BATTER_STATS: "boxScoreBatterStats",       // Cumulative batter stats
  PROPERTY_ACTIVE_PITCHERS: "boxScoreActivePitchers", // Current active pitchers (C3/C4)
  
  // Prefixes for defensive stat associations (NP/E tracking)
  PROPERTY_PREFIX_NP: "NP",  // Nice Play associations
  PROPERTY_PREFIX_E: "E"     // Error associations
};