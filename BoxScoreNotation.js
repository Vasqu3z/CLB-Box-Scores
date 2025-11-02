// ===== BOX SCORE NOTATION PARSER MODULE =====
// Parses at-bat notation into statistical components
// Used by both pitching and hitting modules

/**
 * Parse at-bat notation into stats
 * @param {string} value - At-bat notation (e.g., "2RBI HR", "K", "OUT NP")
 * @return {Object} Stats object with all parsed components
 */
function parseNotation(value) {
  var stats = {
    // Pitching stats
    BF: 0,      // Batters faced
    outs: 0,    // Outs recorded
    H: 0,       // Hits allowed
    HR: 0,      // Home runs allowed
    R: 0,       // Runs allowed
    BB: 0,      // Walks allowed
    K: 0,       // Strikeouts
    
    // Hitting stats
    AB: 0,      // At bats
    TB: 0,      // Total bases
    
    // Defensive/baserunning
    NP: false,  // Nice play occurred
    E: false,   // Error occurred
    SB: false,  // Stolen base
    CS: false,  // Caught stealing
    DP: false   // Double play
  };
  
  // Empty cell = no stats
  if (!value || value === "") return stats;
  
  value = String(value).toUpperCase().trim();
  
  // ===== BATTER FACED =====
  // Every non-empty at-bat = 1 batter faced (for pitcher)
  stats.BF = 1;
  
  // ===== HITS =====
  var isHit = false;
  
  if (value.indexOf("1B") !== -1) {
    stats.H = 1;
    stats.TB = 1;
    isHit = true;
  }
  if (value.indexOf("2B") !== -1) {
    stats.H = 1;
    stats.TB = 2;
    isHit = true;
  }
  if (value.indexOf("3B") !== -1) {
    stats.H = 1;
    stats.TB = 3;
    isHit = true;
  }
  if (value.indexOf("HR") !== -1) {
    stats.H = 1;
    stats.HR = 1;
    stats.TB = 4;
    isHit = true;
  }
  
  // ===== WALKS =====
  var isWalk = false;
  if (value.indexOf("BB") !== -1) {
    stats.BB = 1;
    isWalk = true;
  }
  
  // ===== STRIKEOUTS =====
  if (value.indexOf("K") !== -1) {
    stats.K = 1;
  }
  
  // ===== FIELDER'S CHOICE =====
  var isFC = false;
  if (value.indexOf("FC") !== -1) {
    isFC = true;
    // Check if it's "FC OUT" (with out) or just "FC" (no out)
    if (value.indexOf("FC OUT") !== -1 || value.indexOf("FCOUT") !== -1) {
      stats.outs = 1;
    }
    // Plain FC has no out
  }
  
  // ===== SACRIFICE FLY / SACRIFICE HIT =====
  var isSacrifice = false;
  if (value.indexOf("SF") !== -1) {
    stats.outs = 1;
    isSacrifice = true;
  }
  if (value.indexOf("SH") !== -1) {
    stats.outs = 1;
    isSacrifice = true;
  }
  
  // ===== OUTS (priority order: TP > DP > single out) =====
  if (value.indexOf("TP") !== -1) {
    stats.outs = 3;
    stats.DP = true;  // Triple play counts as DP for hitting stats
  } else if (value.indexOf("DP") !== -1) {
    stats.outs = 2;
    stats.DP = true;
  } else if (value.indexOf("OUT") !== -1 || stats.K === 1) {
    // Generic out or strikeout = 1 out (unless already set by TP/DP/FC OUT/SF/SH)
    if (stats.outs === 0) {
      stats.outs = 1;
    }
  }
  
  // ===== STOLEN BASES / CAUGHT STEALING =====
  if (value.indexOf("SB") !== -1) {
    stats.SB = true;
  }
  if (value.indexOf("CS") !== -1) {
    stats.CS = true;
    stats.outs += 1;  // CS adds an out
  }
  
  // ===== RUNS BATTED IN =====
  if (value.indexOf("4RBI") !== -1) {
    stats.R = 4;
  } else if (value.indexOf("3RBI") !== -1) {
    stats.R = 3;
  } else if (value.indexOf("2RBI") !== -1) {
    stats.R = 2;
  } else if (value.indexOf("RBI") !== -1) {
    stats.R = 1;
  }
  
  // ===== NICE PLAY / ERROR =====
  if (value.indexOf("NP") !== -1) {
    stats.NP = true;
  }
  
  // Check for E with very careful detection
  // Only match: "E" alone, " E " with spaces, " E" at end, or "E " at start
  if (value === "E") {
    stats.E = true;
  } else if (value.length >= 3) {  // Need at least 3 chars for " E " or "E "
    if (value.indexOf(" E ") !== -1 || value.startsWith("E ") || value.endsWith(" E")) {
      stats.E = true;
    }
  }
  
  // ===== AT BATS (for hitting) =====
  // AB counts all plate appearances EXCEPT: walks, sacrifices
  // Hits, outs, strikeouts, FC, errors all count as AB
  stats.AB = 0;
  
  if (!isWalk && !isSacrifice) {
    // This is an at-bat
    stats.AB = 1;
  }
  
  return stats;
}

/**
 * Calculate innings pitched from outs
 * @param {number} outs - Number of outs
 * @return {number} Innings pitched with fractional notation (0.33, 0.67)
 */
function calculateIP(outs) {
  if (outs < 0) outs = 0;
  
  var fullInnings = Math.floor(outs / 3);
  var remainderOuts = outs % 3;
  
  // Convert remainder to fractional IP
  // 0 outs = .00, 1 out = .33, 2 outs = .67
  var fractional = 0;
  if (remainderOuts === 1) fractional = 0.33;
  else if (remainderOuts === 2) fractional = 0.67;
  
  return fullInnings + fractional;
}