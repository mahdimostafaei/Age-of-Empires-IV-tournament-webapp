// =========================================================================
// CivStats.gs - Civilization Statistics & Two-Part Odds Engine
// =========================================================================

const GAMES_API_BASE = "https://aoe4world.com/api/v0/players";

/**
 * Fetches up to 200 matches and constructs a perfectly balanced 100-game pool
 * (50 Solo / 50 Team), backfilling if the player lacks enough of one type.
 */
function fetch_balanced_matches(aoe4_world_id) {
  const clean_id = String(aoe4_world_id).trim();
  const url = `${GAMES_API_BASE}/${clean_id}/games?limit=200`;
  
  try {
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (response.getResponseCode() !== 200) return [];
    
    const json = JSON.parse(response.getContentText());
    const games = Array.isArray(json) ? json : (json.games || []);
    
    if (games.length === 0) return [];

    let allSolo = [];
    let allTeam = [];

    // Separate Ranked/QM games into Solo and Team buckets
    games.forEach(match => {
      if (!match.teams) return;
      const kind = String(match.kind || match.leaderboard || "").toLowerCase();
      
      // Skip Custom Games
      if (kind.indexOf("rm_") === -1 && kind.indexOf("qm_") === -1) return;

      const isSolo = (kind.indexOf("1v1") !== -1 || kind.indexOf("solo") !== -1);
      if (isSolo) {
        allSolo.push(match);
      } else {
        allTeam.push(match);
      }
    });

    // Intelligently construct the 100-game pool with dynamic backfilling
    let selectedSolo = allSolo.slice(0, 50);
    let selectedTeam = allTeam.slice(0, 50);
    let totalSelected = selectedSolo.length + selectedTeam.length;

    if (totalSelected < 100) {
      if (selectedSolo.length < 50 && allTeam.length > 50) {
        let needed = 100 - totalSelected;
        selectedTeam = selectedTeam.concat(allTeam.slice(50, 50 + needed));
      } else if (selectedTeam.length < 50 && allSolo.length > 50) {
        let needed = 100 - totalSelected;
        selectedSolo = selectedSolo.concat(allSolo.slice(50, 50 + needed));
      }
    }

    return selectedSolo.concat(selectedTeam);
  } catch (e) {
    Logger.log(`API Error: ${e.message}`);
    return [];
  }
}

/**
 * Syncs civilization metrics using the Two-Part Modifier:
 * Rewards both Bayesian Win Rate (Performance) and Games Played (Experience).
 */
function sync_civ_stats() {
  const players = get_all_records('players');
  let update_count = 0;
  
  const sheet_name = 'player_civs';
  const sheet = get_sheet(sheet_name);
  if (!sheet) throw new Error(`Please create a tab named '${sheet_name}'.`);
  
  if (sheet.getLastRow() > 1) {
    sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).clearContent();
  }
  
  const all_civ_records = [];

  players.forEach(player => {
    if (!player.aoe4_world_id) return;
    
    const games = fetch_balanced_matches(player.aoe4_world_id);
    const civ_counts = {};
    
    games.forEach(game => {
      const targetPlayer = findTargetPlayer(game, String(player.aoe4_world_id));

      if (targetPlayer) {
        const civRaw = targetPlayer.civilization || targetPlayer.civ;
        if (civRaw) {
          const civKey = formatCivName(civRaw);
          const isWin = (targetPlayer.result === "win" || targetPlayer.result === "victory");

          if (!civ_counts[civKey]) civ_counts[civKey] = { wins: 0, total: 0 };
          
          civ_counts[civKey].total += 1;
          if (isWin) civ_counts[civKey].wins += 1;
        }
      }
    });
    
    Object.keys(civ_counts).forEach(civ => {
      const stats = civ_counts[civ];
      
      // Removed the 4-game minimum. Even 1 game gets logged now.
      if (stats.total < 1) return;
      
      // -------------------------------------------------------------
      // 1. PERFORMANCE MODIFIER (Minor Impact: max +/- 50 Elo)
      // -------------------------------------------------------------
      const phantom_games = 5;
      const phantom_wins = 2.5;
      const bayesian_wr = (stats.wins + phantom_wins) / (stats.total + phantom_games);
      
      // Reduced sensitivity from 300 down to 100
      const perf_sensitivity = 100; 
      const perf_modifier = (bayesian_wr - 0.50) * perf_sensitivity;
      
      // -------------------------------------------------------------
      // 2. EXPERIENCE BONUS (Major Impact: up to +75 Elo for 20 games)
      // -------------------------------------------------------------
      const exp_threshold = 20; 
      // Increased max bonus from 50 to 75
      const max_exp_bonus = 75; 
      const exp_bonus = (Math.min(stats.total, exp_threshold) / exp_threshold) * max_exp_bonus;
      
      // -------------------------------------------------------------
      // 3. FINAL CALCULATION
      // -------------------------------------------------------------
      let elo_modifier = perf_modifier + exp_bonus;
      
      // Apply the +/- 100 cap to keep overall tournament odds balanced
      if (elo_modifier > 100) elo_modifier = 100;
      if (elo_modifier < -100) elo_modifier = -100;
      
      all_civ_records.push({
        aoe4_world_id: player.aoe4_world_id,
        player_name: player.player_name,
        civilization: civ,
        games_played: stats.total,
        wins: stats.wins,
        raw_win_rate: (stats.wins / stats.total).toFixed(3),
        bayesian_win_rate: bayesian_wr.toFixed(3),
        elo_modifier: Math.round(elo_modifier)
      });
    });
    
    update_count++;
    Utilities.sleep(500); 
  });
  
  all_civ_records.forEach(record => {
    append_record(sheet_name, record);
  });
  
  Logger.log(`Successfully mapped ${all_civ_records.length} civs for ${update_count} players.`);
}

// ==========================================
// HELPER FUNCTIONS (From Original Engine)
// ==========================================

/**
 * Recursively hunts for the target player ID inside deeply nested API arrays.
 */
function findTargetPlayer(obj, targetIdStr) {
  if (!obj || typeof obj !== 'object') return null;
  
  if (obj.civilization || obj.civ) {
    if (JSON.stringify(obj).indexOf(targetIdStr) !== -1) {
      return obj;
    }
  }
  
  for (let key in obj) {
    if (typeof obj[key] === 'object') {
      let result = findTargetPlayer(obj[key], targetIdStr);
      if (result) return result;
    }
  }
  return null;
}

/**
 * Formats API slugs into clean Civilization names.
 */
function formatCivName(slug) {
  const civMap = {
    "abbasid_dynasty": "Abbasid Dynasty",
    "ayyubids": "Ayyubids",
    "byzantines": "Byzantines",
    "chinese": "Chinese",
    "delhi_sultanate": "Delhi Sultanate",
    "english": "English",
    "french": "French",
    "holy_roman_empire": "Holy Roman Empire",
    "japanese": "Japanese",
    "jeanne_darc": "Jeanne d'Arc",
    "malians": "Malians",
    "mongols": "Mongols",
    "order_of_the_dragon": "Order of the Dragon",
    "ottomans": "Ottomans",
    "rus": "Rus",
    "zhu_xis_legacy": "Zhu Xi's Legacy",
    "knights_templar": "Knights Templar",
    "golden_horde": "Golden Horde"
  };

  const key = String(slug || "").toLowerCase().trim();
  return civMap[key] || (key.charAt(0).toUpperCase() + key.slice(1));
}