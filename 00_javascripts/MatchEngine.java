// =========================================================================
// MatchEngine.gs - Tournament Match & Bracket Logic
// =========================================================================

function create_match(tournament_id, stage, format, player1_id, player2_id) {
  const match_id = "M-" + Utilities.getUuid().slice(0, 8).toUpperCase();
  
  const new_match = {
    match_id: match_id,
    tournament_id: tournament_id,
    stage: stage,
    format: format,
    player1_id: player1_id,
    player2_id: player2_id,
    p1_score: 0,
    p2_score: 0,
    match_status: "Pending", 
    winner_id: "",
    p1_civ: "",       // NEW: Blank default
    p2_civ: "",       // NEW: Blank default
    map_name: ""      // NEW: Blank default
  };
  
  append_record('matches', new_match);
  Logger.log(`Created match ${match_id}: ${player1_id} vs ${player2_id}`);
  return match_id;
}

/**
 * Handles numeric score submissions (e.g., from Telegram Bot)
 */
function report_score(match_id, p1_wins, p2_wins) {
  const matches = get_all_records('matches');
  const match = matches.find(m => m.match_id === match_id);
  
  if (!match) throw new Error(`Match ID ${match_id} not found.`);
  if (match.match_status === "Settled") throw new Error("Match is already settled.");

  const format_upper = String(match.format).toUpperCase();
  let win_threshold = 1; 
  if (format_upper === "BO3") win_threshold = 2;
  if (format_upper === "BO5") win_threshold = 3;

  let match_status = "Active";
  let winner_id = "";

  if (p1_wins >= win_threshold) {
    match_status = "Settled";
    winner_id = match.player1_id;
  } else if (p2_wins >= win_threshold) {
    match_status = "Settled";
    winner_id = match.player2_id;
  }

  const update_obj = { p1_score: p1_wins, p2_score: p2_wins, match_status: match_status, winner_id: winner_id };
  update_record_by_key('matches', 'match_id', match_id, update_obj);
  Logger.log(`Match ${match_id} updated. Score: ${p1_wins}-${p2_wins}. Status: ${match_status}`);
}

/**
 * Handles dropdown name selections (from the Google Sheet UI)
 */
function settle_match_by_winner(match_id, winner_name) {
  const matches = get_all_records('matches');
  const match = matches.find(m => m.match_id === match_id);
  
  if (!match) throw new Error(`Match ID ${match_id} not found in database.`);

  // FIX 1: If the admin clears the cell, reset the match to Pending
  if (winner_name === "") {
    update_record_by_key('matches', 'match_id', match_id, {
      p1_score: 0, 
      p2_score: 0, 
      match_status: "Pending", 
      winner_id: "", 
      settled_at: ""
    });
    return;
  }

  // Handle "Unplayed" for series sweeps
  if (winner_name.toLowerCase() === "unplayed") {
    update_record_by_key('matches', 'match_id', match_id, {
      p1_score: 0, 
      p2_score: 0, 
      match_status: "Settled", 
      winner_id: "Unplayed", 
      settled_at: new Date()
    });
    return;
  }

  let p1_score = 0;
  let p2_score = 0;

  if (String(match.player1_id).toLowerCase().trim() === winner_name.toLowerCase()) {
    p1_score = 1;
  } else if (String(match.player2_id).toLowerCase().trim() === winner_name.toLowerCase()) {
    p2_score = 1;
  } else {
    throw new Error(`Winner '${winner_name}' does not match P1 or P2.`);
  }

  // FIX 2: Ensure your database headers are exactly p1_score and p2_score
  update_record_by_key('matches', 'match_id', match_id, {
    p1_score: p1_score, 
    p2_score: p2_score, 
    match_status: "Settled", 
    winner_id: winner_name, 
    settled_at: new Date()
  });
}