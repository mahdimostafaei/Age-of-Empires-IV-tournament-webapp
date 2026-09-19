// =========================================================================
// EventRouter.gs - Master Omni-Channel Input Receiver & The Global Dam
// =========================================================================

function router_onEdit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  
  // 🚨 KILL SWITCH: Ignore any sheet that has been renamed by the archive script
  if (sheetName.includes(" - ") && sheetName.includes("(")) return;

  const col = e.range.getColumn();
  const row = e.range.getRow();
  const newValue = e.value === undefined ? "" : e.value;
  
  if (sheetName.toLowerCase().includes("league") || sheetName.toLowerCase().includes("group") || sheetName.toLowerCase().includes("bracket") || sheetName.toLowerCase().includes("playoff")) {
    handle_league_edit(sheet, row, col, newValue);
  }
}

function handle_league_edit(sheet, row, col, editedValue) {
  if (row < 2) return; 
  
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  const winnerColIdx = headers.indexOf("winner") + 1; 
  const matchIdColIdx = headers.findIndex(h => h.includes("match_id")) + 1;
  const p1ColIdx = headers.indexOf("player 1") + 1;
  const p2ColIdx = headers.indexOf("player 2") + 1;
  const stageColIdx = headers.indexOf("stage") + 1; 
  
  if (col !== winnerColIdx || matchIdColIdx === 0 || p1ColIdx === 0 || p2ColIdx === 0) return;
  
  const match_id = sheet.getRange(row, matchIdColIdx).getValue();
  const p1_name = sheet.getRange(row, p1ColIdx).getValue();
  const p2_name = sheet.getRange(row, p2ColIdx).getValue();
  const stage_name = stageColIdx > 0 ? sheet.getRange(row, stageColIdx).getValue() : "";
  const winner_name = String(editedValue).trim();
  
  process_match_resolution(match_id, winner_name);
  
  if (winner_name === "") {
    sheet.getRange(row, 1, 1, sheet.getLastColumn()).setBackground(null).setFontColor(null);
  }
  
  if (winner_name !== "" && winner_name !== "Unplayed") {
    enforce_series_sweep_and_advance(sheet, p1_name, p2_name, matchIdColIdx, winnerColIdx, p1ColIdx, p2ColIdx, row, stage_name, winner_name);
  }

  // GLOBAL SYNC: Only check the dam if we are editing a League
  if (sheet.getName().toLowerCase().includes("league")) {
    check_global_dam_and_sync(sheet.getParent());
  }

  // UPDATE ELO: Recalculate Live Elo leaderboard
  recalculate_global_elo();
}

function process_match_resolution(match_id, winner_name) {
  if (!match_id) return;
  try { settle_match_by_winner(match_id, winner_name); } 
  catch (err) { Logger.log(`❌ Error: ${err.message}`); }
}

function enforce_series_sweep_and_advance(sheet, p1, p2, matchIdCol, winnerCol, p1Col, p2Col, editedRow, stageName, justSelectedWinner) {
  const data = sheet.getDataRange().getValues();
  const seriesRows = [];
  let p1_wins = 0, p2_wins = 0;

  for (let i = 1; i < data.length; i++) {
    let r_p1 = data[i][p1Col - 1]; let r_p2 = data[i][p2Col - 1]; let r_winner = data[i][winnerCol - 1];
    if (r_p1 === p1 && r_p2 === p2) {
      seriesRows.push(i + 1); 
      if (r_winner === p1) p1_wins++; else if (r_winner === p2) p2_wins++;
    }
  }

  const winThreshold = Math.ceil(seriesRows.length / 2);

  if (p1_wins > winThreshold || p2_wins > winThreshold) {
    SpreadsheetApp.getActiveSpreadsheet().toast("❌ Rejected: Series is already won.", "Invalid Action", 5);
    sheet.getRange(editedRow, winnerCol).setValue("Unplayed");
    process_match_resolution(sheet.getRange(editedRow, matchIdCol).getValue(), "Unplayed");
    return; 
  }

  if (p1_wins === winThreshold || p2_wins === winThreshold) {
    const overallSeriesWinner = p1_wins === winThreshold ? p1 : p2;
    if (stageName) auto_advance_bracket(sheet, stageName, overallSeriesWinner);
    
    if (seriesRows.length >= 2) {
      seriesRows.forEach(r => {
        let currentWinner = sheet.getRange(r, winnerCol).getValue();
        if (currentWinner === "") {
          let unplayed_match_id = sheet.getRange(r, matchIdCol).getValue();
          sheet.getRange(r, winnerCol).setValue("Unplayed");
          sheet.getRange(r, 1, 1, sheet.getLastColumn()).setBackground("#e0e0e0").setFontColor("#888888");
          process_match_resolution(unplayed_match_id, "Unplayed");
        }
      });
    }
  }
}

function auto_advance_bracket(sheet, currentStage, winnerName) {
  if (!currentStage || !winnerName || winnerName === "Unplayed") return;
  
  let searchTarget = "";
  if (currentStage.includes("Quarterfinal")) searchTarget = `Winner QF${currentStage.split(" ")[1]}`;
  else if (currentStage.includes("Semifinal")) searchTarget = `Winner SF${currentStage.split(" ")[1]}`;
  else if (currentStage.includes("RO16")) searchTarget = `Winner RO16 ${currentStage.split(" ")[1]}`;
  else return; 

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    let p1 = String(data[i][2]); let p2 = String(data[i][3]); let rowNum = i + 1;

    if (p1.includes(searchTarget)) {
      sheet.getRange(rowNum, 3).setValue(winnerName).setBackground("#d9ead3"); 
      update_playoff_validation(sheet, rowNum, winnerName, sheet.getRange(rowNum, 4).getValue());
    }
    if (p2.includes(searchTarget)) {
      sheet.getRange(rowNum, 4).setValue(winnerName).setBackground("#d9ead3"); 
      update_playoff_validation(sheet, rowNum, sheet.getRange(rowNum, 3).getValue(), winnerName);
    }
  }
}

/**
 * =========================================================================
 * THE GLOBAL DAM: Waits for all groups to finish, then sorts all seeds
 * =========================================================================
 */
function check_global_dam_and_sync(ss) {
  const sheets = ss.getSheets();
  const leagueSheets = sheets.filter(s => s.getName().toLowerCase().includes("league"));
  if (leagueSheets.length === 0) return;

  // 1. IS THE DAM READY TO BREAK?
  // Scan every single match across all groups to see if any are empty
  for (let s of leagueSheets) {
    const data = s.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    const winnerColIdx = headers.indexOf("winner");
    const p1ColIdx = headers.indexOf("player 1");
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][p1ColIdx] === "BYE" || data[i][p1ColIdx+1] === "BYE") continue; // Ignore BYEs
      if (data[i][winnerColIdx] === "") return; // Found an empty match. Dam stays closed.
    }
  }

  // 2. DAM BROKE! Fetch all starting Elo ratings for tiebreakers
  const qualifiers = [];
  const dbSheet = ss.getSheetByName('players');
  const dbData = dbSheet ? dbSheet.getDataRange().getValues() : [];
  const eloMap = {};
  
  if (dbData.length > 1) {
    const h = dbData[0].map(x => String(x).toLowerCase().trim());
    const nameIdx = h.indexOf('player_name');
    const eloIdx = h.indexOf('ranked_elo');
    for (let i = 1; i < dbData.length; i++) {
        eloMap[dbData[i][nameIdx]] = Number(dbData[i][eloIdx]) || 0;
    }
  }

  // 3. Extract the Top 2 players from every League tab
  for (let s of leagueSheets) {
    const data = s.getRange("L2:O3").getValues(); // Grabs Rank, Player, Score, SB Score
    for (let i = 0; i < 2; i++) {
       let pName = data[i][1];
       let pScore = Number(data[i][2]) || 0;
       let pSB = Number(data[i][3]) || 0;
       let pElo = eloMap[pName] || 0;
       
       if (pName && pName !== "") {
         qualifiers.push({ name: pName, score: pScore, sb: pSB, elo: pElo, random: Math.random() });
       }
    }
  }

  // 4. IMPENETRABLE TIEBREAKER MATH
  qualifiers.sort((a, b) => {
    if (b.score !== a.score) return b.score - a.score;       // Check 1: Most Points
    if (b.sb !== a.sb) return b.sb - a.sb;                   // Check 2: Sonneborn-Berger Score
    if (b.elo !== a.elo) return b.elo - a.elo;               // Check 3: Historic Elo Rating
    return b.random - a.random;                              // Check 4: Microscopic RNG Coin Toss
  });

  // 5. Inject Top Players into the Playoff Bracket
  const playoffSheet = ss.getSheetByName("Playoffs");
  if (!playoffSheet) return;

  const playoffData = playoffSheet.getDataRange().getValues();
  let updatesMade = false;

  for (let q = 0; q < qualifiers.length; q++) {
    let searchString = "Overall Seed " + (q + 1);
    let actualName = qualifiers[q].name;

    for (let i = 1; i < playoffData.length; i++) {
      let rowNum = i + 1;
      let p1 = String(playoffData[i][2]);
      let p2 = String(playoffData[i][3]);

      if (p1 === searchString) {
        playoffSheet.getRange(rowNum, 3).setValue(actualName).setBackground("#d9ead3");
        update_playoff_validation(playoffSheet, rowNum, actualName, playoffSheet.getRange(rowNum, 4).getValue());
        updatesMade = true;
      }
      if (p2 === searchString) {
        playoffSheet.getRange(rowNum, 4).setValue(actualName).setBackground("#d9ead3");
        update_playoff_validation(playoffSheet, rowNum, playoffSheet.getRange(rowNum, 3).getValue(), actualName);
        updatesMade = true;
      }
    }
  }

  if (updatesMade) ss.toast("🚨 Group Stage Complete! Global seeds and Byes have been mathematically locked.", "Playoffs Generated", 8);
}

function update_playoff_validation(sheet, rowNum, p1, p2) {
  if (!p1.includes("Overall Seed") && !p2.includes("Overall Seed")) {
    let rule = SpreadsheetApp.newDataValidation().requireValueInList([p1, p2, "Unplayed"], true).setAllowInvalid(false).build();
    sheet.getRange(rowNum, 8).setDataValidation(rule);
  }
}