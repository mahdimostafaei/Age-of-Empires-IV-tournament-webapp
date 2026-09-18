// =========================================================================
// PlayerDashboardEngine.gs - Unified Match Workflow & Elo Engine
// =========================================================================

function authenticate_player_dashboard(playerName, pin) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName("players");
    if (!pSheet) throw new Error("Players sheet not found.");

    const data = pSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const nameCol = headers.indexOf("player_name");
    const pinCol = headers.indexOf("web_pin");
    const eloCol = headers.indexOf("tournament_elo") > -1 ? headers.indexOf("tournament_elo") : headers.indexOf("ranked_elo");
    const teleCol = headers.indexOf("telegram_handle");

    let authenticated = false;
    let playerData = null;
    const playerInfoMap = {};

    // 1. Build Player Map & Authenticate (Pure Player Login - No God PIN)
    for (let i = 1; i < data.length; i++) {
      let rowName = String(data[i][nameCol]).trim();
      let rowPin = String(data[i][pinCol]).trim();
      
      if (rowName) {
        playerInfoMap[rowName.toLowerCase()] = {
          pin: rowPin,
          elo: data[i][eloCol] || 1000,
          telegram: data[i][teleCol] || ""
        };
      }

      if (rowName && rowName.toLowerCase() === String(playerName).trim().toLowerCase()) {
        if (rowPin === String(pin).trim()) {
          authenticated = true;
          playerData = { name: rowName, elo: data[i][eloCol] || 1000, telegram: data[i][teleCol] || "" };
        }
      }
    }

    if (!authenticated) {
      return JSON.stringify({ success: false, error: "Invalid Name or PIN." });
    }

    // 2. Extract Database Data
    const mSheet = ss.getSheetByName("matches");
    if (!mSheet) return JSON.stringify({ success: false, error: "Matches database missing." });
    
    const mData = mSheet.getDataRange().getValues();
    if (mData.length < 2) {
        return JSON.stringify({ success: true, player: playerData, isAdmin: false, matches: [], pendingConfirmation: null, activeMaps: [], civList: [] });
    }

    const mHeaders = mData[0].map(h => String(h).toLowerCase().trim());
    const mIdCol = mHeaders.indexOf("match_id");
    const tIdCol = mHeaders.indexOf("tournament_id");
    const stageCol = mHeaders.indexOf("stage");
    const formatCol = mHeaders.indexOf("format");
    const p1Col = mHeaders.indexOf("player1_id");
    const p2Col = mHeaders.indexOf("player2_id");
    const statusCol = mHeaders.indexOf("match_status");
    const p1CivCol = mHeaders.indexOf("p1_civ");
    const p2CivCol = mHeaders.indexOf("p2_civ");
    const mapCol = mHeaders.indexOf("map_name");
    const subByCol = mHeaders.indexOf("submitted_by");
    const winnerCol = mHeaders.indexOf("winner_id");
    const propP1CivCol = mHeaders.indexOf("proposed_p1_civ");
    const propP2CivCol = mHeaders.indexOf("proposed_p2_civ");
    const propMapCol = mHeaders.indexOf("proposed_map");
    const propWinnerCol = mHeaders.indexOf("proposed_winner");
    
    // 3. Extract Form (W/L) and H2H Tracking
    let statsMap = {};
    for (let i = 1; i < mData.length; i++) {
        let p1 = p1Col > -1 ? String(mData[i][p1Col]).trim() : "";
        let p2 = p2Col > -1 ? String(mData[i][p2Col]).trim() : "";
        let status = statusCol > -1 ? String(mData[i][statusCol]).trim().toLowerCase() : "";
        let winner = winnerCol > -1 ? String(mData[i][winnerCol]).trim() : "";

        if (status === 'settled' || status === 'archived' || status === 'completed') {
            if (winner !== '' && winner.toLowerCase() !== 'unplayed' && p1.toUpperCase() !== 'BYE' && p2.toUpperCase() !== 'BYE') {
                let p1Key = p1.toLowerCase();
                let p2Key = p2.toLowerCase();
                let winKey = winner.toLowerCase();
                
                if (p1Key) {
                    if (!statsMap[p1Key]) statsMap[p1Key] = { history: [], h2h: {} };
                    statsMap[p1Key].history.push(winKey === p1Key ? 'W' : 'L');
                }
                if (p2Key) {
                    if (!statsMap[p2Key]) statsMap[p2Key] = { history: [], h2h: {} };
                    statsMap[p2Key].history.push(winKey === p2Key ? 'W' : 'L');
                }

                if (p1Key && p2Key) {
                    if (!statsMap[p1Key].h2h[p2Key]) statsMap[p1Key].h2h[p2Key] = { wins: 0, losses: 0 };
                    if (!statsMap[p2Key].h2h[p1Key]) statsMap[p2Key].h2h[p1Key] = { wins: 0, losses: 0 };
                    
                    if (winKey === p1Key) {
                        statsMap[p1Key].h2h[p2Key].wins++;
                        statsMap[p2Key].h2h[p1Key].losses++;
                    } else if (winKey === p2Key) {
                        statsMap[p2Key].h2h[p1Key].wins++;
                        statsMap[p1Key].h2h[p2Key].losses++;
                    }
                }
            }
        }
    }

    // 4. Extract Top Civs
    let civStatsMap = {};
    let pcSheet = ss.getSheetByName("players_civs") || ss.getSheetByName("player_civs"); 
    
    if (pcSheet) {
      const pcData = pcSheet.getDataRange().getValues();
      const pcHeaders = pcData[0].map(h => String(h).toLowerCase().trim());
      
      let pcNameCol = pcHeaders.indexOf("player_name");
      if (pcNameCol === -1) pcNameCol = 1; 
      let pcCivCol = pcHeaders.indexOf("civilization");
      if (pcCivCol === -1) pcCivCol = 2; 
      let pcGamesCol = pcHeaders.indexOf("games_played");
      if (pcGamesCol === -1) pcGamesCol = 3; 

      for (let i = 1; i < pcData.length; i++) {
        let name = String(pcData[i][pcNameCol]).trim().toLowerCase();
        let civ = String(pcData[i][pcCivCol]).trim();
        let games = Number(pcData[i][pcGamesCol]) || 0;
        
        if (name && civ) {
          if (!civStatsMap[name]) civStatsMap[name] = [];
          civStatsMap[name].push({ civ: civ, games: games });
        }
      }
    }
    
    for (let playerKey in civStatsMap) {
      civStatsMap[playerKey].sort((a, b) => b.games - a.games);
      civStatsMap[playerKey] = civStatsMap[playerKey].slice(0, 3).map(obj => obj.civ);
    }

    // 5. Build Active Matches Array
    let playerMatches = [];
    let pendingConfirmationMatch = null;
    let activeTourneyIds = new Set();
    let myKey = playerData.name.toLowerCase();
    
    for (let i = 1; i < mData.length; i++) {
      let p1 = p1Col > -1 ? String(mData[i][p1Col]).trim() : "";
      let p2 = p2Col > -1 ? String(mData[i][p2Col]).trim() : "";
      let status = statusCol > -1 ? String(mData[i][statusCol]).trim().toLowerCase() : "pending";
      let submittedBy = subByCol > -1 ? String(mData[i][subByCol]).trim() : "";
      let mapVal = mapCol > -1 ? String(mData[i][mapCol]).trim() : "";
      let winnerVal = winnerCol > -1 ? String(mData[i][winnerCol]).trim() : "";
      
      let isUserInMatch = (p1 !== "" && p2 !== "") && (p1.toLowerCase() === myKey || p2.toLowerCase() === myKey);
      
      if (isUserInMatch) {
        let opponentName = (p1.toLowerCase() === myKey ? p2 : p1);
        let oppKey = opponentName.toLowerCase();
        
        let oppInfo = playerInfoMap[oppKey] || { elo: 1000, telegram: "" };
        let oppStats = statsMap[oppKey] || { history: [], h2h: {} };
        let myStats = statsMap[myKey] || { history: [], h2h: {} };
        
        let recentHistory = oppStats.history.slice(-5);
        let opponentCivsArr = civStatsMap[oppKey] || [];
        let topCivsString = opponentCivsArr.length > 0 ? opponentCivsArr.join(', ') : "No Recorded Data";
        
        let h2hWins = myStats.h2h[oppKey] ? myStats.h2h[oppKey].wins : 0;
        let h2hLosses = myStats.h2h[oppKey] ? myStats.h2h[oppKey].losses : 0;

        if (status === "waiting_confirmation" && submittedBy !== "" && submittedBy.toLowerCase() !== myKey) {
          let viewerIsP1 = (p1.toLowerCase() === myKey);
          
          // Fetch from the PROPOSED columns instead of the official ones
          let proposedP1Civ = propP1CivCol > -1 ? mData[i][propP1CivCol] : "";
          let proposedP2Civ = propP2CivCol > -1 ? mData[i][propP2CivCol] : "";
          let proposedMap = propMapCol > -1 ? mData[i][propMapCol] : "";
          let proposedWinner = propWinnerCol > -1 ? mData[i][propWinnerCol] : "";

          pendingConfirmationMatch = {
            matchId: mIdCol > -1 ? mData[i][mIdCol] : "ERR_ID",
            opponent: submittedBy,
            myCiv: viewerIsP1 ? proposedP1Civ : proposedP2Civ,
            oppCiv: viewerIsP1 ? proposedP2Civ : proposedP1Civ,
            map: proposedMap,
            winner: proposedWinner
          };
        }
        
        if (status === "pending" || status === "active" || status === "in_progress" || status === "waiting_confirmation" || status === "disputed") {
           let tId = tIdCol > -1 ? String(mData[i][tIdCol]).trim() : "";
           if (tId) activeTourneyIds.add(tId);
           
           playerMatches.push({
             matchId: mIdCol > -1 ? mData[i][mIdCol] : "ERR_ID",
             tournamentId: tId,
             stage: stageCol > -1 ? mData[i][stageCol] : "",
             format: formatCol > -1 ? mData[i][formatCol] : "",
             opponent: opponentName,
             opponentElo: oppInfo.elo,
             opponentTelegram: oppInfo.telegram,
             opponentHistory: recentHistory,
             opponentTopCivs: topCivsString,
             myH2HWins: h2hWins,
             opponentH2HWins: h2hLosses,
             isPlayer1: p1.toLowerCase() === myKey,
             isDraftLocked: mapVal !== "",
             isWaitingConfirmation: status === "waiting_confirmation",
             isActive: status === "active" || status === "in_progress",
             isDisputed: status === "disputed"
           });
        }
      }
    }

    let draftLinks = {};
    if (activeTourneyIds.size > 0) {
      const tSheet = ss.getSheetByName("tournaments");
      if (tSheet) {
        const tData = tSheet.getDataRange().getValues();
        const tHeaders = tData[0].map(h => String(h).toLowerCase().trim());
        const tIdColT = tHeaders.indexOf("tournament_id");
        if (tIdColT > -1) {
            for (let i = 1; i < tData.length; i++) {
              let tId = String(tData[i][tIdColT]).trim();
              if (activeTourneyIds.has(tId)) draftLinks[tId] = { headers: tHeaders, rowData: tData[i] };
            }
        }
      }
    }

    playerMatches = playerMatches.map(m => {
       m.civDraftLink = ""; m.mapDraftLink = "";
       if (draftLinks[m.tournamentId]) {
         const tHeaders = draftLinks[m.tournamentId].headers;
         const tRow = draftLinks[m.tournamentId].rowData;
         let cleanFormat = String(m.format).toLowerCase().replace(/[^a-z0-9]/g, '');
         let civColIdx = tHeaders.indexOf("civ_draft_" + cleanFormat);
         let mapColIdx = tHeaders.indexOf("map_draft_" + cleanFormat);
         m.civDraftLink = civColIdx > -1 ? String(tRow[civColIdx]).trim() : "";
         m.mapDraftLink = mapColIdx > -1 ? String(tRow[mapColIdx]).trim() : "";
       }
       return m;
    });

    // 6. Fetch Global Maps & Civs
    let activeMaps = [];
    const mapSheet = ss.getSheetByName("config_maps");
    if (mapSheet) {
      const mapData = mapSheet.getDataRange().getValues();
      const mapHeaders = mapData[0].map(h => String(h).toLowerCase().trim());
      const mapNameCol = mapHeaders.indexOf("map_name");
      const mapStatusCol = mapHeaders.indexOf("map_status");
      if (mapNameCol > -1 && mapStatusCol > -1) {
        for (let i = 1; i < mapData.length; i++) {
          if (String(mapData[i][mapStatusCol]).trim().toLowerCase() === "active") activeMaps.push(String(mapData[i][mapNameCol]).trim());
        }
      }
    }

    let civList = [];
    const civSheet = ss.getSheetByName("config_civs");
    if (civSheet) {
      const civData = civSheet.getDataRange().getValues();
      const civHeaders = civData[0].map(h => String(h).toLowerCase().trim());
      let civNameCol = civHeaders.indexOf("civilization_name");
      if (civNameCol === -1) civNameCol = 1; 
      for (let i = 1; i < civData.length; i++) {
        let civ = String(civData[i][civNameCol]).trim();
        if (civ !== "") civList.push(civ);
      }
    }

    return JSON.stringify({
      success: true, player: playerData, isAdmin: false, matches: playerMatches,
      pendingConfirmation: pendingConfirmationMatch, activeMaps: activeMaps, civList: civList
    });
  } catch (err) {
    return JSON.stringify({ success: false, error: "Backend Auth Error: " + err.toString() });
  }
}

// ---------------------------------------------------------
// CENTRAL ELO CALCULATOR HELPER
// ---------------------------------------------------------
function processEloForMatchRow(ss, mSheet, mRow) {
  SpreadsheetApp.flush(); 

  const headers = mSheet.getRange(1, 1, 1, mSheet.getLastColumn()).getValues()[0].map(h => String(h).toLowerCase().trim());
  
  const statusCol = headers.indexOf("match_status") > -1 ? headers.indexOf("match_status") : headers.indexOf("status");
  const winnerCol = headers.indexOf("winner_id") > -1 ? headers.indexOf("winner_id") : headers.indexOf("winner");
  const p1Col = headers.indexOf("player1_id") > -1 ? headers.indexOf("player1_id") : headers.indexOf("player 1");
  const p2Col = headers.indexOf("player2_id") > -1 ? headers.indexOf("player2_id") : headers.indexOf("player 2");
  
  let dateCol = headers.indexOf("settled_at");
  let shift1Col = headers.indexOf("p1_elo_shift");
  let shift2Col = headers.indexOf("p2_elo_shift");
  
  let nextBlank = headers.length + 1;
  if (dateCol === -1) { dateCol = nextBlank - 1; mSheet.getRange(1, nextBlank++).setValue('settled_at').setFontWeight('bold'); }
  if (shift1Col === -1) { shift1Col = nextBlank - 1; mSheet.getRange(1, nextBlank++).setValue('p1_elo_shift').setFontWeight('bold'); }
  if (shift2Col === -1) { shift2Col = nextBlank - 1; mSheet.getRange(1, nextBlank++).setValue('p2_elo_shift').setFontWeight('bold'); }

  if (statusCol === -1 || winnerCol === -1 || p1Col === -1 || p2Col === -1) return "Core columns (status, winner, p1, p2) missing in Matches sheet.";

  SpreadsheetApp.flush(); 
  const rowData = mSheet.getRange(mRow, 1, 1, mSheet.getLastColumn()).getValues()[0];
  
  const status = String(rowData[statusCol]).trim().toLowerCase();
  const winner = String(rowData[winnerCol]).trim();
  const p1 = String(rowData[p1Col]).trim();
  const p2 = String(rowData[p2Col]).trim();
  
  let shift1 = Number(rowData[shift1Col]) || 0;
  let shift2 = Number(rowData[shift2Col]) || 0;

  const pSheet = ss.getSheetByName("players") || ss.getSheetByName("Players");
  if (!pSheet) return "Players sheet not found.";
  
  const pData = pSheet.getDataRange().getValues();
  const pHeaders = pData[0].map(h => String(h).toLowerCase().trim());
  
  const pNameCol = pHeaders.indexOf("player_name") > -1 ? pHeaders.indexOf("player_name") : pHeaders.indexOf("player");
  const pEloCol = pHeaders.indexOf("tournament_elo") > -1 ? pHeaders.indexOf("tournament_elo") : pHeaders.indexOf("ranked_elo"); 
  
  if (pNameCol === -1 || pEloCol === -1) return "Elo or Name column missing in Players sheet.";

  let p1Row = -1, p2Row = -1;
  let p1Elo = 1000, p2Elo = 1000;

  for (let i = 1; i < pData.length; i++) {
    let name = String(pData[i][pNameCol]).trim().toLowerCase();
    if (name === p1.toLowerCase()) { p1Row = i + 1; p1Elo = Number(pData[i][pEloCol]) || 1000; }
    if (name === p2.toLowerCase()) { p2Row = i + 1; p2Elo = Number(pData[i][pEloCol]) || 1000; }
  }

  if (p1Row === -1 || p2Row === -1) return "Player (" + p1 + " or " + p2 + ") not found in Players sheet."; 

  if (shift1 !== 0 || shift2 !== 0) {
    p1Elo -= shift1;
    p2Elo -= shift2;
    mSheet.getRange(mRow, shift1Col + 1).setValue("");
    mSheet.getRange(mRow, shift2Col + 1).setValue("");
    pSheet.getRange(p1Row, pEloCol + 1).setValue(p1Elo);
    pSheet.getRange(p2Row, pEloCol + 1).setValue(p2Elo);
  }

  if (status === "settled" && winner !== "") {
    const K = 32; 
    let expected1 = 1 / (1 + Math.pow(10, (p2Elo - p1Elo) / 400));
    let expected2 = 1 / (1 + Math.pow(10, (p1Elo - p2Elo) / 400));
    
    let s1 = (winner.toLowerCase() === p1.toLowerCase()) ? 1 : 0;
    let s2 = (winner.toLowerCase() === p2.toLowerCase()) ? 1 : 0;

    let newShift1 = Math.round(K * (s1 - expected1));
    let newShift2 = Math.round(K * (s2 - expected2));

    p1Elo += newShift1;
    p2Elo += newShift2;

    pSheet.getRange(p1Row, pEloCol + 1).setValue(p1Elo);
    pSheet.getRange(p2Row, pEloCol + 1).setValue(p2Elo);
    mSheet.getRange(mRow, shift1Col + 1).setValue(newShift1);
    mSheet.getRange(mRow, shift2Col + 1).setValue(newShift2);
    mSheet.getRange(mRow, dateCol + 1).setValue(new Date());
    return "Success";
  } else {
    mSheet.getRange(mRow, dateCol + 1).setValue("");
    return "Status is not Settled, or Winner is empty.";
  }
}

// =========================================================================
// ADMIN DATA ENDPOINTS
// =========================================================================

function get_admin_queue() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName("matches");
    if (!mSheet) return JSON.stringify([]);

    const data = mSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const matchIdCol = headers.indexOf("match_id");
    const p1Col = headers.indexOf("player1_id");
    const p2Col = headers.indexOf("player2_id");
    const statusCol = headers.indexOf("match_status");
    const propWinnerCol = headers.indexOf("proposed_winner");
    const stageCol = headers.indexOf("stage");

    const queue = [];
    for (let i = 1; i < data.length; i++) {
      const status = String(data[i][statusCol]).trim().toLowerCase();
      if (status === "waiting_confirmation" || status === "disputed") {
        queue.push({
          matchId: data[i][matchIdCol],
          stage: data[i][stageCol],
          matchup: data[i][p1Col] + " vs " + data[i][p2Col],
          status: data[i][statusCol],
          proposedWinner: data[i][propWinnerCol] || "None"
        });
      }
    }
    return JSON.stringify({ success: true, data: queue });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

function admin_force_resolve(matchId, actionType) {
  return submit_dashboard_action(actionType === "approve" ? "confirm_result" : "admin_cancel", matchId, {});
}

function get_all_admin_matches() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName("matches");
    if (!mSheet) return JSON.stringify({success:false, data:[]});

    const data = mSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const mIdCol = headers.indexOf("match_id");
    const stageCol = headers.indexOf("stage");
    const p1Col = headers.indexOf("player1_id");
    const p2Col = headers.indexOf("player2_id");
    const statusCol = headers.indexOf("match_status");
    
    const p1CivCol = headers.indexOf("p1_civ");
    const p2CivCol = headers.indexOf("p2_civ");
    const mapCol = headers.indexOf("map_name");
    const winnerCol = headers.indexOf("winner_id");
    
    const propP1CivCol = headers.indexOf("proposed_p1_civ");
    const propP2CivCol = headers.indexOf("proposed_p2_civ");
    const propMapCol = headers.indexOf("proposed_map");
    const propWinnerCol = headers.indexOf("proposed_winner");

    let allMatches = [];
    for (let i = 1; i < data.length; i++) {
      let statusStr = String(data[i][statusCol]).trim();
      
      if (statusStr.toLowerCase() !== "archived") {
        allMatches.push({
          id: data[i][mIdCol],
          stage: String(data[i][stageCol]).trim() || "Unassigned",
          p1: String(data[i][p1Col]).trim(),
          p2: String(data[i][p2Col]).trim(),
          status: statusStr,
          p1Civ: (propP1CivCol > -1 && data[i][propP1CivCol]) ? data[i][propP1CivCol] : (p1CivCol > -1 ? data[i][p1CivCol] : ""),
          p2Civ: (propP2CivCol > -1 && data[i][propP2CivCol]) ? data[i][propP2CivCol] : (p2CivCol > -1 ? data[i][p2CivCol] : ""),
          map: (propMapCol > -1 && data[i][propMapCol]) ? data[i][propMapCol] : (mapCol > -1 ? data[i][mapCol] : ""),
          winner: (propWinnerCol > -1 && data[i][propWinnerCol]) ? data[i][propWinnerCol] : (winnerCol > -1 ? data[i][winnerCol] : "")
        });
      }
    }

    let civList = [];
    const civSheet = ss.getSheetByName("config_civs");
    if (civSheet) {
      const cData = civSheet.getDataRange().getValues();
      for(let i=1; i<cData.length; i++) {
        if(cData[i][1]) civList.push(String(cData[i][1]).trim());
      }
    }

    let mapList = [];
    const mapSheet = ss.getSheetByName("config_maps");
    if (mapSheet) {
      const mData = mapSheet.getDataRange().getValues();
      const mHeaders = mData[0].map(h => String(h).toLowerCase().trim());
      const mNameCol = mHeaders.indexOf("map_name");
      const mStatCol = mHeaders.indexOf("map_status");
      if (mNameCol > -1) {
        for(let i=1; i<mData.length; i++) {
          if (mStatCol === -1 || String(mData[i][mStatCol]).toLowerCase() === "active") {
            mapList.push(String(mData[i][mNameCol]).trim());
          }
        }
      }
    }

    return JSON.stringify({ success: true, data: allMatches, civs: civList, maps: mapList });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

function admin_save_match(matchId, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName("matches");
    const data = mSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const rowIndex = data.findIndex(r => String(r[headers.indexOf("match_id")]) === matchId);
    if (rowIndex === -1) return JSON.stringify({success: false, error: "Match not found in database."});
    const row = rowIndex + 1;

    let p1CivCol = headers.indexOf("p1_civ") > -1 ? headers.indexOf("p1_civ") : headers.indexOf("player1_civ");
    let p2CivCol = headers.indexOf("p2_civ") > -1 ? headers.indexOf("p2_civ") : headers.indexOf("player2_civ");
    let mapCol = headers.indexOf("map_name") > -1 ? headers.indexOf("map_name") : headers.indexOf("map");
    let winnerCol = headers.indexOf("winner_id") > -1 ? headers.indexOf("winner_id") : headers.indexOf("winner");
    let statusCol = headers.indexOf("match_status") > -1 ? headers.indexOf("match_status") : headers.indexOf("status");

    if (p1CivCol > -1) mSheet.getRange(row, p1CivCol + 1).setValue(payload.p1Civ);
    if (p2CivCol > -1) mSheet.getRange(row, p2CivCol + 1).setValue(payload.p2Civ);
    if (mapCol > -1) mSheet.getRange(row, mapCol + 1).setValue(payload.map);
    if (winnerCol > -1) mSheet.getRange(row, winnerCol + 1).setValue(payload.winner);
    
    let newStatus = "Active";
    if (payload.winner && payload.winner.trim() !== "") {
       newStatus = "Settled";
    }
    if (statusCol > -1) mSheet.getRange(row, statusCol + 1).setValue(newStatus);

    const wipeCols = ["proposed_p1_civ", "proposed_p2_civ", "proposed_map", "proposed_winner", "submitted_by"];
    wipeCols.forEach(col => {
       let idx = headers.indexOf(col);
       if (idx > -1) mSheet.getRange(row, idx + 1).setValue("");
    });

    if (newStatus === "Settled") {
       let eloResult = processEloForMatchRow(ss, mSheet, row);
       if (eloResult !== "Success") {
          throw new Error("Data Saved, but Elo Aborted: " + eloResult);
       }
    }

    const stageColIdx = headers.indexOf("stage");
    const stageName = stageColIdx > -1 ? String(data[row - 1][stageColIdx]).trim() : "";
    if (stageName) {
       sync_to_stage_sheet(ss, stageName, matchId, { p1Civ: payload.p1Civ, p2Civ: payload.p2Civ, map: payload.map, winner: payload.winner });
    }

    return JSON.stringify({success: true});

  } catch(e) {
    return JSON.stringify({success: false, error: e.toString()});
  }
}

function submit_dashboard_action(action, matchId, data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName("matches");
    const mData = mSheet.getDataRange().getValues();
    const headers = mData[0].map(h => String(h).toLowerCase().trim());
    
    const rowIndex = mData.findIndex(r => String(r[headers.indexOf("match_id")]) === matchId);
    if (rowIndex === -1) return JSON.stringify({ success: false, error: "Match not found." });
    const row = rowIndex + 1; 

    const stageColIdx = headers.indexOf("stage");
    const stageName = stageColIdx > -1 ? String(mData[row - 1][stageColIdx]).trim() : "";
    
    if (action === "report_match_start") {
      const propP1Civ = headers.indexOf("proposed_p1_civ");
      const propP2Civ = headers.indexOf("proposed_p2_civ");
      const propMap = headers.indexOf("proposed_map");
      const statusCol = headers.indexOf("match_status");

      if (propP1Civ > -1) mSheet.getRange(row, propP1Civ + 1).setValue(data.p1Civ);
      if (propP2Civ > -1) mSheet.getRange(row, propP2Civ + 1).setValue(data.p2Civ);
      if (propMap > -1) mSheet.getRange(row, propMap + 1).setValue(data.map);
      if (statusCol > -1) mSheet.getRange(row, statusCol + 1).setValue("active");
      
      sync_to_stage_sheet(ss, stageName, matchId, { p1Civ: data.p1Civ, p2Civ: data.p2Civ, map: data.map });

      try {
        if (typeof announceMatch === "function") announceMatch(matchId);
      } catch (tgError) {
        console.error("Telegram Poll Error: " + tgError.message);
      }
      return JSON.stringify({ success: true, message: "Match start reported." });
    }
    
    if (action === "report_winner") {
      const propWinnerCol = headers.indexOf("proposed_winner");
      const subByCol = headers.indexOf("submitted_by");
      const statusCol = headers.indexOf("match_status");
      
      if (propWinnerCol === -1 || subByCol === -1 || statusCol === -1) {
        return JSON.stringify({ success: false, error: "Database error. Please add 'proposed_winner' column." });
      }

      let submitterName = data.playerName ? data.playerName : "Unknown_Player";
      mSheet.getRange(row, propWinnerCol + 1).setValue(data.winnerName);
      mSheet.getRange(row, subByCol + 1).setValue(submitterName);
      mSheet.getRange(row, statusCol + 1).setValue("waiting_confirmation");
      
      return JSON.stringify({ success: true, message: "Winner recorded! Awaiting confirmation." });
    }

    if (action === "confirm_result") {
      const propWinnerCol = headers.indexOf("proposed_winner");
      const winnerCol = headers.indexOf("winner_id");
      const statusCol = headers.indexOf("match_status");
      
      if (propWinnerCol > -1 && winnerCol > -1) {
        const officialWinner = String(mData[row - 1][propWinnerCol]).trim();
        if (officialWinner !== "") {
          mSheet.getRange(row, winnerCol + 1).setValue(officialWinner);
          sync_to_stage_sheet(ss, stageName, matchId, { winner: officialWinner });
        }
        mSheet.getRange(row, propWinnerCol + 1).setValue("");
      }

      if (statusCol > -1) mSheet.getRange(row, statusCol + 1).setValue("Settled");
      processEloForMatchRow(ss, mSheet, row); 
      return JSON.stringify({ success: true, message: "Match confirmed!" });
    }

    if (action === "reject_result") {
      if (headers.indexOf("match_status") > -1) mSheet.getRange(row, headers.indexOf("match_status") + 1).setValue("Disputed");
      return JSON.stringify({ success: true, message: "Match flagged for admin review." });
    }
    
    if (action === "admin_cancel") {
      if (headers.indexOf("p1_civ") > -1) mSheet.getRange(row, headers.indexOf("p1_civ") + 1).setValue("");
      if (headers.indexOf("p2_civ") > -1) mSheet.getRange(row, headers.indexOf("p2_civ") + 1).setValue("");
      if (headers.indexOf("map_name") > -1) mSheet.getRange(row, headers.indexOf("map_name") + 1).setValue("");
      if (headers.indexOf("winner_id") > -1) mSheet.getRange(row, headers.indexOf("winner_id") + 1).setValue("");
      if (headers.indexOf("submitted_by") > -1) mSheet.getRange(row, headers.indexOf("submitted_by") + 1).setValue("");
      if (headers.indexOf("proposed_winner") > -1) mSheet.getRange(row, headers.indexOf("proposed_winner") + 1).setValue("");
      if (headers.indexOf("match_status") > -1) mSheet.getRange(row, headers.indexOf("match_status") + 1).setValue("pending");
      
      sync_to_stage_sheet(ss, stageName, matchId, { clearAll: true });
      processEloForMatchRow(ss, mSheet, row); 
      return JSON.stringify({ success: true, message: "Match reset successfully." });
    }
    
    return JSON.stringify({ success: false, error: "Unknown action." });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

// ---------------------------------------------------------
// UNIFIED ONEDIT: Master Listener
// ---------------------------------------------------------
function onEdit(e) {
  if (!e || !e.range) return;
  
  var sheet = e.source.getActiveSheet();
  if (sheet.getName() !== "matches") return;

  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(h) { 
    return String(h).toLowerCase().trim(); 
  });
  
  var winnerCol = headers.indexOf("winner_id") + 1; 
  var statusCol = headers.indexOf("match_status") + 1;
  var proposedWinnerCol = headers.indexOf("proposed_winner") + 1;

  if (e.range.getColumn() === winnerCol && e.range.getRow() > 1) {
    var newWinner = e.value;
    var row = e.range.getRow();
    
    if (newWinner && newWinner !== "") {
      sheet.getRange(row, statusCol).setValue("Settled");
      if (proposedWinnerCol > 0) sheet.getRange(row, proposedWinnerCol).clearContent();
    } else {
      sheet.getRange(row, statusCol).setValue("Active");
    }
  }
}

// ---------------------------------------------------------
// EloEngine Failsafe - Master Recalibration Engine
// ---------------------------------------------------------
function recalibrate_all_elo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const pSheet = ss.getSheetByName('players');
  if (!pSheet) return SpreadsheetApp.getUi().alert("Players sheet not found.");
  
  const pData = pSheet.getDataRange().getValues();
  const pHeaders = pData[0].map(h => String(h).toLowerCase().trim());
  const nameIdx = pHeaders.indexOf('player_name') > -1 ? pHeaders.indexOf('player_name') : pHeaders.indexOf('player');
  const eloIdx = pHeaders.indexOf("tournament_elo");
  
  if (nameIdx === -1 || eloIdx === -1) return SpreadsheetApp.getUi().alert("Player Name or Elo column missing.");

  const playersMap = {};
  for (let i = 1; i < pData.length; i++) {
    let name = String(pData[i][nameIdx]).trim();
    if (name) {
      playersMap[name] = { row: i + 1, elo: 1000 }; 
    }
  }

  const mSheet = ss.getSheetByName('matches');
  if (!mSheet) return SpreadsheetApp.getUi().alert("Matches sheet not found.");
  
  const mData = mSheet.getDataRange().getValues();
  const mHeaders = mData[0].map(h => String(h).toLowerCase().trim());
  
  const p1Col = mHeaders.indexOf('player 1') > -1 ? mHeaders.indexOf('player 1') : mHeaders.indexOf('player1_id');
  const p2Col = mHeaders.indexOf('player 2') > -1 ? mHeaders.indexOf('player 2') : mHeaders.indexOf('player2_id');
  const winnerCol = mHeaders.indexOf('winner') > -1 ? mHeaders.indexOf('winner') : mHeaders.indexOf('winner_id');
  const statusCol = mHeaders.indexOf('match_status') > -1 ? mHeaders.indexOf('match_status') : mHeaders.indexOf('status');
  
  let shift1Col = mHeaders.indexOf('p1_elo_shift');
  let shift2Col = mHeaders.indexOf('p2_elo_shift');
  
  if (shift1Col === -1 || shift2Col === -1) {
    shift1Col = mHeaders.length;
    shift2Col = mHeaders.length + 1;
    mSheet.getRange(1, shift1Col + 1).setValue('p1_elo_shift').setFontWeight('bold');
    mSheet.getRange(1, shift2Col + 1).setValue('p2_elo_shift').setFontWeight('bold');
  }

  const K_FACTOR = 32;
  const shift1Updates = [];
  const shift2Updates = [];
  
  for (let i = 1; i < mData.length; i++) {
    let status = String(mData[i][statusCol]).trim().toLowerCase();
    let p1 = String(mData[i][p1Col]).trim();
    let p2 = String(mData[i][p2Col]).trim();
    let winner = String(mData[i][winnerCol]).trim();

    if ((status === 'settled' || status === 'archived') && winner !== '' && winner !== 'Unplayed' && p1 !== 'BYE' && p2 !== 'BYE') {
      if (playersMap[p1] && playersMap[p2]) {
        let elo1 = playersMap[p1].elo;
        let elo2 = playersMap[p2].elo;

        let expected1 = 1 / (1 + Math.pow(10, (elo2 - elo1) / 400));
        let expected2 = 1 / (1 + Math.pow(10, (elo1 - elo2) / 400));

        let actual1 = (winner === p1) ? 1 : 0;
        let actual2 = (winner === p2) ? 1 : 0;

        let shift1 = Math.round(K_FACTOR * (actual1 - expected1));
        let shift2 = Math.round(K_FACTOR * (actual2 - expected2));

        playersMap[p1].elo += shift1;
        playersMap[p2].elo += shift2;
        
        shift1Updates.push([shift1]);
        shift2Updates.push([shift2]);
      } else {
        shift1Updates.push([""]);
        shift2Updates.push([""]);
      }
    } else {
      shift1Updates.push([""]);
      shift2Updates.push([""]);
    }
  }

  if (shift1Updates.length > 0) {
    mSheet.getRange(2, shift1Col + 1, shift1Updates.length, 1).setValues(shift1Updates);
    mSheet.getRange(2, shift2Col + 1, shift2Updates.length, 1).setValues(shift2Updates);
  }

  const eloUpdates = [];
  for (let i = 1; i < pData.length; i++) {
    let name = String(pData[i][nameIdx]).trim();
    if (playersMap[name]) {
      eloUpdates.push([playersMap[name].elo]);
    } else {
      eloUpdates.push([""]);
    }
  }
  pSheet.getRange(2, eloIdx + 1, eloUpdates.length, 1).setValues(eloUpdates);
  SpreadsheetApp.getUi().alert("✅ Recalibration Complete! All historical match shifts calculated and current Elos perfectly synchronized.");
}

// ---------------------------------------------------------
// RECONSTRUCTED HELPER: sync_to_stage_sheet
// ---------------------------------------------------------
function sync_to_stage_sheet(ss, stageName, matchId, data) {
  try {
    if (!stageName || !matchId) return;
    const sheet = ss.getSheetByName(stageName);
    if (!sheet) return;
    
    const sData = sheet.getDataRange().getValues();
    const headers = sData[0].map(h => String(h).toLowerCase().trim());
    const mIdCol = headers.indexOf("match_id");
    if (mIdCol === -1) return;
    
    const rowIndex = sData.findIndex(r => String(r[mIdCol]) === String(matchId));
    if (rowIndex === -1) return;
    const row = rowIndex + 1;
    
    if (data.clearAll) {
       if (headers.indexOf("p1 civ") > -1) sheet.getRange(row, headers.indexOf("p1 civ") + 1).setValue("");
       if (headers.indexOf("p2 civ") > -1) sheet.getRange(row, headers.indexOf("p2 civ") + 1).setValue("");
       if (headers.indexOf("map") > -1) sheet.getRange(row, headers.indexOf("map") + 1).setValue("");
       if (headers.indexOf("winner") > -1) sheet.getRange(row, headers.indexOf("winner") + 1).setValue("");
       return;
    }
    
    if (data.p1Civ !== undefined && headers.indexOf("p1 civ") > -1) sheet.getRange(row, headers.indexOf("p1 civ") + 1).setValue(data.p1Civ);
    if (data.p2Civ !== undefined && headers.indexOf("p2 civ") > -1) sheet.getRange(row, headers.indexOf("p2 civ") + 1).setValue(data.p2Civ);
    if (data.map !== undefined && headers.indexOf("map") > -1) sheet.getRange(row, headers.indexOf("map") + 1).setValue(data.map);
    if (data.winner !== undefined && headers.indexOf("winner") > -1) sheet.getRange(row, headers.indexOf("winner") + 1).setValue(data.winner);
  } catch (e) {
    console.error("Stage Sync Error: " + e.message);
  }
}

// =========================================================================
// ADMIN DATA ENDPOINTS
// =========================================================================

function get_admin_queue() {
  // Leave your existing get_admin_queue() function exactly as it is here
  // ...
}

function admin_force_resolve(matchId, actionType) {
  // Leave your existing admin_force_resolve() function exactly as it is here
  // ...
}

function get_all_admin_matches() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName("matches");
    if (!mSheet) return JSON.stringify({success:false, data:[]});

    const data = mSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const mIdCol = headers.indexOf("match_id");
    const stageCol = headers.indexOf("stage");
    const p1Col = headers.indexOf("player1_id");
    const p2Col = headers.indexOf("player2_id");
    const statusCol = headers.indexOf("match_status");
    
    // Official Data Columns
    const p1CivCol = headers.indexOf("p1_civ");
    const p2CivCol = headers.indexOf("p2_civ");
    const mapCol = headers.indexOf("map_name");
    const winnerCol = headers.indexOf("winner_id");
    
    // Holding Pen / Proposed Data Columns
    const propP1CivCol = headers.indexOf("proposed_p1_civ");
    const propP2CivCol = headers.indexOf("proposed_p2_civ");
    const propMapCol = headers.indexOf("proposed_map");
    const propWinnerCol = headers.indexOf("proposed_winner");

    let allMatches = [];
    for (let i = 1; i < data.length; i++) {
      let statusStr = String(data[i][statusCol]).trim();
      
      if (statusStr.toLowerCase() !== "archived") {
        allMatches.push({
          id: data[i][mIdCol],
          stage: String(data[i][stageCol]).trim() || "Unassigned",
          p1: String(data[i][p1Col]).trim(),
          p2: String(data[i][p2Col]).trim(),
          status: statusStr,
          // Prefill with proposed data if it exists; otherwise use official data
          p1Civ: (propP1CivCol > -1 && data[i][propP1CivCol]) ? data[i][propP1CivCol] : (p1CivCol > -1 ? data[i][p1CivCol] : ""),
          p2Civ: (propP2CivCol > -1 && data[i][propP2CivCol]) ? data[i][propP2CivCol] : (p2CivCol > -1 ? data[i][p2CivCol] : ""),
          map: (propMapCol > -1 && data[i][propMapCol]) ? data[i][propMapCol] : (mapCol > -1 ? data[i][mapCol] : ""),
          winner: (propWinnerCol > -1 && data[i][propWinnerCol]) ? data[i][propWinnerCol] : (winnerCol > -1 ? data[i][winnerCol] : "")
        });
      }
    }

    let civList = [];
    const civSheet = ss.getSheetByName("config_civs");
    if (civSheet) {
      const cData = civSheet.getDataRange().getValues();
      for(let i=1; i<cData.length; i++) {
        if(cData[i][1]) civList.push(String(cData[i][1]).trim());
      }
    }

    let mapList = [];
    const mapSheet = ss.getSheetByName("config_maps");
    if (mapSheet) {
      const mData = mapSheet.getDataRange().getValues();
      const mHeaders = mData[0].map(h => String(h).toLowerCase().trim());
      const mNameCol = mHeaders.indexOf("map_name");
      const mStatCol = mHeaders.indexOf("map_status");
      if (mNameCol > -1) {
        for(let i=1; i<mData.length; i++) {
          if (mStatCol === -1 || String(mData[i][mStatCol]).toLowerCase() === "active") {
            mapList.push(String(mData[i][mNameCol]).trim());
          }
        }
      }
    }

    return JSON.stringify({ success: true, data: allMatches, civs: civList, maps: mapList });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

function admin_save_match(matchId, payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mSheet = ss.getSheetByName("matches");
    const data = mSheet.getDataRange().getValues();
    const headers = data[0].map(h => String(h).toLowerCase().trim());
    
    const rowIndex = data.findIndex(r => String(r[headers.indexOf("match_id")]) === matchId);
    if (rowIndex === -1) return JSON.stringify({success: false, error: "Match not found in database."});
    const row = rowIndex + 1;

    let p1CivCol = headers.indexOf("p1_civ");
    let p2CivCol = headers.indexOf("p2_civ");
    let mapCol = headers.indexOf("map_name");
    let winnerCol = headers.indexOf("winner_id");
    let statusCol = headers.indexOf("match_status");

    if (p1CivCol > -1) mSheet.getRange(row, p1CivCol + 1).setValue(payload.p1Civ);
    if (p2CivCol > -1) mSheet.getRange(row, p2CivCol + 1).setValue(payload.p2Civ);
    if (mapCol > -1) mSheet.getRange(row, mapCol + 1).setValue(payload.map);
    if (winnerCol > -1) mSheet.getRange(row, winnerCol + 1).setValue(payload.winner);
    
    let newStatus = "Active";
    if (payload.winner && payload.winner.trim() !== "") {
       newStatus = "Settled";
    }
    if (statusCol > -1) mSheet.getRange(row, statusCol + 1).setValue(newStatus);

    const wipeCols = ["proposed_p1_civ", "proposed_p2_civ", "proposed_map", "proposed_winner", "submitted_by"];
    wipeCols.forEach(col => {
       let idx = headers.indexOf(col);
       if (idx > -1) mSheet.getRange(row, idx + 1).setValue("");
    });

    if (newStatus === "Settled") {
       let eloResult = processEloForMatchRow(ss, mSheet, row);
       if (eloResult !== "Success") {
          throw new Error("Data Saved, but Elo Aborted: " + eloResult);
       }
    }

    // Push updates to the specific generated tournament sheet (League, Group, or Playoff)
    sync_match_to_tournament_sheets(ss, matchId, {
      p1Civ: payload.p1Civ,
      p2Civ: payload.p2Civ,
      map: payload.map,
      winner: payload.winner,
      status: newStatus
    });

    return JSON.stringify({success: true});

  } catch(e) {
    return JSON.stringify({success: false, error: e.toString()});
  }
}

function sync_match_to_tournament_sheets(ss, matchId, updates) {
  try {
    if (!matchId) return;
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      const name = sheet.getName().toLowerCase();
      // Skip system sheets that aren't tournament brackets/groups/leagues
      if (["players", "tournaments", "config_civs", "config_maps", "bets", "bettors", "matches"].includes(name)) return;
      
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return;
      
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).toLowerCase().trim());
      const mIdCol = headers.indexOf("match_id");
      
      if (mIdCol === -1) return;
      
      // Locate the exact row matching this matchId across any group or bracket sheet
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][mIdCol]).trim() === String(matchId).trim()) {
          const row = i + 1;
          
          const p1CivCol = headers.indexOf("p1_civ");
          const p2CivCol = headers.indexOf("p2_civ");
          const mapCol = headers.indexOf("map_name") > -1 ? headers.indexOf("map_name") : headers.indexOf("map");
          const winnerCol = headers.indexOf("winner_id") > -1 ? headers.indexOf("winner_id") : headers.indexOf("winner");
          const statusCol = headers.indexOf("match_status") > -1 ? headers.indexOf("match_status") : headers.indexOf("status");
          
          if (updates.p1Civ !== undefined && p1CivCol > -1) sheet.getRange(row, p1CivCol + 1).setValue(updates.p1Civ);
          if (updates.p2Civ !== undefined && p2CivCol > -1) sheet.getRange(row, p2CivCol + 1).setValue(updates.p2Civ);
          if (updates.map !== undefined && mapCol > -1) sheet.getRange(row, mapCol + 1).setValue(updates.map);
          if (updates.winner !== undefined && winnerCol > -1) sheet.getRange(row, winnerCol + 1).setValue(updates.winner);
          if (updates.status !== undefined && statusCol > -1) sheet.getRange(row, statusCol + 1).setValue(updates.status);
          
          break;
        }
      }
    });
  } catch (e) {
    console.error("Universal Sync Error: " + e.message);
  }
}

function get_betting_data() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Fetch Leaderboard
    let bSheet = ss.getSheetByName("bettors") || ss.getSheetByName("bet_leaderboard") || ss.getSheetByName("leaderboard");
    let leaderboard = [];
    if (bSheet && bSheet.getLastRow() > 1) {
      const bData = bSheet.getDataRange().getValues();
      const bHeaders = bData[0].map(h => String(h).toLowerCase().trim());
      const tIdCol = bHeaders.indexOf("telegram_id");
      const nameCol = bHeaders.indexOf("first_name");
      const userCol = bHeaders.indexOf("username");
      const ptsCol = bHeaders.indexOf("points");
      
      for (let i = 1; i < bData.length; i++) {
        leaderboard.push({
          telegramId: tIdCol > -1 ? bData[i][tIdCol] : "",
          firstName: nameCol > -1 ? bData[i][nameCol] : "",
          username: userCol > -1 ? bData[i][userCol] : "",
          points: ptsCol > -1 ? (Number(bData[i][ptsCol]) || 0) : 0
        });
      }
      leaderboard.sort((a, b) => b.points - a.points);
    }

    // 2. Fetch Wager Log (Capped to latest 20)
    let lSheet = ss.getSheetByName("bet_logs") || ss.getSheetByName("bet_log") || ss.getSheetByName("betslog");
    let betsLog = [];
    if (lSheet && lSheet.getLastRow() > 1) {
      const lData = lSheet.getDataRange().getValues();
      const lHeaders = lData[0].map(h => String(h).toLowerCase().trim());
      const pollCol = lHeaders.indexOf("poll_id");
      const nameCol = lHeaders.indexOf("first_name");
      const userCol = lHeaders.indexOf("username");
      const optCol = lHeaders.indexOf("option_id");
      const payoutCol = lHeaders.indexOf("payout");
      const outcomeCol = lHeaders.indexOf("outcome");
      const dateCol = lHeaders.indexOf("created_at");

      for (let i = 1; i < lData.length; i++) {
        betsLog.push({
          pollId: pollCol > -1 ? lData[i][pollCol] : "",
          firstName: nameCol > -1 ? lData[i][nameCol] : "",
          username: userCol > -1 ? lData[i][userCol] : "",
          optionId: optCol > -1 ? lData[i][optCol] : "",
          payout: payoutCol > -1 ? lData[i][payoutCol] : 0,
          outcome: outcomeCol > -1 ? lData[i][outcomeCol] : "Pending",
          createdAt: dateCol > -1 ? lData[i][dateCol] : ""
        });
      }
      betsLog.reverse();             // Put newest wagers first
      betsLog = betsLog.slice(0, 20); // Restrict to the top 20 most recent entries
    }

    return JSON.stringify({ success: true, leaderboard: leaderboard, logs: betsLog });
  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}
function get_player_telegram(playerName) {
  try {
    const pSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("players");
    if (!pSheet) return "";
    const pData = pSheet.getDataRange().getValues();
    const headers = pData[0].map(h => String(h).toLowerCase().trim());
    
    const nameCol = headers.indexOf("player_name") > -1 ? headers.indexOf("player_name") : headers.indexOf("player");
    let tgCol = headers.indexOf("telegram_username");
    if (tgCol === -1) tgCol = headers.indexOf("telegram");
    if (tgCol === -1) tgCol = headers.indexOf("username");
    
    if (nameCol === -1 || tgCol === -1) return "";
    
    for (let i = 1; i < pData.length; i++) {
      if (String(pData[i][nameCol]).trim() === String(playerName).trim()) {
        return String(pData[i][tgCol]).trim().replace("@", ""); 
      }
    }
    return "";
  } catch (e) {
    return "";
  }
}

function get_registered_players() {
  try {
    // Change "Registrations" to match your actual sheet name if it differs
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Registration"); 
    if (!sheet) return JSON.stringify({ success: false, error: "Sheet not found" });
    
    var data = sheet.getDataRange().getValues();
    // Assuming row 1 contains headers
    if (data.length < 2) return JSON.stringify({ success: true, data: [] });
    
    var players = [];
    for (var i = 1; i < data.length; i++) {
      players.push({
        timestamp: data[i][0] ? data[i][0].toString() : "",
        name: data[i][1] ? data[i][1].toString() : "", 
        contact: data[i][2] ? data[i][2].toString() : ""
      });
    }
    
    return JSON.stringify({ success: true, data: players });
  } catch (e) {
    return JSON.stringify({ success: false, error: e.message });
  }
}