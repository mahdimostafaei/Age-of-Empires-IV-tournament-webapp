// =========================================================================
// MASTER MENU: Tournament Tools (Generator & Archival Combined)
// =========================================================================
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏆 Tournament Tools')
    .addItem('⚙️ Generate Tournament...', 'open_tournament_dialog') 
    .addSeparator() 
    .addItem('📦 Archive Active Tournament', 'archive_active_tournament') 
    .addToUi();
}

// =========================================================================
// ARCHIVAL PIPELINE & METRICS SWEEPER (Batch League Support)
// =========================================================================
function archive_active_tournament() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  const tSheet = ss.getSheetByName("tournaments");
  if (!tSheet) return ui.alert("Error: 'tournaments' sheet not found.");
  
  const tData = tSheet.getDataRange().getValues();
  const headers = tData[0].map(h => String(h).toLowerCase().trim());
  const statusIdx = headers.indexOf("status");
  const idIdx = headers.indexOf("tournament_id");
  const nameIdx = headers.indexOf("tournament_name");
  const dateIdx = headers.indexOf("end_date");
  
  const tPartsCol = headers.indexOf("total_participants");
  const tMatchesCol = headers.indexOf("total_matches");
  const tWinCol = headers.indexOf("winner_id");
  const tRunCol = headers.indexOf("runner_up_id");
  const tMapCol = headers.indexOf("fav_map_id");
  const tCivCol = headers.indexOf("most_played_civ");
  
  let activeTournaments = [];
  let isLeagueBatch = false;
  let targetBaseName = "";
  let singleTargetId = "";
  
  // 1. Identify active tournaments and detect if it is a League Batch
  for (let i = 1; i < tData.length; i++) {
    if (String(tData[i][statusIdx]).trim().toLowerCase() === "active") {
      let tName = String(tData[i][nameIdx]).trim();
      let tId = String(tData[i][idIdx]).trim();
      
      if (tName.includes(" - League")) {
        isLeagueBatch = true;
        targetBaseName = tName.split(" - League")[0];
      } else {
        targetBaseName = tName;
        singleTargetId = tId;
      }
      break; 
    }
  }
  
  if (!targetBaseName) return ui.alert("No active tournament found to archive.");
  
  // 2. Group all related leagues into the batch
  if (isLeagueBatch) {
    for (let i = 1; i < tData.length; i++) {
      if (String(tData[i][statusIdx]).trim().toLowerCase() === "active" && 
          String(tData[i][nameIdx]).trim().startsWith(targetBaseName + " - League")) {
        activeTournaments.push({ rowIdx: i, id: String(tData[i][idIdx]).trim(), name: String(tData[i][nameIdx]).trim() });
      }
    }
  } else {
    for (let i = 1; i < tData.length; i++) {
      if (String(tData[i][idIdx]).trim() === singleTargetId) {
        activeTournaments.push({ rowIdx: i, id: singleTargetId, name: targetBaseName });
        break;
      }
    }
  }
  
  // 3. Dynamic Safety Prompt
  let promptMsg = isLeagueBatch 
    ? `You are about to permanently archive ALL leagues for:\nName: ${targetBaseName}\n\nTo proceed, type the exact BASE NAME:\n"${targetBaseName}"`
    : `You are about to permanently archive:\nName: ${targetBaseName}\n\nTo proceed, type the exact TOURNAMENT ID:\n"${singleTargetId}"`;
  
  const requiredInput = isLeagueBatch ? targetBaseName : singleTargetId;
  const response = ui.prompt("⚠️ CRITICAL: Confirm Archival", promptMsg, ui.ButtonSet.OK_CANCEL);

  if (response.getSelectedButton() !== ui.Button.OK) return ui.alert("Archival aborted.");
  if (response.getResponseText().trim() !== requiredInput) return ui.alert("Mismatch! Incorrect input. Archival aborted.");
  
  // 4. FREEZE ELO (Global)
  const pSheet = ss.getSheetByName("players");
  if (pSheet) {
    const pData = pSheet.getDataRange().getValues();
    const pHeaders = pData[0].map(h => String(h).toLowerCase().trim());
    const startEloIdx = pHeaders.indexOf("starting_elo");
    const liveEloIdx = pHeaders.indexOf("live_elo");
    
    if (startEloIdx > -1 && liveEloIdx > -1) {
      const eloUpdates = [];
      for (let i = 1; i < pData.length; i++) {
        let live = pData[i][liveEloIdx];
        eloUpdates.push([live !== "" ? live : pData[i][startEloIdx]]);
      }
      if (eloUpdates.length > 0) pSheet.getRange(2, startEloIdx + 1, eloUpdates.length, 1).setValues(eloUpdates);
    }
  }

  const timestamp = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), "MMdd");
  const me = Session.getEffectiveUser();
  const allSheets = ss.getSheets();
  const mSheet = ss.getSheetByName("matches");

  // 5. PROCESS EACH TOURNAMENT IN THE BATCH
  for (let t of activeTournaments) {
    let matchCount = 0;
    let participants = new Set();
    let mapCounts = {};
    let civCounts = {};
    let tourneyWinner = "";
    let tourneyRunnerUp = "";

    // Quarantine Matches & Gather Metrics
    if (mSheet) {
      const mData = mSheet.getDataRange().getValues();
      const mHeaders = mData[0].map(h => String(h).toLowerCase().trim());
      
      const mStatusCol = mHeaders.indexOf("match_status") > -1 ? mHeaders.indexOf("match_status") : mHeaders.indexOf("status");
      const tIdCol = mHeaders.indexOf("tournament_id");
      const p1Col = mHeaders.indexOf("player1_id") > -1 ? mHeaders.indexOf("player1_id") : mHeaders.indexOf("player 1");
      const p2Col = mHeaders.indexOf("player2_id") > -1 ? mHeaders.indexOf("player2_id") : mHeaders.indexOf("player 2");
      const winnerCol = mHeaders.indexOf("winner_id") > -1 ? mHeaders.indexOf("winner_id") : mHeaders.indexOf("winner");
      const stageCol = mHeaders.indexOf("stage");
      const mapCol = mHeaders.indexOf("map_name") > -1 ? mHeaders.indexOf("map_name") : mHeaders.indexOf("map");
      const p1CivCol = mHeaders.indexOf("p1_civ");
      const p2CivCol = mHeaders.indexOf("p2_civ");

      if (mStatusCol > -1 && tIdCol > -1) {
        const statusUpdates = [];
        for (let i = 1; i < mData.length; i++) {
          let currentStatus = String(mData[i][mStatusCol]).trim().toLowerCase();
          let rTid = String(mData[i][tIdCol]).trim();
          let newStatus = String(mData[i][mStatusCol]).trim();

          if (rTid === t.id) {
            if (currentStatus === "settled" || currentStatus === "completed" || currentStatus === "pending") {
              newStatus = "Archived";
            }
            
            let w = winnerCol > -1 ? String(mData[i][winnerCol]).trim() : "";
            let p1 = p1Col > -1 ? String(mData[i][p1Col]).trim() : "";
            let p2 = p2Col > -1 ? String(mData[i][p2Col]).trim() : "";
            let stage = stageCol > -1 ? String(mData[i][stageCol]).trim().toLowerCase() : "";

            if (w !== "" && w.toLowerCase() !== "unplayed" && p1.toLowerCase() !== "bye" && p2.toLowerCase() !== "bye" && (currentStatus === "settled" || currentStatus === "completed" || currentStatus === "archived")) {
              matchCount++;
              if (p1) participants.add(p1);
              if (p2) participants.add(p2);

              let map = mapCol > -1 ? String(mData[i][mapCol]).trim() : "";
              if (map) mapCounts[map] = (mapCounts[map] || 0) + 1;
              let civ1 = p1CivCol > -1 ? String(mData[i][p1CivCol]).trim() : "";
              if (civ1) civCounts[civ1] = (civCounts[civ1] || 0) + 1;
              let civ2 = p2CivCol > -1 ? String(mData[i][p2CivCol]).trim() : "";
              if (civ2) civCounts[civ2] = (civCounts[civ2] || 0) + 1;

              if (stage.includes("final") && !stage.includes("quarter") && !stage.includes("semi")) {
                 tourneyWinner = w;
                 tourneyRunnerUp = (w === p1) ? p2 : p1;
              }
            }
          }
          statusUpdates.push([newStatus]);
        }
        if (statusUpdates.length > 0) mSheet.getRange(2, mStatusCol + 1, statusUpdates.length, 1).setValues(statusUpdates);
      }
    }

    // Fallback Winner Detection (Targets specific active league name)
    if (!tourneyWinner) {
       let leagueNameMarker = t.name.split(" - ")[1]; 
       let targetSheetName = leagueNameMarker ? leagueNameMarker.toLowerCase() : "league";
       
       let leagueSheet = allSheets.find(s => s.getName().toLowerCase().includes(targetSheetName) && !s.getName().toLowerCase().includes("archived"));
       if (leagueSheet) {
          const lData = leagueSheet.getDataRange().getValues();
          for(let r = 0; r < lData.length; r++) {
             let rowStr = lData[r].map(c => String(c).toLowerCase()).join(",");
             if (rowStr.includes("rank") && rowStr.includes("player")) {
                tourneyWinner = String(lData[r+1][1]).trim(); 
                tourneyRunnerUp = String(lData[r+2][1]).trim();
                break;
             }
          }
       }
    }

    // Finalize Ledger for this specific tournament row
    let topMap = Object.keys(mapCounts).sort((a,b) => mapCounts[b] - mapCounts[a])[0] || "N/A";
    let topCiv = Object.keys(civCounts).sort((a,b) => civCounts[b] - civCounts[a])[0] || "N/A";

    tSheet.getRange(t.rowIdx + 1, statusIdx + 1).setValue("Archived");
    
    if (dateIdx > -1) tSheet.getRange(t.rowIdx + 1, dateIdx + 1).setValue(new Date());
    if (tPartsCol > -1) tSheet.getRange(t.rowIdx + 1, tPartsCol + 1).setValue(participants.size);
    if (tMatchesCol > -1) tSheet.getRange(t.rowIdx + 1, tMatchesCol + 1).setValue(matchCount);
    if (tWinCol > -1 && tourneyWinner) tSheet.getRange(t.rowIdx + 1, tWinCol + 1).setValue(tourneyWinner);
    if (tRunCol > -1 && tourneyRunnerUp) tSheet.getRange(t.rowIdx + 1, tRunCol + 1).setValue(tourneyRunnerUp);
    if (tMapCol > -1) tSheet.getRange(t.rowIdx + 1, tMapCol + 1).setValue(topMap);
    if (tCivCol > -1) tSheet.getRange(t.rowIdx + 1, tCivCol + 1).setValue(topCiv);
  }

  // 6. CLEANUP & PROTECT STAGE SHEETS
  allSheets.forEach(sheet => {
    let name = sheet.getName().toLowerCase();
    if ((name.includes("league") || name.includes("knockout") || name.includes("playoff") || name.includes("group")) && !name.includes("archived")) {
      
      let protection = sheet.protect().setDescription('Archived Stage - Locked');
      protection.removeEditors(protection.getEditors());
      if (protection.canDomainEdit()) protection.setDomainEdit(false);
      protection.addEditor(me);
      
      let cleanName = sheet.getName().replace(/ - .*/g, "");
      sheet.setName(`${cleanName} - ${isLeagueBatch ? targetBaseName : singleTargetId} (${timestamp})`);
      sheet.hideSheet();
    }
  });

  ui.alert("✅ Archival Complete!", `"${targetBaseName}" has been successfully archived. Stage sheets are now protected and read-only.`, ui.ButtonSet.OK);
}