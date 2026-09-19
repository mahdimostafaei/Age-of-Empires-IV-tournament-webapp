// =========================================================================
// TournamentViewData.gs - Universal Data Bridge with Code-Driven Standings
// =========================================================================

function get_tournament_page_data() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = { hasActive: false, tournament: null, groups: [], bracketRounds: {}, rawMatches: [], history: [] };

    // 1. STRICT TOURNEY LEDGER FETCH
    const tSheet = ss.getSheetByName("tournaments");
    if (tSheet) {
      const tData = tSheet.getDataRange().getValues();
      if (tData.length > 1) {
        const headers = tData[0].map(h => String(h).toLowerCase().trim());
        const nameIdx = headers.indexOf("tournament_name");
        const typeIdx = headers.indexOf("tournament_type");
        const statusIdx = headers.indexOf("status");
        const winnerIdx = headers.indexOf("winner_id");
        const matchesIdx = headers.indexOf("total_matches");
        const startDateIdx = headers.indexOf("start_date");
        const dateIdx = headers.indexOf("end_date");
        const civIdx = headers.indexOf("most_played_civ");

        for (let i = 1; i < tData.length; i++) {
          if (statusIdx === -1) continue; 
          const status = String(tData[i][statusIdx]).trim().toLowerCase();
          
          if (status === "active") {
            result.hasActive = true;
            result.tournament = {
              name: nameIdx > -1 ? String(tData[i][nameIdx]).trim() : "Unknown Name",
              type: typeIdx > -1 ? String(tData[i][typeIdx]).trim() : "Tournament",
              status: "Active"
            };
          } 
          else if (status === "completed" || status === "archived") {
            let safeEndDate = "N/A";
            if (dateIdx > -1 && tData[i][dateIdx]) {
              let rawDate = tData[i][dateIdx];
              safeEndDate = rawDate instanceof Date ? Utilities.formatDate(rawDate, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy") : String(rawDate);
            }

            let safeStartDate = "";
            if (startDateIdx > -1 && tData[i][startDateIdx]) {
              let rawStart = tData[i][startDateIdx];
              safeStartDate = rawStart instanceof Date ? Utilities.formatDate(rawStart, ss.getSpreadsheetTimeZone(), "MM/dd/yyyy") : String(rawStart);
            }

            result.history.push({
              name: nameIdx > -1 ? String(tData[i][nameIdx]).trim() : "",
              type: typeIdx > -1 ? String(tData[i][typeIdx]).trim() : "",
              winner: winnerIdx > -1 ? String(tData[i][winnerIdx]).trim() : "",
              matches: matchesIdx > -1 ? String(tData[i][matchesIdx]).trim() : "-",
              startDate: safeStartDate,
              endDate: safeEndDate,
              topCiv: civIdx > -1 ? String(tData[i][civIdx]).trim() : "-"
            });
          }
        }
      }
    }

    if (!result.hasActive) return JSON.stringify(result);

    const allSheets = ss.getSheets();

    // 2. UNIVERSAL GROUP & LEAGUE STAGE FETCH (Dynamic Standings Calculation)
    const groupSheets = allSheets.filter(s => {
      const name = s.getName().toLowerCase();
      return (name.includes("league") || name.includes("group")) && !s.isSheetHidden() && !name.includes("(");
    });
    
    groupSheets.forEach(gSheet => {
      try {
        const fullData = gSheet.getDataRange().getValues();
        if (fullData.length > 1) {
          const mHeaders = fullData[0].map(h => String(h).toLowerCase().trim());
          const p1Idx = mHeaders.findIndex(h => h.includes("player 1") || h.includes("player1") || h === "player1_id");
          const p2Idx = mHeaders.findIndex(h => h.includes("player 2") || h.includes("player2") || h === "player2_id");
          const winIdx = mHeaders.findIndex(h => h.includes("winner") || h === "winner_id");
          const statusIdx = mHeaders.findIndex(h => h.includes("status"));

          let stats = {};

          for (let r = 1; r < fullData.length; r++) {
            if (!fullData[r]) continue;
            let p1 = p1Idx > -1 ? String(fullData[r][p1Idx] || "").trim() : "";
            let p2 = p2Idx > -1 ? String(fullData[r][p2Idx] || "").trim() : "";
            let winner = winIdx > -1 ? String(fullData[r][winIdx] || "").trim() : "";
            let status = statusIdx > -1 ? String(fullData[r][statusIdx] || "").trim().toLowerCase() : "";

            if (p1 && p1 !== "BYE") {
              if (!stats[p1]) stats[p1] = { player: p1, wins: 0, losses: 0 };
            }
            if (p2 && p2 !== "BYE") {
              if (!stats[p2]) stats[p2] = { player: p2, wins: 0, losses: 0 };
            }

            if (status === "settled" && winner && winner !== "Unplayed") {
              let winKey = winner.toLowerCase();
              if (p1 && p1.toLowerCase() === winKey) {
                if (stats[p1]) stats[p1].wins++;
                if (stats[p2]) stats[p2].losses++;
              } else if (p2 && p2.toLowerCase() === winKey) {
                if (stats[p2]) stats[p2].wins++;
                if (stats[p1]) stats[p1].losses++;
              }
            }
          }

          // Convert to array and sort strictly by Wins (descending), then Win Rate
          let standings = Object.values(stats).map(s => {
            let total = s.wins + s.losses;
            let winRate = total > 0 ? (s.wins / total) : 0;
            return {
              player: s.player,
              score: s.wins,
              losses: s.losses,
              winRate: winRate
            };
          });

          standings.sort((a, b) => {
            if (b.score !== a.score) return b.score - a.score;
            return b.winRate - a.winRate;
          });

          // Add ranks
          standings.forEach((st, idx) => {
            st.rank = idx + 1;
          });

          if (standings.length > 0) {
            result.groups.push({ groupName: gSheet.getName(), standings: standings });
          }

          // Extract Matches
          extractMatchesFromSheet(gSheet, result);
        }
      } catch (sheetErr) {}
    });

    // 3. UNIVERSAL KNOCKOUT / PLAYOFF / BRACKET FETCH
    const bracketSheets = allSheets.filter(s => {
      const name = s.getName().toLowerCase();
      return (name.includes("knockout") || name.includes("playoff") || name.includes("bracket")) && !s.isSheetHidden() && !name.includes("(");
    });

    bracketSheets.forEach(bSheet => {
      try {
        extractMatchesFromSheet(bSheet, result);

        const bData = bSheet.getDataRange().getValues();
        if (bData.length > 1) {
          const headers = bData[0].map(h => String(h).toLowerCase().trim());
          const stageIdx = headers.indexOf("stage");
          const p1Idx = headers.indexOf("player 1") > -1 ? headers.indexOf("player 1") : headers.indexOf("player1_id");
          const p2Idx = headers.indexOf("player 2") > -1 ? headers.indexOf("player 2") : headers.indexOf("player2_id");
          const winnerIdx = headers.indexOf("winner") > -1 ? headers.indexOf("winner") : headers.indexOf("winner_id");

          for (let i = 1; i < bData.length; i++) {
            if (stageIdx === -1 || !bData[i]) continue;
            let rawStage = String(bData[i][stageIdx] || "").trim();
            if (!rawStage || rawStage === "undefined" || rawStage === "") continue;

            let cleanStage = rawStage.replace(/\s+\d+$/, "");
            const p1 = p1Idx > -1 ? String(bData[i][p1Idx] || "").trim() : "";
            const p2 = p2Idx > -1 ? String(bData[i][p2Idx] || "").trim() : "";
            const winner = winnerIdx > -1 ? String(bData[i][winnerIdx] || "").trim() : "";

            if (!result.bracketRounds[cleanStage]) result.bracketRounds[cleanStage] = [];
            result.bracketRounds[cleanStage].push({ player1: p1, player2: p2, winner: winner, rawStage: rawStage });
          }
        }
      } catch(e) {}
    });

    return JSON.stringify(result);

  } catch (err) {
    return JSON.stringify({ error: err.toString() });
  }
}

function extractMatchesFromSheet(sheet, result) {
  const fullData = sheet.getDataRange().getValues();
  if (fullData.length <= 1) return;
  
  const mHeaders = fullData[0].map(h => String(h).toLowerCase().trim());
  const mIdIdx = mHeaders.findIndex(h => h === "match_id" || h === "match id");
  const p1Idx = mHeaders.findIndex(h => h.includes("player 1") || h.includes("player1") || h === "player1_id");
  const p1CivIdx = mHeaders.findIndex(h => h.includes("p1 civ") || h.includes("p1_civ"));
  const p2CivIdx = mHeaders.findIndex(h => h.includes("p2 civ") || h.includes("p2_civ"));
  const p2Idx = mHeaders.findIndex(h => h.includes("player 2") || h.includes("player2") || h === "player2_id");
  const mapIdx = mHeaders.findIndex(h => h === "map" || h === "map_name");
  const winIdx = mHeaders.findIndex(h => h.includes("winner") || h === "winner_id");

  for (let r = 1; r < fullData.length; r++) {
    if (!fullData[r]) continue;
    let p1 = p1Idx > -1 ? String(fullData[r][p1Idx] || "").trim() : "";
    let p2 = p2Idx > -1 ? String(fullData[r][p2Idx] || "").trim() : "";
    
    if (p1 && p1 !== "" && p1 !== "undefined") {
      let matchId = (mIdIdx > -1 && fullData[r][mIdIdx]) ? String(fullData[r][mIdIdx]).trim() : "-";

      result.rawMatches.push({
        league: sheet.getName(),
        matchId: matchId,
        player1: p1,
        p1Civ: (p1CivIdx > -1 && fullData[r][p1CivIdx]) ? String(fullData[r][p1CivIdx]).trim() : "-",
        p2Civ: (p2CivIdx > -1 && fullData[r][p2CivIdx]) ? String(fullData[r][p2CivIdx]).trim() : "-",
        player2: p2,
        map: (mapIdx > -1 && fullData[r][mapIdx]) ? String(fullData[r][mapIdx]).trim() : "-",
        winner: (winIdx > -1 && fullData[r][winIdx]) ? String(fullData[r][winIdx]).trim() : ""
      });
    }
  }
}