// =========================================================================
// ApiAoE4.gs - High-Speed AoE4 World API Integration & Elo Sync
// =========================================================================

const AOE4_API_BASE = "https://aoe4world.com/api/v0";

/**
 * Synchronizes player Elo using parallel batch processing.
 * Priority: Ranked Matchmaking Elo (True MMR) -> Quick Match Elo
 */
function sync_all_players_elo() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('players');
  if (!sheet) throw new Error("Could not find 'players' sheet.");

  const data = sheet.getDataRange().getValues();
  const headers = data[0].map(h => String(h).toLowerCase().trim());
  
  const idColIdx = headers.indexOf('aoe4_world_id');
  if (idColIdx === -1) throw new Error("Could not find 'aoe4_world_id' column.");

  // Map header indexes dynamically so we don't hardcode column letters
  const colMap = {
    pc_1v1: headers.indexOf('pc_1v1'),
    pc_2v2: headers.indexOf('pc_2v2'),
    pc_3v3: headers.indexOf('pc_3v3'),
    pc_4v4: headers.indexOf('pc_4v4'),
    console_1v1: headers.indexOf('console_1v1'),
    last_api_sync: headers.indexOf('last_api_sync')
  };

  const requests = [];
  const rowMapping = []; // Tracks which array index corresponds to which API request

  // 1. Build parallel fetch requests
  for (let i = 1; i < data.length; i++) {
    const aoe4_id = data[i][idColIdx];
    if (aoe4_id) {
      requests.push({
        url: `${AOE4_API_BASE}/players/${aoe4_id}`,
        muteHttpExceptions: true
      });
      rowMapping.push(i);
    }
  }

  if (requests.length === 0) {
    Logger.log("No valid AoE4 World IDs found to sync.");
    return;
  }

  // 2. Fire all API requests simultaneously
  Logger.log(`Firing ${requests.length} API requests in parallel...`);
  const responses = UrlFetchApp.fetchAll(requests);
  
  let update_count = 0;
  const syncTime = new Date().toISOString();

  // 3. Process responses and update memory array
  for (let j = 0; j < responses.length; j++) {
    if (responses[j].getResponseCode() !== 200) continue;
    
    try {
      const api_data = JSON.parse(responses[j].getContentText());
      const modes = api_data.modes || {};
      
      const pc_1v1 = modes.rm_1v1_elo?.rating || modes.rm_solo_elo?.rating || modes.qm_1v1?.rating || "";
      const pc_2v2 = modes.rm_2v2_elo?.rating || modes.qm_2v2?.rating || "";
      const pc_3v3 = modes.rm_3v3_elo?.rating || modes.qm_3v3?.rating || "";
      const pc_4v4 = modes.rm_4v4_elo?.rating || modes.qm_4v4?.rating || "";
      const console_1v1 = modes.rm_1v1_console_elo?.rating || modes.rm_solo_console_elo?.rating || modes.qm_1v1_console?.rating || "";
      
      const rowIndex = rowMapping[j];
      
      if (colMap.pc_1v1 !== -1) data[rowIndex][colMap.pc_1v1] = pc_1v1;
      if (colMap.pc_2v2 !== -1) data[rowIndex][colMap.pc_2v2] = pc_2v2;
      if (colMap.pc_3v3 !== -1) data[rowIndex][colMap.pc_3v3] = pc_3v3;
      if (colMap.pc_4v4 !== -1) data[rowIndex][colMap.pc_4v4] = pc_4v4;
      if (colMap.console_1v1 !== -1) data[rowIndex][colMap.console_1v1] = console_1v1;
      if (colMap.last_api_sync !== -1) data[rowIndex][colMap.last_api_sync] = syncTime;
      
      update_count++;
    } catch (e) {
      Logger.log(`Failed to parse data for row index ${rowMapping[j]}`);
    }
  }

  // 4. Batch write all updates to the database in a single keystroke
  sheet.getDataRange().setValues(data);
  Logger.log(`Successfully synced Elo for ${update_count} players in batch.`);
}

/**
 * Run this once to automatically schedule nightly Elo syncs between 2:00 AM and 3:00 AM.
 */
function create_nightly_sync_trigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'sync_all_players_elo') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('sync_all_players_elo')
    .timeBased()
    .everyDays(1)
    .atHour(2) 
    .create();

  Logger.log("Nightly trigger created successfully for 2:00 AM - 3:00 AM.");
}