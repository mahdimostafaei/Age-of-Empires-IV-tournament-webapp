function process_tournament_registration(payload) {
  try {
    // 0. Master Switch Check
    if (PropertiesService.getScriptProperties().getProperty("REG_OPEN") !== "true") {
      return JSON.stringify({ success: false, error: "Registration is currently closed by the administrator." });
    }
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let pSheet = ss.getSheetByName("players");
    if (!pSheet) return JSON.stringify({ success: false, error: "Players sheet not found." });

    let rawName = payload.name.trim();
    let rawTele = payload.telegram.trim();
    let rawAoeId = payload.aoeId.trim();

    // 1. Sanitize Telegram to standard https://t.me/username
    let cleanTele = rawTele;
    if (cleanTele) {
      cleanTele = cleanTele.replace(/https?:\/\/(www\.)?t\.me\//i, '').replace('@', '').trim();
      if (cleanTele.includes('/')) cleanTele = cleanTele.split('/')[0];
      cleanTele = "https://t.me/" + cleanTele;
    }

    // 2. Sanitize AoE4 ID
    let cleanAoeId = rawAoeId.replace(/\D/g, ''); 

    // 2.5 API Micro-Fetch for Instant Elo
    let pc_1v1 = "", pc_2v2 = "", pc_3v3 = "", pc_4v4 = "", console_1v1 = "";
    if (cleanAoeId) {
      try {
        let res = UrlFetchApp.fetch(`https://aoe4world.com/api/v0/players/${cleanAoeId}`, {muteHttpExceptions: true});
        if (res.getResponseCode() === 200) {
          let api_data = JSON.parse(res.getContentText());
          let modes = api_data.modes || {};
          pc_1v1 = modes.rm_1v1_elo?.rating || modes.rm_solo_elo?.rating || modes.qm_1v1?.rating || "";
          pc_2v2 = modes.rm_2v2_elo?.rating || modes.qm_2v2?.rating || "";
          pc_3v3 = modes.rm_3v3_elo?.rating || modes.qm_3v3?.rating || "";
          pc_4v4 = modes.rm_4v4_elo?.rating || modes.qm_4v4?.rating || "";
          console_1v1 = modes.rm_1v1_console_elo?.rating || modes.rm_solo_console_elo?.rating || modes.qm_1v1_console?.rating || "";
        }
      } catch(e) {
        // Fail silently so registration still succeeds even if API is down
      }
    }

    const pData = pSheet.getDataRange().getValues();
    const headers = pData[0].map(h => String(h).toLowerCase().trim());
    
    const aoeCol = headers.indexOf("aoe4_world_id");
    const nameCol = headers.indexOf("player_name");
    const teleCol = headers.indexOf("telegram_handle");
    const pinCol = headers.indexOf("web_pin");
    const statusCol = headers.indexOf("status");
    const eloCol = headers.indexOf("tournament_elo");
    
    // Elo Columns
    const c1v1 = headers.indexOf("pc_1v1");
    const c2v2 = headers.indexOf("pc_2v2");
    const c3v3 = headers.indexOf("pc_3v3");
    const c4v4 = headers.indexOf("pc_4v4");
    const cCon = headers.indexOf("console_1v1");

    let found = false;
    let playerPin = "";
    let isNew = false;
    let finalName = rawName;

    const normalize = str => String(str).toLowerCase().replace(/\s+/g, '');
    const targetName = normalize(rawName);

    // 3. Check for existing player
    for (let i = 1; i < pData.length; i++) {
      let rowName = String(pData[i][nameCol]);
      let rowTele = String(pData[i][teleCol] || "").trim();
      let rowAoeId = String(pData[i][aoeCol] || "").trim();

      // EDGE CASE 1: Telegram matches, but Name DOES NOT
      if (cleanTele !== "" && rowTele.toLowerCase() === cleanTele.toLowerCase() && normalize(rowName) !== targetName) {
        return JSON.stringify({ 
          success: false, 
          error: `This Telegram is already linked to the player '${rowName}'. Please register using that exact name, or contact an admin to update your profile.` 
        });
      }

      // EDGE CASE 2: AoE4 World ID matches, but Name DOES NOT
      if (cleanAoeId !== "" && rowAoeId === cleanAoeId && normalize(rowName) !== targetName) {
        return JSON.stringify({ 
          success: false, 
          error: `The AoE4 World ID '${cleanAoeId}' is already linked to the player '${rowName}'. You cannot register duplicate in-game IDs.` 
        });
      }

      // STANDARD CASE: Name matches (Returning Player)
      if (normalize(rowName) === targetName) {
        found = true;
        finalName = rowName;
        playerPin = pData[i][pinCol];
        
        if (statusCol > -1) pSheet.getRange(i + 1, statusCol + 1).setValue("Active");
        if (teleCol > -1 && !rowTele && cleanTele) pSheet.getRange(i + 1, teleCol + 1).setValue(cleanTele);
        if (aoeCol > -1 && !pData[i][aoeCol] && cleanAoeId) pSheet.getRange(i + 1, aoeCol + 1).setValue(cleanAoeId);
        
        // Patch Elo instantly if fetched
        if (c1v1 > -1 && pc_1v1) pSheet.getRange(i + 1, c1v1 + 1).setValue(pc_1v1);
        if (c2v2 > -1 && pc_2v2) pSheet.getRange(i + 1, c2v2 + 1).setValue(pc_2v2);
        if (c3v3 > -1 && pc_3v3) pSheet.getRange(i + 1, c3v3 + 1).setValue(pc_3v3);
        if (c4v4 > -1 && pc_4v4) pSheet.getRange(i + 1, c4v4 + 1).setValue(pc_4v4);
        if (cCon > -1 && console_1v1) pSheet.getRange(i + 1, cCon + 1).setValue(console_1v1);
        if (eloCol > -1 && pc_1v1) pSheet.getRange(i + 1, eloCol + 1).setValue(pc_1v1);
        break;
      }
    }

    // 4. Create New Player
    if (!found) {
      isNew = true;
      playerPin = Math.floor(1000 + Math.random() * 9000).toString(); 
      
      let newRow = new Array(headers.length).fill("");
      if (aoeCol > -1) newRow[aoeCol] = cleanAoeId;
      if (nameCol > -1) newRow[nameCol] = rawName;
      if (teleCol > -1) newRow[teleCol] = cleanTele;
      if (pinCol > -1) newRow[pinCol] = playerPin;
      if (statusCol > -1) newRow[statusCol] = "Active";
      
      // Assign fetched Elo or default
      if (eloCol > -1) newRow[eloCol] = pc_1v1 || 1000; 
      if (c1v1 > -1) newRow[c1v1] = pc_1v1;
      if (c2v2 > -1) newRow[c2v2] = pc_2v2;
      if (c3v3 > -1) newRow[c3v3] = pc_3v3;
      if (c4v4 > -1) newRow[c4v4] = pc_4v4;
      if (cCon > -1) newRow[cCon] = console_1v1;
      
      pSheet.appendRow(newRow);
    }

    // 5. Log to the Registration Sheet
    let regSheet = ss.getSheetByName("Registration");
    if (!regSheet) {
      regSheet = ss.insertSheet("Registration");
      regSheet.appendRow(["Timestamp", "Player Name", "AoE4 ID", "Telegram", "Type"]);
      regSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#1e5b43").setFontColor("white");
    }
    regSheet.appendRow([new Date(), finalName, cleanAoeId, cleanTele, isNew ? "New Account" : "Returning"]);

    return JSON.stringify({ success: true, pin: playerPin, name: finalName, isNew: isNew });

  } catch (err) {
    return JSON.stringify({ success: false, error: err.toString() });
  }
}

// --- GLOBAL REGISTRATION TOGGLE ---
function get_registration_status() {
  const props = PropertiesService.getScriptProperties();
  return props.getProperty("REG_OPEN") === "true";
}

function toggle_registration_status(isOpen) {
  // 1. Flip the Master Switch
  PropertiesService.getScriptProperties().setProperty("REG_OPEN", isOpen ? "true" : "false");
  
  // 2. Broadcast and Pin to Telegram
  try {
    var botToken = "8813949251:AAEIzAI4y8Ni2PTZAP0Lo86VG4qOaWPR8h8";
    var telegramGroupId = "@AgeofEmpiresivPersian"; 
    var webAppUrl = ScriptApp.getService().getUrl(); 
    var textMessage = "";
    
    // Clean, standard text to prevent API crashes
    if (isOpen) {
      textMessage = "🟢 *TOURNAMENT REGISTRATION OPEN!* 🟢\n\n" +
                    "The registration phase for the upcoming tournament has officially started!\n\n" +
                    "👉 [Click here to register](" + webAppUrl + ")";
    } else {
      textMessage = "🔴 *REGISTRATION CLOSED* 🔴\n\n" +
                    "The registration is finished. Thank you to everyone who signed up!\n\n" +
                    "The group stage and layout will be given to you soon. Stay tuned.";
    }
    
    var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendMessage", {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify({ 
        chat_id: telegramGroupId, 
        text: textMessage, 
        parse_mode: "Markdown",
        disable_web_page_preview: true 
      }),
      muteHttpExceptions: true
    });
    
    var result = JSON.parse(response.getContentText());
    
    // Pin the announcement
    if (result.ok) {
      var messageIdToPin = result.result.message_id;
      UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/pinChatMessage", {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify({ chat_id: telegramGroupId, message_id: messageIdToPin })
      });
    }
  } catch (e) {
    // Fail silently so the script finishes and unlocks the button even if Telegram lags
  }

  return isOpen;
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