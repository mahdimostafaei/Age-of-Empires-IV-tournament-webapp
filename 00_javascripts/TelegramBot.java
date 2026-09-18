// ==========================================
// 1. SETTINGS & BROADCAST
// ==========================================
var botToken = "8813949251:AAEIzAI4y8Ni2PTZAP0Lo86VG4qOaWPR8h8";
var telegramGroupId = "@AgeofEmpiresivPersian";

// ==========================================
// 2. WEBHOOK LISTENER (COMMANDS, VOTES, CONFIRMATIONS)
// ==========================================
function doPost(e) {
  // 1. RAW LOGGER: Placed before parsing to guarantee we catch the raw data even if JSON crashes
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var debugSheet = ss.getSheetByName("webhook_debug");
    if (debugSheet && e && e.postData) {
      debugSheet.appendRow([new Date(), e.postData.contents]);
    }
  } catch(logErr) {
    // Ignore logger errors so it doesn't break the webhook
  }

  // 2. MAIN LOGIC
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // --- A. HANDLE TEXT COMMANDS (/pin, /betrank, /dashboard) ---
    if (data.message && data.message.text) {
      var text = data.message.text.trim();
      var chatId = data.message.chat.id;
      var userId = String(data.message.from.id);
      
      if (text.indexOf("/betrank") === 0 || text.indexOf("/dashboard") === 0) {
        var webAppUrl = ScriptApp.getService().getUrl(); 
        var replyMessage = "🏆 *Tournament Dashboard*\nClick below to view the Live Leaderboard and Active Matches!";
        var keyboard = { inline_keyboard: [[{ text: "📊 Open Dashboard", url: webAppUrl }]] };
        
        UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendMessage", {
          method: "post",
          contentType: "application/json",
          payload: JSON.stringify({ chat_id: chatId, text: replyMessage, reply_markup: keyboard, parse_mode: "Markdown" }),
          muteHttpExceptions: true
        });
      }
      
      if (text === "/pin" || text.indexOf("/pin") === 0) {
        var playerSheet = ss.getSheetByName("players");
        if (playerSheet) {
          var playerData = playerSheet.getDataRange().getValues();
          var headers = playerData[0].map(function(h) { return String(h).toLowerCase().trim(); });
          
          var nameCol = headers.indexOf("player_name") > -1 ? headers.indexOf("player_name") : headers.indexOf("player");
          var tgCol = headers.findIndex(function(h) { return h.includes("telegram_chat_id") || h.includes("telegram_id") || h.includes("telegram_numeric"); });
          var pinCol = headers.findIndex(function(h) { return h.includes("web_pin") || h === "pin"; });
          
          var foundPin = null;
          var playerName = "Player";
          
          if (tgCol > -1 && pinCol > -1 && nameCol > -1) {
            for (var i = 1; i < playerData.length; i++) {
              var rawCell = playerData[i][tgCol];
              if (!rawCell) continue;
              
              var rowTgId = String(rawCell).trim().split(".")[0];
              if (rowTgId === userId) {
                playerName = String(playerData[i][nameCol] || "Player");
                foundPin = String(playerData[i][pinCol] || "").trim();
                break;
              }
            }
          }
          
          var privateMessage = foundPin ? 
            "🔐 Hello " + playerName + "!\n\nYour tournament dashboard PIN is: *" + foundPin + "*\n\nKeep it secret and use it to log into the web dashboard!" : 
            "❌ We couldn't find a PIN linked to your Telegram ID (`" + userId + "`).\n\nPlease make sure your ID is correctly saved in the database!";
          
          UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendMessage", {
            method: "post",
            contentType: "application/json",
            payload: JSON.stringify({ chat_id: userId, text: privateMessage, parse_mode: "Markdown" }),
            muteHttpExceptions: true
          });
        }
      }
    }
    
    // --- B. HANDLE POLL VOTES (MIMICKING PREVIOUS WORKING PROJECT) ---
    if (data.poll_answer) {
      var pollId = String(data.poll_answer.poll_id);
      var bettorId = String(data.poll_answer.user.id);
      var firstName = String(data.poll_answer.user.first_name || "Player");
      var username = String(data.poll_answer.user.username || "");
      
      // Ensure option_ids exists and isn't empty
      if (data.poll_answer.option_ids && data.poll_answer.option_ids.length > 0) {
        var chosenOption = data.poll_answer.option_ids[0]; 
        
        var betLogSheet = ss.getSheetByName("bet_logs");
        var pollSheet = ss.getSheetByName("poll_logs");
        
        if (pollSheet && betLogSheet) {
          var pollData = pollSheet.getDataRange().getValues();
          var isActive = false;
          for (var p = 1; p < pollData.length; p++) {
            if (String(pollData[p][0]) === pollId && String(pollData[p][6]) === "Active") {
              isActive = true; break;
            }
          }
          
          if (isActive) {
            updateLeaderboardUser(bettorId, firstName, username);
            betLogSheet.appendRow([pollId, bettorId, firstName, username, chosenOption, "", "", new Date()]);
          }
        }
      }
    }   
  } catch (error) {
    console.error("Webhook Error: " + error.message);
  }
  
  return HtmlService.createHtmlOutput("OK");
}

// ==========================================
// 3. READ THE SPREADSHEET (VIA PIN)
// ==========================================
function getMatchesByPin(pin) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var cleanPin = String(pin).trim();
  var adminPins = ["9999", "5215"];
  var isAdmin = adminPins.indexOf(cleanPin) !== -1;
  
  var playerIdentifier = null, playerName = null;
  
  if (isAdmin) {
    playerName = "Admin (God Mode)";
  } else {
    var playerSheet = ss.getSheetByName("players"); 
    if (!playerSheet) throw new Error("Could not find 'players' sheet.");
    
    var playerData = playerSheet.getDataRange().getValues();
    var headers = playerData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    
    var nameCol = headers.indexOf("player_name") > -1 ? headers.indexOf("player_name") : headers.indexOf("player");
    var pinCol = headers.findIndex(function(h) { return h.includes("web_pin") || h === "pin"; });
    
    for (var i = 1; i < playerData.length; i++) {
      var sheetPin = String(playerData[i][pinCol]).trim();
      if (sheetPin === cleanPin && sheetPin !== "") { 
        playerName = String(playerData[i][nameCol]); 
        playerIdentifier = String(playerData[i][nameCol]); 
        break;
      }
    }
    if (!playerIdentifier) throw new Error("Invalid PIN. Please check your number and try again.");
  }

  var matchSheet = ss.getSheetByName("matches");
  if (!matchSheet) throw new Error("Could not find 'matches' sheet.");
  
  var matchData = matchSheet.getDataRange().getValues();
  var matches = [];
  
  for (var j = 1; j < matchData.length; j++) {
    var row = matchData[j];
    var p1 = String(row[1]), p2 = String(row[2]), status = String(row[4]); 

    if (status === "Active" || status === "active") {
      if (isAdmin) {
        matches.push({ id: row[0], opponent: p1 + " & " + p2 });
      } else if (p1 === playerIdentifier || p2 === playerIdentifier) {
        var opponent = (p1 === playerIdentifier) ? p2 : p1;
        matches.push({ id: row[0], opponent: opponent });
      }
    }
  }
  return { name: playerName, matches: matches };
}

// ==========================================
// 4. SEND OUTBOUND MESSAGE & ANNOUNCE MATCH
// ==========================================
function announceMatch(matchId) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var matchSheet = ss.getSheetByName("matches");
  if (!matchSheet) throw new Error("Could not find 'matches' sheet.");
  
  var matchData = matchSheet.getDataRange().getValues();
  var headers = matchData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  
  var matchIdCol = headers.indexOf("match_id") > -1 ? headers.indexOf("match_id") : 0;
  var p1Col = headers.indexOf("player1_id") > -1 ? headers.indexOf("player1_id") : 1;
  var p2Col = headers.indexOf("player2_id") > -1 ? headers.indexOf("player2_id") : 2;
  var p1CivCol = headers.indexOf("proposed_p1_civ");
  var p2CivCol = headers.indexOf("proposed_p2_civ");
  var statusCol = headers.indexOf("match_status");
  var winnerCol = headers.indexOf("winner_id");
  
  var player1 = "Player 1", player2 = "Player 2", p1Civ = "Unknown", p2Civ = "Unknown", matchFound = false;
  
  for (var i = 1; i < matchData.length; i++) {
    if (String(matchData[i][matchIdCol]) === String(matchId)) {
      player1 = String(matchData[i][p1Col]); 
      player2 = String(matchData[i][p2Col]); 
      p1Civ = p1CivCol > -1 ? String(matchData[i][p1CivCol] || "Unknown") : "Unknown";
      p2Civ = p2CivCol > -1 ? String(matchData[i][p2CivCol] || "Unknown") : "Unknown";
      matchFound = true;
      break;
    }
  }
  
  if (!matchFound) throw new Error("Could not find Match ID " + matchId);

  // Anti-flood guard
  var pollSheet = ss.getSheetByName("poll_logs");
  if (pollSheet) {
    var existingPolls = pollSheet.getDataRange().getValues();
    for (var p = 1; p < existingPolls.length; p++) {
      if (String(existingPolls[p][1]) === String(matchId) && String(existingPolls[p][6]) === "Active") {
        return false; 
      }
    }
  }

  // Calculate Last 5 Form & Head-to-Head Record
  var p1History = [];
  var p2History = [];
  var h2hP1Wins = 0;
  var h2hP2Wins = 0;
  
  for (var m = 1; m < matchData.length; m++) {
    var mStatus = statusCol > -1 ? String(matchData[m][statusCol]).trim().toLowerCase() : "";
    var mWinner = winnerCol > -1 ? String(matchData[m][winnerCol]).trim() : "";
    var mp1 = String(matchData[m][p1Col]).trim();
    var mp2 = String(matchData[m][p2Col]).trim();
    
    if (mStatus === "settled" || mStatus === "archived" || mStatus === "completed") {
      if (mWinner !== "" && mWinner.toLowerCase() !== "unplayed") {
        if (mp1.toLowerCase() === player1.toLowerCase() || mp2.toLowerCase() === player1.toLowerCase()) {
          p1History.push(mWinner.toLowerCase() === player1.toLowerCase() ? "W" : "L");
        }
        if (mp1.toLowerCase() === player2.toLowerCase() || mp2.toLowerCase() === player2.toLowerCase()) {
          p2History.push(mWinner.toLowerCase() === player2.toLowerCase() ? "W" : "L");
        }
        if ((mp1.toLowerCase() === player1.toLowerCase() && mp2.toLowerCase() === player2.toLowerCase()) ||
            (mp1.toLowerCase() === player2.toLowerCase() && mp2.toLowerCase() === player1.toLowerCase())) {
          if (mWinner.toLowerCase() === player1.toLowerCase()) h2hP1Wins++;
          else if (mWinner.toLowerCase() === player2.toLowerCase()) h2hP2Wins++;
        }
      }
    }
  }
  
  var p1RecentForm = p1History.length > 0 ? p1History.slice(-5).join(" ") : "No Data";
  var p2RecentForm = p2History.length > 0 ? p2History.slice(-5).join(" ") : "No Data";

  // Get Elo for Win Probabilities
  var playerSheet = ss.getSheetByName("players"); 
  var p1Elo = 1200, p2Elo = 1200;
  if (playerSheet) {
    var playerData = playerSheet.getDataRange().getValues();
    var pHeaders = playerData[0].map(function(h) { return String(h).toLowerCase().trim(); });
    var nameCol = pHeaders.indexOf("player_name") > -1 ? pHeaders.indexOf("player_name") : pHeaders.indexOf("player");
    var eloCol = pHeaders.indexOf("tournament_elo");
    
    for (var j = 1; j < playerData.length; j++) {
      var pName = String(playerData[j][nameCol]).trim();
      if (pName === player1 && eloCol > -1) p1Elo = Number(playerData[j][eloCol]) || 1200;
      if (pName === player2 && eloCol > -1) p2Elo = Number(playerData[j][eloCol]) || 1200;
    }
  }
  
  var expectedP1 = 1 / (1 + Math.pow(10, (p2Elo - p1Elo) / 400));
  var expectedP2 = 1 / (1 + Math.pow(10, (p1Elo - p2Elo) / 400));
  var p1Percent = Math.round(expectedP1 * 100);
  var p2Percent = Math.round(expectedP2 * 100);

  // 1. Send the Text Message First
  var textMessage = "🔥 *MATCH IS STARTING!* 🔥\n\n" + 
                    "*" + player1 + "* vs *" + player2 + "*\n" +
                    "_" + p1Civ + "_ ⚔️ _" + p2Civ + "_\n\n" +
                    "📊 *Win Probability:*\n" + 
                    "• " + player1 + ": *" + p1Percent + "%* - Form: " + p1RecentForm + "\n" +
                    "• " + player2 + ": *" + p2Percent + "%* - Form: " + p2RecentForm + "\n\n" +
                    "*All-time Head-to-Head:*\n" +
                    player1 + " " + h2hP1Wins + " - " + h2hP2Wins + " " + player2;

  UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: telegramGroupId, text: textMessage, parse_mode: "Markdown" }),
    muteHttpExceptions: true
  });
  
  // 2. Send the Poll Second
  var pollPayload = { 
    chat_id: telegramGroupId, 
    question: "👇 Place your bets below! 👇", 
    options: [player1, player2],
    is_anonymous: false,
    open_period: 600
  };
  
  var pollResponse = UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/sendPoll", { 
    method: "post", 
    contentType: "application/json", 
    payload: JSON.stringify(pollPayload),
    muteHttpExceptions: true
  });
  
  var result = JSON.parse(pollResponse.getContentText());
  if (result.ok !== true) throw new Error("Telegram rejected it: " + result.description);
  
  // 3. Pin the Poll using the Poll's message_id
  var pollMessageId = result.result.message_id;
  UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/pinChatMessage", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({ chat_id: telegramGroupId, message_id: pollMessageId }),
    muteHttpExceptions: true
  });
  
  // 4. Log to database
  if (pollSheet) {
    pollSheet.appendRow([String(result.result.poll.id), matchId, player1, player2, expectedP1, expectedP2, "Active"]);
  }
  return true;
}

// ==========================================
// 5. SETTLE PENDING BETS & HANDLE REVERSALS
// ==========================================
function settlePendingBets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var pollSheet = ss.getSheetByName("poll_logs");
  var matchSheet = ss.getSheetByName("matches");
  var betLogSheet = ss.getSheetByName("bet_logs");
  
  // Safely find the leaderboard sheet regardless of its exact name
  var leaderboardSheet = ss.getSheetByName("bet_leaderboard") || ss.getSheetByName("bettors"); 
  
  if (!pollSheet || !matchSheet || !betLogSheet || !leaderboardSheet) return;
  
  var pollData = pollSheet.getDataRange().getValues();
  var matchData = matchSheet.getDataRange().getValues();
  
  // DYNAMIC HEADERS: Safely locate Match ID and Winner ID columns
  var mHeaders = matchData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  var mIdCol = mHeaders.indexOf("match_id");
  var mWinCol = mHeaders.indexOf("winner_id");
  
  var matchWinners = {};
  for (var m = 1; m < matchData.length; m++) {
    var idIdx = mIdCol > -1 ? mIdCol : 0;
    var winIdx = mWinCol > -1 ? mWinCol : 9;
    
    matchWinners[String(matchData[m][idIdx])] = String(matchData[m][winIdx] || "").trim(); 
  }
  
  var betData = betLogSheet.getDataRange().getValues();
  var lbData = leaderboardSheet.getDataRange().getValues();
  var lbHeaders = lbData[0].map(function(h) { return String(h).toLowerCase().trim(); });
  
  var userCol = lbHeaders.indexOf("telegram_id"); 
  var pointsCol = lbHeaders.indexOf("points");
  
  var userPoints = {};
  for (var l = 1; l < lbData.length; l++) {
    var uId = String(lbData[l][userCol]);
    userPoints[uId] = { row: l + 1, points: Number(lbData[l][pointsCol]) };
  }

  for (var p = 1; p < pollData.length; p++) {
    var pollId = String(pollData[p][0]);
    var matchId = String(pollData[p][1]);
    var status = String(pollData[p][6]); 
    var currentActualWinner = matchWinners[matchId] || "";
    
    // Reverse settled bets if a winner is wiped
    if (status === "Settled" && currentActualWinner === "") {
      for (var b = 1; b < betData.length; b++) {
        if (String(betData[b][0]) === pollId) {
          var bettorId = String(betData[b][1]);
          var potentialPayout = Number(betData[b][5]); 
          var previousOutcome = String(betData[b][6]); 
          
          if (userPoints[bettorId]) {
            if (previousOutcome === "Won") userPoints[bettorId].points -= (potentialPayout - 50);
            else if (previousOutcome === "Lost") userPoints[bettorId].points += 50;
          }
          betLogSheet.getRange(b + 1, 6).setValue("");
          betLogSheet.getRange(b + 1, 7).setValue("");
        }
      }
      pollSheet.getRange(p + 1, 7).setValue("Active");
      continue;
    }
    
    // Settle active bets when a winner is declared
    if (status === "Active" && currentActualWinner !== "") {
      var p1 = String(pollData[p][2]), p2 = String(pollData[p][3]);
      var expP1 = Number(pollData[p][4]), expP2 = Number(pollData[p][5]);
      var oddsP1 = expP1 > 0 ? (1 / expP1) : 2, oddsP2 = expP2 > 0 ? (1 / expP2) : 2;
      
      for (var b = 1; b < betData.length; b++) {
        if (String(betData[b][0]) === pollId && String(betData[b][6]) === "") {
          var bettorId = String(betData[b][1]);
          var chosenOption = Number(betData[b][4]); 
          var choiceName = (chosenOption === 0) ? p1 : p2;
          
          var won = (choiceName === currentActualWinner);
          var payout = 0, outcomeStr = "Lost";
          
          if (won) {
            payout = Math.round(50 * ((choiceName === p1) ? oddsP1 : oddsP2));
            outcomeStr = "Won";
            if (userPoints[bettorId]) userPoints[bettorId].points += (payout - 50);
          } else {
            if (userPoints[bettorId]) userPoints[bettorId].points -= 50;
          }
          
          betLogSheet.getRange(b + 1, 6).setValue(payout);
          betLogSheet.getRange(b + 1, 7).setValue(outcomeStr);
        }
      }
      pollSheet.getRange(p + 1, 7).setValue("Settled");
    }
  }
  
  for (var uId in userPoints) {
    leaderboardSheet.getRange(userPoints[uId].row, pointsCol + 1).setValue(userPoints[uId].points);
  }
}

// ==========================================
// 6. LEADERBOARD UTILITIES
// ==========================================
function updateLeaderboardUser(telegramId, firstName, username) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("bet_leaderboard");
  if (!sheet) return;
  
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (String(data[i][0]) === String(telegramId)) {
      sheet.getRange(i + 1, 2).setValue(firstName);
      sheet.getRange(i + 1, 3).setValue(username);
      return;
    }
  }
  sheet.appendRow([String(telegramId), firstName, username, 1000, new Date()]);
}

function getLeaderboardData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("bet_leaderboard");
  if (!sheet) return [];
  var data = sheet.getDataRange().getValues();
  var leaderboard = [];
  for (var i = 1; i < data.length; i++) {
    leaderboard.push({ firstName: String(data[i][1]), points: Number(data[i][3]) });
  }
  return leaderboard.sort((a, b) => b.points - a.points);
}

function fixWebhookUrl() {
  var botToken = "8813949251:AAEIzAI4y8Ni2PTZAP0Lo86VG4qOaWPR8h8";
  
  // PASTE YOUR /exec URL INSIDE THESE QUOTES:
  var liveUrl = "https://script.google.com/macros/s/AKfycbzIvh3JCWQ8Z8wPdXMaKR4Pbu7TFF6AIpAGy7PelBt0yidYKB4_ceTup_XN4Jy-GjO-8g/exec"; 
  
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/setWebhook", {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify({
      url: liveUrl,
      allowed_updates: ["message", "poll_answer", "callback_query"]
    }),
    muteHttpExceptions: true
  });
  console.log(response.getContentText());
}

function checkWebhookStatus() {
  var botToken = "8813949251:AAEIzAI4y8Ni2PTZAP0Lo86VG4qOaWPR8h8";
  var response = UrlFetchApp.fetch("https://api.telegram.org/bot" + botToken + "/getWebhookInfo", { muteHttpExceptions: true });
  console.log(response.getContentText());
}