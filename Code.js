// ==========================================
// 🎛️ GLOBAL CAMPAIGN CONFIGURATION DASHBOARD
// ==========================================
var CONFIG = {
  SPREADSHEET_ID: "  THE_GOOGLE_SPREADSHEET_ID_HERE",
  GOAL_AMOUNT: 5000.00,
  
  // 🌐 Centralized Webhook Routes
  URL_SANDBOX_CHAT: "https://chat.googleapis.com/v1/spaces/YOUR_SANDBOX_SPACE_ID/messages?key=THE_KEY&token=THE_TOKEN",
  URL_MAIN_CHAT: "https://chat.googleapis.com/v1/spaces/YOUR_MAIN_SPACE_ID/messages?key=THE_KEY&token=THE_TOKEN"
};


// ==========================================
// 🧪 LIVE WEBHOOK SIMULATOR / TESTER
// ==========================================
function testMyWebhook() {
  var fakeEvent = {
    postData: {
      contents: JSON.stringify({
        data: {
          first_name: "Lisa",
          last_name: "Alexander",
          email: "lisa.alexander@example.com",
          amount: "200.00",
          team_member: "Atlanta"
        }
      })
    }
  };
  
  doPost(fakeEvent);
  Logger.log("Test execution completed! Check your spreadsheet and Google Chat Sandbox.");
}


// ==========================================
//  1. LIVE WEBHOOK TRAFFIC COP
// ==========================================
function doPost(e) {
  var json = JSON.parse(e.postData.contents);
  var data = json.data;
  
  if (json.event === "transaction.success" && (!data.amount || parseFloat(data.amount) === 0)) {
    return ContentService.createTextOutput(JSON.stringify({"status": "skipped"})).setMimeType(ContentService.MimeType.JSON);
  }
  
  var fullName = (data.first_name || "") + " " + (data.last_name || "");
  var email = data.email || "";
  var amount = (data.amount && !isNaN(data.amount)) ? parseFloat(data.amount) : 0.00;
  
  var teamName = data.team_member || data.team || "";
  if (!teamName || teamName.trim().toLowerCase() === "our team") {
    teamName = "General Campaign";
  }
  
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var currentDate = new Date(); 
  var cleanIncomingEmail = email.toString().toLowerCase().trim();
  
  if (amount > 0) {
    var donorSheet = ss.getSheetByName("Donors");
    var donorLastRow = donorSheet.getLastRow();
    if (donorLastRow > 1) {
      var donorEmails = donorSheet.getRange("C1:C" + donorLastRow).getValues().flat().map(function(eStr) {
        return eStr.toString().toLowerCase().trim();
      });
      if (donorEmails.indexOf(cleanIncomingEmail) !== -1) {
        return ContentService.createTextOutput(JSON.stringify({"status": "duplicate donation skipped"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    donorSheet.appendRow([currentDate, fullName, email, teamName, amount]);
    sendGoogleChatNotification("💰 **NEW DONATION!** " + fullName + " just supported " + teamName + " with $" + amount.toFixed(2) + "! 🔥");
    
  } else {
    var volunteerSheet = ss.getSheetByName("Volunteers");
    var volunteerLastRow = volunteerSheet.getLastRow();
    if (volunteerLastRow > 1) {
      var volunteerEmails = volunteerSheet.getRange("C1:C" + volunteerLastRow).getValues().flat().map(function(eStr) {
        return eStr.toString().toLowerCase().trim();
      });
      if (volunteerEmails.indexOf(cleanIncomingEmail) !== -1) {
        return ContentService.createTextOutput(JSON.stringify({"status": "duplicate volunteer skipped"})).setMimeType(ContentService.MimeType.JSON);
      }
    }
    
    volunteerSheet.appendRow([currentDate, fullName, email, teamName]);
    
    var reportSheet = ss.getSheetByName("Weekly Report");
    var liveVolunteerCount = reportSheet.getRange("G2").getValue();
    
    var volunteerMessage = "🙋🏽‍♂️ **NEW VOLUNTEER REGISTERED!** \n\n" +
                           "Huge thank you to " + fullName + " for signing up to serve during our **2026 Day Of Service**! 💛🔥\n\n" +
                           "🚀 This takes our community total to **" + liveVolunteerCount + "** superstar volunteers mobilized!";
                           
    sendGoogleChatNotification(volunteerMessage);
  }
  
  return ContentService.createTextOutput(JSON.stringify({"status": "success"})).setMimeType(ContentService.MimeType.JSON);
}


// ==========================================
//  2. AUTOMATED FRIDAY RECAP SCRIPT
// ==========================================
function sendWeeklyRecap() {
  var ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  var reportSheet = ss.getSheetByName("Weekly Report");
  
  var today = new Date();
  if (today.getDay() !== 5) return; 
  
  var totalRaised = reportSheet.getRange("B2").getValue();
  var weeklyDifference = reportSheet.getRange("E2").getValue();
  var totalVolunteers = reportSheet.getRange("G2").getValue();
  var progressPercentage = (totalRaised / CONFIG.GOAL_AMOUNT) * 100;
  
  var leaderboard = getTopTeamData(reportSheet);
  var winningTeam = leaderboard.teamName; 
  var winningAmount = leaderboard.amountRaised;
  var contributorCount = leaderboard.contributors; 

  var advancedData = getAdvancedFundraisingData(ss);
  
  var top5String = "";
  for (var i = 0; i < advancedData.top5Rows.length; i++) {
    top5String += (i + 1) + "️⃣ " + advancedData.top5Rows[i].name + " — $" + advancedData.top5Rows[i].amount.toFixed(2) + "\n";
  }
  if (top5String === "") { top5String = "No individual donations registered yet! 🌱\n"; }

  var wowSentence = "";
  if (weeklyDifference > 0) {
    wowSentence = "📈 **GROWTH MOMENTUM:** We pushed past last week's numbers by adding an incredible **$" + weeklyDifference.toFixed(2) + "** to our total this week alone! 🎉\n\n";
  } else if (weeklyDifference === 0 && totalRaised > 0) {
    wowSentence = "✊ **HOLDING STRONG:** We are maintaining our steady momentum from last week! Let's push for a big breakout weekend! 🚀\n\n";
  }

  var message = "Good evening SLS Family! 👋🏽💙\n\n" +
                "I hope y'all had an incredible week! 😄 With the holiday weekend behind us, let's keep that absolute fire burning hot for our campaign goals! 🔥\n\n" +
                "Let's take a look at this week's recap for our 2026 DOS Campaign:\n\n" +
                "                                 🚨 **SLS DAY OF SERVICE: WEEKLY SNAPSHOT** 🚨\n" +
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                "💰 TOTAL FUNDS RAISED: $" + totalRaised.toFixed(2) + " / $" + CONFIG.GOAL_AMOUNT.toFixed(2) + " 🚀\n" +
                "📈 PROGRESS TO GOAL: We are officially " + progressPercentage.toFixed(1) + "% of the way to our $5,000 finish line 🙌🏽!\n" +
                "🙋🏽‍♂️ VOLUNTEERS MOBILIZED: " + totalVolunteers + " leaders signed up to serve!\n\n" +
                wowSentence +
                "🏆 COHORT LEADERBOARD TOPPER:\n" +
                "👉 *" + winningTeam + " is officially leading the race with $" + winningAmount.toFixed(2) + " raised from " + contributorCount + " contributors!* 🔥⚔️\n\n" +
                "🎖️ COHORT MVP SHOUTOUTS:\n" +
                "🍑 **Atlanta Top Supporter:** " + advancedData.mvpAtlanta.name + " ($" + advancedData.mvpAtlanta.amount.toFixed(2) + ")\n" +
                "🏛️ **Athens Top Supporter:** " + advancedData.mvpAthens.name + " ($" + advancedData.mvpAthens.amount.toFixed(2) + ")\n\n" +
                "💎 INDIVIDUAL TOP 5 LEADERBOARD:\n" +
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n" +
                top5String + 
                "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n" +
                "I want to take a moment to acknowledge all the incredible hard work every single one of you has poured into this week.\nLet's keep spreading the word and pushing this momentum forward over the weekend.\n\n" +
                "Have a beautiful, restful weekend ahead! 💛💙";
                
  //  ROUTE OUTPUT: Sends directly to the Sandbox Chat room for layout validation
  sendGoogleChatNotification(message);

  // ==========================================
  //  ARCHIVE ROUTINE: TARGETS HISTORICAL SEPARATE TAB
  // ==========================================
  var historySheet = ss.getSheetByName("Snapshot History");
  var archiveRow = historySheet.getLastRow() + 1;
  if (archiveRow <= 1) { archiveRow = 2; }
  var weekLabel = "Week ending " + Utilities.formatDate(today, Session.getScriptTimeZone(), "MM/dd");
  
  historySheet.getRange(archiveRow, 1).setValue(weekLabel);
  historySheet.getRange(archiveRow, 2).setValue(totalRaised);
  historySheet.getRange(archiveRow, 3).setValue(totalVolunteers);
  historySheet.getRange(archiveRow, 4).setValue(today);
  
  reportSheet.getRange("D2").setValue(totalRaised);
}


// ==========================================
// HELPER FUNCTION: DYNAMIC PIVOT TABLE SCANNER
// ==========================================
function getTopTeamData(sheet) {
  var lastRow = sheet.getLastRow();
  var defaultResult = { teamName: "No Cohort Yet", amountRaised: 0.00, contributors: 0 };
  
  if (lastRow < 3) return defaultResult;
  
  var dataRange = sheet.getRange("I3:K" + lastRow).getValues();
  var maxAmount = -1;
  var topTeam = "";
  var topContributors = 0;
  
  for (var i = 0; i < dataRange.length; i++) {
    var team = dataRange[i][0] ? dataRange[i][0].toString().trim() : "";
    var amount = dataRange[i][1] && !isNaN(dataRange[i][1]) ? parseFloat(dataRange[i][1]) : 0;
    var count = dataRange[i][2] && !isNaN(dataRange[i][2]) ? parseInt(dataRange[i][2]) : 0;
    
    if (team === "" || team === "0" || team.toLowerCase().includes("grand total") || team.toLowerCase().includes("team / member")) {
      continue;
    }
    
    if (amount > maxAmount) {
      maxAmount = amount;
      topTeam = team;
      topContributors = count;
    }
  }
  
  if (topTeam !== "") {
    return { teamName: topTeam, amountRaised: maxAmount, contributors: topContributors };
  }
  return defaultResult;
}


// ==========================================
//  HELPER FUNCTION: RAW DATA AGGREGATION & MVP ANALYSIS
// ==========================================
function getAdvancedFundraisingData(ss) {
  var donorSheet = ss.getSheetByName("Donors");
  var lastRow = donorSheet.getLastRow();
  
  var result = {
    top5Rows: [],
    mvpAtlanta: { name: "No donations yet", amount: 0.00 },
    mvpAthens: { name: "No donations yet", amount: 0.00 }
  };
  
  if (lastRow < 2) return result;
  
  var rawData = donorSheet.getRange("A2:E" + lastRow).getValues();
  var individualMap = {};
  
  var maxAtlanta = 0;
  var maxAthens = 0;
  
  for (var i = 0; i < rawData.length; i++) {
    var name = rawData[i][1] ? rawData[i][1].toString().trim() : "";
    var team = rawData[i][3] ? rawData[i][3].toString().trim().toLowerCase() : "";
    var amount = rawData[i][4] && !isNaN(rawData[i][4]) ? parseFloat(rawData[i][4]) : 0.00;
    
    if (name === "" || amount <= 0) continue;
    
    if (!individualMap[name]) {
      individualMap[name] = 0.00;
    }
    individualMap[name] += amount;
    
    if (team === "atlanta" && amount > maxAtlanta) {
      maxAtlanta = amount;
      result.mvpAtlanta = { name: name, amount: amount };
    }
    if (team === "athens" && amount > maxAthens) {
      maxAthens = amount;
      result.mvpAthens = { name: name, amount: amount };
    }
  }
  
  var sortedList = [];
  for (var key in individualMap) {
    sortedList.push({ name: key, amount: individualMap[key] });
  }
  sortedList.sort(function(a, b) { return b.amount - a.amount; });
  
  result.top5Rows = sortedList.slice(0, 5);
  return result;
}


// ==========================================
// 3. GOOGLE CHAT SENDER (SANDBOX ALERTS)
// ==========================================
function sendGoogleChatNotification(text) {
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({"text": text})
  };
  UrlFetchApp.fetch(CONFIG.URL_SANDBOX_CHAT, options);
}


// ==========================================
// 4. DEDICATED MAIN CHANNEL SENDER (PRODUCTION RECAPS)
// ==========================================
function sendMainChannelNotification(text) {
  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify({"text": text})
  };
  UrlFetchApp.fetch(CONFIG.URL_MAIN_CHAT, options);
}
