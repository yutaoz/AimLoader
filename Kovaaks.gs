function onOpen(e) {
  createUI();
  var spreadsheetProperties = PropertiesService.getDocumentProperties();
  var setupCompleted = spreadsheetProperties.getProperty('AimSetupCompleted');
  if (!setupCompleted) { // prompt setup if it has never been done
    displaySetupInstructions();
  }
}

function createTriggers() {
  ScriptApp.newTrigger('writeScores')
    .timeBased()
    .everyMinutes(1)
    .create();
}

// takes scenid, gets kovaaks score
function fetchLeaderboardScores(scen_id, steamid) {
  var url = "https://kovaaks.com/sa_leaderboard_scores_steam_ids_get";
  
  var headers = {
    "Accept": "*/*",
    "User-Agent": "X-UnrealEngine-Agent",
    "Content-Type": "application/json",
    "Authorization": "Bearer 140000007407386b2c618e1e3eb4425b01001001a9a90b66180000000100000002000000f15b782a2e4453214f67070001000000b200000032000000040000003eb4425b01001001ce930c00c30e5fae0138a8c000000000aba90b662b5927660100b62e080000000000c9f0e8dc401c8594ca19563a6a6989c4d0d865d8538663ff4329600d3bcbc221ef3223dd25966fb3ccb71fd2a3ee94a8331ed3373c77b3fe4eab5aca10f564e4f9b9fa9e2808581ec851b9966dc30f204f6a43865979ab316c097e27109dcd87b15d72acf74e61511abbb8973c414dba51331987af9815d626367076bfd5f574",
    "GSTVersion": "3.4.2.2024-02-28-14-22-08-791139f13a"
  };
  
  var payload = JSON.stringify({
    "leaderboard_id": scen_id,
    "steam_id": steamid,
    "steam_ids": [
      steamid
    ]
  });
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true
  };
  
  var response = UrlFetchApp.fetch(url, options);
  var result = response.getContentText();
  var data = JSON.parse(result);
  //Logger.log(data[0].score);
  return data[0] ? data[0].score : null;
}


// takes scen name, gets scen id from leaderboard
function getLeaderboardId(scen_name) {
  var url = "https://kovaaks.com/sa_leaderboard_put";
  
  var headers = {
    "Accept": "*/*",
    "User-Agent": "X-UnrealEngine-Agent",
    "Content-Type": "application/json",
    "GSTVersion": "3.4.2.2024-02-28-14-22-08-791139f13a"
  };
  
  var payload = JSON.stringify({
    "scenario_name": scen_name,
    "value_type": "time"
  });
  
  var options = {
    "method": "post",
    "contentType": "application/json",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true
  };
  
  // try {
  //   var response = UrlFetchApp.fetch(url, options);
  //   Logger.log(response.getContentText()); 
  // } catch (error) {
  //   Logger.log("LBID: " + error.toString());
  // }

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(url, options);
  var result = response.getContentText();
  var data = JSON.parse(result);
  return data;
}

// create ui menu
function createUI() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Aim Loader')
    .addItem('Setup Instructions', 'displaySetupInstructions')
    .addItem('Setup', 'setupLoader')
    .addItem('Reload', 'writeScores')
    .addToUi();
}

function displaySetupInstructions() {
  var message = "To setup, select the range of cells with the scenario names you want updated, go to the Aim Loader menu at the top, "
  + "and press Setup";
  SpreadsheetApp.getUi().alert('Aim Loader Setup', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

// do setup, records aimlab username, which sheet to update, and scen range to update
function setupLoader() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var activeRange = spreadsheet.getActiveRange();
  var activeSheet = spreadsheet.getActiveSheet();

  var selectedRange = activeRange.getA1Notation();
  var selectedSheetName = activeSheet.getName();

  var message = "Selected Range: " + selectedRange + "\n" +
                "Selected Sheet: " + selectedSheetName;

  // prompt username
  SpreadsheetApp.getUi().alert('Selected Range and Sheet', message, SpreadsheetApp.getUi().ButtonSet.OK);
  var prompt = SpreadsheetApp.getUi().prompt('Setup SteamID', "Enter your SteamID64: ", SpreadsheetApp.getUi().ButtonSet.OK);
  var response = prompt.getResponseText();
  var spreadsheetProperties = PropertiesService.getDocumentProperties();
  
  if (response !== '' && prompt.getSelectedButton() === SpreadsheetApp.getUi().Button.OK) { // store properties and complete setup
    Logger.log('Username entered: ' + response);
    spreadsheetProperties.setProperty('username', response);
    

    // make sure to store range
    spreadsheetProperties.setProperty('range', selectedRange);

    // store that setup has been completed so we dont realert each load
    spreadsheetProperties.setProperty('AimSetupCompleted', true);

    // store which sheet we are updating
    var spreadsheetId = spreadsheet.getId();
    spreadsheetProperties.setProperty('sheetid', spreadsheetId);
    spreadsheetProperties.setProperty('sheetname', selectedSheetName);
    createTriggers();
    SpreadsheetApp.getUi().alert("Setup Complete", SpreadsheetApp.getUi().ButtonSet.OK);
    writeScores(); // write score after setup
  } else {
    Logger.log('No SteamID entered.');
    SpreadsheetApp.getUi().alert("response", SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

// writes scores to sheet
function writeScores() {
  var spreadsheetProperties = PropertiesService.getDocumentProperties();
  var sheetid = spreadsheetProperties.getProperty('sheetid');
  var spreadsheet = SpreadsheetApp.openById(sheetid);
  var sheetname = spreadsheetProperties.getProperty('sheetname');
  var sheet = spreadsheet.getSheetByName(sheetname);
  
  var rangeProp = spreadsheetProperties.getProperty('range');
  var range = sheet.getRange(rangeProp);
  
  var steam_id = spreadsheetProperties.getProperty('username');

  var cells = range.getValues();
  console.log("B: " + cells);
  var startRow = range.getRow();
  var startColumn = range.getColumn();

  var scens = [];

  // add benchmark scens to list, create a map where key = scen name, value = position of scen name in an [x, y] array
  for (var i = 0; i < cells.length; i++) {
    for (var j = 0; j < cells[i].length; j++) {
      scens.push(cells[i][j].trim());

      var scen_id = getLeaderboardId(scens[i]).leaderboard_id;
      //Logger.log(scen_id); 
      var userScore = fetchLeaderboardScores(scen_id, steam_id) / 100;
      var cellCoords = [startRow + i, startColumn + j];
      var row = cellCoords[0];
      var col = cellCoords[1] + 2;
      sheet.getRange(row, col).setValue(userScore);
    }
  }
}
