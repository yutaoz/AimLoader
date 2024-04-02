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

// takes username, gets aimlabs user id
function getUserID(username) {
  var url = "https://api.aimlab.gg/graphql";

  var headers = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "accept-language": "en-US,en;q=0.6",
    "origin": "https://app.voltaic.gg",
    "sec-fetch-site": "cross-site",
    "sec-fetch-mode": "cors",
    "sec-fetch-dest": "empty",
    "referer": "https://app.voltaic.gg/"
  };

  var payload = JSON.stringify({
    "query": "\n  query GetProfile($username: String) {\n    aimlabProfile(username: $username) {\n      username\n      user {\n        id\n      }\n      ranking {\n        rank {\n          displayName\n        }\n        skill\n      }\n    }\n  }\n",
    "variables": {
      "username": username
    }
  });

  var options = {
    "method": "POST",
    "headers": headers,
    "payload": payload,
    "muteHttpExceptions": true
  };

  var response = UrlFetchApp.fetch(url, options);
  
  var result = response.getContentText();
  if (JSON.parse(result).data.aimlabProfile === null) {
    SpreadsheetApp.getUi().alert('Error', 'Invalid Username, please setup again.', SpreadsheetApp.getUi().ButtonSet.OK);
    return null;
  }
  var userid = JSON.parse(result).data.aimlabProfile.user.id;
  Logger.log(userid);
  
  return userid; 

}

// takes userid, gets a list of all scens played by user
function getScens(userId) {
  var url = "https://api.aimlab.gg/graphql";

  var headers = {
    "accept": "application/json, text/plain, */*",
    "content-type": "application/json",
    "accept-language": "en-US,en;q=0.6",
    "origin": "https://app.voltaic.gg",
    "sec-fetch-site": "cross-site",
    "sec-fetch-mode": "cors",
    "sec-fetch-dest": "empty",
    "referer": "https://app.voltaic.gg/"
  };

  var payload = JSON.stringify({
    "query": "\n  query GetAimlabProfileAgg($where: AimlabPlayWhere!) {\n    aimlab {\n      plays_agg(where: $where) {\n        group_by {\n          task_id\n          task_name\n        }\n        aggregate {\n          count\n          avg {\n            score\n            accuracy\n          }\n          max {\n            score\n            accuracy\n            created_at\n          }\n        }\n      }\n    }\n  }\n",
    "variables": {
      "where": {
        "is_practice": {
          "_eq": false
        },
        "score": {
          "_gt": 0
        },
        "user_id": {
          "_eq": userId
        }
      }
    }
  });

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
  + "and press Setup Loader";
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
  var prompt = SpreadsheetApp.getUi().prompt('Setup Username', "Enter your Aimlabs username", SpreadsheetApp.getUi().ButtonSet.OK);
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
    Logger.log('No username entered.');
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
    

  var cells = range.getValues();
  console.log("B: " + cells);
  var startRow = range.getRow();
  var startColumn = range.getColumn();

  var scens = [];
  var scenMap = new Map();

  // add benchmark scens to list, create a map where key = scen name, value = position of scen name in an [x, y] array
  for (var i = 0; i < cells.length; i++) {
    for (var j = 0; j < cells[i].length; j++) {
      scens.push(cells[i][j].trim().toLowerCase());
      scenMap.set(cells[i][j].trim().toLowerCase(), [startRow + i, startColumn + j]);
    }
  }

  // get user info and selected range
  let username = spreadsheetProperties.getProperty('username');
  let userid = getUserID(username);
  let userData = getScens(userid);
  let userScens = userData.data.aimlab.plays_agg;

  // go through all played scens, update sheet with high scores
  for (var i = 0; i < userScens.length; i++) {
    var uscen = userScens[i].group_by.task_name.trim().toLowerCase(); // scen name
    if (scens.includes(uscen)) { // if scen matches a benchmark scen...
      var maxScore = userScens[i].aggregate.max.score;
      var cellCoords = scenMap.get(uscen);
      // adjust coords to fill-in box
      var row = cellCoords[0];
      var col = cellCoords[1] + 3;
      // write high score to coords
      sheet.getRange(row, col).setValue(maxScore);
      
    }
  }

}

