/*
- Make "Initialization" function. Runs only if a property is set to some boolean. Triggered by edit. Removes its own trigger once done
  -Turns out, copying files doesn't preserve properties. Current Solution: star as first character in title
*/

// REQUIRED TRIGGERS: autoIdCheck; sendHoursViaButton
// NOTE: Ghost proerties ARE NOT NORMALLY DEFINED. Define if needed.

function initialize() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName('Activation Page').getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  
  defaultIdState();
  defaultAttendanceState();
  eventSwitch();
  setSheetId();
  
  ScriptApp.newTrigger('autoIdCheck')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
  
  ScriptApp.newTrigger('sendHoursViaButton')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
  
  ScriptApp.newTrigger('eventSwitch')
  .forSpreadsheet(ss)
  .onEdit()
  .create();
  
  ss.deleteSheet(s);
}

function buttonInitialize() { // RESERVED FUNCTION NAME! Insert any function which must run on boot.
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var file;
  var filename;
  
  file = DriveApp.getFileById(SpreadsheetApp.getActiveSpreadsheet().getId());
  filename = file.getName();
  
  if (filename.slice(0,1) == "*") {
  
  initialize();
  
  file.setName(filename.slice(1, filename.length));
  
  SpreadsheetApp.getUi().alert("Sheet Initialized!\nAutomatic functionality may take time to load.\n(Google servers can be slow sometimes.)");
  }
}

function setSheetId() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.setProperty("sheetId", ss.getId());
}


// Options boxes functionality

function activateParticipationMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.setProperty('eventType', "Participation");
  s.getRange(2,9).setValue("P Hours");
  s.getRange(9,10).setValue("Send P Hours:");
}

function activateVolunteeringMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.setProperty('eventType', "Volunteering");
  s.getRange(2,9).setValue("V Hours");
  s.getRange(9,10).setValue("Send V Hours:");
}

function activateParticipationMode() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var scriptProperties = PropertiesService.getScriptProperties();
  
  scriptProperties.setProperty('eventType', "Participation");
  s.getRange(2,9).setValue("P Hours");
  s.getRange(9,10).setValue("Send P Hours:");
}

function eventSwitch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var eventType;
  
  
  eventType = s.getRange(3,11).getValue();
  switch (eventType) {
  case "Participation":
    activateParticipationMode();
    break;
  case "General Meeting":
    
    break;
  default:
    activateVolunteeringMode();
    break;
  }

}


function trackingSwitch() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var trackingType;
  
  
  trackingType = s.getRange(5,11).getValue();
  switch (trackingType) {
  case "":
    
    break;
  case "":
    
    break;
  default:
    
    break;
  
  }

}



function te() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  
  Logger.log(s.getRange(3,9).getValue());
}


function lookupStudent(email) {
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto")); // ADMIN SHEET ID
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var endRow = s.getLastRow();
  var range;
  
  if (email == "@bergenfield.org") {
    return '';
  }
  for (var i = 2; i <= endRow; i++) {
    if (s.getRange(i, 2).getValue() == email) {
      return s.getRange(i, 3).getValue();
    }
  }
  return "Student Not Found";
}



function ttttttttt() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range = s.getRange(3,3,endRow - 2,4);
  var currentAttendanceState = range.getValues();
  var scriptProperties = PropertiesService.getScriptProperties();
  var properties;
  var isRowFinishedList;
  var isRowEditedList;
  Logger.log(true ^ true);
  Logger.log(true ^ false);
  Logger.log(true && true);
  Logger.log(true && (true ^ false))
  if (true && (true ^ false)) {
    Logger.log("Numbers Work!")
  }
  Logger.log(1 && 0);
  Logger.log(endRow);
  
}

function idCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = 62;
  var row;
  var column = 2;
  var scriptProperties = PropertiesService.getScriptProperties();
  var currentList = [];
  var lastList = scriptProperties.getProperty('lastIdState').split(',');
  var exit = true;
  var increment = 0;
  
      for (var i = 3; i <= endRow; i++) {
        currentList.push(s.getRange(i,2).getValue());
      }
      do {
        if (currentList[increment] != lastList[increment]) {
          row = increment + 3;
          s.getRange(row,1).setValue(lookupStudent(s.getRange(row,column).getValue() + "@bergenfield.org"));
        }
        if (increment == 59){
          exit = false;
        }
        increment ++;
      } while (exit);
  scriptProperties.setProperty('lastIdState', currentList.toString());
}

function autoIdCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  
  if (s.getSheetName() == "Attendance") {
    idCheck();
  }
}

function defaultIdState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = 62;
  var scriptProperties = PropertiesService.getScriptProperties();
  var currentState = '';
  var list;
  
  for (var i = 3; i <= endRow; i++) {
    currentState += '-';
    if (i == endRow) {
      continue;
    }
    currentState += ',';
  }
  scriptProperties.setProperty('lastIdState', currentState);
  return;
}

function defaultAttendanceState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = 62;
  var scriptProperties = PropertiesService.getScriptProperties();
  var currentState = '';
  var list;
  
  for (var i = 0; i < (endRow - 2) * 4; i++) {
    currentState += 'false';
    if (i == (endRow - 2) * 4 - 1) {
      continue;
    }
    currentState += ',';
  }
  scriptProperties.setProperty('lastAttendanceState', currentState);
  return;
}

function setAttendanceState() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = 62;
  var scriptProperties = PropertiesService.getScriptProperties();
  var currentState = '';
  var list;
  
  for (var i = 0; i < (endRow - 2) * 4; i++) {
    currentState += s.getRange((Math.floor((i)/4) + 3), (i % 4 + 3)).getValue();
    if (i == (endRow - 2) * 4 - 1) {
      continue;
    }
    currentState += ',';
  }
  scriptProperties.setProperty('lastAttendanceState', currentState);
  return;
}


// A "Ghost Check" is a check box that was edited very quickly after another; these are NOT processed
// editedRange is the FIRST EDITED RANGE; not the ghost check
function fixGhostChecks(row, column) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = 62;
  var range = s.getRange(3,3,endRow - 2,4);
  var currentAttendanceState = range.getValues();
  var lastAttendanceState;
  var scriptProperties = PropertiesService.getScriptProperties();
  var properties;
  var originalCheck = ((row - 3) * 4) + (column - 2);
  /*
  if (scriptProperties.getProperty('ghostCheckIsFinished' == true)) {
    scriptProperties.setProperty('ghostCheckRow', row);
    scriptProperties.setProperty('ghostCheckColumn', column);
    scriptProperties.setProperty('ghostCheckIsFinished', false);
  }
  row = scriptProperties.getProperty('ghostCheckRow');
  column = scriptProperties.getProperty('ghostCheckColumn');
  */
  
  
  if (currentAttendanceState.toString() != scriptProperties.getProperty('lastAttendanceState')) {
    lastAttendanceState = scriptProperties.getProperty('lastAttendanceState').split(",");
    currentAttendanceState = currentAttendanceState.toString().split(",");
    for (var i = 0; i < (endRow - 2) * 4; i++) {
      if (i == originalCheck - 1) {
        continue;
      }
      s.getRange((Math.floor((i)/4) + 3), (i % 4 + 3)).setValue(lastAttendanceState[i]);
    }
  }
  range = s.getRange(3,3,endRow - 2,4);
  currentAttendanceState = range.getValues();
  scriptProperties.setProperty('lastAttendanceState', currentAttendanceState.toString());
}

function onOpen() {

}

function onEdit(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var editedRange = e.range;
  var editedRangeValues = editedRange.getValues();
  var row;
  var column;
  var PSbox;
  var PEbox;
  var Ibox;
  var Obox;
  var PSboxValue;
  var PEboxValue;
  var IboxValue;
  var OboxValue;
  
  var calculateBox;
  var sendBox;
  var calculateBoxValue;
  var sendBoxValue;
  // Detect page
  if (s.getSheetName() == "Attendance") {
    // Detect if edited range is a single cell or not. If yes, continue.
    if (editedRangeValues.length == 1 && editedRangeValues[0].length == 1) {
      row = editedRange.getRow();
      column = editedRange.getColumn();
      
      
      // ID Edit
      // NOT POSSIBLE WITH onEdit(e)
      // ID Edit END
      
      // Checkbox edited
      if (row > 2 && (column >= 3 && column <= 6)) {
        PSbox = s.getRange(row,3);
        Ibox = s.getRange(row,4);
        Obox = s.getRange(row,5);
        PEbox = s.getRange(row,6);
        PSboxValue = PSbox.getValue();
        PEboxValue = PEbox.getValue();
        IboxValue = Ibox.getValue();
        OboxValue = Obox.getValue();
        
        switch(column) {
          case 3:
            if (PSboxValue == true && PEboxValue == false && IboxValue == false && OboxValue == false && s.getRange(row,8).getValue() == '') {
              s.getRange(row,7).setValue(Utilities.formatDate(new Date(s.getRange(1,7).getValue()),Session.getScriptTimeZone(),"HH:mm"));
              editedRange.setBackgroundRGB(0,255,0);
            } else if (PSboxValue == false && PEboxValue == false && IboxValue == false && OboxValue == false && s.getRange(row,8).getValue() == '') {
              s.getRange(row,7).clearContent();
              Ibox.setValue(false);
              editedRange.clearFormat();
            } /*else
            // Error Detection
            
            {
              SpreadsheetApp.getUi().alert(
              "INVALID INPUT!"+
              "\nOnly one of each check-in type and one of each check-out type may be selected."+
              "\nA check-in type must be selected before a check-out type."
              );
            }
            */
            break;
          case 4:
            if (PSboxValue == false && PEboxValue == false && IboxValue == true && OboxValue == false && s.getRange(row,8).getValue() == '') {
              s.getRange(row,7).setValue(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"HH:mm"));
              editedRange.setBackgroundRGB(0,255,0);
            } else if (PSboxValue == false && PEboxValue == false && IboxValue == false && OboxValue == false) {
              s.getRange(row,7).clearContent();
              editedRange.clearFormat();
            } /*else
            
            {
              SpreadsheetApp.getUi().alert(
              "INVALID INPUT!"+
              "\nOnly one of each check-in type and one of each check-out type may be selected."+
              "\nA check-in type must be selected before a check-out type."
              );
            }
            */
            break;
          case 5:
            if (PEboxValue == false && (PSboxValue == true || IboxValue == true) && OboxValue == true && s.getRange(row,7).getValue() != '') {
              s.getRange(row,8).setValue(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"HH:mm"));
              editedRange.setBackgroundRGB(255,0,0);
            } else if (PEboxValue == false && (PSboxValue == true || IboxValue == true) && OboxValue == false) {
              s.getRange(row,8).clearContent();
              editedRange.clearFormat();
            } /*else
            
            {
              SpreadsheetApp.getUi().alert(
              "INVALID INPUT!"+
              "\nOnly one of each check-in type and one of each check-out type may be selected."+
              "\nA check-in type must be selected before a check-out type."
              );
            }
            */
            break;
          case 6:
            if (PEboxValue == true && (PSboxValue ^ IboxValue) && OboxValue == false && s.getRange(row,7).getValue() != '') {
              s.getRange(row,8).setValue(Utilities.formatDate(new Date(s.getRange(1,8).getValue()),Session.getScriptTimeZone(),"HH:mm"));
              editedRange.setBackgroundRGB(255,0,0);
            } else if (PEboxValue == false && (PSboxValue == true || IboxValue == true) && OboxValue == false && s.getRange(row,7).getValue() != '') {
              s.getRange(row,8).clearContent();
              editedRange.clearFormat();
            } /*else
            
            {
              SpreadsheetApp.getUi().alert(
              "INVALID INPUT!"+
              "\nOnly one of each check-in type and one of each check-out type may be selected."+
              "\nA check-in type must be selected before a check-out type."
              );
            }
            */
            break;
          default:
            SpreadsheetApp.getUi().alert(
            "INVALID INPUT!"+
            "\nOnly one of each check-in type and one of each check-out type may be selected."+
            "\nA check-in type must be selected before a check-out type."
            );
        }
        
        
        
        
        /*
        if (column == 3 && PSboxValue == true && PEboxValue == false && IboxValue == false && OboxValue == false && s.getRange(row,8).getValue() == '') {
          s.getRange(row,7).setValue(Utilities.formatDate(new Date(s.getRange(1,7).getValue()),Session.getScriptTimeZone(),"HH:mm"));
          editedRange.setBackgroundRGB(0,255,0);
        } else if (column == 3 && PSboxValue == false && PEboxValue == false && IboxValue == false && OboxValue == false) {
          s.getRange(row,7).clearContent();
          Ibox.setValue(false);
          editedRange.clearFormat();
        } else
        
        if (column == 4 && PEboxValue == true && (PSboxValue == true || IboxValue == true) && OboxValue == false && s.getRange(row,7).getValue() != '') {
          s.getRange(row,8).setValue(Utilities.formatDate(new Date(s.getRange(1,8).getValue()),Session.getScriptTimeZone(),"HH:mm"));
          editedRange.setBackgroundRGB(255,0,0);
        } else if (column == 4 && PEboxValue == false && (PSboxValue == true || IboxValue == true) && OboxValue == false) {
          s.getRange(row,8).clearContent();
          editedRange.clearFormat();
        } else
        
        if (column == 5 && PSboxValue == false && PEboxValue == false && IboxValue == true && OboxValue == false && s.getRange(row,8).getValue() == '') {
          s.getRange(row,7).setValue(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"HH:mm"));
          editedRange.setBackgroundRGB(0,255,0);
        } else if (column == 5 && PSboxValue == false && PEboxValue == false && IboxValue == false && OboxValue == false) {
          s.getRange(row,7).clearContent();
          editedRange.clearFormat();
        } else
        
        if (column == 6 && PEboxValue == false && (PSboxValue == true || IboxValue == true) && OboxValue == true && s.getRange(row,7).getValue() != '') {
          s.getRange(row,8).setValue(Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"HH:mm"));
          editedRange.setBackgroundRGB(255,0,0);
        } else if (column == 6 && PEboxValue == false && (PSboxValue == true || IboxValue == true) && OboxValue == false) {
          s.getRange(row,8).clearContent();
          editedRange.clearFormat();
        } else
        // Edge Cases to prevent lock up
        /*
        // PS
        if (column == 3 && PSboxValue == true && (PEboxValue == true || OboxValue == true)) {
        } else
        // PE
        if (column == 4 && PSboxValue == false && PEboxValue == false && IboxValue == false && OboxValue == false) {
        } else
        //I
        if (column == 5 && PSboxValue == true && (PEboxValue == true || OboxValue == true)) {
        } else
        //O
        if (column == 6 && PSboxValue == false && PEboxValue == false && IboxValue == false && OboxValue == false) {
          s.getRange(row,8).clearContent();
          editedRange.clearFormat();
        } else
        ***
        
        
        if (column == 3 && PSboxValue == true && (PEboxValue == true || OboxValue == true)) {
        } else
        {
          SpreadsheetApp.getUi().alert(
          "INVALID INPUT!"+
          "\nOnly one of each check-in type and one of each check-out type may be selected."+
          "\nA check-in type must be selected before a check-out type."
          );
        }
        */
      }
      // Checkbox edited END
      
      // Manual editing
       //if (row > 2 && (column == 7 || column == 8)) {
       //  editedRange.getValue();
       //}
       
      // Manual Stuff
      // Calcualte Hours
      if (row == 7 && column == 11) {
        calculateBox = s.getRange(row,column);
        calculateBoxValue = calculateBox.getValue();
        if (calculateBoxValue == true) {
          calculateBox.setBackground("#ffa500");
          calculateHours();
          calculateBox.setBackground(null);
          calculateBox.setValue(false);
        }
      }
      // Send Hours
      // DID NOT WORK EARLIER????
      
    }
  }
  //scriptProperties.setProperty('ghostCheckIsFinished', true);
  return;
}

function calculateHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range = s.getRange(3,3,endRow - 2,4);
  
  for (var i = 3; i <= endRow; i++) {
    if (s.getRange(i,8).getValue() != '' && s.getRange(i,7).getValue() != '') {
      if (s.getRange(i,8).getValue() > s.getRange(i,7).getValue()) {
        //s.getRange(i,9).setValue(s.getRange(i,8).getValue()-s.getRange(i,7).getValue());
        s.getRange(i,9).setValue((s.getRange(i,8).getValue()-s.getRange(i,7).getValue()) / 3600000);
      } else {
        s.getRange(i,9).setValue(24 - ((s.getRange(i,8).getValue()-s.getRange(i,7).getValue()) / -3600000));
      }
    }
  }
}




//TO BE COMPLETED! IMPORTANT FUNCTIONALITY!

function sendHours() {
  //SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto")); // ADMIN SHEET ID
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range;
  var data = [];
  var timestamp = Utilities.formatDate(new Date, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  var sig;
  
  var targetSheet;
  var scriptProperties = PropertiesService.getScriptProperties();
  var eventType = scriptProperties.getProperty("eventType");
  
  var vFlag = false;
  
  if (eventType == "Volunteering") {
    targetSheet = "Master Hours Log";
    vFlag = true;
  } else if (eventType == "Participation") {
    targetSheet = "Master Participation Hours Log";
  }
  
  for (var i = 3; i <= endRow; i++) {
    range = s.getRange(i,9).getValue();
    Logger.log(range);
    if (range != '') {
      data.push(timestamp,(s.getRange(i,2).getValue()+"@bergenfield.org"), s.getRange(i,9).getValue());
      SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
      ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
      ssIndex = ss.getSheetByName(targetSheet).getIndex();
      saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
      s = ss.getSheets()[saIndex];
      endRow = s.getLastRow();
      s.getRange(endRow+1,2).setValue(data[0]);
      s.getRange(endRow+1,3).setValue(data[1]);
      s.getRange(endRow+1,7).setValue(data[2]);
      if (vFlag) {
        ssIndex = ss.getSheetByName("Signature List").getIndex();
        saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
        s = ss.getSheets()[saIndex];
        endRow = s.getLastRow();
        
        for (var j = 2; j < endRow; j++) {
          if (s.getRange(j,1).getValue() == "UNUSED") {
            sig = s.getRange(j,2).getValue();
            s.getRange(j,1).setValue("USED");
            ssIndex = ss.getSheetByName(targetSheet).getIndex();
            saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
            s = ss.getSheets()[saIndex];
            endRow = s.getLastRow();
            s.getRange(endRow,8).setValue(sig);
          }
        }
      }
      
      SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(scriptProperties.getProperty("sheetId")));
        ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
        ssIndex = ss.getSheetByName("Attendance").getIndex();
        saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
        s = ss.getSheets()[saIndex];
        endRow = s.getLastRow();
        range = '';
        data = [];
    }
  }
}

function sendHoursViaButton() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  
  if (s.getRange(9,11).getValue() == true) {
    s.getRange(9,11).setBackground("#ffa500");
    sendHours();
    s.getRange(9,11).setBackground(null);
    s.getRange(9,11).setValue(false);
  }
}

function TESTsetEventTimeFrame() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  
  range = s.getRange(1,7);
  range.setValue("Fri Aug 17 8:00:00 GMT-04:00 2018");
  range = s.getRange(1,8);
  range.setValue("Fri Aug 17 15:00:00 GMT-04:00 2018");
  
}





function BLAHBLAH() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range = s.getRange(3,3,endRow - 2,4);
  var currentAttendanceState = range.getValues();
  var scriptProperties = PropertiesService.getScriptProperties();
  var properties;
  var isRowFinishedList;
  var isRowEditedList;
  
  if (currentAttendanceState.toString() != scriptProperties.getProperty('lastAttendanceState')) {
    
    for (var i = 3; i <= endRow; i++) {
      if (!isRowFinishedList[i]) {
        if (isRowEditedList[i]) {
          
        }
      }
    }
    Logger.log('DIFFERENT!');
  }
  Logger.log('SAME!');
  range = s.getRange(3,3,endRow - 2,4);
  currentAttendanceState = range.getValues();
  scriptProperties.setProperty('lastAttendanceState', currentAttendanceState.toString());
}

function generateSubArray(input) {
  var output = [];
  var length = input.length;
  var currentItem = '';
  
  for (var i = 0; i < length; i++) {
    if (input.charAt(i) == ',') {
      output.push(currentItem);
      currentNumber = '';
    } else {
      currentItem += input.charAt(i);
    }    
  }
}

function nextLevel(input) {
  return input[0];
}

// FINDING! Javascript makes SHALLOW COPIES (reference based) of Arrays
function arrayTesting() {
  var array = [];
  var c = array;
  var length = 2;
  for (var i = 0; i < length; i ++) {
    c.push([]);
    c = c[0];
  }
  Logger.log(array[0][0][0]);
  
}


function qqqq(a) {
  Logger.log(arguments[0]);
  Logger.log(arguments[1]);
  Logger.log(arguments[2]);
  Logger.log(arguments[3]);
  Logger.log(arguments[4]);
  Logger.log(arguments.length);
  Logger.log(a);
}

function zzzz() {
  qqqq(1,2,3,4,5,6);
}


// input, 2, 4, 5
// length - 1 = 2
// expected array: [][][]

// input, 1
// input, 0
// length = 0 or 1
// expected array: []

function multiSplit(input, separator, dimensions) {
  var output = [];
  var numberOfInputElements = input.split(",");
  var numberOfOutputElements = 1;
  var numberOfDimensions = (arguments.length == 2) ? (1) : (arguments.length - 2);
  var currentArray = output;
  var currentItem = '';
  
  
  input.split(',');

}

function logqqq() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range = s.getRange(3,3,endRow - 2,4);
  var currentAttendanceState = range.getValues();
  Logger.log(currentAttendanceState.toString());
}

// false,x
// true,x

function listIfRowIsEdited(currentState, lastState) {
  var output = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastState = scriptProperties.getProperty('lastAttendanceState');
  var cPos = 0;
  var lPos = 0;
  
  var range = s.getRange(3,3,endRow - 2,4);
  currentState = range.getValues();
  var longestStringLength = (currentState.length > lastState.length) ? currentState.length : lastState.length;
  
  for (var i = 0; i < endRow - 2; i++) {
    for (var j = 0; j < 4; j++) {
      if (currentState.charAt(cPos) == lastState.charAt(lPos)) {
        output.push()
      }
      if (currentState.charAt(cPos) == f) {
        cPos += 6;
      } else {
        cPos += 5;
      }
      if (lastState.charAt(cPos) == f) {
        lPos += 6;
      } else {
        lPos += 5;
      }
    }
  }
}



function test() {
var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var editedRange = s.getRange(3,2);
  var editedRangeValues = editedRange.getValues();
  var row;
  var column;
  var PSbox;
  var PEbox;
  var Ibox;
  var Obox;
  var PSboxValue;
  var PEboxValue;
  var IboxValue;
  var OboxValue;
  var temp;
  
  row = editedRange.getRow();
  column = editedRange.getColumn();
  s.getRange(row,1).setValue(lookupStudent(s.getRange(row,column).getValue()));
}

// BARE MINIMUM NECESSARY:
