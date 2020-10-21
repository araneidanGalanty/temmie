/* New Admin Welcome and Checklist:
 * 
 * Hello, future system admin.
 * This is TEMMIE: the backend to the SNHS administration system.
 * Important instructions at "Steps to add new members:"
 * 
 * Contact: james@jameslabayna.com (will retire this email eventually)
 * 
 * WARNING: Syncing process is horribly inefficient (too much getValue). May exceed max execution time at some point. Current fix (2019-06-18: seperate sync and summing; one process per hour)
 * POSSIBLE FIX: Copy every year's master hours log into a master master hours log
 * BETTER FIX: getRange of all new hours/relevant students. Iteratre through ranges as multidimensional arrays.
 *
 * BUG: Multiple submits of attendance are valid
 * FIX: Atendance buffering:
 * Create a TEMMIE sheet that lists active sheets for maybe the last a day or so.
 * All attendance submissions are stored on this list page somehow. Old submissions are overwrittern.
 * After the time limit, submissions are sent to hours log.
 *
 * Note: All participation hours are not checked for validity because really only admins are going to send them (as of 2019-06-18). Just don't mess up :P (or implement a fix)
*/

/* MESSAGE TO FUTURE TEMMIE ADMIN (programmer/IT, not general administrator):
 * This might not be necessary since student drives aren't really delete don graduation, but maybe do it anyway? or no? 
 * Create a folder in your google drive. DO NOT DELETE THIS FOLDER OR ADD OTHER USERS TO IT.
 * Make note of the folder id.
 * Then, copy and paste the folder id, your school account's email and username exactly as was original into the getStudentNameByEmail(email).
 */
 
// NEED ABILITY TO ADD NEW STUDENTS INTO SHEETS. NON MANUAL

/* Idea: QR clock-in
 * 1. Create clock in and clock out form (different for each event maybe)
      Should be 1 response only. No questions, just submit button
 * 2. Make QR code for each
 * 3. Students scan code on entry/exit
 * 4. Responses go to TEMMIE or TEMMIE attendence sheet
 * 5. Someone takes attendance to check for conflicts? maybe?
 * 6. Hours compilated and sent to TEMMIE master logs
 * 
 * Possible issue: students take pictures of QR code and send it.
 * Fix: Just take attendence dammit. That's why we have supervisors.
 */


/* TODO:
 * Set active spreadsheet to fixed id
 * QR Code Clock-in and Clock-Out Functionality
 *   - Via google forms via QR code?
 * TIDY EVERYTHING UP
 *
 * IMPORTANT!!!!!
 * Anything with Student Reference Sheet must be rechecked for functionality!
 * 
 * Form --> Seperate Sheet or Spreadsheet --> Master Sheet
 */

/* GENERAL NOTES:
 * Event Supervisors want the simplicity of a signature
 * Resultant problem: Supervisors won't manually input student hours
 * Solution: Signatures
 *
 * Really hacky way to avoid having to use a web app for clocking in and out:
 *   - Use Spreadsheet
 *   - Identify user on open
 *   - THE CATCH: May not work on phones.
 */

// Kinda hacky. Dependent on 195300@bergenfield.org's drive existing.

/* Steps to add new members: NEW ADMINISTRATION LOOK HERE!!!!!!!!! Do this every year when you have new members or restart the society.
0. HAVE ALL MEMBERS (including old ones) RESUMBIT FORM. Here's how to reset the form:
  a. Open https://docs.google.com/forms/d/1o78fKdcN0gOg4JCSqo0dOC8JXXSCU3BMfUA4zxcr6gM/edit#responses or a copy.
  b. Delete all responses
  c. Unlink from sheet.
  d. Check TEMMIE to see if the associated form is also unlinked there. You may need to right click the sheet and unlink in Temmie.
  e. Delete the old response sheet.
  f. Open https://docs.google.com/forms/d/1o78fKdcN0gOg4JCSqo0dOC8JXXSCU3BMfUA4zxcr6gM/edit#responses or the copy.
  g. Link to spreadsheet TEMMIE
  h. RESEND TO EVERYONE.
  i. When done, PUT COPY OF ALL RESPONSES UNDER OLD NAMES IN master student reference sheet. There should be a continuous line of all SNHS members since 2018
  j. Automate this in the future :P
1. run recordMemberNames() once ALL NAMES are submitted to form and copied.
2. run organizeYearlyStudentList()
3. createStudentSheets() (This auto deletes old). USE THIS AGAIN IF YOU EVER DELETE A STUDENT
4. createSharedStudentFolders()
5. createSharedHoursLogs()
6. syncStudentSheets() (Should be automatic? May need manual activation)
7. syncParticipationHours()
8. sumhours
9. sumParticipationHours

RESET PROCESS (if there is no added/deleted student):
1. Set lastSyncedParticipationEntry & lastSyncedEntry to 1.0
   Set all "signatures" to "UNUSED"
2. createStudentSheets()
3. syncStudentSheets()
4. syncParticipationHours()
5. sumhours
6. sumParticipationHours

RESET if a studnet was deleted: (Delete row from yearly reference and delete sheet or do the following)
0. Delete student FROM YEARLY STUDENT REFERENCE SHEET ONLY.
1. Set lastSyncedParticipationEntry & lastSyncedEntry to 1.0
   Set all "signatures" to "UNUSED"
2. createStudentSheets() (auto deletes deleted student)
3. syncStudentSheets()
4. syncParticipationHours()
5. sumhours
6. sumParticipationHours

RESET if student was added:
1. set script property "lastRecordedMemberRow" to the ROW IMMEDIATELY BEFORE CURRENT YEAR ROSTER SUBMISSIONS. (i.e. if last year there were 10 students and this year 6, and these
   are the only students ever, set "lastRecordedMemberRow" to 11 because the title bar counts as a submission)
   This value is usually "prevLastRecordedMemberRow".
2. follow "Steps to add new members" from step 1.

How to add admins:
1. add admin to drive using email. Must be manager

How to Nullify or remove hours:
1. Submit negative hours. May require manual or special entry to master log

*/

// WHOEVER WRITES THE CODE AND "OWNS" TEMMIE SHOULD BE LISTED HERE! NO ONE ELSE! Might not be necessary as long as 195300 is listed here.
function getStudentNameByEmail(email){
  var middlemanFile = DriveApp.getFileById('1GPFu6_hxlCYE1-7g8ktLgzSbSmTOYUAF'); // Folder in 195300@bergenfield.org's drive. change as appropriate, but never delete the folder
  var knownUserEmail = "195300@bergenfield.org"; // This variable is EXTREMELY IMPORTANT! Since the sheet owner (programmer/IT) can't really be added as a viewer, the owner MUST enter their email here.
  var KnownUnformattedUserName = "JAMES-ALBERT.LABAYNA 195300" // Similarly, enter the default school google account name here.
  var viewerList;
  var targetUser;
  var unformattedName;
  var formattedName;
  var buffer;
  var hyphenIndexes = [];
  
  if (email == knownUserEmail) {
    // Trim off ID
    unformattedName = KnownUnformattedUserName.substring(0, KnownUnformattedUserName.indexOf(' '));
    unformattedName = unformattedName.toLowerCase();
    unformattedName = unformattedName.charAt(0).toUpperCase() + unformattedName.slice(1);
    
    for(var i = 0; i < unformattedName.length; i++) {
      if (unformattedName[i] == "-") hyphenIndexes.push(i);
    }
    if (hyphenIndexes[0] !== '') {
      for (var j = 0; j < hyphenIndexes.length; j++) {
        unformattedName = unformattedName.substring(0, hyphenIndexes[j] + 1) +  unformattedName.charAt(hyphenIndexes[j] + 1).toUpperCase() + unformattedName.slice(hyphenIndexes[j] + 2);
      }
    }
    unformattedName = unformattedName.substring(0, unformattedName.indexOf('.') + 1) + unformattedName.charAt(unformattedName.indexOf('.') + 1).toUpperCase() + unformattedName.slice(unformattedName.indexOf('.') + 2);
    buffer = unformattedName.substring(unformattedName.indexOf('.') + 1);
    formattedName = buffer + ", " + unformattedName.substring(0, unformattedName.indexOf('.'))
    Logger.log(formattedName);
    return formattedName;
  }
  
  middlemanFile.addViewer(email);
  viewerList = middlemanFile.getViewers();
  targetUser = viewerList[0];
  middlemanFile.removeViewer(targetUser);
  unformattedName = targetUser.getName();
  // Trim off ID
  unformattedName = unformattedName.substring(0, unformattedName.indexOf(' '));
  unformattedName = unformattedName.toLowerCase();
  unformattedName = unformattedName.charAt(0).toUpperCase() + unformattedName.slice(1);
  
  for(var i = 0; i < unformattedName.length; i++) {
    if (unformattedName[i] == "-") hyphenIndexes.push(i);
  }
  if (hyphenIndexes[0] !== '') {
    for (var j = 0; j < hyphenIndexes.length; j++) {
      unformattedName = unformattedName.substring(0, hyphenIndexes[j] + 1) +  unformattedName.charAt(hyphenIndexes[j] + 1).toUpperCase() + unformattedName.slice(hyphenIndexes[j] + 2);
    }
  }
  unformattedName = unformattedName.substring(0, unformattedName.indexOf('.') + 1) + unformattedName.charAt(unformattedName.indexOf('.') + 1).toUpperCase() + unformattedName.slice(unformattedName.indexOf('.') + 2);
  buffer = unformattedName.substring(unformattedName.indexOf('.') + 1);
  formattedName = buffer + ", " + unformattedName.substring(0, unformattedName.indexOf('.'))
  Logger.log(formattedName);
  return formattedName;
}

// Finds names of new members in the Master Student Reference Sheet and lists those names in the propper field
function recordMemberNames() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex]; // WTF is getSheets() INDEXED AT 0?
  var range;
  var startRow;
  var endRow = s.getLastRow();
  var scriptProperties = PropertiesService.getScriptProperties();
  
  var currentName;
  
  
  startRow = parseInt(scriptProperties.getProperty("lastRecordedMemberRow")) + 1; // Default value of lastRecordedMemberRow is 1.0
  for (var i = startRow; i <= endRow; i++) {
    currentName = getStudentNameByEmail(s.getRange(i,2).getValue());
    s.getRange(i,3).setValue(currentName);
  }
  scriptProperties.setProperty("prevLastRecordedMemberRow", scriptProperties.getProperty("lastRecordedMemberRow"));
  scriptProperties.setProperty("lastRecordedMemberRow", endRow);
}

// Creates a new list of current members in a seperate Yearly Sheet
function organizeYearlyStudentList() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex]; // WTF is getSheets() INDEXED AT 0?
  var range;
  var scriptProperties = PropertiesService.getScriptProperties();
  var startRow = scriptProperties.getProperty("lastRecordedMemberRow") + 1;//findCurrentStudents(); // findCurrentStudents() is broken 1/30/2019 // 2019-06-17: Fixed?
  var endRow = s.getLastRow(); // Subtract 1 from endRow to get name array length
  var currentNames = s.getRange(startRow,3,(endRow - startRow) + 1).getValues(); // 2d array
  var currentIds = s.getRange(startRow,2,(endRow - startRow) + 1).getValues();
  var namesAndIds;
  
  // Clear old
  range = s.getRange(2, 1, 200, 4); // If this ever breaks, then you must have changed the sheet dimensions of the yearly list.
  range.clear({contentsOnly: true});
  // Reset values
  ssIndex = ss.getSheetByName("Master Student Reference Sheet").getIndex();
  saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  s = ss.getSheets()[saIndex]; // WTF is getSheets() INDEXED AT 0?
  endRow = s.getLastRow(); // Subtract 1 from endRow to get name array length
  currentNames = s.getRange(startRow,3,(endRow - startRow) + 1).getValues(); // 2d array
  currentIds = s.getRange(startRow,2,(endRow - startRow) + 1).getValues();
  // Clear old end
  
  for (var i = 0; i < currentIds.length; i++) {
    currentIds[i][0] = currentIds[i][0].replace("@bergenfield.org",'');
  }
  namesAndIds = currentNames;
  for (var i = 0; i < namesAndIds.length; i++) {
    namesAndIds[i].push(i);
  }
  namesAndIds.sort();
  for (var i = 0; i < namesAndIds.length; i++) {
    namesAndIds[i][1] = currentIds[namesAndIds[i][1]][0];
  }
  
  ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  saIndex = ssIndex - 1;
  s = ss.getSheets()[saIndex];
  s.getRange(2,1,namesAndIds.length,2).setValues(namesAndIds);
  
}

// Disabled as of 2019-06-17
// IMPORTANT NEEDS TO BE REDONE
// Finds first student reference sheet row of current year
// Child of createStudentSheets()
function findCurrentStudents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex]; // WTF is getSheets() INDEXED AT 0? I mean, it's nice, but inconsistent
  var endRow = s.getLastRow(); // Subtract 1 from endRow to get name array length
  var studentsNotFound = true;
  var tempRange;
  var joinDates;
  var currentRow = 2; // Subtract 2 from currentRow to get usable array value
  /*
  var now = new Date();
  
  tempRange = s.getRange(2, 1, (endRow - 1));
  joinDates = tempRange.getValues();

  while (studentsNotFound) {
    if ((joinDates[currentRow - 2][0]).getYear() == now.getYear() || currentRow == endRow) {
      studentsNotFound = false;
    } else {
      currentRow++;
    }
  }
  return currentRow;
  */
  // Broken bit? ^^^
  
  }

// WARNING! DELETES ALL SHEETS AFTER "Yearly Student Reference Sheet"
function deleteStudentSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var allSheets = ss.getSheets();
  var ssLength = allSheets.length;
  
  for (var i = ssIndex; i < ssLength; i++) {
    ss.deleteSheet(allSheets[i]);
  }
}


// IMPORTANT NEEDS TO BE FINISHED
// HACKED TOGETHER 2018-07-09: 7:39!!!!! PLEASE FIX ME LATER
function createStudentSheets() {
  /* Spreadsheet indexes start at 1
   * Sheet array indexes start at 0
   * WTF Google!
   */
   // NOTE TO SELF: Probably should create seperate spreadsheets for each user.
   //               Keep 1 big reference for admins.
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex]; // WTF is getSheets() INDEXED AT 0?
  var startRow = 2;//findCurrentStudents(); // Start row; 2 by default // findCurrentStudents() NEEDS TO BE FIXED!!!!
  var endRow = s.getLastRow(); // Subtract 1 from endRow to get name array length
  var range = s.getRange(2, 1, (endRow - 1)); // row coodinate should be startrow. Plz change later; Change endrow - 1 back to endrow - 2
  var names = range.getValues();
  var ids = s.getRange(2,2,(endRow -1)).getValues();
  var titleRange; // Used to populate new student sheet titles
  var hoursChartTitles; // Hours
  
  var allSheets = ss.getSheets();
  var ssLength = allSheets.length;
  
  
  var idCell; // Actually a merged range
  var nameCell;
  var summaryTitleCell;
  var entryTitleRow;
  var blah;
  
  var sTemp;
  
  /*
  //Fancy indicator thingy
  var indicatorCell = s.getRange("F1");
  indicatorCell.setBackgroundRGB(250,100,100);
  indicatorCell.setValue("WORKING");
  */
  
  deleteStudentSheets();
  saIndex++;
  for (var i = 0; i < (endRow - startRow + 1); i++) { // Change middle parenthesis to endRow - startRow + 1 
    // New
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1_FHSPuMGs2uy44FpnKTzfV4jTTL9m4M3NW4F7sTywRA"));
    ss = SpreadsheetApp.getActiveSpreadsheet();
    sTemp = ss.getSheetByName("PLACEHOLDER");
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
    ss = SpreadsheetApp.getActiveSpreadsheet();
    sTemp.copyTo(ss);
    
    s = ss.getSheets()[saIndex];
    s.setName(names[i][0]);
    
    idCell = s.getRange(1,1,2,2);
    idCell.setValue(ids[i]);
    idCell = s.getRange(1,4,2,2);
    idCell.setValue(ids[i]);
    idCell = s.getRange(1,7,2,2);
    idCell.setValue(ids[i]);
    
    nameCell = s.getRange(1,9,2,9);
    nameCell.setValue(names[i][0]);
    nameCell = s.getRange(1,19,2,9);
    nameCell.setValue(names[i][0]);
    nameCell = s.getRange(1,29,2,9);
    nameCell.setValue(names[i][0]);
    saIndex++;
    //New End
    //Old
    /*
    ss.insertSheet(names[i][0], ssIndex); // ssIndex is 1 greater than current real index
    s = ss.getSheets()[ssIndex];      // ssIndex is real index of new sheet
    
    // Column and Row Deletion
    s.deleteColumns(12,15);
    s.deleteRows(47,951);
    // School ID Cell
    idCell = s.getRange(1,1,2,2)
    idCell.merge();
    idCell.setBackground("#004182");
    idCell.setFontColor("#FFFFFF");
    idCell.setFontWeight("bold");
    idCell.setHorizontalAlignment("center");
    idCell.setFontSize(25);
    idCell.setValue(ids[i]);
    // Name Box
    nameCell = s.getRange(1,3,2,9);
    nameCell.merge();
    nameCell.setBackground("#004182");
    nameCell.setFontColor("#FFFFFF");
    nameCell.setFontWeight("bold");
    nameCell.setFontSize(25);
    nameCell.setValue(names[i][0]);
    
    // Hours Summary Title
    summaryTitleCell = s.getRange(3,1,1,2);
    summaryTitleCell.merge();
    summaryTitleCell.setBackground("#004182");
    summaryTitleCell.setFontColor("#FFFFFF");
    summaryTitleCell.setFontWeight("bold");
    summaryTitleCell.setHorizontalAlignment("center");
    summaryTitleCell.setFontSize(10);
    summaryTitleCell.setValue("HOURS SUMMARY");
    
    // Entry Titles
    entryTitleRow = s.getRange(3,3,1,9);
    entryTitleRow.setValues([["Entry ID", "Status", "Timestamp", "Email", "Activity Date", "Location", "Duties/Work", "Hours", "Signature"]]);
    
    // Banding
    s.getRange(3,3,47,9).applyRowBanding(SpreadsheetApp.BandingTheme.GREY);
    
    // Sumary Sections
    for (var j = 1; j <= 9; j++) {
      s.getRange((j*5)-1,1,1,2).merge();
      s.getRange((j*5)-1,1,1,2).setBackground("#7fa0c0")
      s.getRange((j*5)-1,1,1,2).setFontWeight("bold");
      s.getRange((j*5)-1,1,1,2).setFontSize(10);
      if (j <= 4) {
        s.getRange((j*5)-1,1,1,2).setValue("Y1 "+"Q"+j);
      } else if (j <=8) {
        s.getRange((j*5)-1,1,1,2).setValue("Y2 "+"Q"+j);
      } else {
        s.getRange((j*5)-1,1,1,2).setValue("OVERALL");
      }
      s.getRange(j*5,1,3,2).setValues([["Submitted:",''],["Pending:",''],["Approved",'']])
      s.getRange((j*5)+3,1,1,2).merge();
      if (j <= 8) {
        s.getRange((j*5)+3,1).setValue("=SPARKLINE("+"B"+((j*5)+2)+",{\"charttype\",\"bar\";\"max\",8;\"empty\",\"zero\";\"color1\",if("+"B"+((j*5)+2)+" < 8/3,\"#FF0000\",if("+"B"+((j*5)+2)+" < 16/3,\"#ffa500\",if("+"B"+((j*5)+2)+" < 8,\"#e6e600\",\"#00FF00\")))})");
      } else {
        s.getRange((j*5)+3,1).setValue("=SPARKLINE("+"B"+((j*5)+2)+",{\"charttype\",\"bar\";\"max\",64;\"empty\",\"zero\";\"color1\",if("+"B"+((j*5)+2)+" < 64/3,\"#FF0000\",if("+"B"+((j*5)+2)+" < 64/3,\"#ffa500\",if("+"B"+((j*5)+2)+" < 64,\"#e6e600\",\"#00FF00\")))})");
      }
    }
    
    // Bottom Cosmetic thing that makes the bottom of the summary look like it's got no cells
    
    s.getRange(49,1,1,2).setBackground("#f3f3f3");
    s.getRange(49,1,1,2).setBorder(true, false, true, false, true, true, "#f3f3f3", null);
    
    
    
    s.setFrozenRows(3);
    s.setFrozenColumns(2);
    */
    // END OLD
    /*
    hoursChartTitles = s.getRange(1,1,2,4);
    hoursChartTitles.setValues([["Hours:", '', '', ''],["Approved:", '', "Submitted:", '']]);
    
    titleRange = s.getRange("A3:I3");  // DEPENDENT ON HOURS INPUT CATEGORIES. EDIT AS NEEDED ***
    titleRange.setValues([["Entry ID","Status","Timestamp", "Email", "Activity Date", "Location", "Duties/Work", "Hours", "Signature"]]);
    s.setFrozenRows(3);
    ssIndex++;
    */
  }
  
  /*
  indicatorCell.setBackgroundRGB(100,250,100);
  indicatorCell.setValue("Done!");
  */
}


// No idea what I was doing when I wrote "ent.statusTargetCells". Not sure what field/method I was referring to...
/* How to implement:
 * 1. Obtain string in propper format
 * 2. remove all spaces
 * 3. split string into array; seperator is ";"
 * 4. Iterate through array
 *   a. if there is a "-", then set start and finish of inner iteration to start/finish IF AND ONLY IF start <= finish
 *      else, start = finish = array element
 *      MAKE SURE ALL ENTRIES EXIST
 *   b. In the inner iteration (for loop), approve/pend/whatever from start to finish
 * 5. Store array in a script property
 *    If sync hasn't rerun is true, append new array to old
 *    Else, replace old
 * 6. Add code to sync functions so that:
 *   a. Sync first syncs entries in stored array that are OLDER THAN LAST SYNCED ENTRY
 *      Sync ONLY the LAST instance of a number in an array
 *   b. If there are no array elements above last synced entry, then sync as normal
 *      Else, sync all non-manually changed entries as normal. Follow a for changed entries.
 */
function approveEntry(ent) {
  ent = ent.statusTargetCells;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var statusCell = s.getRange(ent, 1);
  
  statusCell.setValue("APPROVED");
  statusCell.setBackgroundRGB(0,255,0);
}

function rejectEntry(ent) {
  ent = ent.statusTargetCells;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var statusCell = s.getRange(ent, 1);
  
  statusCell.setValue("REJECTED");
  statusCell.setBackgroundRGB(255,0,0);
}

function pendEntry(ent) {
  Logger.log(ent);
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  ent = ent.statusTargetCells;
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var statusCell = s.getRange(ent, 1);
  
  statusCell.setValue("PENDING");
  statusCell.setBackgroundRGB(255,165,0);
}

function pendAll() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  
  for (var i = 2; i <= endRow; i++) {
    pendEntryFirstTime(i);
  }
}

function pendEntryFirstTime(ent) {
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var statusCell = s.getRange(ent, 1);
  
  statusCell.setValue("PENDING");
  statusCell.setBackgroundRGB(255,165,0);
}

function applyStatus() { // ON FORM SUBMIT
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range = s.getRange(2, 1, (endRow - 1));
  var statusArray = range.getValues(); // True row - 2 = array row ; True column = 0
  
  for (var i = 2; i <= endRow; i++) {
    Logger.log(statusArray[i - 2][0] == '');
    if (statusArray[i - 2][0] == '') {
      pendEntryFirstTime(i);
    }
  }
}

function pendParticipationEntryFirstTime(ent) {
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Master Participation Hours Log").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var statusCell = s.getRange(ent, 1);
  
  statusCell.setValue("PENDING");
  statusCell.setBackgroundRGB(255,165,0);
}

function applyParticipationStatus() { 
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Participation Hours Log").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var range = s.getRange(2, 1, (endRow - 1));
  var statusArray = range.getValues(); // True row - 2 = array row ; True column = 0
  
  for (var i = 2; i <= endRow; i++) {
    Logger.log(statusArray[i - 2][0] == '');
    if (statusArray[i - 2][0] == '') {
      pendParticipationEntryFirstTime(i);
    }
  }
}

function validateSignature(signature) {
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ssIndex = ss.getSheetByName("Signature List").getIndex();
  var saIndex = ssIndex - 1;
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  
  for (var i = 2; i <= endRow; i++) {
    if (s.getRange(i,1).getValue() == "UNUSED" && s.getRange(i,2).getValue() == signature) {
      s.getRange(i,1).setValue("USED");
      return 1;
    }
  }
  return 0;
}

// TEMPORARY FIX

function sumHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var currentEntry;
  var currentEntryValues;
  var endRow = s.getLastRow();
  var range;
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty('lastSyncedEntry');
  var entryEmail;
  var currentStudent;
  var currentTotalSum = 0;
  var currentApprovedSum = 0;
  
  
  for (var i = 2; i <= endRow; i++) {
    range = s.getRange(i, 1);
    currentStudent = range.getValue();
    ssIndex = ss.getSheetByName(currentStudent).getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
    for (var j = 5; j <= s.getLastRow(); j++) {
      range = s.getRange(j, 16);
      currentTotalSum += range.getValue();
      if (s.getRange(j, 10).getValue() == "APPROVED") {
        currentApprovedSum += range.getValue();
      }
    }
    s.getRange(48,2).setValue(currentApprovedSum);
    s.getRange(46,2).setValue(currentTotalSum);
    currentApprovedSum = 0;
    currentTotalSum = 0;
    ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
  }
  
}

function sumParticipationHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var currentEntry;
  var currentEntryValues;
  var endRow = s.getLastRow();
  var range;
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty('lastParticipationSyncedEntry');
  var entryEmail;
  var currentStudent;
  var currentTotalSum = 0;
  var currentApprovedSum = 0;
  
  
  for (var i = 2; i <= endRow; i++) {
    range = s.getRange(i, 1);
    currentStudent = range.getValue();
    ssIndex = ss.getSheetByName(currentStudent).getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
    for (var j = 5; j <= s.getLastRow(); j++) {
      range = s.getRange(j, 26);
      currentTotalSum += range.getValue();
      if (s.getRange(j, 20).getValue() == "APPROVED") {
        currentApprovedSum += range.getValue();
      }
    }
    s.getRange(48,5).setValue(currentApprovedSum);
    s.getRange(46,5).setValue(currentTotalSum);
    currentApprovedSum = 0;
    currentTotalSum = 0;
    ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
  }
  
}

// Volunteer Syncing:
function syncParticipationHours() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Participation Hours Log").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var currentEntry;
  var currentEntryValues;
  var endRow = s.getLastRow();
  var lastEntryRow;
  var currentEntryRow;
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = parseInt(scriptProperties.getProperty('lastSyncedParticipationEntry'));
  var entryEmail;
  var currentStudent;
  var statusCell;
  var j;
  
  var studentEnd;
  
  //applyParticipationStatus(); // This is buggy. No idea why this starts executing but finishes after next block of code MAY MEAN REGULAR HOURS WILL BUG TOO
  
  if (endRow > lastSyncedEntry) {
    for (var i = (lastSyncedEntry + 1); i <= endRow; i++) {
      
      //Validation process first. May be moved
      statusCell = s.getRange(i,1);
      if (true) {
        statusCell.setValue("APPROVED");
        statusCell.setBackgroundRGB(0,255,0);
      } else {
        statusCell.setValue("REJECTED");
        statusCell.setBackgroundRGB(255,0,0);
      }
      // Validation process END
      range = s.getRange(i , 3);
      entryEmail = range.getValue();
      currentStudent = lookupStudent(entryEmail);
      currentEntry = s.getRange(i, 1, 1, 8);
      Logger.log(i);
      Logger.log("student: " + currentStudent);
      
      // Handle invalid email
      if (currentStudent != "Student Not Found") {
        ssIndex = ss.getSheetByName(currentStudent).getIndex();
        saIndex = ssIndex - 1;
        // s now refers to student sheet
        s = ss.getSheets()[saIndex];
        studentEnd = s.getLastRow();
        if (s.getRange(50,19).getValue() == '') {
          // Find last row without messing up the look
          j = 4;
          while (s.getRange(j,19).getValue() != '') {
            j++;
          }
          j--; // Subtraction because blank row is not last row
          lastEntryRow = j;
          currentEntryRow = lastEntryRow + 1;
        } else {
          // Find last row without messing up the look
          if (s.getRange(studentEnd,19).getValue() != '') {
            s.insertRowAfter(studentEnd);
            currentEntryRow = studentEnd + 1;
            lastEntryRow = currentEntryRow - 1;
            s.getRange(50,1,1,8).copyTo(s.getRange(currentEntryRow,1));
          } else {
            // Find last row without messing up the look
            j = 50;
            while (s.getRange(j,19).getValue() != '') {
              j++;
            }
            j--; // Subtraction because blank row is not last row
            lastEntryRow = j;
            currentEntryRow = lastEntryRow + 1;
          }
        }
        s.getRange(currentEntryRow,19).setValue(i);
        currentEntry.copyTo(s.getRange((currentEntryRow), 20, 1, 8));
        ssIndex = ss.getSheetByName("Master Participation Hours Log").getIndex();
        saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
        s = ss.getSheets()[saIndex];
      // Invalid student branch
      } else {
        statusCell.setValue("STUDENT NOT FOUND");
        statusCell.setBackgroundRGB(255,255,0);
      }
    }
    ssIndex = ss.getSheetByName("Master Participation Hours Log").getIndex();
    saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
    s = ss.getSheets()[saIndex];
    scriptProperties.setProperty('lastSyncedParticipationEntry', s.getLastRow());
  }
  
}

// Needs protection against duplicate entries.
function syncStudentSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var currentEntry;
  var currentEntryValues;
  var endRow = s.getLastRow();
  var lastEntryRow;
  var currentEntryRow;
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = parseInt(scriptProperties.getProperty('lastSyncedEntry'));
  var entryEmail;
  var currentStudent;
  var statusCell;
  var j;
  
  var studentEnd;
  
  applyStatus();
  if (endRow > lastSyncedEntry) {
    for (var i = (lastSyncedEntry + 1); i <= endRow; i++) {
      
      //Validation process first. May be moved
      statusCell = s.getRange(i,1);
      if (validateSignature(s.getRange(i , 8).getValue())) {
        statusCell.setValue("APPROVED");
        statusCell.setBackgroundRGB(0,255,0);
      } else {
        statusCell.setValue("REJECTED");
        statusCell.setBackgroundRGB(255,0,0);
      }
      // Validation process END
      range = s.getRange(i , 3);
      entryEmail = range.getValue();
      currentStudent = lookupStudent(entryEmail);
      currentEntry = s.getRange(i, 1, 1, 8);
      Logger.log(i);
      Logger.log("student: " + currentStudent);
      
      // Handle invalid email
      if (currentStudent != "Student Not Found") {
        ssIndex = ss.getSheetByName(currentStudent).getIndex();
        saIndex = ssIndex - 1;
        // s now refers to student sheet
        s = ss.getSheets()[saIndex];
        studentEnd = s.getLastRow();
        if (s.getRange(50,9).getValue() == '') {
          // Find last row without messing up the look
          j = 4;
          while (s.getRange(j,9).getValue() != '') {
            j++;
          }
          j--; // Subtraction because blank row is not last row
          lastEntryRow = j;
          currentEntryRow = lastEntryRow + 1;
        } else {
          // Find last row without messing up the look
          if (s.getRange(studentEnd,9).getValue() != '') {
            s.insertRowAfter(studentEnd);
            currentEntryRow = studentEnd + 1;
            lastEntryRow = currentEntryRow - 1;
            s.getRange(50,1,1,8).copyTo(s.getRange(currentEntryRow,1));
          } else {
            // Find last row without messing up the look
            j = 50;
            while (s.getRange(j,9).getValue() != '') {
              j++;
            }
            j--; // Subtraction because blank row is not last row
            lastEntryRow = j;
            currentEntryRow = lastEntryRow + 1;
          }
        }
        s.getRange(currentEntryRow,9).setValue(i);
        currentEntry.copyTo(s.getRange((currentEntryRow), 10, 1, 8));
        ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
        saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
        s = ss.getSheets()[saIndex];
      // Invalid student branch
      } else {
        statusCell.setValue("STUDENT NOT FOUND");
        statusCell.setBackgroundRGB(255,255,0);
      }
    }
    ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
    saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
    s = ss.getSheets()[saIndex];
    scriptProperties.setProperty('lastSyncedEntry', s.getLastRow());
  }
  // Commented out b/c max execution time reached (2019-06-18)
  /*
  syncParticipationHours();
  sumHours();
  sumParticipationHours();
  */
}

function resetSync() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty('lastSyncedEntry');
  scriptProperties.setProperty('lastSyncedEntry', 1);
}

function lookupStudent(email) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var endRow = s.getLastRow();
  var range;
  var id = email.substring(0,email.indexOf("@bergenfield.org"));
  
  for (var i = 2; i <= endRow; i++) {
    if (s.getRange(i, 2).getValue() == id) {
      return s.getRange(i, 1).getValue();
    }
  }
  return "Student Not Found";
}

// Signature Stuff
// NEEDS MASSIVE OVERHAUL!!!! USE OWN SIGNATURE FUNCTIONS! GOOGLE'S ARE WEIRD AND MYSTERIOUS
// Darned PKCS#8 format


function generateRandomString(length) {
    var charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!\"#$%&'()*+,-./:;<=>?@[\\]^_`{|}~", // Some charcters duplicated to escape them
        retVal = "";
    for (var i = 0, n = charset.length; i < length; ++i) {
        retVal += charset.charAt(Math.floor(Math.random() * n));
    }
    return retVal;
}

function generateSignatures(amount) {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Signature List").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var endRow = s.getLastRow();
  var date = new Date();
  var duplicateFound = 0;
  
  var currentSignature;
  
  amount = amount.amountOfSignatures;
  amount = +amount;
  for (var i = endRow + 1; i <= endRow + amount; i++) {
    do {
      duplicateFound = 0;
      do {
        s.getRange(i,1,1,2).setValues([["UNUSED", generateRandomString(16)]]);
        currentSignature = s.getRange(i,2).getValue();
      } while (currentSignature.slice(0,1) == "=");
      if (s.getLastRow() > 2 && s.getLastRow() <= 44012666865176569775543212890625 + 1) { // Good for all 16 digit strings of 95 character set
        for (var j = 2; j < i; j++) {
          if (s.getRange(i,2).getValue() == s.getRange(j,2).getValue()) {
            duplicateFound = 1;
          }
        }
      }
    } while (duplicateFound);
    duplicateFound = 0;
  }
}


//TEST FUNCTION
function fdsfdsfdsffdsfd() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Signature List").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var date = new Date();
  var duplicateFound = 0;
  var charset = "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789!#$%&()*+,-./:;<=>?@[]^_`{|}~"
    
    Logger.log(charset.length);
    Logger.log(charset);
}

function padLeading (stringtoPad , targetLength , padWith) {
  return (stringtoPad.length < targetLength ? Array(1+targetLength-stringtoPad.length).join(padWith | "0") : "" ) + stringtoPad ;
}

function byteToHexString(bytes) {
  return bytes.reduce(function (p,c) {
    return p += padLeading((c < 0 ? c+256 : c).toString(16), 2 );
  },'');
}

function testes() {
 var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "101010100101010103920924u09348309843084309843904539ire9043u09");
 Logger.log(digest);
 Logger.log(byteToHexString(digest));
 
 // This will write an array of bytes to the log.
 /*
 -----BEGIN RSA PRIVATE KEY-----
MIICXgIBAAKBgQCMXZYNnZ1PMyRbUN6d8L087AcC9beuiegSUjASS/o0rcs3y0Im
V16PNJSjys7E1abHn3Dk7JkXWWZusTSbML2n/X4hBng4MCLP8cOWSqxmOMKv2B+0
GVTkVGOo9WvpA0s1B/dMZKfe/sxbswDWpKyIJYL3NPTz/xijb96yRjT6MwIDAQAB
AoGBAIYf2wVRqYKHbOMw6DflVP5EzwJeB1Fph28SR8sD/KaftwKuX5xBeiK+7JWC
coeVXBN94CNvjW3JSF7XR1xPe7j5LL8o9Mp4Bo6BMrloMiTBC/4ykuReo7B83QTf
o+DCXp+CQhvj6MbywQrfoEjuXV7xKYEwJZtKERIfXJjtCO2RAkEA7wUG1T22SXye
RAQRrhtpKAqYt5+SOnAJeJFdnjtES0JCjn0HO3SLxJXzDunsMJLRi3WCssPk49M0
G9MW2jFeWQJBAJZWXqNdTFuRlaJg6op9w1nM7gMXZi2gszQIiuOlHaXTXN0xTK/R
dcXEZa/OEPJ3UF42FGwv4DTuCrOXFTdDg2sCQQCJW1wn40UESicxcx0t7vapWh2V
OJByILxwmykvq2N91GAnPlaPplRD7uA1K9zdtSHSgP9Q+B5rho4lh1NUpJZRAkAr
iMNLB19vPM9aADqq9BQ30vIxjvsVx21dagPePBhDxtsjan1MhJlYNbFEoaWisQ5i
2cI8OfjxGuWab+vC3xgVAkEApf28JQ6kJWQ4fDqXwERR/gnEeEqURrNaoXU2/XXw
1lKE36YblUuyAa5eu2PI5wnXRNHwOqRpjpfeXAu9CAS26A==
-----END RSA PRIVATE KEY-----

-----BEGIN PUBLIC KEY-----
MIGfMA0GCSqGSIb3DQEBAQUAA4GNADCBiQKBgQCMXZYNnZ1PMyRbUN6d8L087AcC
9beuiegSUjASS/o0rcs3y0ImV16PNJSjys7E1abHn3Dk7JkXWWZusTSbML2n/X4h
Bng4MCLP8cOWSqxmOMKv2B+0GVTkVGOo9WvpA0s1B/dMZKfe/sxbswDWpKyIJYL3
NPTz/xijb96yRjT6MwIDAQAB
-----END PUBLIC KEY-----
 
  */
 var signature = Utilities.computeRsaSha256Signature("this is my input",
     "-----BEGIN PRIVATE KEY-----\nMIICeAIBADANBgkqhkiG9w0BAQEFAASCAmIwggJeAgEAAoGBAIxdlg2dnU8zJFtQ3p3wvTzsBwL1t66J6BJSMBJL+jStyzfLQiZXXo80lKPKzsTVpsefcOTsmRdZZm6xNJswvaf9fiEGeDgwIs/xw5ZKrGY4wq/YH7QZVORUY6j1a+kDSzUH90xkp97+zFuzANakrIglgvc09PP/GKNv3rJGNPozAgMBAAECgYEAhh/bBVGpgods4zDoN+VU/kTPAl4HUWmHbxJHywP8pp+3Aq5fnEF6Ir7slYJyh5VcE33gI2+NbclIXtdHXE97uPksvyj0yngGjoEyuWgyJMEL/jKS5F6jsHzdBN+j4MJen4JCG+PoxvLBCt+gSO5dXvEpgTAlm0oREh9cmO0I7ZECQQDvBQbVPbZJfJ5EBBGuG2koCpi3n5I6cAl4kV2eO0RLQkKOfQc7dIvElfMO6ewwktGLdYKyw+Tj0zQb0xbaMV5ZAkEAllZeo11MW5GVomDqin3DWczuAxdmLaCzNAiK46UdpdNc3TFMr9F1xcRlr84Q8ndQXjYUbC/gNO4Ks5cVN0ODawJBAIlbXCfjRQRKJzFzHS3u9qlaHZU4kHIgvHCbKS+rY33UYCc+Vo+mVEPu4DUr3N21IdKA/1D4HmuGjiWHU1SkllECQCuIw0sHX288z1oAOqr0FDfS8jGO+xXHbV1qA948GEPG2yNqfUyEmVg1sUShpaKxDmLZwjw5+PEa5Zpv68LfGBUCQQCl/bwlDqQlZDh8OpfARFH+CcR4SpRGs1qhdTb9dfDWUoTfphuVS7IBrl67Y8jnCddE0fA6pGmOl95cC70IBLbo\n-----END PRIVATE KEY-----\n");
 Logger.log(signature);
 Logger.log(byteToHexString(signature));
 var test = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, "this is my input");
 Logger.log(byteToHexString(test));
}





function TESTFUNCTION() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Signature List").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range = s.getRange(1, 1, 1, 1);
  var names = range.getValues();
  var now = new Date();
  
  //Logger.log(ssIndex = ss.getSheetByName("Hours Log").getIndex());
  //s.getRange("A4" + 2).setValue("TESThhh");
  //Logger.log(s.getRange(1,1).getValue());
  //s.getRange(439,1).setValue("T"); // Conclusion: GAS can autogenerate rows via setValue but not via copyTo
  
}




// Sidebar Functionality
function onOpen() { // RESERVED FUNCTION NAME! Insert any function which must run on boot.
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .createMenu('Admin Tools')
      .addItem('Show sidebar', 'showSidebar')
      .addToUi();
  SpreadsheetApp.getUi().alert("You are " + getUserEmail() + "\nAdmin Tools Loaded!");
}

function getUserEmail() { // Necessary because Google's library function returns blank
  var userEmail = PropertiesService.getUserProperties().getProperty("userEmail");
  if(!userEmail) {
    var protection = SpreadsheetApp.getActive().getRange("A1").protect();
    // trick: the owner and user can not be removed
    protection.removeEditors(protection.getEditors());
    var editors = protection.getEditors();
    if(editors.length === 2) {
      var owner = SpreadsheetApp.getActive().getOwner();
      editors.splice(editors.indexOf(owner),1); // remove owner, take the user
    }
    userEmail = editors[0];
    protection.remove();
    // saving for better performance next run
    PropertiesService.getUserProperties().setProperty("userEmail",userEmail);
  }
  return userEmail;
}

function showSidebar() {
  var html = HtmlService.createHtmlOutputFromFile('Page')
      .setTitle('My custom sidebar')
      .setWidth(300);
  SpreadsheetApp.getUi() // Or DocumentApp or FormApp.
      .showSidebar(html);
}













// Time Tracking Stuff

//NOTE: Drive.Files.insert(metadata, mediaBody???, params)

function fromTest() {
  var form = FormApp.openById('1Iw_as2m87qkkVfSesLjTYMg5g0fgRmnBzN-C7orArWQ');
  var name = 'Form Link Test SS'
  var folderId = '15s-_b4MFv0qf2K7xiajXkmcIabGc4Yab'
  var metadata = {
    title: name,
    mimeType: MimeType.GOOGLE_SHEETS,
    parents: [{ id: folderId }]
  }
  var params = {
    supportsTeamDrives: true,
    includeTeamDriveItems: true
  };
  var fileJson = Drive.Files.insert(metadata, null, params)
  var fileId = fileJson.id
  form.setDestination(FormApp.DestinationType.SPREADSHEET, fileId);
}




function getTest() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range = s.getRange(4,3);
  Logger.log();
}

// Make me able to be triggered by check
function generateAttendanceSheet() { 
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Attendance Sheet Generator").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var generatorStatusBox = s.getRange(2,2);
  var title;
  var eventType;
  var startDate;
  var fileNameDate;
  var startTime;
  var endDate;
  var endTime;
  var dataTimer;
  var trackingType;
  var file;
  var docFile;
  
  var formattedStartDateTime;
  var formattedEndDateTime;
    
  
  if (s.getRange(16,2).getValue() == true) {
    generatorStatusBox.setBackground("#ffa500");
    generatorStatusBox.setValue("WORKING");
    
    title = s.getRange(4,2).getValue();
    eventType = s.getRange(5,2).getValue();
    
    startDate = s.getRange(7,2).getDisplayValue();
    fileNameDate = startDate;
    endDate = s.getRange(8,2).getDisplayValue();
    startTime = s.getRange(9,2).getDisplayValue();
    endTime = s.getRange(10,2).getDisplayValue();
    
    trackingType = s.getRange(12,2).getValue();
    
    dataTimer = s.getRange(14,2).getValue();
    
    Logger.log(getGenericFormattedDate(startDate, startTime));
    
    // File creation and switch to file
    file = DriveApp.getFileById("1QM8iy05ZL--6E_kFng4PSma2Vwitx3Q45XihakzK1c0").makeCopy(DriveApp.getFolderById("1hEF2Cq8kJ4zCFD56n7Xv8CUMIDBt6k9x"));
    file.setName("*(" + fileNameDate + ") " + title);
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(file.getId()));
    ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
    ssIndex = ss.getSheetByName("Attendance").getIndex();
    saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
    s = ss.getSheets()[saIndex];
    
    // Edit values
    formattedStartDateTime = getGenericFormattedDate(startDate, startTime);
    formattedEndDateTime = getGenericFormattedDate(endDate, endTime);
    Logger.log(formattedStartDateTime);
    
    s.getRange(1,1,1,8).setValues([[title,"","","","","",formattedStartDateTime,formattedEndDateTime]]);
    s.getRange(3,11,3,1).setValues([[eventType],[""],[trackingType]]);
    
    
    // Status Flash Sequence
    generatorStatusBox.setValue("DONE");
    generatorStatusBox.setBackground("#00ff00");
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    generatorStatusBox.setBackground(null);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    generatorStatusBox.setBackground("#00ff00");
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    generatorStatusBox.setBackground(null);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    generatorStatusBox.setBackground("#00ff00");
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    generatorStatusBox.setBackground(null);
    SpreadsheetApp.flush();
    Utilities.sleep(100);
    generatorStatusBox.setBackground("#00ff00");
    SpreadsheetApp.flush();
    
    Utilities.sleep(5000);
    generatorStatusBox.setValue("IDLE");
    generatorStatusBox.setBackground("#cccccc");
  }
  
}

function buttonGenerateAttendanceSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = SpreadsheetApp.getActiveSheet().getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  
  if (s.getSheetName() == "Attendance Sheet Generator") {
    if (s.getRange(16,2).getValue() == true) {
      generateAttendanceSheet();
      s.getRange(16,2).setValue(false);
    }
  }
}

function getGenericFormattedDate(date, time) {
  // Input formatting
  if (time.slice(0, time.lastIndexOf(":") + 3).length < 5) {
    time = "0" + time;
  }
  
  var zone = Session.getScriptTimeZone();
  var temp = Utilities.formatDate(new Date(),Session.getScriptTimeZone(),"Z");
  var hours = time.slice(0, time.lastIndexOf(":"));
  var output;
  
  if (time.slice(0, time.lastIndexOf(":") + 3).length < 5) {
    time = "0" + time;
  }
  
  if (time.slice(-2) == "PM")  {
    hours = ((parseInt(time.slice(0,2), 10)+ 12) % 12) + 12;
    time = time.slice(time.lastIndexOf(":"), time.lastIndexOf(":") + 3) + ":00";
    time = hours + time;
  } else if (time.slice(-2) == "AM") {

    if (hours == 12) {
      hours = "00";
    }
    time = time.slice(time.lastIndexOf(":"), time.lastIndexOf(":") + 3) + ":00";
    time = hours + time;
  }
  date = date + "T" + time;
  output = new Date(date);
  output = output.toString();
  output = output.slice(0, output.lastIndexOf("(") - 1);
  return output;
}

function rr() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Attendance Sheet Generator").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var date = "2018-10-01";
  var time = "2:00 AM";
  Logger.log(time.slice(0, time.lastIndexOf(":")));
  
}

function triggerTest() {
var file = DriveApp.getFileById("16yH5ZBosH1o3Lhk93ojwXExDd-h2W5AUx_leH9G0cn0");
SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(file.getId()));
var ss = SpreadsheetApp.openById(file.getId()); // ss means spreadsheet, or all the sheets combined
var ssIndex = ss.getSheetByName("Attendance").getIndex();
var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
var s = ss.getSheets()[saIndex];

Logger.log(ScriptApp.getUserTriggers(ss));
}



// Remote Form Functionality

function relayVHours() {
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("V Hours Radio").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastVHoursRadioed = parseInt(scriptProperties.getProperty("lastVHoursRadioed"));
  
  var data;
  var height = endRow - lastVHoursRadioed;
  
  if (lastVHoursRadioed < endRow) {
    data = s.getRange((lastVHoursRadioed + 1), 1, (endRow - lastVHoursRadioed), 7).getValues();
    
    lastVHoursRadioed = endRow;
    
    ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
    endRow = s.getLastRow();
    
    s.getRange((endRow + 1), 2, height, 7).setValues(data);
    
    applyStatus();
    
    scriptProperties.setProperty("lastVHoursRadioed", lastVHoursRadioed);
  }
}

function resetVHoursRadioSync() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty("lastVHoursRadioed");
  scriptProperties.setProperty("lastVHoursRadioed", 1);
}

function relayAdminVHours() {
  SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("V Hours Admin Radio").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var endRow = s.getLastRow();
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastAdminVHoursRadioed = parseInt(scriptProperties.getProperty("lastAdminVHoursRadioed"));
  
  var data;
  var height = endRow - lastAdminVHoursRadioed;
  
  if (lastAdminVHoursRadioed < endRow) {
    data = s.getRange((lastAdminVHoursRadioed + 1), 1, (endRow - lastAdminVHoursRadioed), 8).getValues();
    
    for (var i = 0; i < (endRow - lastAdminVHoursRadioed); i++) {
      data[i].splice(1,1);
      data[i][1] = data[i][1] + "@bergenfield.org";
    }
    
    
    lastAdminVHoursRadioed = endRow;
    
    ssIndex = ss.getSheetByName("Master Hours Log").getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
    endRow = s.getLastRow();
    
    s.getRange((endRow + 1), 2, height, 7).setValues(data);
    
    applyStatus();
    
    scriptProperties.setProperty("lastAdminVHoursRadioed", lastAdminVHoursRadioed);
  }
}

function resetAdminVHoursRadioSync() {
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty("lastAdminVHoursRadioed");
  scriptProperties.setProperty("lastAdminVHoursRadioed", 1);
}

// MANUAL SHARING REQUIRED. DO ONCE A YEAR. DO NOT DELETE SHETS IN TEMMIE W/O REGENERATING ID LIST
function createSharedStudentFolders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var currentEntry;
  var currentEntryValues;
  var endRow = s.getLastRow();
  var range;
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty('lastParticipationSyncedEntry');
  var entryEmail;
  var currentStudent;
  var studentName;
  var folder;
  var folderId;
  var ids = "";
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty("studentFileIds");
  var file;

  for (var i = 2; i <= endRow; i++) {
    range = s.getRange(i, 1);
    currentStudent = range.getValue();
    ssIndex = ss.getSheetByName(currentStudent).getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
    studentName = s.getName();
    folder = DriveApp.getFolderById("1KCqIpdfXmLd22qE24DC8ItOEcFpJ6ZsX").createFolder(studentName);
    file = DriveApp.getFileById("1NkrCAUsYjdYfFHYF_MjDUnRtD1wB7ZsBWD9N7ksME9s").makeCopy(DriveApp.getFolderById(folder.getId()));
    file.setName(studentName);
    if (i == endRow) {
      ids += file.getId();
    } else {
      ids += file.getId() + ",";
    }
    ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
    saIndex = ssIndex - 1;
    s = ss.getSheets()[saIndex];
  }
  scriptProperties.setProperty("studentFileIds", ids);
}

// NOTE! Add copied sheet then delete old sheet. Use s.copyTo(x). Ids stored in studentFileIds seperated by commas
function createSharedHoursLogs() {
  var ss = SpreadsheetApp.getActiveSpreadsheet(); // ss means spreadsheet, or all the sheets combined
  var ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex();
  var saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
  var s = ss.getSheets()[saIndex];
  var range;
  var currentEntry;
  var currentEntryValues;
  var endRow = s.getLastRow();
  var range;
  var scriptProperties = PropertiesService.getScriptProperties();
  var lastSyncedEntry = scriptProperties.getProperty('lastParticipationSyncedEntry');
  var entryEmail;
  var currentStudent;
  var studentName;
  var file;
  var scriptProperties = PropertiesService.getScriptProperties();
  var rawIds = scriptProperties.getProperty("studentFileIds");
  var idArray = rawIds.split(",");
  var dest;
  // var destination = SpreadsheetApp.openById('ID_GOES HERE');
  
  for (var i = 0; i < idArray.length; i++) {
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById("1PoRi2Up-s5AHFe9gZFnuaQPUFlcXQJ7NlDwPUzYiAto"));
    ss = SpreadsheetApp.getActiveSpreadsheet();
    ssIndex = ss.getSheetByName("Yearly Student Reference Sheet").getIndex() + i  + 1;
    saIndex = ssIndex - 1; // Spreadsheets are 1-indexed; sa = sheet actual
    s = ss.getSheets()[saIndex];
    dest = SpreadsheetApp.openById(idArray[i]);
    s.copyTo(dest);
    SpreadsheetApp.setActiveSpreadsheet(SpreadsheetApp.openById(idArray[i]));
    ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.deleteSheet(ss.getSheets()[0]);
    s = ss.getSheets()[0];
    s.setName(ss.getName());
  }
}

 
