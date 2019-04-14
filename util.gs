/*update the log sheet*/
function updateLog(taskId, initials, status){
  var userTimeZone = CalendarApp.getTimeZone();
  var userTime = Utilities.formatDate(new Date(), CalendarApp.getTimeZone(), DATE_FORMAT);
  var logSheet = getSheet(SH_LOG, false);
  var lastRow = logSheet.getDataRange().getHeight()+1;
  var range = logSheet.getRange(lastRow, 1, 1, logSheet.getLastColumn());
  var rowValues = range.getValues();
  rowValues[0][getColumn(logSheet, HDR_LOG_TASKID)-1] = taskId;
  rowValues[0][getColumn(logSheet, HDR_LOG_INITIALS)-1]  = initials;
  rowValues[0][getColumn(logSheet, HDR_LOG_STATUS)-1]  = status;
  rowValues[0][getColumn(logSheet, HDR_LOG_TIME)-1]  = userTime;
  rowValues[0][getColumn(logSheet, HDR_LOG_TIME_ZONE)-1]  = userTimeZone;
  range.setValues(rowValues);
}

function otherUserSheet(){
  return !isOwnUserSheet() && !isMainSheet();
}

/*verifying task id*/
function isValidTaskID(taskid) {
  if (taskid.indexOf(ID_PREFIX) === 0)
    return true;
  return false;
}

/*creating task id*/
function createID(){
  var ID_KEY = "Next_ID";
  var p = PropertiesService.getDocumentProperties(); 
  var id = parseInt(p.getProperty(ID_KEY));
  id = (id > 0) ? id + 1 : 1;
  p.setProperty(ID_KEY, id);
  return id;
}

/*set default values for a task*/
function setDefaultProps(taskrow, sheet){
    var col = getColumn(sheet, HDR_TASKID);
     var rng = sheet.getRange(taskrow, col);
    if (rng.getValue() === "")
      rng.setValue(ID_PREFIX + createID());
    
    var col = getColumn(sheet, HDR_PRIORITY);
     var rng = sheet.getRange(taskrow, col);
    if (rng.getValue() === "")
      rng.setValue("2: M");
  
}

/*If the description is set*/
function isDescSet(taskrow, sheet){
    var col = getColumn(sheet, HDR_DESCRIPTION);
     var rng = sheet.getRange(taskrow, col);
    if (rng.getValue() === "")
      return false;
    
  return true;
}

/*If the user in his/her sheet*/
function isCurrentUserSheet(){
  return !(SpreadsheetApp.getActiveSheet().getName().indexOf(SH_TASKS_USERS_PREFIX) != 0);
}

function isOwnUserSheet(){
  var initials = getCurrentUserDetails()[0];
  return !(SpreadsheetApp.getActiveSheet().getName().indexOf(SH_TASKS_USERS_PREFIX+initials) != 0);
}

/*if the user is in main sheet*/
function isMainSheet(){
   return !(SpreadsheetApp.getActiveSheet().getName().indexOf(SH_MASTER) != 0);
}

/*if the user is in section 1*/
function isSectionOne(){
  return (SpreadsheetApp.getActiveSheet().getActiveRange().getLastRow() < SEC1_RESERVED_LAST+1);
}

/*Returns an array of row numbers corresponding the IDs of tasks in the given range*/
function getMainIdCellArr(sheet, activeRange){
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
  var firstRow = activeRange.getRow();
  var numRows = activeRange.getNumRows();
  var idCol = getColumn(mainSheet, HDR_TASKID);
  var resultArr = [];
  
  for(var i = firstRow; i < firstRow + numRows; i++){
    var taskIdValue = sheet.getRange(i , idCol).getValue();    
    //corresponding cell of task in main sheet
    var cellTaskIDMain = findFirstCell(taskIdValue, mainSheet.getRange(2, idCol, mainSheet.getLastRow(), 1));
    resultArr.push(cellTaskIDMain);
  }
  return resultArr;
}

function getWIPFolderName(taskid){
   return WIP_FOLDERNAME.replace("[task_id]", taskid);
}

function getTaskSpecFileName(taskid){
   return SPEC_FILENAME.replace("[task_id]", taskid);
}

function getRFISpecFileName(taskid){
   return RFI_FILENAME.replace("[task_id]", taskid);
}

/*returns a date representation of the date string in 'dd-MMM-yyyy HH:mm' format*/
function strToDate(dateStr){
  
  var months = {
    'Jan' : '00', 'Feb' : '01', 'Mar' : '02', 'Apr' : '03', 'May' : '04','Jun' : '05', 'Jul' : '06',
    'Aug' : '07', 'Sep' : '08', 'Oct' : '09', 'Nov' : '10','Dec' : '11' 
  };
  
  var dateParts = dateStr.split(' ');
  var date = dateParts[0];
  var time = dateParts[1];
 
  var splitDate = date.split('-');
  var dd = splitDate[0];
  var MMM = splitDate[1];
  var MM = months[MMM];
  var yyyy = splitDate[2];
  
  var splitTime = time.split(':');
  var HH = splitTime[0];
  var mm = splitTime[1];
  
  return new Date(Date.UTC(yyyy,MM,dd,HH,mm,00));
  
}

function FileExists(name, folder){
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName() == name) {
      return true;
    }
  }
  return false;
}

/*sets current time into the specified cell*/
function getTime(){
    return Utilities.formatDate(new Date(), "UTC" , DATE_FORMAT) + " UTC";
}

/*returns the two letter initials from the "Users" sheet, for the currently logged in user*/
function getCurrentUserDetails(){
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var sheet_user = sheets.getSheetByName(SH_USERS);
  
  
  var email = Session.getEffectiveUser().getEmail();
  var colInitials = getColumn(sheet_user, HDR_INITIALS);
  var colName = getColumn(sheet_user, HDR_NAME);
  
  var arrMatch = findFirstCell(email, sheet_user.getDataRange());

  if (arrMatch[0] == -1){
    Logger.log("Could not find email in Users sheet");
    return "Unset";
  } 

  var arr = [];
  arr.push((sheet_user.getRange(arrMatch[0], colInitials)).getValue());
  arr.push((sheet_user.getRange(arrMatch[0], colName)).getValue());
  arr.push(email);
  return arr;
}


///////////////////////////////////////////////////////////////////////////
//returns column number for the header name passed in
function getColumn(sheet, hdrArray){
  if (hdrArray[1] > 0)
    return hdrArray[1];
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  hdrArray[1] = (headers[0].indexOf(hdrArray[0])) + 1;
  return hdrArray[1];
}

///////////////////////////////////////////////////////////////////////////
//returns array with row and col of first match of containingText in the range
function findFirstCell(containingText, inRng){
  var rng = inRng;
  var data = rng.getValues();
  var startRow = rng.getRow();  var startCol = rng.getColumn();
  if(!(containingText==undefined || containingText==null || containingText=='')){
    containingText = containingText.toString().toLowerCase();
    for (var r = 0; r < data.length; r++){
      for (var c=0; c < data[0].length; c++){
        if (data[r][c]==''){ continue };
        if (data[r][c].toString().toLowerCase().indexOf(containingText) > -1){
          return [r+startRow, c+startCol];
        }
      }
    }
  }
  return [-1, -1];
}

/*returns array of array[row, coll] for the matches containing the containingText*/
function findContainingCells(containingText, inRng){
  var rng = inRng;
  var data = rng.getValues();
  var startRow = rng.getRow();  var startCol = rng.getColumn();
  var results = [];
  if(!(containingText==undefined || containingText==null || containingText=='')){
    containingText = containingText.toString().toLowerCase();
    for (var r = 0; r < data.length; r++){
      for (var c=0; c < data[0].length; c++){
        if (data[r][c]==''){ continue };
        if (data[r][c].toString().toLowerCase().indexOf(containingText) > -1){
          results.push([r+startRow, c+startCol]);
        }
      }
    }
  }
  return results;
}

/*Returns the sheet of the current user*/
function getCurrentUserSheet(){
  var initials = getCurrentUserDetails()[0];
  return getSheet(SH_TASKS_USERS_PREFIX+initials, false);
}

/*returns sheet object with the given name, creates sheet if not present*/
function getSheet(name, addIfNotPresent){
  var sheets = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = sheets.getSheetByName(name);
  if (sheet != null)  return sheet;
  if (!addIfNotPresent) return null;
  sheet = activeSpreadsheet.insertSheet();
  sheet.setName(name);
  return sheet;
}

function isSubTask(btaskID) {
  pos = subTaskID.indexOf(".");
  return pos != -1;
}

function getParentTaskID(subtaskID) {
  pos = subTaskID.indexOf(".");
  if (pos == -1)
    return "";
  return subTaskID.substr(0, pos);
}
  
function onStatusChanged(taskID, newValue, user, sheet){
  var master_sh = getSheet(SH_MASTER, false);
  var colStatus = getColumn(master_sh, HDR_STATUS);
  var colID = getColumn(master_sh, HDR_TASKID);
  var arrTaskIDCell = findFirstCell(taskID, master_sh.getRange(1, colID, master_sh.getLastRow()));
  master_sh.getRange(arrTaskIDCell[0], colStatus).setValue(newValue);            
} 
    
/*set the value of sub-tasks*/
function setSubTaskValue(taskId, column, value){
  var userSheet = getCurrentUserSheet();
  var colID = getColumn(userSheet, HDR_TASKID);
  var subTaskCell = findFirstCell(taskId, userSheet.getRange(SEC1_RESERVED_LAST+1, colID, userSheet.getLastRow()));
  var subTaskRow = subTaskCell[0];
  /*If subtask is not present*/
  if(subTaskRow < 0)
    return;
  userSheet.getRange(subTaskRow, column).setValue(value);
}
