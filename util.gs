var HDR_BL_ID = ['ID', -1];
var HDR_BL_Status = ['Status', -1];
var SH_BL = 'Backlog';
var STATUS_BL_CLOSED_COMPLETE = 'Closed-Complete';
var STATUS_BL_WORKABLE = 'Workable';

function updateBLISSStatus(blissId, projectCode, status){
    if(projectCode === '')
        return;
    var blSheet = getBacklogSheet(projectCode);
    var blRowRange = getMatchingRowRange(blSheet, HDR_BL_ID, blissId);
    if(blRowRange!=null){
        var blRow = blRowRange.getRow();
        blSheet.getRange(blRow, getColumn(blSheet, HDR_BL_Status)).setValue(status);
    }
}

function getBacklogSheet(projectCode){
    var projectSheet = getSheet(SH_PROJECTS);
    var projectCodeRow = getMatchingRowRange(projectSheet, HDR_PROJECT_CODE, projectCode);
    if(projectCodeRow == null)
        throw Error("No such project - "+ projectCode);
    var productColumn = getColumn(projectSheet, HDR_PRODUCT);
    var backlogCol = getColumn(projectSheet, HDR_BACKLOG);
    var projectCodeRowValues = projectCodeRow.getValues();
    var productCode = projectCodeRowValues[0][productColumn-1];
    var backlogFormula = projectSheet.getRange(projectCodeRow.getRow(), backlogCol).getFormula();
    if(backlogFormula === '')
        throw Error('Backlog entry is missing in projects sheet for ' + projectCode);
    var backlogID = backlogFormula.split('"')[1].split('/')[5];
    return SpreadsheetApp.openById(backlogID).getSheetByName(productCode);
}

function getMatchingRowRange(sheet, columnHeader, value){
    var column = getColumn(sheet, columnHeader);
    for(var i = 2; i <= sheet.getLastRow(); i++){
        var cellValue = sheet.getRange(i, column).getValue();
        Logger.log(cellValue);
        if(cellValue === value){
            return sheet.getRange(i, 1, 1, sheet.getLastColumn());
        }
    }
    return null;
}

function getActiveSheet(){
    return SpreadsheetApp.getActiveSheet();
}

function isUserSheet(){
    return SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().
    getName().indexOf(SH_TASKS_USERS_PREFIX) == 0;
}

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
    return isUserSheet() && !isOwnUserSheet();
}

/*verifying task id*/
function isValidTaskID(taskid) {
    if (taskid!=null && taskid.indexOf(ID_PREFIX) === 0)
        return true;
    return false;
}

/*creating task id*/
function createID(){
    if(isMainSheet()){
        var ID_KEY = "Next_ID";
        var stateSheet = getSheet('State');
        var counterCell = stateSheet.getRange(1,1);
        var counter = counterCell.getValue();
        var id = parseInt(counter);
        id = (id > 0) ? id + 1 : 1;
        counterCell.setValue(id);
        return id;
    } else
        return createSubTaskId();
}

function createSubTaskId(){
    var prevRowId = getActiveSheet().getRange((getActiveSheet().getActiveRange().getRow() - 1),
                                              getColumn(getActiveSheet(), HDR_TASKID)).getValue();
    if(isSubTaskId(prevRowId)){
        var splitPrevRowId = prevRowId.split('.');
        return splitPrevRowId[1] + '.' + (parseInt(splitPrevRowId[2])+1);
    }else{
        return '';
    }
    
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
        rng.setValue("5");
    
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

/*If the user in his/her sheet*/
function isUserListSheet(){
    return !(SpreadsheetApp.getActiveSheet().getName().indexOf(SH_USERS) != 0);
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

function getTimeAt(time){
    return Utilities.formatDate(new Date(time), "UTC" , DATE_FORMAT) + " UTC";
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

/*Returns the sheet of the current user*/
function getUserSheet(initials){
    return getSheet(SH_TASKS_USERS_PREFIX+initials, false);
}

/*returns sheet object with the given name, creates sheet if not present*/
function getSheet(name, addIfNotPresent){
    var sheets = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = sheets.getSheetByName(name);
    if (sheet != null)  return sheet;
    if (!addIfNotPresent) return null;
    sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet();
    sheet.setName(name);
    return sheet;
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

function isSubTaskId(id){
    return id.split('.').length > 2;
}

function isSubTaskOf(taskId, id){
    if(isSubTaskId(id)){
        var splitId = id.split('.');
        return taskId == splitId[0] + '.' +splitId[1];
    } else
        return false;
}

function deleteEmptySubtasks(id, sheet){
    var col = getColumn(sheet, HDR_TASKID);
    var range = sheet.getRange(SEC1_RESERVED_LAST+1, 1, sheet.getLastRow());
    var resultArr = findContainingCells(id, range);
    var toDelRowArr = [];
    for(var i = 0 ; i < resultArr.length; i++){
        var row = (resultArr[i])[0];
        var id_i = sheet.getRange(row, col).getValue();
        if(isSubTaskOf(id, id_i)){
            if(!isDescSet(row, sheet))
                toDelRowArr.push(id_i);
            else if(!isCompleted(row, sheet)){
                subTaskStatusChange(sheet, row);
            }
        }
    }
    while(toDelRowArr.length>0){
        var row = findFirstCell(toDelRowArr.pop(), range)[0];
        if(row > 0)
            sheet.deleteRow(row);
    }
}

function subTaskStatusChange(sheet, row){
    var range = sheet.getRange(row, sheet.getLastColumn(), 1);
    doStatusChangeCOMPLETE(getTime(), sheet, range);
}

function isCompleted(row, sheet){
    var col = getColumn(sheet, HDR_STATUS);
    var rng = sheet.getRange(row, col);
    if (rng.getValue() === STATUS_COMPLETED)
        return true;
    return false;
}

function isProgressed(row, sheet){
    var col = getColumn(sheet, HDR_STATUS);
    var rng = sheet.getRange(row, col);
    if (rng.getValue() === STATUS_INPROGRESS)
        return true;
    return false;
    
}

function getUserDetailsAsJson(initials){
    var userSheet = getSheet(SH_USERS, false);
    var emailColumn = getColumn(userSheet, HDR_EMAIL);
    var nameColumn = getColumn(userSheet, HDR_NAME);
    var roleColumn = getColumn(userSheet, HDR_ROLES);
    var initialsColumn = getColumn(userSheet, HDR_INITIALS);
    var lastRow = userSheet.getLastRow();
    var startThreshold = 2;
    var range = userSheet.getRange(startThreshold, initialsColumn, lastRow);
    var matchingCell = findFirstCell(initials, range);
    var row = (matchingCell[0]),
    email = userSheet.getRange(row, emailColumn).getValue(),
    name = userSheet.getRange(row, nameColumn).getValue(),
    roles = userSheet.getRange(row, roleColumn).getValue(),
    json = {
        'name' : name,
        'initials' : initials,
        'email' : email,
    };
    var allRoles = ['R1','R2','R3','R4'];
    for(var i = 0; i < allRoles.length; i++){
        if(roles.indexOf(allRoles[i]) >= 0)
            json[allRoles[i]] = true;
    }
    Logger.log(json);
    return json;
    
}

function getProjectDetailsAsJson(code){
    var projectSheet = getSheet(SH_PROJECTS, false),
    codeCol = getColumn(projectSheet, HDR_PROJECT_CODE),
    nameCol = getColumn(projectSheet, HDR_PROJECT_NAME),
    clientCol = getColumn(projectSheet, HDR_PROJECT_CLIENT);
    var lastRow = projectSheet.getLastRow();
    var startThreshold = 2;
    var range = projectSheet.getRange(startThreshold, codeCol, lastRow);
    var matchingCell = findFirstCell(code, range);
    var row = (matchingCell[0]);
    if(row <= 0)
        return {};
    var json = {
        'code' : code,
        'name' : projectSheet.getRange(row, nameCol).getValue(),
        'client' : projectSheet.getRange(row, clientCol).getValue()
    };
    Logger.log(json);
    //TODO fill in
    return json;
    
}

function copyFile(targetFileId, destnDirId, targetfileName){
    var file = DriveApp.getFileById(targetFileId),
    destn = DriveApp.getFolderById(destnDirId),
    targetFile = file.makeCopy(destn);
    if(targetfileName!=null)
        targetFile.setName(targetfileName);
}

function isSheetExists(sheetName){
    var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
    for(var i = 0; i < sheets.length; i++){
        if(sheets[i].getName() === sheetName){
            Logger.log(true);
            return true;
        }
    }
    Logger.log(false);
    return false;
}

function setSpecialSheet(){
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
    sheet.setTabColor(SH_SPECIAL_COLOR);
    
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_USERS);
    sheet.setTabColor(SH_SPECIAL_COLOR);
    
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CLOCK);
    sheet.setTabColor(SH_SPECIAL_COLOR);
    
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_PROJECTS);
    sheet.setTabColor(SH_SPECIAL_COLOR);
    
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_WAITING);
    sheet.setTabColor(SH_SPECIAL_COLOR);
    
    sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_LISTS);
    sheet.setTabColor(SH_SPECIAL_COLOR);
}

function getActiveBlissID(){
    return getActiveSheet().getRange(getActiveSheet().getActiveCell().getRow(), 17).getValue();
}

function getActiveProject(){
    return getActiveSheet().getRange(getActiveSheet().getActiveCell().getRow(), 2).getValue();
}

function syncHtefToBacklog(){
    Logger.log("called");
    if(!isMainSheet()){
        return;
    }
    
    var blissID = getActiveBlissID();
    var project = getActiveProject();
    var blSheetName = getBacklogSheet(project).getSheetName();
    var col = getActiveSheet().getActiveCell().getColumn();
    
    switch (col){
        case 5:
            syncDescHtefToBacklog(blissID, blSheetName);
            break;
        case 6:
            syncPriorityHtefToBacklog(blissID, blSheetName);
            break;
        case 8:
            syncAsigneeHtefToBacklog(blissID, blSheetName);
            break;
        default:
            Logger.log("syncHtefToBacklog: Active cell is not from Desc, Priority or Assignee");
            break;
    }
}

function syncDescHtefToBacklog(blissID, blSheetName){
    var blSpreadSheet = SpreadsheetApp.openById(ID_BACKLOG).getSheetByName(blSheetName);
    if (blissID != null){
        var val = getActiveSheet().getActiveCell().getValue();
        var row = findFirstCell(blissID, blSpreadSheet.getRange(1, 1, blSpreadSheet.getLastRow()))[0,0];
        blSpreadSheet.getRange(row, 3).setValue(val);
        Logger.log(val);
    }
}

function syncPriorityHtefToBacklog(blissID, blSheetName){
    var blSpreadSheet = SpreadsheetApp.openById(ID_BACKLOG).getSheetByName(blSheetName);
    if (blissID != null){
        var val = getActiveSheet().getActiveCell().getValue();
        var row = findFirstCell(blissID, blSpreadSheet.getRange(1, 1, blSpreadSheet.getLastRow()))[0,0];
        blSpreadSheet.getRange(row, 4).setValue(val);
        Logger.log(val);
    }
}

function syncAsigneeHtefToBacklog(blissID, blSheetName){
    var blSpreadSheet = SpreadsheetApp.openById(ID_BACKLOG).getSheetByName(blSheetName);
    if (blissID != null){
        var val = getActiveSheet().getActiveCell().getValue();
        var row = findFirstCell(blissID, blSpreadSheet.getRange(1, 1, blSpreadSheet.getLastRow()))[0,0];
        blSpreadSheet.getRange(row, 6).setValue(val);
        Logger.log(val);
    }
}


function syncArtifactToBacklog(blissID, blSheetName, atfHeader, atfValue){
    var blSpreadSheet = SpreadsheetApp.openById(ID_BACKLOG).getSheetByName(blSheetName);
    var row = findFirstCell(blissID, blSpreadSheet.getRange(1, 1, blSpreadSheet.getLastRow()))[0,0];
    Logger.log("row: "+row);
    
    if (blissID != null){
        if (atfHeader[0,0] === HDR_RFI_SPEC[0,0]){
            blSpreadSheet.getRange(row, 10).setValue(atfValue);
        }
        if (atfHeader[0,0] === HDR_TASK_SPEC[0,0]){
            blSpreadSheet.getRange(row, 9).setValue(atfValue);
        }
        if (atfHeader[0,0] === HDR_WIP_FOLDER[0,0]){
            blSpreadSheet.getRange(row, 8).setValue(atfValue);
        }
    }
}

/*Runtime Recorder functions */
function runtimeCountStop(start) {
    start = new Date();
    var props = PropertiesService.getScriptProperties();
    var currentRuntime = props.getProperty("runtimeCount");
    var stop = new Date();
    var newRuntime = Number(stop) - Number(start) + Number(currentRuntime);
    var setRuntime = {
    runtimeCount: newRuntime,
    }
    props.setProperties(setRuntime);
}

function recordRuntime(functionName) {
    
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = "Runtime";
    try {
        ss.getSheetByName("Runtime");
    } catch (e) {
        ss.insertSheet(sheetName);
    }
    var sheet = ss.getSheetByName("Runtime");
    var props = PropertiesService.getScriptProperties();
    var runtimeCount = props.getProperty("runtimeCount");
    var recordTime = new Date();
    
    sheet.appendRow([recordTime, runtimeCount,functionName]);
    props.deleteProperty("runtimeCount");
}
/*end runtime recorder functions*/
