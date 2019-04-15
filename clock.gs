/*returns the latest time stamp of the specified Type(In/Out)*/
function getLatestClockActivity(initials, type){
    var clockSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_CLOCK);
    var lastRow = clockSheet.getDataRange().getLastRow();
    var lastCol = clockSheet.getDataRange().getLastColumn();
    var typeCol = getColumn(clockSheet, HDR_CLOCK_TYPE);
    var nameCol = getColumn(clockSheet, HDR_CLOCK_NAME);
    var initialsCol = getColumn(clockSheet, HDR_CLOCK_INITIALS);
    var timeCol = getColumn(clockSheet, HDR_CLOCK_TIME);
    for(var i=lastRow; i>0; i--){
        var rowValues = (clockSheet.getRange(i, nameCol, 1, lastCol).getValues())[0];
        var timeStamp = rowValues[timeCol-1];
        var initials_i = rowValues[initialsCol-1];
        var type_i = rowValues[typeCol-1];
        if(initials_i == initials && type_i == type){
            return timeStamp;
        } else
            continue;
    }
    return null;
}

/*set lock status to the given value*/
function setClockStatus(value, initials){
    if(initials == null)
        initials = getCurrentUserDetails()[0];
    var sheet = getSheet(SH_USERS, false);
    var cell = findFirstCell(initials, sheet.getRange(2, getColumn(sheet, HDR_INITIALS), sheet.getLastRow()));
    var row = cell[0];
    var col = getColumn(sheet, HDR_CLOCK_STATUS);
    var status = sheet.getRange(row, col).setValue(value);
}

/*Returns the current clock status*/
function getClockStatus(){
    var initials = getCurrentUserDetails()[0];
    var sheet = getSheet(SH_USERS, false);
    var cell = findFirstCell(initials, sheet.getRange(2, getColumn(sheet, HDR_INITIALS), sheet.getLastRow()));
    var row = cell[0];
    var col = getColumn(sheet, HDR_CLOCK_STATUS);
    var status = sheet.getRange(row, col).getValue();
    return status;
}

/*Checks if user is clocked in*/
function isClockedIn(initials){
    if(getClockStatus() === STATUS_CLOCK_IN)
        return true;
    else
        return false;
}

/*clock in the current user*/
function clockIn(triggeredBy){
    var currentTime = getTime();
    updateLastProgressTime();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SH_CLOCK);
    var lastRow = sheet.getLastRow();
    var u = getCurrentUserDetails();
    
    var values = [[u[1], u[0], STATUS_CLOCK_IN, triggeredBy, currentTime]];
    var range = sheet.getRange(lastRow + 1, 1,1,5);
    range.setValues(values);
    setClockStatus(STATUS_CLOCK_IN);
}

/*clock out the current user*/
function clockOut(triggeredBy, name, initials, time){
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SH_CLOCK);
    var lastRow = sheet.getLastRow();
    if(name==null || initials==null){
        var u = getCurrentUserDetails();
        name = u[1];
        initials = u[0];
    }
    if(time==null)
        time=getTime();
    var values = [[name, initials, STATUS_CLOCK_OUT, triggeredBy, time]];
    var range = sheet.getRange(lastRow + 1, 1,1,5);
    range.setValues(values);
    setClockStatus(STATUS_CLOCK_OUT, initials);
}

function getAutoClockOutTime(waitingTimeInMin){
    var currentTimeInMs = new Date().getTime();
    var waitingTimeInMs = waitingTimeInMin * 60 * 1000;
    return getTimeAt(currentTimeInMs - waitingTimeInMs/2);
}

function updateLastProgressTime(){
    var scriptProperties = PropertiesService.getScriptProperties();
    var initials = getCurrentUserDetails()[0];
    scriptProperties.setProperty(LAST_PROGRESS_TIME+":"+initials, getTime());
}

function setLastProgresstime(initials){
    var scriptProperties = PropertiesService.getScriptProperties();
    return scriptProperties.setProperty(LAST_PROGRESS_TIME+":"+initials, getTime());
}

function getLastProgressTime(initials){
    var scriptProperties = PropertiesService.getScriptProperties();
    return scriptProperties.getProperty(LAST_PROGRESS_TIME+":"+initials);
}
