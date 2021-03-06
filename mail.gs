/*Send email to all the emails in the array*/
function sendEmail(emailArr, sub, message){
    var len = emailArr.length;
    for(var i=0; i< len; i++){
        MailApp.sendEmail(emailArr[i], sub, message);
    }
}

/*Returns an array of email address from the Users sheet for the given role*/
function getMultipleEmail(role){
    var results = [];
    var userSheet = getSheet(SH_USERS, false);
    var emailColumn = getColumn(userSheet, HDR_EMAIL)
    var roleColumn = getColumn(userSheet, HDR_ROLES);
    var stColumn = getColumn(userSheet, HDR_STATUS_ACTIVE);
    
    var lastRow = userSheet.getLastRow();
    var startThreshold = 2;
    var range = userSheet.getRange(startThreshold, roleColumn, lastRow);
    var matchingCells = findContainingCells(role, range);
    for(var i =0; i<matchingCells.length; i++){
        var row = (matchingCells[i])[0];
        var emails = userSheet.getRange(row, emailColumn).getValue();
        var status = userSheet.getRange(row, stColumn).getValue();
        if (status == "Active")
            results.push(emails.split(",")[0]);
        else
            Logger.log("Ignoring InActive user:" + emails);
    }
    return results;
}

/*Returns a single email address belonging to the given name*/
//NOTE: returns the first email ID if there are multiple entries
function getEmail(initials){
    var userSheet = getSheet(SH_USERS, false);
    var emailColumn = getColumn(userSheet, HDR_EMAIL)
    var initialsColumn = getColumn(userSheet, HDR_INITIALS);
    var lastRow = userSheet.getLastRow();
    var startThreshold = 2;
    var range = userSheet.getRange(startThreshold, initialsColumn, lastRow);
    var matchingCell = findFirstCell(initials, range);
    var row = (matchingCell[0]);
    var emails = userSheet.getRange(row, emailColumn).getValue();
    /*returns the first email alone*/
    return emails.split(",")[0];
}

/*Returns email address of all experts*/
function getExpertEmail(){
    var roleId = getRoleId(ROLE_EXPERT);
    return getMultipleEmail(roleId);
}

/*Returns email address of all experts*/
function getTaskMasterEmail(){
    var roleId = getRoleId(ROLE_TASK_MASTER);
    return getMultipleEmail(roleId);
}

function getRoleId(role){
    var userSheet = getSheet(SH_USERS, false);
    var idCol = getColumn(userSheet, HDR_ID);
    var roleColumn = getColumn(userSheet, HDR_NAME);
    Logger.log(roleColumn);
    var matchingCell = findFirstCell(role, userSheet.getRange(2, roleColumn, userSheet.getLastRow()));
    var matchingRow = matchingCell[0];
    var roleId =  userSheet.getRange(matchingRow, idCol).getValue();
    return roleId;
}
