var TENTATIVE_LAST_COLUMN = 25;

function onSetupUser() {
    if(!isUserListSheet())
        return;
    var currentRow = SpreadsheetApp.getActiveRange().getRow();
    var userInitials = SpreadsheetApp.getActiveSheet().getRange(currentRow, getColumn(getSheet(SH_USERS), HDR_INITIALS)).getValues()[0][0];
    Logger.log("Found userinitials:" + userInitials);
    var userSheet = getSheet(SH_TASKS_USERS_PREFIX+userInitials, true);
    clearSection1(userSheet);
    setUpHeader(userSheet);
    setUpSection1(userSheet);
}

function setUpHeader(activeSheet){
    var formula = '=query(T.Main!A1:AC1, "Select * ", 0)';
    var firstCell = activeSheet.getRange(1, getColumn(getSheet(SH_MASTER), HDR_TASKID));
    firstCell.setFormula(formula);
    var headerRange = activeSheet.getRange(1, 1, 1, 20);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#a4c2f4");
}

function setUpSection1(activeSheet){
    var userInitials = getInitialsFromSheet(activeSheet.getName());
    setSection1Query(activeSheet, userInitials);
    protectSection1(activeSheet, userInitials);
}

function clearSection1(activeSheet){
    var section1Range = activeSheet.getRange(1, 1 , SEC1_RESERVED_LAST, TENTATIVE_LAST_COLUMN);
    section1Range.clearContent();
}

function setSection1Query(activeSheet, userInitials){
    var query = "=iferror(query(T.Main!A2:AC, \"Select * Where A != '' AND ((H='' AND G='Open') OR (H='" + userInitials
    + "' AND (G='In Progress' OR G='Open'))) ORDER BY G, F LIMIT 1\", 0),\"No Workables taks, please check your backlog\")";
    var queryHolderCell = activeSheet.getRange(2, getColumn(getSheet(SH_MASTER), HDR_TASKID));
    queryHolderCell.setFormula(query);
    var queryHolderRange = activeSheet.getRange(2, getColumn(getSheet(SH_MASTER), HDR_TASKID), 1, activeSheet.getLastColumn());
    queryHolderRange.setFontColor('#ffffff');
    queryHolderRange.setBackground('#434343');
    activeSheet.setFrozenRows(2);
}

function protectSection1(activeSheet, userInitials){
    clearProtections(activeSheet);
    var section1Range = activeSheet.getRange(1, 1 , SEC1_RESERVED_LAST, TENTATIVE_LAST_COLUMN);
    var protectObj = section1Range.protect().setDescription(userInitials +" - Section 1");
    protectObj.removeEditors(protectObj.getEditors()).addEditors(getMultipleEmail('R5')).setDomainEdit(false);
}

function getInitialsFromSheet(sheetName){
    if(sheetName.indexOf(SH_TASKS_USERS_PREFIX) == 0)
        return sheetName.split('.')[2];
}

function onCleanTasks(){
    copyFile(ID_SELF, ID_ARCHIVAL_FOLDER, getTime()+'_HTEF-V3');
    deleteCompletedTasks();
    deleteCancelledTasks();
}

function deleteCompletedTasks(){
    var mainSheet = getSheet(SH_MASTER),
    lastRow = mainSheet.getLastRow(),
    statusCol = getColumn(mainSheet, HDR_STATUS),
    timeCompletedCol = getColumn(mainSheet, HDR_OPENED_ON), //NOT a bug! too lazy to change variable names throughout - NV
    lastCol = mainSheet.getLastColumn();
    
    for(var i=2; i<=lastRow; i++){
        var row = mainSheet.getRange(i, 1, 1, lastCol).getValues(),
        result = row[0];
        if (result[statusCol-1] === STATUS_DELETED){
            mainSheet.deleteRow(i);
            --i;
            --lastRow;
        }
        else if(result[statusCol-1] === STATUS_COMPLETED){
            var completedTimeStr = result [timeCompletedCol-1];
            if(completedTimeStr!==''){
                var completedTime = strToDate(completedTimeStr).getTime(),
                currentDate = new Date().getTime(),
                timeDiff = (currentDate - completedTime)/(1000*60*60*24);
                if(timeDiff > 15){
                    mainSheet.deleteRow(i);
                    --i;
                    --lastRow;
                }
            }
        }
    }
}

function deleteCancelledTasks(){
    var mainSheet = getSheet(SH_MASTER),
    lastRow = mainSheet.getLastRow(),
    statusCol = getColumn(mainSheet, HDR_STATUS),
    timeCompletedCol = getColumn(mainSheet, HDR_OPENED_ON), //NOT a bug! too lazy to change variable names throughout - NV
    lastCol = mainSheet.getLastColumn();
    
    for(var i=2; i<=lastRow; i++){
        var row = mainSheet.getRange(i, 1, 1, lastCol).getValues(),
        result = row[0];
        if (result[statusCol-1] === STATUS_DELETED){
            mainSheet.deleteRow(i);
            --i;
            --lastRow;
        }
        else if(result[statusCol-1] === STATUS_CANCEL){
            var completedTimeStr = result [timeCompletedCol-1];
            if(completedTimeStr!==''){
                var completedTime = strToDate(completedTimeStr).getTime(),
                currentDate = new Date().getTime(),
                timeDiff = (currentDate - completedTime)/(1000*60*60*24);
                if(timeDiff > 15){
                    mainSheet.deleteRow(i);
                    --i;
                    --lastRow;
                }
            }
        }
    }
}

function onCleanLog(){
    deleteLogs();
}

function deleteLogs(){
    var ui = SpreadsheetApp.getUi();
    var response = ui.prompt('Log cleanup', 'Till which report would you like to delete?', ui.ButtonSet.OK_CANCEL);
    var lastReport = response.getResponseText();
    if(response.getSelectedButton() != ui.Button.OK || lastReport == '')
        return;
    var logSheet = getSheet(SH_LOG);
    var max = logSheet.getLastRow();
    var count = 0;
    for(var i= 2; i < logSheet.getLastRow(); i++){
        var report = logSheet.getRange(i,6).getValue();
        if(report == lastReport)
            break;
        else
            count++;
    }
    if(count!=0)
        logSheet.deleteRows(2, count);
    
}

function clearProtections(sheet){
    var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.RANGE);
    for (var i = 0; i < protections.length; i++) {
        var protection = protections[i];
        if (protection.canEdit()) {
            protection.remove();
        }
    }
}

