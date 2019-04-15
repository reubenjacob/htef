/////////////////////////////////
//VARIABLES AND CONSTANTS
//Open Time Start Time  Stuck Time  Completed Time  Defer Time  Last Status Update by Last Reset time Last Reset by
var HDR_TASKID = ["ID", -1];

var HDR_PROJECT = ["Project", -1];
var HDR_TYPE = ["Type", -1];
var HDR_STATUS = ["Status", -1];
var HDR_DESCRIPTION = ["Task Description", -1];
var HDR_ASSIGNER = ["Assigner", -1];
var HDR_ASSIGNEE = ["Assignee", -1];
var HDR_PRIORITY = ["Priority", -1];
var HDR_OPENED_ON = ["Opened On", -1];
var HDR_STARTED_ON = ["Started On", -1];
var HDR_COMPLETED_ON = ["Completed On", -1];
var HDR_WAITING_FROM = ["Waiting from", -1];
var HDR_WIP_FOLDER = ["WIP Folder", -1];
var HDR_TASK_SPEC = ["Task Spec", -1];
var HDR_RFI_SPEC = ["RFI Spec", -1];
var HDR_TIME_TAKEN = ["Duration (min)",-1];
var HDR_BLISS = ["BL/ISS",-1];

///////////////////////////////////////////////////
//User sheet headers
var HDR_INITIALS = ["Initials", -1];
var HDR_ROLES = ["Roles",-1];
var HDR_STATUS_ACTIVE = ["Status",-1];
var HDR_NAME = ["Name",-1];
var HDR_EMAIL = ["Email", -1];
var HDR_ID = ["ID", -1];
var ROLE_EXPERT = "Expert";
var ROLE_TASK_MASTER = "Task Master"
var HDR_CLOCK_STATUS = ["Clock Status", -1];

///////////////////////////////////////////////////////
//Log sheet headers
var HDR_LOG_TASKID = ["ID", -1];
var HDR_LOG_STATUS = ["Status", -1];
var HDR_LOG_INITIALS = ["Initials", -1];
var HDR_LOG_TIME = ["Time", -1];
var HDR_LOG_TIME_ZONE = ["Timezone", -1];

//Project sheet headers
var HDR_PROJECT_CODE= ["Project Code", -1];
var HDR_PROJECT_NAME = ["Project Name", -1];
var HDR_PROJECT_CLIENT = ["Client Name", -1];
var HDR_BACKLOG = ["Product / Dept Backlog", -1];
var HDR_PLANNING_MOM = ["Planning Session MOM", -1];
var HDR_PRODUCT = ["Product", -1];

//has to be defined after all array elems are declared..
var TASK_HEADERS = [HDR_TASKID, HDR_PROJECT, HDR_TYPE, HDR_STATUS, HDR_DESCRIPTION, HDR_ASSIGNER, HDR_ASSIGNEE, HDR_PRIORITY, HDR_OPENED_ON, HDR_STARTED_ON, HDR_COMPLETED_ON, HDR_WAITING_FROM, HDR_WIP_FOLDER, HDR_TASK_SPEC, HDR_RFI_SPEC];
var SUB_TASK_HEADERS_TO_EMPTY = [HDR_STATUS, HDR_DESCRIPTION, HDR_OPENED_ON, HDR_STARTED_ON, HDR_COMPLETED_ON, HDR_WAITING_FROM, HDR_WIP_FOLDER, HDR_TASK_SPEC, HDR_RFI_SPEC];
var HDR_INITIALS = ["Initials", -1];

/*status*/
var STATUS_OPEN = "Open";
var STATUS_INPROGRESS = "In Progress";
var STATUS_WAIT_RFI = "Waiting-RFI";
var STATUS_WAIT_STA = "Waiting-STA";
var STATUS_COMPLETED = "Completed";
var STATUS_DEFER = "Waiting-Pause";
var STATUS_CANCEL = "Cancelled";
var STATUS_DELETED = "Deleted";
var STATUS_CLOCK_IN = "In";
var STATUS_CLOCK_OUT = "Out";

/*sheet names*/
var SH_MASTER = "T.Main";
var SH_USERS = "Users";
var SH_TASKS_USERS_PREFIX = "T.User.";
var SH_CLOCK = "Clock";
var SH_LOG = "Log";
var SH_PROJECTS = "Projects";
var SH_WAITING="T.Waiting";
var SH_LISTS="Lists";

var SH_SPECIAL_COLOR = "FFDE00";

var ID_SELF = '1jBEellzk45PZopydkz4TclV_vjqqzbJVo2H7kpcBhNw';
var ID_ARCHIVAL_FOLDER = '0B2sB0DUvFeWQbGdTWmxyS2lfdVk';
var ID_TASK_ARTIFACT_FOLDER_ROOT = '0B2sB0DUvFeWQZGhWZ2tXX3IzZlU';
var ID_RFI_TEMPLATE = '15N_JBWDgpZXkv-P3xH2x_yhWgDZvS6odpd9Ni5_TUtk';
var ID_SPEC_TEMPLATE = '1zHZAcu9YN7A7ECC5UQzYMLXD_q_9B5h3OjO899Gt2oE';
var ID_BACKLOG = '18FV21vyY8UmsJH0ufSAW62oAKKuRExQoPM3m_JqQTEw';

var WIP_FOLDERNAME = "WIP-[task_id]";
var RFI_FILENAME = "RFI - [task_id]";
var SPEC_FILENAME = "TaskSpec-[task_id]";
var ID_PREFIX = "T.";

var SEC1_RESERVED_LAST = 2;
var DATE_FORMAT = "dd-MMM-yyyy HH:mm";

var LAST_PROGRESS_TIME = "LAST_PROGRESS_TIME";
var CLOCKOUT_WAITING_TIME_IN_MIN = 120;

////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ******* START REGION: Triggers. These are event handlers for events from the googlespreadsheets
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function onEdit(e) {
    var runtimeCountStart = new Date();
    /* event details */
    var sheets = e.source;
    var u = e.user;
    var rng = e.range;
    var sh = rng.getSheet();
    //var value = e.value;
    var value = rng.getValue();
    
    /* Extract the task ID from the query in ID column*/
    var colID = getColumn(sh, HDR_TASKID);
    var colStart = rng.getColumn();
    var rowStart = rng.getRow();
    var taskId = sh.getRange(rowStart, colID).getValue();
    
    if(typeof value == 'undefined')
        value = '';
    
    
    /* If user is in main sheet */
    if (isMainSheet()){
        if(colStart == colID){
            SpreadsheetApp.getUi().alert("ID must not be entered manually. Please click on Open Task to generate task ID");
            rng.clearContent();
            return;
        }
        var userInitials = sh.getRange(rowStart, getColumn(sh, HDR_ASSIGNEE)).getValue();
        var userSheet = getSheet(SH_TASKS_USERS_PREFIX+userInitials, false);
        if(userSheet != null){
            /*reflect update on section 2 of user sheet*/
            if(isSubTaskId(taskId))
                return;
            var subTaskCell = findFirstCell(taskId, userSheet.getRange(SEC1_RESERVED_LAST+1, 1, userSheet.getLastRow()));
            if(subTaskCell[0]<1)
                return;
            userSheet.getRange(subTaskCell[0], colStart).setValue(value);
        }
    }
    else if(isCurrentUserSheet()){
        /* If the user is editing in section 1 */
        if(isSectionOne()){
            SpreadsheetApp.getUi().alert("Not allowed to edit in Section 1. Please undo your changes and edit in section 2");
            //e.range.clearContent();
            return;
        }
        
        /* If the User is in section 2 */
        /* Use the taskID to update the content on the Main sheet in the corresponding cell */
        var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
        var cellTaskIDMain = getMainIdCellArr(sh, SpreadsheetApp.getActiveSheet().getActiveRange());
        var mainRow = (cellTaskIDMain[0])[0];
        if(mainRow<1)
            return;
        mainSheet.getRange(mainRow, rng.getColumn()).setValue(value);
    }
    runtimeCountStop(runtimeCountStart);
}


/*inserts menu items into the spreadsheet menu and declares menu callbacks*/
function onOpen() {
    
    var menuTasks = SpreadsheetApp.getUi().createMenu('TIG Tasks');
    
    menuTasks.addItem('Start Task', 'onStartTask');
    menuTasks.addItem('Complete Task', 'onCompleteTask');
    
    menuTasks.addSeparator();
    menuTasks.addItem('Open RFI', 'onOpenRFI');
    menuTasks.addItem('Close RFI', 'onCloseRFI');
    
    menuTasks.addSeparator();
    menuTasks.addItem('Request Approval for Subtasks', 'onRequestSubTaskApproval');
    menuTasks.addItem('Approve Subtasks', 'onApproveSubTasks');
    menuTasks.addItem('Breakdown Subtasks', 'onBreakdownSubTasks');
    
    menuTasks.addSeparator();
    menuTasks.addItem('Create RFI Spec', 'onCreateRFI');
    menuTasks.addItem('Create Task Spec', 'OnCreateTaskInputSpec');
    menuTasks.addItem('Create WIP Folder', 'OnCreateWIPFolder');
    
    menuTasks.addSeparator();
    menuTasks.addItem('Open Task', 'onOpenTask');
    menuTasks.addItem('Cancel Task', 'OnCancelTask');
    menuTasks.addItem('Pause Task', 'OnDeferTask');
    
    menuTasks.addToUi();
    
    SpreadsheetApp.getUi()
    .createMenu('TIG Options')
    .addItem('Jump to last row', 'jumpToLastRow')
    .addItem('Jump to task breakdown', 'jumpToSubTaskBreakDown')
    .addItem('Jump to my sheet', 'jumpToMySheet')
    .addSeparator()
    .addItem('Clock In', 'onClockIn')
    .addItem('Clock Out', 'onClockOut')
    .addToUi();
    
    SpreadsheetApp.getUi()
    .createMenu('TIG Admin')
    .addItem('Add User Sheet1', 'onSetupUser')
    .addSeparator()
    .addItem('Clean Tasks', 'onCleanTasks')
    .addToUi();
    
    setSpecialSheet();
};

// ******* END REGION: Triggers
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


///////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ******* START REGION: Menu Callbacks
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
function onTest() {
    SpreadsheetApp.getActiveSpreadsheet().toast(SpreadsheetApp.getActiveSheet().getActiveCell().get);
}

function onClockIn() {
    if(isClockedIn())
        SpreadsheetApp.getUi().alert("Already Clocked in");
    else
        clockIn("User");
}

function onClockOut() {
    if(isClockedIn())
        clockOut("User");
    else
        SpreadsheetApp.getUi().alert("Not Clocked in");
}

function onIdle(){
    var userSheet = getSheet(SH_USERS,false);
    var lastColumn = userSheet.getLastColumn();
    var lastRow = userSheet.getLastRow();
    var idCol = getColumn(userSheet, HDR_ID);
    var clockStatusCol = getColumn(userSheet, HDR_CLOCK_STATUS);
    var initialsCol = getColumn(userSheet, HDR_INITIALS);
    var nameCol = getColumn(userSheet, HDR_NAME);
    for (var i = 1; i<=lastRow; i++){
        var vals = userSheet.getRange(i, idCol, 1, lastColumn).getValues()[0];
        var ind = vals[0].indexOf("U");
        var isUserEntry = (ind === 0);
        if(isUserEntry){
            var isUserClockedIn = (vals[clockStatusCol - 1].indexOf("In")===0);
            if(isUserClockedIn){
                var initials = vals[initialsCol - 1];
                var lpTime = getLastProgressTime(initials);
                if(lpTime==null){
                    /*for the case of first time users*/
                    setLastProgresstime(initials);
                    continue;
                }
                //  tickDurationClock(initials); //TODO: RJ review
                var startTime = strToDate(lpTime.replace(' UTC',''));
                var endTime = strToDate(getTime().replace(' UTC',''));
                var timeTaken = endTime.getTime() - startTime.getTime();
                var minutesTaken = Math.floor(timeTaken/(60*1000));
                //if last progress time was more than 2 hours ago, clock out
                if(minutesTaken >= CLOCKOUT_WAITING_TIME_IN_MIN)
                    clockOut("Auto", vals[nameCol-1], initials, getAutoClockOutTime(CLOCKOUT_WAITING_TIME_IN_MIN));
            }
        }
    }
}

function tickDurationClock(initials){
    var mainSheet = getSheet(SH_MASTER),
    lastRow = mainSheet.getLastRow(),
    assigneeCol= getColumn(mainSheet, HDR_ASSIGNEE),
    timeTakenCol = getColumn(mainSheet, HDR_TIME_TAKEN),
    statusCol = getColumn(mainSheet, HDR_STATUS),
    lastCol = mainSheet.getLastColumn();
    
    for(var i=2; i<=lastRow; i++){
        var row = mainSheet.getRange(i, 1, 1, lastCol).getValues(),
        result = row[0];
        if(result[assigneeCol-1] === initials && result[statusCol-1] === STATUS_INPROGRESS){
            var duration = result[timeTakenCol-1];
            if(duration === '')
                doStatusChange(null, null, HDR_TIME_TAKEN, mainSheet, mainSheet.getRange(i, timeTakenCol), 5, initials);
            else
                doStatusChange(null, null, HDR_TIME_TAKEN, mainSheet, mainSheet.getRange(i, timeTakenCol), duration+5, initials);
        }
    }
}

function onStartTask() {
    var runtimeCountStart = new Date();
    if(otherUserSheet() || !isCurrentUserSheet()) {
        SpreadsheetApp.getUi().alert("Cannot start task. Please start the task from your task sheet [" + SH_TASKS_USERS_PREFIX + getCurrentUserDetails()[0] + "]");
        return;
    }
    /*Automatic clock In*/
    if(!isClockedIn()){
        clockIn("Auto");
        SpreadsheetApp.getActiveSpreadsheet().toast("You have been Clocked in automatically");
    }
    /*status updates*/
    var arr = doStatusChangeSTART(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_STARTED_ON,  SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onStartTask');
}

function onBreakdownSubTasks() {
    
    var sheet = SpreadsheetApp.getActiveSheet();
    var rng = SpreadsheetApp.getActiveSheet().getActiveRange();
    //statusValue, ByCol, OnCol, sheet, rng, tm;
    
    var sheet_main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
    var colIDMain = getColumn(sheet_main, HDR_TASKID);
    var colID = getColumn(sheet, HDR_TASKID);
    var cHDR_DESCRIPTION = getColumn(sheet, HDR_DESCRIPTION);
    var cHDR_STATUS = getColumn(sheet, HDR_STATUS);
    var cHDR_OPENED_ON = getColumn(sheet, HDR_OPENED_ON);
    var cHDR_STARTED_ON = getColumn(sheet, HDR_STARTED_ON);
    var cHDR_COMPLETED_ON = getColumn(sheet, HDR_COMPLETED_ON);
    var cHDR_WAITING_FROM = getColumn(sheet, HDR_WAITING_FROM);
    var cHDR_WIP_FOLDER = getColumn(sheet, HDR_WIP_FOLDER);
    var cHDR_TASK_SPEC = getColumn(sheet, HDR_TASK_SPEC);
    var numRows = rng.getNumRows();
    var startRow = rng.getRow();
    
    /*loop over active range*/
    for (var i = startRow; i < startRow + numRows; i++){
        var taskid = sheet.getRange(i , colID).getValue();
        var cellTaskIDMain = findFirstCell(taskid, sheet_main.getRange(2, colIDMain, sheet_main.getLastRow(), 1));
        
        if(cellTaskIDMain[0] < 0){
            continue;
        }
        
        /*if subtasks have already been generated, skip inserting again*/
        var subTaskCell = findFirstCell(taskid, sheet.getRange(SEC1_RESERVED_LAST+1, colID, sheet.getLastRow()));
        if(subTaskCell[0]>0)
            continue;
        
        /*****Section-2 sub task insert****/
        var taskvalues = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues();
        
        //--changes here--
        sheet.insertRows(SEC1_RESERVED_LAST+1, 5);
        
        var lastRow = SEC1_RESERVED_LAST + 2;
        
        /*ensure it is written beyond the reserved ranges of section 1*/
        //if(lastRow < SEC1_RESERVED_LAST+1){
        //  lastRow = lastRow + (SEC1_RESERVED_LAST+1-lastRow);
        //}
        
        /*clone of task record*/
        
        
        var insertedTask = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
        insertedTask.setBackground("#E0F2F1");
        insertedTask.setValues(taskvalues);
        
        /*sub task 1*/ //=CONCATENATE($A$20, "." ,ROW() - ROW($A$20))
        var insertedSubTask1 = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());
        taskvalues[0][colID - 1] = taskid + "." + 1;
        taskvalues[0][cHDR_DESCRIPTION - 1] = "";
        taskvalues[0][cHDR_STATUS- 1] = "";
        taskvalues[0][cHDR_OPENED_ON- 1] = "";
        taskvalues[0][cHDR_STARTED_ON- 1] = "";
        taskvalues[0][cHDR_COMPLETED_ON - 1] = "";
        taskvalues[0][cHDR_WAITING_FROM- 1] = "";
        taskvalues[0][cHDR_WIP_FOLDER - 1] = "";
        taskvalues[0][cHDR_TASK_SPEC - 1] = "";
        insertedSubTask1.setValues(taskvalues);
        
        /*sub task 2*/
        insertedSubTask1 = sheet.getRange(lastRow + 2, 1, 1, sheet.getLastColumn());
        taskvalues[0][colID - 1] = taskid + "." + 2;
        insertedSubTask1.setValues(taskvalues);
        
        /*sub task 3*/
        insertedSubTask1 = sheet.getRange(lastRow + 3, 1, 1, sheet.getLastColumn());
        taskvalues[0][colID - 1] = taskid + "." + 3;
        insertedSubTask1.setValues(taskvalues);
        
    }
    
}


function onCompleteTask() {
    var runtimeCountStart = new Date();
    
    if(otherUserSheet() || !isCurrentUserSheet()) {
        SpreadsheetApp.getUi().alert("Cannot complete task. Please complete the task from your task sheet [" + SH_TASKS_USERS_PREFIX + getCurrentUserDetails()[0] + "]");
        return;
    }
    /*update status and time taken*/
    var time = getTime();
    var arr = doStatusChangeCOMPLETE(time, SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
    
    var sh = SpreadsheetApp.getActiveSheet();
    var rng = SpreadsheetApp.getActiveSheet().getActiveRange();
    var colID = getColumn(sh, HDR_TASKID);
    var rowStart = rng.getRow();
    var taskId = sh.getRange(rowStart, colID).getValue();
    
    //if not maintask id,
    if(isSubTaskId(taskId))
        return;
    
    var waitingSheet = getSheet(SH_WAITING);
    var lastRow = waitingSheet.getLastRow();
    var lastColumn = waitingSheet.getLastColumn();
    var statusColumn = getColumn(waitingSheet, HDR_STATUS);
    var idColumn = getColumn(waitingSheet, HDR_TASKID);
    var assigneeColumn = getColumn(waitingSheet, HDR_ASSIGNEE);
    var currentUserInitials = getCurrentUserDetails()[0];
    
    for(var r = 2; r <= lastRow; r++){
        var status = waitingSheet.getRange(r, statusColumn).getValue();
        var assignee = waitingSheet.getRange(r, assigneeColumn).getValue();
        if(status === STATUS_DEFER && assignee === currentUserInitials){
            var mainSheet = getSheet(SH_MASTER);
            var deferedTaskRange = waitingSheet.getRange(r, 1, 1, lastColumn);
            updateStatusInactiveRange(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_WAITING_FROM, "", waitingSheet, deferedTaskRange);
            break;
        }
    }
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onCompleteTask');
}

function onCreateRFI() {
    var runtimeCountStart = new Date();
    
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    createArtifacts(false, true);
    /*SpreadsheetApp.getActiveSpreadsheet().toast("RFI document created. Please open the document from column '" + HDR_RFI_SPEC + "' of the selected task(s). Please edit the document and then use the 'Open RFI' menu item to get the  RFI published");*/
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onCreateRFI');
}

function onOpenRFI() {
    var runtimeCountStart = new Date();
    
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    var arr;
    /*fetch the assigner*/
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
    var assignerCol = getColumn(mainSheet, HDR_ASSIGNER);
    var idCol = getColumn(mainSheet, HDR_TASKID);
    var descCol = getColumn(mainSheet, HDR_DESCRIPTION);
    var rfiSpecCol = getColumn(mainSheet, HDR_RFI_SPEC);
    /*iterate over each row in main sheet and update the status + send an email*/
    var mainTasks = getMainIdCellArr(SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
    for(var i = 0; i < mainTasks.length; i++){
        
        var row = (mainTasks[i])[0];
        var assigner = mainSheet.getRange(row, assignerCol).getValue();
        var taskId = mainSheet.getRange(row, idCol).getValue();
        var desc = mainSheet.getRange(row, descCol).getValue();
        var rfiSpec = mainSheet.getRange(row, rfiSpecCol).getValue();
        
        /*email recipients*/
        var emailArr = getExpertEmail();
        if(assigner !== '')
            emailArr.push(getEmail(assigner));
        emailArr.concat(getTaskMasterEmail());
        
        /*build email body & subject*/
        var currUserName = (getCurrentUserDetails())[1];
        var message = "RFI raised by " + currUserName+" on task " + taskId + ": " + desc + "\nLink: "+rfiSpec;
        var sub = "RFI - "+taskId;
        
        /*make status change to task*/
        arr = doStatusChange(STATUS_WAIT_RFI, HDR_ASSIGNEE, HDR_WAITING_FROM,  mainSheet, mainSheet.getRange(row, idCol), getTime());
        
        /*send mail*/
        sendEmail(emailArr, sub, message);
        
    }
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onOpenRFI');
}


function onCloseRFI() {
    var runtimeCountStart = new Date();
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    var arr = updateStatus(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_WAITING_FROM, '');
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onCloseRFI');
}

function onRequestSubTaskApproval() {
    var runtimeCountStart = new Date();
    
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    //var arr = doStatusChange(STATUS_WAIT_STA, HDR_ASSIGNEE, HDR_WAITING_FROM,
    //                       SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
    updateStatus(STATUS_WAIT_STA, HDR_ASSIGNEE, HDR_WAITING_FROM, getTime());
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onRequestSubTaskApproval');
}

function onApproveSubTasks() {
    var runtimeCountStart = new Date();
    
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    // var arr = doStatusChange(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_WAITING_FROM,
    //                         SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), "");
    updateStatus(STATUS_OPEN, null, HDR_WAITING_FROM, "");
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onApproveSubTasks');
}

function OnCreateWIPFolder() {
    var runtimeCountStart = new Date();
    
    createArtifacts(false, false);
    runtimeCountStop(runtimeCountStart);
    recordRuntime('OnCreateWIPFolder');
}

function OnCreateTaskInputSpec() {
    var runtimeCountStart = new Date();
    
    createArtifacts(true, false);
    runtimeCountStop(runtimeCountStart);
    recordRuntime('OnCreateTaskInputSpec');
}

function onOpenTask() {
    var runtimeCountStart = new Date();
    
    /*if not on main sheet, alert!*/
    if(!isMainSheet()){
        SpreadsheetApp.getUi().alert("Task must be opened from main sheet");
        return;
    }
    var arr = doStatusChange(STATUS_OPEN, HDR_ASSIGNER, HDR_OPENED_ON,
                             SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
    runtimeCountStop(runtimeCountStart);
    recordRuntime('onOpenTask');
}

function onOpenSubtask() {
    //generateSubTaskIDs();
    //var arr = doStatusChange(STATUS_OPEN, HDR_ASSIGNER, HDR_OPENED_ON, SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
}

function OnCancelTask() {
    var runtimeCountStart = new Date();
    
    if(!getCurrentUserDetails == "RJ") {
        if ((otherUserSheet() || !isCurrentUserSheet())) {
            SpreadsheetApp.getUi().alert("Cannot cancel task. Please -Cancel- the task from your task sheet [" + SH_TASKS_USERS_PREFIX + getCurrentUserDetails()[0] + "]");
            return;
        }
    }
    
    /*update status and time taken*/
    var time = getTime();
    var arr = doStatusChangeCANCEL(time, SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
    
    var sh = SpreadsheetApp.getActiveSheet();
    var rng = SpreadsheetApp.getActiveSheet().getActiveRange();
    var colID = getColumn(sh, HDR_TASKID);
    var rowStart = rng.getRow();
    var taskId = sh.getRange(rowStart, colID).getValue();
    
    //if not maintask id,
    if(isSubTaskId(taskId))
        return;
    
    var waitingSheet = getSheet(SH_WAITING);
    var lastRow = waitingSheet.getLastRow();
    var lastColumn = waitingSheet.getLastColumn();
    var statusColumn = getColumn(waitingSheet, HDR_STATUS);
    var idColumn = getColumn(waitingSheet, HDR_TASKID);
    var assigneeColumn = getColumn(waitingSheet, HDR_ASSIGNEE);
    var currentUserInitials = getCurrentUserDetails()[0];
    
    for(var r = 2; r <= lastRow-1; r++){
        var status = waitingSheet.getRange(r, statusColumn).getValue();
        var assignee = waitingSheet.getRange(r, assigneeColumn).getValue();
        if(status === STATUS_DEFER && assignee === currentUserInitials){
            var mainSheet = getSheet(SH_MASTER);
            var deferedTaskRange = waitingSheet.getRange(r, 1, 1, lastColumn);
            updateStatusInactiveRange(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_WAITING_FROM, "", waitingSheet, deferedTaskRange);
            break;
        }
    }
    runtimeCountStop(runtimeCountStart);
    recordRuntime('OnCancelTask');
}

function OnDeferTask() {
    var runtimeCountStart = new Date();
    updateStatus(STATUS_DEFER, HDR_ASSIGNEE, HDR_WAITING_FROM, getTime());
    runtimeCountStop(runtimeCountStart);
    recordRuntime('OnDeferTask');
}

// ******* END REGION: Menu Callbacks
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ******* START REGION: Utility Functions
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

/* update the status on main sheet and section 2 */
function updateStatus(statusValue, ByCol, OnCol, OnColValue){
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    var arr;
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
    var mainTasks = getMainIdCellArr(SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
    if((mainTasks[0])[0] < 0){
        doStatusChange(statusValue, ByCol, OnCol, SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), OnColValue);
        return;
    }
    for(var i = 0; i < mainTasks.length; i++){
        var row = (mainTasks[i])[0];
        var idCol = getColumn(mainSheet, HDR_TASKID);
        var assigneeCol = getColumn(mainSheet, HDR_ASSIGNEE);
        var initials = mainSheet.getRange(row, assigneeCol).getValue();
        if(initials == '')
            initials = null;
        arr = doStatusChange(statusValue, ByCol, OnCol, mainSheet, mainSheet.getRange(row, idCol), OnColValue, initials);
    }
    return arr;
}

function updateStatusInactiveRange(statusValue, ByCol, OnCol, OnColValue, sheet, range){
    var mainTasks = getMainIdCellArr(sheet, range);
    for(var i = 0; i < mainTasks.length; i++){
        var row = (mainTasks[i])[0];
        var idCol = getColumn(sheet, HDR_TASKID);
        var mainSheet = getSheet(SH_MASTER);
        arr = doStatusChange(statusValue, ByCol, OnCol, mainSheet, mainSheet.getRange(row, idCol), OnColValue);
    }
    return arr;
}

/* create artifacts and update the column values on main sheet and section 2*/
function createArtifacts(bCreateTaskSpec, bCreateRFIDoc){
    if(otherUserSheet()){
        SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
        return;
    }
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
    if(isOwnUserSheet()){
        var mainTasks = getMainIdCellArr(SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
        if((mainTasks[0])[0] < 0){
            SpreadsheetApp.getUi().alert("Cannot create subtask artifacts");
            return;
        }
        for(var i = 0; i < mainTasks.length; i++){
            var row = (mainTasks[i])[0];
            var idCol = getColumn(mainSheet, HDR_TASKID);
            doCreateTaskArtifacts(bCreateTaskSpec, bCreateRFIDoc, mainSheet, mainSheet.getRange(row, idCol));
        }
    } else{
        doCreateTaskArtifacts(bCreateTaskSpec, bCreateRFIDoc,
                              SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
    }
    
}

/*Important function: does the actual setting of status/time/duration value(s)*/
function doStatusChange(statusValue, ByCol, OnCol, sheet, rng, OnColValue, initials) {
    
    var numRows = rng.getNumRows();
    var startRow = rng.getRow();
    var info = [];
    var ignored = "";
    if(initials == null)
        initials = getCurrentUserDetails()[0];
    var colStatus = (statusValue == null || typeof(statusValue)==='undefined' )? null : getColumn(sheet, HDR_STATUS);
    var colBy = (ByCol == null || typeof(ByCol)==='undefined') ? null : getColumn(sheet, ByCol);
    var colOn = (OnCol == null || typeof(OnCol)==='undefined') ? null : getColumn(sheet, OnCol);
    var colID = getColumn(sheet, HDR_TASKID);
    var userSheet = getSheet(SH_TASKS_USERS_PREFIX + initials, false);
    
    for (i = startRow; i < startRow + numRows; i++){
        /*If task description is not set, ignore*/
        if (!isDescSet(i, sheet)){
            ignored += "[" + i + "] ";
            continue;
        }
        
        /*set defaults*/
        setDefaultProps(i, sheet);
        var taskId = sheet.getRange(i , colID).getValue();
        
        /*reflect update on section 2 of user sheet*/
        var subTaskCell = findFirstCell(taskId, userSheet.getRange(SEC1_RESERVED_LAST+1, colID, userSheet.getLastRow()));
        
        /*if the subtask exists in section 2*/
        if(subTaskCell[0] != -1){
            var subTaskRow = subTaskCell[0];
            if(colStatus !=null)
                userSheet.getRange(subTaskRow, colStatus).setValue(statusValue);
            if (colBy != null)
                userSheet.getRange(subTaskRow, colBy).setValue(initials);
            if (colOn != null)
                userSheet.getRange(subTaskRow, colOn).setValue(OnColValue);
        }
        
        /*set the values*/
        if(colStatus !=null)
            sheet.getRange(i, colStatus).setValue(statusValue);
        if (colBy != null && sheet.getRange(i, colBy).getValue()!='')
            sheet.getRange(i, colBy).setValue(initials);
        if (colOn != null)
            sheet.getRange(i, colOn).setValue(OnColValue);
        
        /*update sheet log*/
        if(statusValue!=null && statusValue.length>0)
            updateLog(taskId, initials, statusValue);
        
        info.push(["rows", i]);
    }
    
    if (ignored != "")
        SpreadsheetApp.getActiveSpreadsheet().toast("Cannot change status: Description is empty for row(s) " + ignored);
    info["ignored"] = ignored;
    return info;
}


/*important function: does the actual setting of the status, time and initials*/
function doStatusChangeSTART(statusValue, ByCol, OnCol, sheet, rng, tm) {
    
    //Logger.log("started status-change-start " + " | " + statusValue  + " | " +  ByCol  + " | " + OnCol  + " | " + sheet + " | " + rng + " | " + tm) ;
    
    var sheet_main = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
    var colIDMain = getColumn(sheet_main, HDR_TASKID);
    var colID = getColumn(sheet, HDR_TASKID);
    var cHDR_DESCRIPTION = getColumn(sheet, HDR_DESCRIPTION);
    var cHDR_STATUS = getColumn(sheet, HDR_STATUS);
    var cHDR_OPENED_ON = getColumn(sheet, HDR_OPENED_ON);
    var cHDR_STARTED_ON = getColumn(sheet, HDR_STARTED_ON);
    var cHDR_COMPLETED_ON = getColumn(sheet, HDR_COMPLETED_ON);
    var cHDR_WAITING_FROM = getColumn(sheet, HDR_WAITING_FROM);
    var cHDR_WIP_FOLDER = getColumn(sheet, HDR_WIP_FOLDER);
    var cHDR_TASK_SPEC = getColumn(sheet, HDR_TASK_SPEC);
    var numRows = rng.getNumRows();
    var startRow = rng.getRow();
    
    //Logger.log("initialized status-change-start");
    
    /*loop over active range*/
    for (var i = startRow; i < startRow + numRows; i++){
        var taskid = sheet.getRange(i , colID).getValue();
        
        //Logger.log("Looking for matching tasks in main sheet. Processing " + taskid);
        
        var cellTaskIDMain = findFirstCell(taskid, sheet_main.getRange(2, colIDMain, sheet_main.getLastRow(), 1));
        
        /*if it is a sub-task*/
        if(cellTaskIDMain[0] < 0){
            //Logger.log(taskid + " is a subtask, applying status change");
            doStatusChange(statusValue, ByCol, OnCol, sheet, rng, tm);
            continue;
        }
        
        /*if subtasks have already not been generated, then generate & insert sub tasks*/
        var subTaskCell = findFirstCell(taskid, sheet.getRange(SEC1_RESERVED_LAST+1, colID, sheet.getLastRow()));
        if(subTaskCell[0]<0){
            
            /*****Section-2 sub task insert****/
            var taskvalues = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues();
            //--changes here--
            sheet.insertRows(SEC1_RESERVED_LAST+1, 5);
            
            var lastRow = SEC1_RESERVED_LAST + 2;
            
            /*ensure it is written beyond the reserved ranges of section 1*/
            //if(lastRow < SEC1_RESERVED_LAST+1){
            //  lastRow = lastRow + (SEC1_RESERVED_LAST+1-lastRow);
            //}
            
            /*clone of task record*/
            
            
            var insertedTask = sheet.getRange(lastRow, 1, 1, sheet.getLastColumn());
            insertedTask.setBackground("#E0F2F1");
            insertedTask.setValues(taskvalues);
            
            /*sub task 1*/ //=CONCATENATE($A$20, "." ,ROW() - ROW($A$20))
            var insertedSubTask1 = sheet.getRange(lastRow + 1, 1, 1, sheet.getLastColumn());
            taskvalues[0][colID - 1] = taskid + "." + 1;
            taskvalues[0][cHDR_DESCRIPTION - 1] = "";
            taskvalues[0][cHDR_STATUS- 1] = "";
            taskvalues[0][cHDR_OPENED_ON- 1] = "";
            taskvalues[0][cHDR_STARTED_ON- 1] = "";
            taskvalues[0][cHDR_COMPLETED_ON - 1] = "";
            taskvalues[0][cHDR_WAITING_FROM- 1] = "";
            taskvalues[0][cHDR_WIP_FOLDER - 1] = "";
            taskvalues[0][cHDR_TASK_SPEC - 1] = "";
            insertedSubTask1.setValues(taskvalues);
            
            /*sub task 2*/
            insertedSubTask1 = sheet.getRange(lastRow + 2, 1, 1, sheet.getLastColumn());
            taskvalues[0][colID - 1] = taskid + "." + 2;
            insertedSubTask1.setValues(taskvalues);
            
            /*sub task 3*/
            insertedSubTask1 = sheet.getRange(lastRow + 3, 1, 1, sheet.getLastColumn());
            taskvalues[0][colID - 1] = taskid + "." + 3;
            insertedSubTask1.setValues(taskvalues);
        }
        /*status update on main sheet*/
        doStatusChange(statusValue,  ByCol, OnCol, sheet_main, sheet_main.getRange(cellTaskIDMain[0], cellTaskIDMain[1]), tm );
        
    }
}

/* Updates the time taken column for a task on main sheet*/
function doStatusChangeCOMPLETE(time, sheet, range){
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER),
    numRows = range.getNumRows(),
    startRow = range.getRow(),
    timeTakenCol = getColumn(sheet, HDR_TIME_TAKEN),
    idCol = getColumn(sheet, HDR_TASKID),
    blissCol = getColumn(sheet, HDR_BLISS),
    projectCol = getColumn(sheet, HDR_PROJECT),
    mainIdCol = getColumn(mainSheet, HDR_TASKID),
    startDateCol = getColumn(mainSheet, HDR_STARTED_ON),
    info = [],
    result = [],
    ignored = "",
    userInitials = sheet.getRange(startRow, getColumn(sheet, HDR_ASSIGNEE)).getValue(),
    userSheet = getSheet(SH_TASKS_USERS_PREFIX+userInitials, false),
    rowArray = getMainIdCellArr(sheet, range);
    
    /*fetch the task IDs in current position - as these maybe disttorted after change*/
    for (var i = startRow; i < startRow + numRows; i++){
        var taskid = sheet.getRange(i , idCol).getValue();
        var blissId = sheet.getRange(i, blissCol).getValue();
        var projectCode = sheet.getRange(i, projectCol).getValue();
        result.push(taskid + "|" + i + "|" + blissId + "|" + projectCode);
    }
    
    var display = "";
    for(var j = 0; j < result.length; j++){
        var i = result[j].split("|");
        var taskid = i[0];
        var row = i[1];
        var blissId = i[2];
        var projectCode = i[3];
        var cellTaskIDMain = findFirstCell(taskid, mainSheet.getRange(2, idCol, mainSheet.getLastRow(), 1));
        var rowMain = cellTaskIDMain[0];
        
        /*if it is a sub task*/
        if(rowMain < 0){
            doStatusChange(STATUS_COMPLETED, null, HDR_COMPLETED_ON, sheet, sheet.getRange(row, timeTakenCol), time);
        } else{
            /*skip if description is missing*/
            if (!isDescSet(rowMain, mainSheet)){
                ignored += "[" + rowMain + "] ";
                continue;
            }
            doStatusChange(STATUS_COMPLETED, null, HDR_COMPLETED_ON, mainSheet, mainSheet.getRange(cellTaskIDMain[0], timeTakenCol), time);
            deleteEmptySubtasks(taskid, sheet);
            updateBLISSStatus(blissId, projectCode, STATUS_BL_CLOSED_COMPLETE);
        }
    }
    info["ignored"] = ignored;
    return info;
}


/* Updates the time taken column for a task on main sheet*/
function doStatusChangeCANCEL(time, sheet, range){
    var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER),
    numRows = range.getNumRows(),
    startRow = range.getRow(),
    timeTakenCol = getColumn(sheet, HDR_TIME_TAKEN),
    idCol = getColumn(sheet, HDR_TASKID),
    blissCol = getColumn(sheet, HDR_BLISS),
    projectCol = getColumn(sheet, HDR_PROJECT),
    mainIdCol = getColumn(mainSheet, HDR_TASKID),
    startDateCol = getColumn(mainSheet, HDR_STARTED_ON),
    info = [],
    result = [],
    ignored = "",
    userInitials = sheet.getRange(startRow, getColumn(sheet, HDR_ASSIGNEE)).getValue(),
    userSheet = getSheet(SH_TASKS_USERS_PREFIX+userInitials, false),
    rowArray = getMainIdCellArr(sheet, range);
    
    /*fetch the task IDs in current position - as these maybe disttorted after change*/
    for (var i = startRow; i < startRow + numRows; i++){
        var taskid = sheet.getRange(i , idCol).getValue();
        var blissId = sheet.getRange(i, blissCol).getValue();
        var projectCode = sheet.getRange(i, projectCol).getValue();
        result.push(taskid + "|" + i + "|" + blissId + "|" + projectCode);
    }
    
    var display = "";
    for(var j = 0; j < result.length; j++){
        var i = result[j].split("|");
        var taskid = i[0];
        var row = i[1];
        var blissId = i[2];
        var projectCode = i[3];
        var cellTaskIDMain = findFirstCell(taskid, mainSheet.getRange(2, idCol, mainSheet.getLastRow(), 1));
        var rowMain = cellTaskIDMain[0];
        
        /*if it is a sub task*/
        if(rowMain < 0){
            doStatusChange(STATUS_CANCEL, null, HDR_COMPLETED_ON, sheet, sheet.getRange(row, timeTakenCol), time);
        } else{
            /*skip if description is missing*/
            if (!isDescSet(rowMain, mainSheet)){
                ignored += "[" + rowMain + "] ";
                continue;
            }
            doStatusChange(STATUS_CANCEL, null, HDR_COMPLETED_ON, mainSheet, mainSheet.getRange(cellTaskIDMain[0], timeTakenCol), time);
            deleteEmptySubtasks(taskid, sheet);
        }
        updateBLISSStatus(blissId, projectCode, STATUS_BL_WORKABLE);
    }
    info["ignored"] = ignored;
    return info;
}

/*important function: creates gdrive folders, spec and rfi files*/
function doCreateTaskArtifacts(bCreateTaskSpec, bCreateRFIDoc, sheet, rng) {
    
    if(sheet==null)
        sheet = SpreadsheetApp.getActiveSheet();
    
    var rng_sel;
    
    if(rng==null)
        rng_sel = sheet.getActiveRange();
    else
        rng_sel = rng;
    
    var numRows = rng_sel.getNumRows();
    var startRow = rng_sel.getRow();
    
    var colID = getColumn(sheet, HDR_TASKID);
    var colDesc = getColumn(sheet, HDR_DESCRIPTION);
    var colFolder = getColumn(sheet, HDR_WIP_FOLDER);
    var colTaskSpec = bCreateTaskSpec === false ? null : getColumn(sheet, HDR_TASK_SPEC);
    var colRFIspec = bCreateRFIDoc  === false ? null : getColumn(sheet, HDR_RFI_SPEC);
    var colBliss = getColumn(sheet, HDR_BLISS);
    var colProject = getColumn(sheet, HDR_PROJECT);
    var userSheet = getCurrentUserSheet();
    
    var taskParentFolder = DriveApp.getFolderById(ID_TASK_ARTIFACT_FOLDER_ROOT);
    for (i = startRow; i < startRow + numRows; i++){
        setDefaultProps(i, sheet);
        
        /*check if WIP folder already exists*/
        var taskid = sheet.getRange(i, colID).getValue();
        var description = sheet.getRange(i, colDesc).getValue();
        var foldername = getWIPFolderName(taskid);
        var subFolders = taskParentFolder.getFolders();
        var bExists = false;
        var taskFolder = null;
        while (subFolders.hasNext()) {
            var subFolder = subFolders.next();
            if (subFolder.getName() == foldername) {
                taskFolder = subFolder;
                break;
            }
        }
        
        var blissID = sheet.getRange(i, colBliss).getValue();
        var project = sheet.getRange(i, colProject).getValue();
        var blSheetName = getBacklogSheet(project).getSheetName();
        var value = "";
        
        /*create WIP folder if it doesnt exist*/
        if (taskFolder == null){
            taskFolder = taskParentFolder.createFolder(foldername);
            sheet.getRange(i, colFolder).setValue(taskFolder.getUrl());
            setSubTaskValue(taskid, colFolder, taskFolder.getUrl());
            
            value = sheet.getRange(i, colFolder).getValue();
            syncArtifactToBacklog(blissID, blSheetName, HDR_WIP_FOLDER, value);
        }
        /*task spec file*/
        var taskSpecFileName = getTaskSpecFileName(taskid);
        if (colTaskSpec !== null && FileExists(taskSpecFileName, taskFolder) == false) {
            var objTaskSpecTemplate = DriveApp.getFileById(ID_SPEC_TEMPLATE);
            var objTaskFile = objTaskSpecTemplate.makeCopy(taskSpecFileName, taskFolder);
            sheet.getRange(i, colTaskSpec).setValue(objTaskFile.getUrl());
            setSubTaskValue(taskid, colTaskSpec, objTaskFile.getUrl());
            
            value = sheet.getRange(i, colTaskSpec).getValue();
            syncArtifactToBacklog(blissID, blSheetName, HDR_TASK_SPEC, value);
        }
        
        var RFIfileName = getRFISpecFileName(description);
        if (colRFIspec !== null && FileExists(RFIfileName, taskFolder) == false) {
            var objRFISpecTemplate = DriveApp.getFileById(ID_RFI_TEMPLATE);
            var objRFIfile  = objRFISpecTemplate.makeCopy(RFIfileName, taskFolder);
            
            var rfiDoc = DocumentApp.openById(objRFIfile.getId());
            
            rngBody = rfiDoc.getBody();
            
            var arr_headers = String(sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]).split(",");
            var arr_attribs = String(sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()[0]).split(",");
            
            var arr_key_vals = [];
            
            for (var n = 0; n < arr_headers.length; n++)
                arr_key_vals[arr_headers[n]] = arr_attribs[n];
            
            for (var m = 0; m < TASK_HEADERS.length; m++)
                rngBody.replaceText("<" + TASK_HEADERS[m][0] + ">", arr_key_vals[TASK_HEADERS[m][0]]);
            
            sheet.getRange(i, colRFIspec).setValue(objRFIfile.getUrl());
            setSubTaskValue(taskid, colRFIspec, objRFIfile.getUrl());
            
            value = sheet.getRange(i, colRFIspec).getValue();
            syncArtifactToBacklog(blissID, blSheetName, HDR_RFI_SPEC, value);
        }
    }
}

// ******* END REGION: Utility Functions
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
