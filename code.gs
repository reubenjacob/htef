/////////////////////////////////
//VARIABLES AND CONSTANTS
//Open Time	Start Time	Stuck Time	Completed Time	Defer Time	Last Status Update by	Last Reset time	Last Reset by
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

///////////////////////////////////////////////////
//User sheet headers
var HDR_INITIALS = ["Initials", -1];
var HDR_ROLES = ["Roles",-1];
var HDR_NAME = ["Name",-1];
var HDR_EMAIL = ["Email", -1];
var HDR_ID = ["ID", -1];
var ROLE_EXPERT = "Expert";
var HDR_CLOCK_STATUS = ["Clock Status", -1];

///////////////////////////////////////////////////////
//Log sheet headers
var HDR_LOG_TASKID = ["ID", -1];
var HDR_LOG_STATUS = ["Status", -1];
var HDR_LOG_INITIALS = ["Initials", -1];
var HDR_LOG_TIME = ["Time", -1];
var HDR_LOG_TIME_ZONE = ["Timezone", -1];

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
var STATUS_DEFER = "Deferred";
var STATUS_CANCEL = "Cancelled";
var STATUS_CLOCK_IN = "In";
var STATUS_CLOCK_OUT = "Out";

/*sheet names*/
var SH_MASTER = "T.Main";
var SH_USERS = "Users";
var SH_TASKS_USERS_PREFIX = "T.User.";
var SH_CLOCK = "Clock";
var SH_LOG = "Log";

var ID_TASK_ARTIFACT_FOLDER_ROOT = '0B2sB0DUvFeWQfkdIRnRtb0hoY2NWQ3YyWEs2WUdZNktQMjBEWW5PRW5XcmtzQy1ULVlyb1k';
var WIP_FOLDERNAME = "WIP-[task_id]";
var ID_RFI_TEMPLATE = "1MmVOSKtItcFktUP2WU9TK8IURyowHe-8H_VAx9Cl8zA";
var RFI_FILENAME = "RFI-[task_id]";
var ID_SPEC_TEMPLATE = "1i4T6a5_d1IiIp_OeRmqAPFO2LHHibuL4QM0e0F1kj5c";
var SPEC_FILENAME = "TaskSpec-[task_id]";
var ID_PREFIX = "JT.";

var SEC1_RESERVED_LAST = 11;
var DATE_FORMAT = "dd-MMM-yyyy HH:mm";

var LAST_PROGRESS_TIME = "LAST_PROGRESS_TIME";
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
// ******* START REGION: Triggers. These are event handlers for events from the googlespreadsheets
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


function onEdit(e) {
  
  /* event details */
  var sheets = e.source;
  var u = e.user; 
  var rng = e.range;
  var sh = rng.getSheet();
  var value = e.value;
  
  /* Extract the task ID from the query in ID column*/  
  var colID = getColumn(sh, HDR_TASKID);
  var colStart = rng.getColumn(); 
  var rowStart = rng.getRow(); 
  var taskId = sh.getRange(rowStart, colID).getValue();  
    
  /* If user is in main sheet */
  if (isMainSheet()){
    if(colStart == colID){
      SpreadsheetApp.getUi().alert("ID must not be entered manually. Please click on Open Task to generate task ID");
      //rng.clearContent();
      return;
    }
    var userInitials = sh.getRange(rowStart, getColumn(sh, HDR_ASSIGNEE)).getValue();
    var userSheet = getSheet(SH_TASKS_USERS_PREFIX+userInitials, false);
    if(userSheet != null){
      /*reflect update on section 2 of user sheet*/
      var subTaskCell = findFirstCell(taskId, userSheet.getRange(SEC1_RESERVED_LAST+1, 1, userSheet.getLastRow()));
      if(subTaskCell[0]<1)
        return;
      userSheet.getRange(subTaskCell[0], colStart).setValue(value);
    }
  }
  else if(isCurrentUserSheet()){
    /* If the user is editing in section 1 */
    if(isSectionOne()){
      SpreadsheetApp.getUi().alert("Not allowed to edit in Section 1. Please edit in section 2");
      e.range.clearContent();
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
}


function onTimer(e) {
//  Logger.log("timer fired");
//  sh = SpreadsheetApp.openById("1D_Z7d8caQDxWr2mmug_BhX6XTDengS0eeKT8WBYw-6I");
//  sh.getActiveSheet().getDataRange().setBackgroundRGB(255,0,0);
}

/*inserts menu items into the spreadsheet menu and declares menu callbacks*/
function onOpen() { 
  var menuTasks = SpreadsheetApp.getUi().createMenu('Jovian Tasks');
  
  menuTasks.addItem('Start Task', 'onStartTask');
  menuTasks.addItem('Complete Task', 'onCompleteTask');
  
  menuTasks.addSeparator();
  menuTasks.addItem('Create RFI Spec & Link', 'onCreateRFI');  
  menuTasks.addItem('Open RFI', 'onOpenRFI');
  menuTasks.addItem('Close RFI', 'onCloseRFI');
  
  menuTasks.addSeparator();
  menuTasks.addItem('Request Approval for Subtasks', 'onRequestSubTaskApproval');
  menuTasks.addItem('Approve Subtasks', 'onApproveSubTasks');
  menuTasks.addItem('Breakdown Subtasks', 'onBreakdownSubTasks');
  
  menuTasks.addSeparator();
  menuTasks.addItem('Create WIP Folder & Link', 'OnCreateWIPFolder');
  menuTasks.addItem('Create Task Spec & Link', 'OnCreateTaskInputSpec');
  
  menuTasks.addSeparator();
  menuTasks.addItem('Open Task', 'onOpenTask');
  menuTasks.addItem('Cancel Task', 'OnCancelTask');
  menuTasks.addItem('Defer Task', 'OnDeferTask');
  
  menuTasks.addToUi();
  
  SpreadsheetApp.getUi()
  .createMenu('Jovian Clock')
  .addItem('Clock In', 'onClockIn')   //.addSubMenu(SpreadsheetApp.getUi().createMenu('Clock')
  .addItem('Clock Out', 'onClockOut') //        .addItem('Clock In', 'onClockIn')
  .addToUi();
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
        var startTime = strToDate(lpTime.replace(' UTC',''));
        var endTime = strToDate(getTime().replace(' UTC',''));
        var timeTaken = endTime.getTime() - startTime.getTime();
        var minutesTaken = Math.floor(timeTaken/(60*1000));
        //if the task has taken more than 2 hours, clock out
        if(minutesTaken>=120) 
          clockOut("Auto", vals[nameCol-1], initials);
      }
    }
  }
}

function onStartTask() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot start task for somebody else. Please start the task from your task sheet [" + SH_TASKS_USERS_PREFIX + getCurrentUserDetails()[0] + "]");
    return;
  }
  /*If not in user sheet, alert the user*/
  if (!isCurrentUserSheet()) {
    SpreadsheetApp.getUi().alert("Cannot start task. Please start the task from your task sheet [" + SH_TASKS_USERS_PREFIX + getCurrentUserDetails()[0] + "]");
    return;
  }
  /*Automatic clock In*/
  if(!isClockedIn()){
    clockIn("Auto");
    SpreadsheetApp.getUi().alert("You have been Clocked in automatically");
  }
  /*status updates*/
  var arr = doStatusChangeSTART(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_STARTED_ON,  SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
  
  /*last progress time*/
  updateLastProgressTime();
  
  /*if (arr["ignored"] == "") 
    SpreadsheetApp.getActiveSpreadsheet().toast("You can now work on this task. If the task spec link is available for your task"+
                                                ", it contains important information about your task. Further, "+
                                               "remember to breakdown tasks assigned to you and to progress the task frequently"); 
                                               */
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
    var lastRow = sheet.getDataRange().getLastRow();
  
    /*ensure it is written beyond the reserved ranges of section 1*/
    if(lastRow < SEC1_RESERVED_LAST+1){
      lastRow = lastRow + (SEC1_RESERVED_LAST+1-lastRow);
    }
    
    /*clone of task record*/
    var insertedTask = sheet.getRange(lastRow + 2, 1, 1, sheet.getLastColumn());
    insertedTask.setBackground("#3c9959");
    insertedTask.setValues(taskvalues);
   
    /*sub task 1*/ //=CONCATENATE($A$20, "." ,ROW() - ROW($A$20))
    var insertedSubTask1 = sheet.getRange(sheet.getDataRange().getLastRow() + 1, 1, 1, sheet.getLastColumn());
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
    insertedSubTask1 = sheet.getRange(sheet.getDataRange().getLastRow() + 1, 1, 1, sheet.getLastColumn());
    taskvalues[0][colID - 1] = taskid + "." + 2;
    insertedSubTask1.setValues(taskvalues);
    
    /*sub task 3*/
    insertedSubTask1 = sheet.getRange(sheet.getDataRange().getLastRow() + 1, 1, 1, sheet.getLastColumn());
    taskvalues[0][colID - 1] = taskid + "." + 3;
    insertedSubTask1.setValues(taskvalues);
    
  }
  
}


function onCompleteTask() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot complete task for somebody else.");
    return;
  }
  if (!isCurrentUserSheet()) {
    SpreadsheetApp.getUi().alert("Cannot complete task. Please complete the task from your task sheet [" + SH_TASKS_USERS_PREFIX + getCurrentUserDetails()[0] + "]");
    return;
  }
  /*update status and time taken*/
  var time = getTime();
  var arr = doStatusChangeCOMPLETE(time, SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
  /*if (arr["ignored"] == "") 
    SpreadsheetApp.
    getActiveSpreadsheet().
    toast("Congratulations, you have completed your task! You are expected to have updated"+
          " all task artifacts and requested results to the Task Spec, or to the WIP folder");
          */
}

function onCreateRFI() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
    return;
  }
  createArtifacts(false, true);
  /*SpreadsheetApp.getActiveSpreadsheet().toast("RFI document created. Please open the document from column '" + HDR_RFI_SPEC + "' of the selected task(s). Please edit the document and then use the 'Open RFI' menu item to get the  RFI published");*/
}
    
function onOpenRFI() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
    return;
  }
  var arr; 
  /*fetch the assigner*/
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
  var assignerCol = getColumn(mainSheet, HDR_ASSIGNER);
  var idCol = getColumn(mainSheet, HDR_TASKID);
  /*iterate over each row in main sheet and update the status + send an email*/
  var mainTasks = getMainIdCellArr(SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange());
  for(var i = 0; i < mainTasks.length; i++){
    var row = (mainTasks[i])[0];
    var assigner = mainSheet.getRange(row, assignerCol).getValue();
    var taskId = mainSheet.getRange(row, idCol).getValue();
    var emailArr = getExpertEmail();
    var currUserName = (getCurrentUserDetails())[1];
    emailArr.push(getEmail(assigner));
    var message = "RFI raised by "+currUserName+" on task "+taskId;
    var sub = "RFI - "+taskId;
    arr = doStatusChange(STATUS_WAIT_RFI, HDR_ASSIGNEE, HDR_WAITING_FROM,  mainSheet, mainSheet.getRange(row, idCol), getTime());
    sendEmail(emailArr, sub, message);
  }
  /*if (arr["ignored"] == "") 
    SpreadsheetApp.getActiveSpreadsheet().toast("The RFI is now open. You are expected to have created and edited an"+
                                                " RFI document prior to Opening the RFI. Once the RFI is closed, the task "+
                                                "will be highlighted to notify you. Now, you may proceed to work on your next task.");
                                                */
}


function onCloseRFI() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
    return;
  }
  var arr = updateStatus(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_WAITING_FROM, "");
  /*if (arr["ignored"] == "") 
    SpreadsheetApp.getActiveSpreadsheet().toast("Thank you for your input. The RFI is now closed. You are expected "
                                                +"to have filled the the response section of the RFI document prior to closing the RFI");
                                                */
}

function onRequestSubTaskApproval() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
    return;
  }
   var arr = doStatusChange(STATUS_WAIT_STA, HDR_ASSIGNEE, HDR_WAITING_FROM, 
                            SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
}

function onApproveSubTasks() {
  if(otherUserSheet()){
    SpreadsheetApp.getUi().alert("Cannot do operation for somebody else.");
    return;
  }
   var arr = doStatusChange(STATUS_INPROGRESS, HDR_ASSIGNEE, HDR_WAITING_FROM, 
                            SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), "");
}

function OnCreateWIPFolder() {
  createArtifacts(false, false);
}

function OnCreateTaskInputSpec() {
  createArtifacts(true, false);
}

function onOpenTask() {
  /*if not on main sheet, alert!*/
  if(!isMainSheet()){
    SpreadsheetApp.getUi().alert("Task must be opened from main sheet");
    return;
  }
  var arr = doStatusChange(STATUS_OPEN, HDR_ASSIGNER, HDR_OPENED_ON,
                           SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
}

function onOpenSubtask() {
  //generateSubTaskIDs();
  //var arr = doStatusChange(STATUS_OPEN, HDR_ASSIGNER, HDR_OPENED_ON, SpreadsheetApp.getActiveSheet(), SpreadsheetApp.getActiveSheet().getActiveRange(), getTime());
}

function OnCancelTask() {
  updateStatus(STATUS_CANCEL, HDR_ASSIGNEE, null, ""); 
}

function OnDeferTask() {
  updateStatus(STATUS_DEFER, HDR_ASSIGNEE, null, ""); 
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
}

/*Important function: does the actual setting of status/time/duration value(s)*/
function doStatusChange(statusValue, ByCol, OnCol, sheet, rng, OnColValue) {
  var numRows = rng.getNumRows();
  var startRow = rng.getRow();
  var info = [];
  var ignored = "";
  var initials = getCurrentUserDetails()[0];
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

    /*set deafaults*/
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
    
    /*update sheet log*/
    updateLog(taskId, initials, statusValue);
    
    /*set the values*/
    if(colStatus !=null)
      sheet.getRange(i, colStatus).setValue(statusValue);
    if (colBy != null)
      sheet.getRange(i, colBy).setValue(initials);
    if (colOn != null)
      sheet.getRange(i, colOn).setValue(OnColValue);
    info.push(["rows", i]);
  }
 
  if (ignored != "") 
    SpreadsheetApp.getActiveSpreadsheet().toast("Cannot change status: Description is empty for row(s) " + ignored);
  info["ignored"] = ignored;
  return info;
}    


/*important function: does the actual setting of the status, time and initials*/
function doStatusChangeSTART(statusValue, ByCol, OnCol, sheet, rng, tm) {

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
    initiateTWT(taskid, getCurrentUserDetails()[0]);
    
    var cellTaskIDMain = findFirstCell(taskid, sheet_main.getRange(2, colIDMain, sheet_main.getLastRow(), 1));
    
     if(cellTaskIDMain[0] < 0){
      doStatusChange(statusValue, ByCol, OnCol, sheet, rng, tm);
      continue;
    }
    
    /*if subtasks have already been generated, skip inserting again*/
    var subTaskCell = findFirstCell(taskid, sheet.getRange(SEC1_RESERVED_LAST+1, colID, sheet.getLastRow()));
    if(subTaskCell[0]>0)
      continue;
      
    /*****Section-2 sub task insert****/
    var taskvalues = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues();
    var lastRow = sheet.getDataRange().getLastRow();
  
    /*ensure it is written beyond the reserved ranges of section 1*/
    if(lastRow < SEC1_RESERVED_LAST+1){
      lastRow = lastRow + (SEC1_RESERVED_LAST+1-lastRow);
    }
    
    /*clone of task record*/
    var insertedTask = sheet.getRange(lastRow + 2, 1, 1, sheet.getLastColumn());
    insertedTask.setBackground("#3c9959");
    insertedTask.setValues(taskvalues);
   
    /*sub task 1*/ //=CONCATENATE($A$20, "." ,ROW() - ROW($A$20))
    var insertedSubTask1 = sheet.getRange(sheet.getDataRange().getLastRow() + 1, 1, 1, sheet.getLastColumn());
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
    insertedSubTask1 = sheet.getRange(sheet.getDataRange().getLastRow() + 1, 1, 1, sheet.getLastColumn());
    taskvalues[0][colID - 1] = taskid + "." + 2;
    insertedSubTask1.setValues(taskvalues);
    
    /*sub task 3*/
    insertedSubTask1 = sheet.getRange(sheet.getDataRange().getLastRow() + 1, 1, 1, sheet.getLastColumn());
    taskvalues[0][colID - 1] = taskid + "." + 3;
    insertedSubTask1.setValues(taskvalues);
    
    /*status update on main sheet*/
    doStatusChange(statusValue,  ByCol, OnCol, sheet_main, sheet_main.getRange(cellTaskIDMain[0], cellTaskIDMain[1]), tm );
    
  }
}    

/* Updates the time taken column for a task on main sheet*/
function doStatusChangeCOMPLETE(time, sheet, range){
  var mainSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SH_MASTER);
  var numRows = range.getNumRows();
  var startRow = range.getRow();
  var timeTakenCol = getColumn(sheet, HDR_TIME_TAKEN);
  var idCol = getColumn(sheet, HDR_TASKID);
  var mainIdCol = getColumn(mainSheet, HDR_TASKID);
  var startDateCol = getColumn(mainSheet, HDR_STARTED_ON);
  var info = [];
  var result = [];
  var ignored = "";
  var userInitials = sheet.getRange(startRow, getColumn(sheet, HDR_ASSIGNEE)).getValue();
  var userSheet = getSheet(SH_TASKS_USERS_PREFIX+userInitials, false);
  var rowArray = getMainIdCellArr(sheet, range);
  updateAllTWT(userInitials);
  
  /*iterate over the tasks and compute all the time taken*/
  for (var i = startRow; i < startRow + numRows; i++){
    var taskid = sheet.getRange(i , idCol).getValue();
    var clockOutDuration = getTWT(taskid, userInitials);
    removeTWT(taskid, userInitials);
    var startDateStr = sheet.getRange(i, startDateCol).getValue();
    var startDate = strToDate(startDateStr.replace(' UTC',''));
    var endDate = strToDate(time.replace(' UTC',''));
    var timeTaken = endDate.getTime() - startDate.getTime() - (clockOutDuration==null?0:parseInt(clockOutDuration));
    var minutesTaken = Math.floor(timeTaken/(60*1000));  
    
    result.push(taskid+"|"+minutesTaken+"|"+i);
  }
  
  var display = "";
  /*iterate over the results*/
  for(var j = 0; j < result.length; j++){
    var i = result[j].split("|");
    var timeTakenStr = i[1];
    var taskid = i[0];
    var row = i[2];
    var cellTaskIDMain = findFirstCell(taskid, mainSheet.getRange(2, idCol, mainSheet.getLastRow(), 1));
    var rowMain = cellTaskIDMain[0];
    
    /*if it is a sub task*/
    if(rowMain < 0){
      doStatusChange(STATUS_COMPLETED, null, HDR_COMPLETED_ON, sheet, sheet.getRange(row, timeTakenCol), time);
      doStatusChange(null, null, HDR_TIME_TAKEN, sheet, sheet.getRange(row, timeTakenCol), timeTakenStr);
    } else{
      /*skip if description is missing*/
      if (!isDescSet(rowMain, mainSheet)){
        ignored += "[" + rowMain + "] "; 
        continue;
      }
      doStatusChange(STATUS_COMPLETED, null, HDR_COMPLETED_ON, mainSheet, mainSheet.getRange(cellTaskIDMain[0], timeTakenCol), time);
      doStatusChange(null, null, HDR_TIME_TAKEN, mainSheet, mainSheet.getRange(cellTaskIDMain[0], timeTakenCol), timeTakenStr);
    }
    display += "Time taken for task "+taskid+" will be updated as "+timeTakenStr+"\n";
  }
  if(display != "") 
    SpreadsheetApp.getActive().toast(display);
  /*if (ignored != "") 
    SpreadsheetApp.getActiveSpreadsheet.toast("Cannot change status: Description is empty for row(s) " + ignored);*/
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
  var colFolder = getColumn(sheet, HDR_WIP_FOLDER);
  var colTaskSpec = bCreateTaskSpec === false ? null : getColumn(sheet, HDR_TASK_SPEC);
  var colRFIspec = bCreateRFIDoc  === false ? null : getColumn(sheet, HDR_RFI_SPEC);
  var userSheet = getCurrentUserSheet();
  
  var taskParentFolder = DriveApp.getFolderById(ID_TASK_ARTIFACT_FOLDER_ROOT);
  for (i = startRow; i < startRow + numRows; i++){
    setDefaultProps(i, sheet);
    
    /*check if WIP folder already exists*/
    var taskid = sheet.getRange(i, colID).getValue();
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
    /*create WIP folder if it doesnt exist*/
    if (taskFolder == null){
      taskFolder = taskParentFolder.createFolder(foldername);
      sheet.getRange(i, colFolder).setValue(taskFolder.getUrl()); 
      setSubTaskValue(taskid, colFolder, taskFolder.getUrl());
    }    
    /*task spec file*/
    var taskSpecFileName = getTaskSpecFileName(taskid);
    if (colTaskSpec !== null && FileExists(taskSpecFileName, taskFolder) == false) {
      var objTaskSpecTemplate = DriveApp.getFileById(ID_SPEC_TEMPLATE);
      var objTaskFile = objTaskSpecTemplate.makeCopy(taskSpecFileName, taskFolder);
      sheet.getRange(i, colTaskSpec).setValue(objTaskFile.getUrl()); 
      setSubTaskValue(taskid, colTaskSpec, objTaskFile.getUrl());
    } 
    
    var RFIfileName = getRFISpecFileName(taskid);
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
    }
    
  }
}   

// ******* END REGION: Utility Functions
////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
