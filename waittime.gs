/*
* <--INTRO--> 
* TWT (Task Wait time Tracker) - tracks the time in which a task waits.
* A task is said to be in wait state when it is in 'In Progress' state when the Asignee is Clocked Out.
*/
var TWT_PRE = "TWT"; 
var HDR_CLOCK_TYPE = ["Type (In/Out)", -1];
var HDR_CLOCK_TIME = ["Time (UTC)",-1];
var HDR_CLOCK_INITIALS = ["Initials",-1];
var HDR_CLOCK_NAME = ["Name", -1];


function initiateTWT(taskId, initials){
  setTWT(taskId, initials, 0);
}

function setTWT(taskId, initials, value){
  var key = getTWTKey(taskId, initials);
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty(key, value);
}

/*called to get the value of clock outs during task progressions*/
function getTWT(taskId, initials){
  var key = getTWTKey(taskId, initials);
  var scriptProperties = PropertiesService.getScriptProperties();
  return scriptProperties.getProperty(key);
}

/*called after task completion*/
function removeTWT(taskId, initials){
  var key = getTWTKey(taskId, initials);
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty(key);
}

/*returns date2 - date1*/
function dateDiff(date1, date2){
  var startTime = strToDate(date1.replace(' UTC',''));
  var endTime = strToDate(date2.replace(' UTC',''));
  var timeTaken = endTime.getTime() - startTime.getTime();
  return timeTaken;
}

function dateDiffInMin(date1, date2){
  return dateDiff(date1, date2)/(60*1000);
}

function getTWTKey(taskId, initials){
  return TWT_PRE + ":" + initials + ":" + taskId;
}

/*
* <--Functionality--> Function to update each task's TWT
* <--Method--> It iterates over all the tasks of the user and updates each one of them individually by 
* incrementing the existing wait time with the caluclated wait time (previously clocked out time stamp - currentTime)
*/
function updateAllTWT(initials, currentTime){
 
  /*determine the necessity to update the TWTs accroding to current state*/
  var currentClockStatus = getClockStatus();
  if(currentTime == null){
    /*clocked out means TWT needs to be updated*/
    if(currentClockStatus == STATUS_CLOCK_OUT)
      currentTime = getTime();
    else
      /*clocked in means TWT has already been updated*/
      return;
  }
  
  /*obtained the duration for which the user was clocked out*/
  var lastClockOut = getLatestClockActivity(initials, STATUS_CLOCK_OUT);
  
  /*if task doesnt have a TWT associated with it*/
  if(lastClockOut == null)
    return;
  
  /*calculate the clock out period and increment the existing wait period with it*/
  var clockOutPeriod = dateDiff(lastClockOut, currentTime);
  var scriptProperties = PropertiesService.getScriptProperties();
  var prefix = TWT_PRE + ":" + initials;
  var keyArray = scriptProperties.getKeys();
  for(var i = 0; i< keyArray.length; i++){
    var key = keyArray[i];
    if(key.indexOf(prefix)==0){
      var splitKey = key.split(":");
      var taskId = splitKey[2];
      var value = scriptProperties.getProperty(key);
      setTWT(taskId, initials, parseInt(value) + clockOutPeriod);
    }
  }
  
}

