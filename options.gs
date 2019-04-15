function jumpToLastRow(){
    var currentSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var newEmptyRow = currentSheet.getLastRow()+1;
    var descriptionCol = getColumn(currentSheet, HDR_DESCRIPTION);
    currentSheet.setActiveCell(currentSheet.getRange(newEmptyRow,descriptionCol));
}

function jumpToSubTaskBreakDown(){
    
    if(isMainSheet() || (isCurrentUserSheet() && isSectionOne())){
        
        //initialize
        var currentSheet = SpreadsheetApp.getActiveSheet();
        var userSheet = getCurrentUserSheet();
        var currentRow = SpreadsheetApp.getActiveRange().getRow();
        var colID = getColumn(currentSheet, HDR_TASKID);
        
        //look for matching breakdown
        var taskid = currentSheet.getRange(currentRow , colID).getValue();
        var subTaskCell =findFirstCell(taskid, userSheet.getRange(SEC1_RESERVED_LAST+1, colID, userSheet.getLastRow()));
        var targetRow = subTaskCell[0];
        
        //if a breakdown exists, jump to it
        if(targetRow > 0){
            var descriptionCol = getColumn(userSheet, HDR_DESCRIPTION);
            userSheet.setActiveCell(userSheet.getRange(targetRow,descriptionCol));
        }
        
    }
    
}

function jumpToMySheet(){
    var userSheet = getCurrentUserSheet();
    userSheet.setActiveCell("A1");
}
