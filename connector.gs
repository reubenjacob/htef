
/*START DISPOSABLE CODE*/
TEST_URL = 'http://52.74.157.228:7474';

function testPush(){
    var url= TEST_URL + '/db/data/',
    nodeUrl=url+'node';
    var headers = {
        "Accept": "application/json"
    };
    
    var options = {
        "method" : "post",
        "headers" : headers
    };
    
    var resp = UrlFetchApp.fetch(nodeUrl, options);
    
    Logger.log(resp.getResponseCode());
    Logger.log(resp.getAllHeaders());
    Logger.log(resp.getContentText());
    
}

var headers = {
    "Accept" : "application/json; charset=UTF-8",
};

function getOptions(){
    return {
        "headers" : headers,
        "contentType" : "application/json"
    };
}

function getPostOptions(){
    var options = getOptions()
    options["method"] = "post";
    return options;
}

function getGetOptions(){
    var options = getOptions()
    options["method"] = "get";
    return options;
}

function getPutOptions(){
    var options = getOptions()
    options["method"] = "put";
    return options;
}

function addLabel(createResp, label){
    var labelUrl = createResp.labels,
    payLoad = label,
    options = getPostOptions();
    //No need to stringify as this is a string
    options['payload'] = payLoad;
    UrlFetchApp.fetch(labelUrl, options);
}

function addProperties(createResp, id, status, description){
    var propertyUrl = createResp.properties,
    properties = {
        'ID' : id,
        'Status' : status,
        'Description' : description
    },
    payLoad = JSON.stringify(properties),
    options = getPutOptions();
    options['payload'] = payLoad;
    var resp = UrlFetchApp.fetch(propertyUrl, options);
}

function createNode(label, properties){
    var baseUrl = TEST_URL + '/db/data/node',
    options = getPostOptions()
    ;
    options['payload'] = JSON.stringify(properties);
    var httpResp = UrlFetchApp.fetch(baseUrl, options),
    jsonResp = JSON.parse(httpResp.getContentText());
    addLabel(jsonResp, '\"'+label+'\"');
    return jsonResp;
}

function addNodeToIndex(indexName, nodeUrl, key, value){
    var url = TEST_URL + '/db/data/index/node/'+indexName+'/';
    var payLoad = {
        "value" : value,
        "uri" : nodeUrl,
        "key" : key
    }, options = getPostOptions();
    options['payload'] = JSON.stringify(payLoad);
    var httpResp =  UrlFetchApp.fetch(url, options);
    return JSON.parse(httpResp.getContentText());
}

function getUserNode(initials){
    var userIndex = 'initials',
    baseUrl = TEST_URL + '/db/data/index/node/User/'+userIndex+'/',
    options = getGetOptions(),
    url = baseUrl + encodeURIComponent(initials);
    var httpResp = UrlFetchApp.fetch(url, options);
    var jsonArrResp = JSON.parse(httpResp.getContentText());
    if(jsonArrResp.length === 0){
        var userProperties = getUserDetailsAsJson(initials);
        var createJsonResp = createNode('User', userProperties);
        return addNodeToIndex('User', createJsonResp.self, userIndex, initials);
    }
    else{
        return jsonArrResp[0];
    }
}

function getProjectNode(projectCode){
    var projectIndex = 'code',
    baseUrl = TEST_URL + '/db/data/index/node/Project/'+projectIndex+'/',
    options = getGetOptions(),
    url = baseUrl + encodeURIComponent(projectCode);
    var httpResp = UrlFetchApp.fetch(url, options);
    var jsonArrResp = JSON.parse(httpResp.getContentText());
    if(jsonArrResp.length === 0){
        Logger.log(jsonArrResp);
        Logger.log("creating project " + projectCode)
        var projectProperties = getProjectDetailsAsJson(projectCode);
        var createJsonResp = createNode('Project', projectProperties);
        return addNodeToIndex('Project', createJsonResp.self, projectIndex, projectCode);
    }
    else{
        return jsonArrResp[0];
    }
}

function createTaskNode(properties){
    var taskIndex = 'id',
    baseUrl = TEST_URL + '/db/data/index/node/Task/' + taskIndex + '/',
    options = getGetOptions(),
    url = baseUrl + encodeURIComponent(properties.id);
    var httpResp = UrlFetchApp.fetch(url, options);
    var jsonArrResp = JSON.parse(httpResp.getContentText());
    if(jsonArrResp.length === 0){
        var createJsonResp = createNode('Task', properties);
        return addNodeToIndex('Task', createJsonResp.self, taskIndex, properties.id);
    }
    else{
        return null;
    }
}

function setAssigner(taskNode, userInitials){
    var userNode = getUserNode(userInitials);
    return setRelation(taskNode, userNode, 'ASSIGNED_BY');
}

function setAssignee(taskNode, userInitials){
    var userNode = getUserNode(userInitials);
    return setRelation(taskNode, userNode, 'ASSIGNED_TO');
}

function setProject(taskNode, projectCode){
    if(projectCode == "")
        return;
    var projectNode = getProjectNode(projectCode);
    return setRelation(taskNode, projectNode, 'TASK_OF');
}

function setRelation(startNode, endNode, relationType){
    var url = startNode.create_relationship,
    options = getPostOptions(),
    payLoad = {
        "to" : endNode.self,
        "type" : relationType
    };
    options['payload'] = JSON.stringify(payLoad);
    var httpResp = UrlFetchApp.fetch(url, options);
    return JSON.parse(httpResp.getContentText());
}

function pushDataToNeo4j(){
    var arr = fetchTaskData('JT.1');
    for(var i=0; i<arr.length; i++){
        var json = arr[i],
        id = json.ID,
        assignee = json.Assignee,
        assigner = json.Assigner,
        status = json.Status,
        project = json.Project,
        description = json['Task Description'];
        var createResp = createTaskNode({
            'id' : id,
            'status' : status,
            'description' : description
        });
        if(createResp != null){
            setAssigner(createResp, assigner);
            setAssignee(createResp, assignee);
            setProject(createResp, project);
        }
    }
}
/*END DISPOSABLE CODE*/

function fetchTaskData(startId) {
    
    var mainSheet = getSheet(SH_MASTER),
    lastRow = mainSheet.getLastRow(),
    lastColumn = mainSheet.getLastColumn(),
    data = mainSheet.getRange(2, 1, lastRow, lastColumn).getValues(),
    headers = getHeaders(),
    startIdNum = getTaskNumber(startId),
    resultArr = [];
    
    for(var i = 0; i < data.length; i++){
        var tId = data[i][getColumn(mainSheet, HDR_TASKID) - 1];
        if(getTaskNumber(tId) > startIdNum){
            var result = arrToJson(headers, data[i]);
            resultArr.push(result);
        }
    }
    
    return resultArr;
}

function arrToJson(headers, arr){
    var json = {};
    for(var i =0; i< arr.length; i++){
        json[headers[i]] = arr[i];
    }
    return json;
}

function getArrPos(HDR_COL){
    var mainSheet = getSheet(SH_MASTER)
    return getColumn(mainSheet, HDR_COL);
}


function getHeaders(){
    var mainSheet = getSheet(SH_MASTER),
    lastColumn = mainSheet.getLastColumn(),
    data = mainSheet.getRange(1, 1, 1, lastColumn).getValues();
    return data[0];
}

function getTaskNumber(taskId){
    if(taskId != '' && typeof taskId == 'string'){
        var splitId = taskId.split('.');
        if(splitId.length > 1){
            return parseInt(splitId[1]);
        }
    }
    return -1;
}
