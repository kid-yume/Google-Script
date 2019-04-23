/*
* This Script is intended to help develop ticket tracking and will be used to do so. 
*authored by Chrome 
*
*/


function addMenu() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Ticket Tracker Menu')
      .addItem('Refresh Open Github Tickets','callOpenGithubIssues')
      .addItem('Refresh Closed Github Tickets', 'callClosedGithubIssues')
      .addToUi();
}
function GetFormattedDate(x) {
  var todayTime = new Date(x);
  var year = todayTime.getFullYear();

  var month = (1 + todayTime.getMonth()).toString();
  month = month.length > 1 ? month : '0' + month;

  var day = todayTime.getDate().toString();
  day = day.length > 1 ? day : '0' + day;
  
  return month + '/' + day + '/' + year;

}

function callOpenGithubIssues() {
  var openIssues = [];
  var startOfCellTitle = "";
  var CellTitleRowNumber = 0;
  
  // Call the Github API for list of issues
  var data = {
  'Authorization': 'Basic ZWdnYm95eEBpY2xvdWQuY29tOmJhY29uYm95MQ==',
  };
  var options = {
  'method' : 'get',
  'contentType': 'application/json; charset=utf-8',
  'headers' : data
  };
  var response = UrlFetchApp.fetch("https://api.github.com/repos/google/loaner/issues?state=open",options);
  var data = JSON.parse(response.getContentText());
  
  for(var i = 0; i < data.length-1; i++)
  {
    //var issueToAdd = {state : data[i]["state"],created : data[i]["created_at"], closed : data[i]["closed_at"], response :data[i]["updated_at"],numberOfResponses : data[i]["comments"],name : data[i]["title"] }
    //Logger.log(data[12]);
    openIssues.push([data[i]["state"],GetFormattedDate(data[i]["created_at"]) ,"N/A",GetFormattedDate(data[i]["updated_at"]),data[i]["comments"],data[i]["title"] ])
  
  }
  
  
   var sortedIssues = openIssues.sort(function(a,b){
      var c = new Date(a.created);
      var d = new Date(b.created);
      return c+d;
   });//Sort issues newest to oldest 
   
   for(var i = 0; i < openIssues.length-1; i++)
  {
    Logger.log(openIssues[i].created)
  
  }
  
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  //var employeeName = sheet.getRange("C2").getValue();
  for(var i = 0; i<data.length;i++){
    if(data[i][1] == "Grab N Go Github Ticket Requests"){ //[1] because column B
      CellTitleRowNumber = i+1;
      startOfCellTitle = sheet.getRange(CellTitleRowNumber,2).getValue();//Normally Counted not like in array
      Logger.log("found " + startOfCellTitle);
    }
  }
  
  // clear any previous content
  sheet.getRange(CellTitleRowNumber+1,1,500,7).clearContent();
  
  var len = sortedIssues.length;
  // paste in the values
  sheet.getRange(CellTitleRowNumber+1,2,len,6).setValues(sortedIssues);
  var RangeFormattingToCopy = sheet.getRange(CellTitleRowNumber+1,2,1,6)
  //copyFormatToRange(sheet, column, columnEnd, row, rowEnd)
  RangeFormattingToCopy.copyFormatToRange(sheet, 2, 6, (CellTitleRowNumber+1), (CellTitleRowNumber+len))
  
  var RangeOfEachRowToMerge; 
  for(var i = 0; i<sortedIssues.length;i++){
    var RangeFormattingToMerge = sheet.getRange(((CellTitleRowNumber+1)+i),7,1,5)
    RangeFormattingToMerge.merge();
  }
  
}

function callClosedGithubIssues() {
  var closedIssues = [];
  var startOfCellTitle = "";
  var CellTitleRowNumber = 0;
  
  // Call the Github API for list of issues
  var data = {
  'Authorization': 'Basic ZWdnYm95eEBpY2xvdWQuY29tOmJhY29uYm95MQ==',
  };
  var options = {
  'method' : 'get',
  'contentType': 'application/json; charset=utf-8',
  'headers' : data
  };
  var response = UrlFetchApp.fetch("https://api.github.com/repos/google/loaner/issues?state=closed",options);
  var data = JSON.parse(response.getContentText());
  
  for(var i = 0; i < data.length-1; i++)
  {
    closedIssues.push([data[i]["state"],GetFormattedDate(data[i]["created_at"]),GetFormattedDate(data[i]["closed_at"]),GetFormattedDate(data[i]["updated_at"]),data[i]["comments"],data[i]["title"] ])
   
  }
  
   var sortedIssues = closedIssues.sort(function(a,b){
      var c = new Date(a.created);
      var d = new Date(b.created);
      return c+d;
   });//Sort issues newest to oldest 
   
  //Start Sheet Manipulation
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  for(var i = 0; i<data.length;i++){
    if(data[i][1] == "Grab N Go Github Ticket Requests"){ //[1] because column B
      CellTitleRowNumber = i+1;
      startOfCellTitle = sheet.getRange(CellTitleRowNumber,2).getValue();//Normally Counted not like in array
      Logger.log("found " + startOfCellTitle);
    }
  }
  
  // clear any previous content
  sheet.getRange(CellTitleRowNumber+1,1,500,7).clearContent();
  
  var len = sortedIssues.length;
  // paste in the values
  sheet.getRange(CellTitleRowNumber+1,2,len,6).setValues(sortedIssues);
  var RangeFormattingToCopy = sheet.getRange(CellTitleRowNumber+1,2,1,6)
  //copyFormatToRange(sheet, column, columnEnd, row, rowEnd)
  RangeFormattingToCopy.copyFormatToRange(sheet, 2, 6, (CellTitleRowNumber+1), (CellTitleRowNumber+len))
  
  var RangeOfEachRowToMerge; 
  for(var i = 0; i<sortedIssues.length;i++){
    var RangeFormattingToMerge = sheet.getRange(((CellTitleRowNumber+1)+i),7,1,5)
    RangeFormattingToMerge.merge();
  }
  

}



