function executeQuery() {
   var result = Browser.msgBox("クエリを実行します", Browser.Buttons.OK_CANCEL);
   if (result == "ok"){ 
     runInstallQuery();
     Browser.msgBox("完了しました", Browser.Buttons.OK_CANCEL)
   }
}

function runInstallQuery() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startTime = sheet.getRange(2, 3).getValue();
  var endTime = sheet.getRange(2, 5).getValue();
  var startTimeStr = Utilities.formatDate(startTime, 'JST', 'yyyy-MM-dd');
  var endTimeStr = Utilities.formatDate(endTime, 'JST', 'yyyy-MM-dd');
  var datasource = sheet.getRange(2, 8).getValue();
  var app_id = sheet.getRange(2, 10).getValue();
  
  var table = datasource + "." +  datasource + "_" + app_id;
  
  var pt_cond = "_PARTITIONTIME BETWEEN TIMESTAMP('" + startTimeStr + "') AND TIMESTAMP('" + endTimeStr + "')";  
  var query = 'SELECT COUNT(*) FROM ' + table + ' WHERE installed_at IS NOT NULL AND ' + pt_cond;
  
  var request = {query: query};
  var queryResults = BigQuery.Jobs.query(request, 'spinapptest-151310');
  var jobId = queryResults.jobReference.jobId;

  var sleepTimeMs = 500;
  while (!queryResults.jobComplete) {
    Utilities.sleep(sleepTimeMs);
    sleepTimeMs *= 2;
    queryResults = BigQuery.Jobs.getQueryResults(projectId, jobId);
  }
  
  var rows = queryResults.rows;
  var installCount = rows[0]['f'][0]['v']
  
  sheet.getRange(8, 3).setValue(installCount)
  Browser.msgBox("クエリ完了");
}
