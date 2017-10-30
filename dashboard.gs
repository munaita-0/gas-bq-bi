Dashboard = function() {
  this.sheet = SpreadsheetApp.getActiveSheet();
  this.startTime = this.sheet.getRange(3, 2).getValue();
  this.endTime = this.sheet.getRange(3, 3).getValue();
  this.startTimeStr = Utilities.formatDate(this.startTime, 'JST', 'yyyy-MM-dd');
  this.endTimeStr = Utilities.formatDate(this.endTime, 'JST', 'yyyy-MM-dd');
  this.datasource = this.sheet.getRange(3, 4).getValue();
  this.app_id = this.sheet.getRange(3, 5).getValue();
  this.table = this.datasource + "." +  this.datasource + "_" + this.app_id;
  this.pt_cond = "_PARTITIONTIME BETWEEN TIMESTAMP('" + this.startTimeStr + "') AND TIMESTAMP('" + this.endTimeStr + "')";
  this.bq = initBqApi();
}

// Dashboardクラスが参照できなかったためinit関数を定義
function initDashboard() {
  return new Dashboard();
}

Dashboard.prototype.update = function() {
  this.updateInstalls();
  this.updateRetentions();
  this.updateEvents();
}

Dashboard.prototype.updateInstalls = function() {
  // organic install count
  var install_query = "SELECT" +
    " SUM(CASE WHEN network_name = 'Organic' THEN 1 ELSE 0 END)," +
    " SUM(CASE WHEN network_name != 'Organic'  THEN 1 ELSE 0 END)" +
    " FROM " + this.table + 
    " WHERE activity_kind = 'install'" +
    " AND network_name = 'Organic'" +
    " AND " + this.pt_cond;

  var rows = this.bq.executeQuery(install_query);
  for(k in rows) {
    row = rows[k];
    this.sheet.getRange(8, 3).setValue(row['f'][0]['v'])
      this.sheet.getRange(9, 3).setValue(row['f'][0]['v'])
  }
}

Dashboard.prototype.updateRetentions = function() {
  // organic classic 2 days retention
  var rows = this.bq.executeQuery(this.getRetentionQuery(2, true));  
  this.sheet.getRange(8, 4).setValue(rows[0]['f'][0]['v']);

  // paid classic 2 days retention
  var rows = this.bq.executeQuery(this.getRetentionQuery(2, false));  
  this.sheet.getRange(9, 4).setValue(rows[0]['f'][0]['v']);

  // organic classic 7 days retention
  var rows = this.bq.executeQuery(this.getRetentionQuery(7, true));  
  this.sheet.getRange(8, 6).setValue(rows[0]['f'][0]['v']);

  // paid classic 7 days retention
  var rows = this.bq.executeQuery(this.getRetentionQuery(7, false));  
  this.sheet.getRange(9, 6).setValue(rows[0]['f'][0]['v']);
  Browser.msgBox("4update done.", Browser.Buttons.OK_CANCEL);

  // organic classic 14 days retention
  var rows = this.bq.executeQuery(this.getRetentionQuery(14, true));  
  this.sheet.getRange(8, 8).setValue(rows[0]['f'][0]['v']);
  Browser.msgBox("5update done.", Browser.Buttons.OK_CANCEL);

  // paid classic 14 days retention
  var rows = this.bq.executeQuery(this.getRetentionQuery(14, false));  
  this.sheet.getRange(9, 8).setValue(rows[0]['f'][0]['v']);
  Browser.msgBox("6update done.", Browser.Buttons.OK_CANCEL);
}

Dashboard.prototype.getRetentionQuery = function(day, isOrganic) {
  eq = (isOrganic) ? '=' : '!=';
  return "SELECT COUNT(DISTINCT(idfa))" + 
    " FROM " + this.table + 
    " WHERE activity_kind = 'session' " + 
    "AND network_name " + eq + " 'Organic' " +
    "AND (created_at - installed_at) BETWEEN ( 24 * 60 * 60 ) AND (24 * 60 * 60 * " + day + ") " +
    "AND installed_at > UNIX_SECONDS(TIMESTAMP('" + this.startTimeStr + "')) " +
    "AND " + this.pt_cond;
}

Dashboard.prototype.updateEvents = function() {
  var eq_array = ['=', '!=']

    var rows = [];
  for(i in eq_array) {
    var events_query = "SELECT event_name, COUNT(distinct(idfa)) AS uu, COUNT(idfa) AS total" + 
      " FROM " + this.table + 
      " WHERE network_name " + eq_array[i] + " 'Organic' " +
      " AND " + this.pt_cond +
      " GROUP BY event_name"

      rows = this.bq.executeQuery(events_query);
    this.setRowsToSheet(rows, (eq_array[i] === '='))
  }
  range = this.sheet.getRange(12, 2, 4,  (rows.length * 2) + 1)
    range.setBorder(true, true, true, true, true, true)
    range = this.sheet.getRange(12, 2, 2,  (rows.length * 2) + 1)
    range.setBackgroundRGB(170, 170, 170)
    Browser.msgBox("6update done.", Browser.Buttons.OK_CANCEL);
}

Dashboard.prototype.setRowsToSheet = function(rows, isOrganic) {
  var firstLine = isOrganic ? 12:13;
  var rowCount = 1;
  for(key in rows) {
    row = rows[key];

    if (isOrganic) {
      this.sheet.getRange(firstLine, rowCount * 2).setValue(row['f'][0]['v']);
      this.sheet.getRange(firstLine + 1, rowCount * 2).setValue('UU');
      this.sheet.getRange(firstLine + 1, (rowCount * 2) + 1).setValue('total');
    }

    this.sheet.getRange(firstLine + 2, rowCount * 2).setValue(row['f'][1]['v']); 
    this.sheet.getRange(firstLine + 2, (rowCount * 2) + 1).setValue(row['f'][2]['v']);

    rowCount++;
  }
}
