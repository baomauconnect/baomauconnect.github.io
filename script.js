function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    
    var headers = data[0];
    var jsonData = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var item = {};
      var hasData = false;
      for (var j = 0; j < headers.length; j++) {
        if (row[j] !== "") hasData = true;
        if (headers[j] === "time" && row[j] instanceof Date) {
          item[headers[j]] = Utilities.formatDate(row[j], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else {
          item[headers[j]] = row[j];
        }
      }
      if (hasData) jsonData.push(item);
    }
    return ContentService.createTextOutput(JSON.stringify(jsonData)).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var payload = {};
    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch(x) { payload = e.parameter; }
    } else { payload = e.parameter; }
    
    if (payload.action === "delete") {
      var rowsToDelete = typeof payload.rows === "string" ? JSON.parse(payload.rows) : payload.rows;
      rowsToDelete.sort(function(a, b) { return b - a; });
      for (var i = 0; i < rowsToDelete.length; i++) {
        var r = parseInt(rowsToDelete[i]);
        if (r > 1) sheet.deleteRow(r);
      }
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    } else {
      var time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.appendRow([
        time, 
        payload.role || "", 
        payload.category || "", 
        payload.name || "", 
        payload.phone || "", 
        payload.address || "", 
        payload.age || "", 
        payload.avatar || "", 
        payload.childAge || "", 
        payload.bmAgeReq || "", 
        payload.workTime || "", 
        payload.careNeonatal || "", 
        payload.detail || ""
      ]);
      return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}