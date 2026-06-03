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

function doPost(e) {
  var p = e.parameter;
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0]; // Chọn sheet đầu tiên
  
  if (p.action === 'delete') {
     // Mã xử lý lệnh xóa cũ của bạn...
  } 
  else if (p.action === 'update') {
     var rowData = JSON.parse(p.rowData);
     var row = parseInt(rowData.sheetRow);
     
     // Ví dụ ánh xạ các cột dữ liệu theo thứ tự cột trên Google Sheets của bạn:
     // Giả định: Cột 3: Tên, Cột 4: Điện thoại, Cột 5: Khung giờ... vv
     sheet.getRange(row, 3).setValue(rowData.name);
     sheet.getRange(row, 4).setValue(rowData.phone);
     sheet.getRange(row, 5).setValue(rowData.workTime);
     sheet.getRange(row, 7).setValue(rowData.address);
     sheet.getRange(row, 8).setValue(rowData.detail);
     
     if (rowData.age !== undefined) {
         sheet.getRange(row, 6).setValue(rowData.age); // Ví dụ Cột 6 là Tuổi BM / Bé
         // Cập nhật thêm các cột careNeonatal, avatar tùy kiến trúc bảng của bạn...
     } else {
         sheet.getRange(row, 6).setValue(rowData.childAge);
     }
     
     return ContentService.createTextOutput("Update Success");
  }
}