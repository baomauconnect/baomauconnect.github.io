/**
 * ============================================================
 *  BaomauConnect - Google Apps Script Backend (Code.gs)
 * ============================================================
 *  Cấu trúc cột trên Google Sheet (đúng theo thứ tự ghi dữ liệu):
 *  1.time | 2.role | 3.category | 4.name | 5.phone | 6.address
 *  7.age | 8.avatar | 9.childAge | 10.bmAgeReq | 11.workTime
 *  12.careNeonatal | 13.detail | 14.userEmail
 *
 *  LƯU Ý: Nếu project Apps Script của bạn còn file "script.js"
 *  (chứa các hàm doGet/doPost trùng tên), hãy XÓA file đó đi.
 *  Vì các hàm cùng tên (doGet, doPost) chỉ có MỘT bản định nghĩa
 *  sau cùng được Apps Script sử dụng, nên để 2 file cùng lúc rất
 *  dễ gây xung đột / mất chức năng. Chỉ giữ lại file Code.gs này.
 * ============================================================
 */

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }

    var headers = data[0];
    var jsonData = [];
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var item = { sheetRow: i + 1 }; // Số dòng thực tế trên Sheet, dùng để update/delete
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
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var payload = {};

    if (e.postData && e.postData.contents) {
      try { payload = JSON.parse(e.postData.contents); } catch (x) { payload = e.parameter; }
    } else {
      payload = e.parameter;
    }

    if (payload.action === "delete") {
      return handleDelete(sheet, payload);
    } else if (payload.action === "update") {
      return handleUpdate(sheet, payload);
    } else {
      return handleCreate(sheet, payload);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * THÊM MỚI 1 ĐĂNG KÝ + LƯU EMAIL GOOGLE + GỬI EMAIL XÁC NHẬN TỰ ĐỘNG
 */
function handleCreate(sheet, payload) {
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
    payload.detail || "",
    payload.userEmail || ""   // Email Google của người vừa đăng nhập & đăng ký
  ]);

  // Tự động gửi email thông báo đăng ký thành công nếu có email người dùng
  if (payload.userEmail) {
    try {
      sendConfirmationEmail(payload);
    } catch (mailErr) {
      // Không để lỗi gửi email làm hỏng việc lưu dữ liệu vào Sheet
      Logger.log("Lỗi gửi email xác nhận: " + mailErr.toString());
    }
  }

  return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * XÓA CÁC DÒNG (dùng cho trang quản trị)
 * payload.rows = mảng số dòng cần xóa, ví dụ [3, 5, 7]
 */
function handleDelete(sheet, payload) {
  var rowsToDelete = typeof payload.rows === "string" ? JSON.parse(payload.rows) : payload.rows;
  rowsToDelete.sort(function (a, b) { return b - a; }); // Xóa từ dòng lớn -> nhỏ để tránh lệch chỉ số
  for (var i = 0; i < rowsToDelete.length; i++) {
    var r = parseInt(rowsToDelete[i]);
    if (r > 1) sheet.deleteRow(r); // Không cho xóa dòng tiêu đề (dòng 1)
  }
  return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * CẬP NHẬT 1 DÒNG (dùng cho trang quản trị)
 * payload.rowData = { sheetRow, role, category, name, phone, address, age,
 *                      avatar, childAge, bmAgeReq, workTime, careNeonatal, detail }
 */
function handleUpdate(sheet, payload) {
  var rowData = typeof payload.rowData === "string" ? JSON.parse(payload.rowData) : payload.rowData;
  var row = parseInt(rowData.sheetRow);

  if (rowData.role !== undefined)         sheet.getRange(row, 2).setValue(rowData.role);
  if (rowData.category !== undefined)     sheet.getRange(row, 3).setValue(rowData.category);
  if (rowData.name !== undefined)         sheet.getRange(row, 4).setValue(rowData.name);
  if (rowData.phone !== undefined)        sheet.getRange(row, 5).setValue(rowData.phone);
  if (rowData.address !== undefined)      sheet.getRange(row, 6).setValue(rowData.address);
  if (rowData.age !== undefined)          sheet.getRange(row, 7).setValue(rowData.age);
  if (rowData.avatar !== undefined)       sheet.getRange(row, 8).setValue(rowData.avatar);
  if (rowData.childAge !== undefined)     sheet.getRange(row, 9).setValue(rowData.childAge);
  if (rowData.bmAgeReq !== undefined)     sheet.getRange(row, 10).setValue(rowData.bmAgeReq);
  if (rowData.workTime !== undefined)     sheet.getRange(row, 11).setValue(rowData.workTime);
  if (rowData.careNeonatal !== undefined) sheet.getRange(row, 12).setValue(rowData.careNeonatal);
  if (rowData.detail !== undefined)       sheet.getRange(row, 13).setValue(rowData.detail);

  return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * GỬI EMAIL THÔNG BÁO ĐĂNG KÝ THÀNH CÔNG TỚI GMAIL NGƯỜI DÙNG
 * (Gmail lấy được từ tài khoản Google mà người dùng vừa đăng nhập trên web)
 */
function sendConfirmationEmail(data) {
  var subject = "✅ BaomauConnect - Đăng ký thành công: " + data.category;
  var body =
    "Chào " + data.name + ",\n\n" +
    "Hệ thống BaomauConnect đã nhận được thông tin đăng ký của bạn.\n\n" +
    "THÔNG TIN CHI TIẾT:\n" +
    "- Vai trò: " + data.role + "\n" +
    "- Loại đăng ký: " + data.category + "\n" +
    "- Số điện thoại: " + data.phone + "\n" +
    "- Địa chỉ: " + data.address + "\n" +
    "- Nội dung: " + data.detail + "\n\n" +
    "Đội ngũ BaomauConnect sẽ kiểm tra và liên hệ lại với bạn trong thời gian sớm nhất.\n\n" +
    "Trân trọng,\n" +
    "Đội ngũ BaomauConnect.asia";

  GmailApp.sendEmail(data.userEmail, subject, body);
}