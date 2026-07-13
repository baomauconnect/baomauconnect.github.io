/**
 * ============================================================
 *  BaomauConnect - Google Apps Script Backend (Code.gs)
 * ============================================================
 *  Cấu trúc cột trên Google Sheet (đúng theo thứ tự ghi dữ liệu):
 *  1.time | 2.role | 3.(giữ trống, không dùng category) | 4.name | 5.phone | 6.address
 *  7.age | 8.avatar | 9.childAge | 10.bmAgeReq | 11.workTime
 *  12.careNeonatal | 13.detail | 14.userEmail
 *
 *  LƯU Ý: Project Apps Script chỉ nên có DUY NHẤT 1 file code
 *  (file này). Nếu còn file "script.js" hay bất kỳ file .gs nào
 *  khác cũng định nghĩa doGet/doPost, hãy XÓA nó đi để tránh
 *  xung đột tên hàm.
 * ============================================================
 */

// Số lượng hồ sơ phù hợp tối đa gửi kèm trong 1 email ghép nối
var MAX_MATCHES = 5;

function doGet(e) {
  try {
    createDailyBackupTrigger();
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var data = sheet.getDataRange().getValues();
    
    // Nếu chỉ có header hoặc không có dữ liệu
    if (data.length <= 1) {
      Logger.log("ℹ️ No data to retrieve - only header present");
      return ContentService.createTextOutput(JSON.stringify([])).setMimeType(ContentService.MimeType.JSON);
    }

    var headers = data[0];
    var jsonData = [];
    
    // Chuyển đổi dữ liệu từ Sheet sang JSON
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var item = { sheetRow: i + 1 }; // Số dòng thực tế trên Sheet
      var hasData = false;
      
      for (var j = 0; j < headers.length; j++) {
        if (row[j] !== "") hasData = true;
        
        // Xử lý đặc biệt cho cột thời gian
        if (headers[j] === "time" && row[j] instanceof Date) {
          item[headers[j]] = Utilities.formatDate(row[j], Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
        } else {
          item[headers[j]] = row[j];
        }
      }
      
      // Chỉ thêm dòng nếu có dữ liệu
      if (hasData) jsonData.push(item);
    }
    
    Logger.log("✅ Retrieved " + jsonData.length + " records from sheet");
    return ContentService.createTextOutput(JSON.stringify(jsonData)).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    var errorMsg = err.toString();
    Logger.log("❌ Error in doGet: " + errorMsg);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: errorMsg })).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var payload = {};

    // Cố gắng phân tích dữ liệu từ các nguồn khác nhau
    if (e.postData && e.postData.contents) {
      try { 
        payload = JSON.parse(e.postData.contents); 
      } catch (x) { 
        payload = e.parameter; 
      }
    } else {
      payload = e.parameter;
    }

    Logger.log("📥 Incoming request - Action: " + (payload.action || 'unknown'));

    // Định tuyến yêu cầu dựa trên action
    if (payload.action === "delete") {
      Logger.log("🗑️ Processing DELETE request");
      return handleDelete(sheet, payload);
    } else if (payload.action === "update") {
      Logger.log("✏️ Processing UPDATE request");
      return handleUpdate(sheet, payload);
    } else if (payload.action === "backup") {
      Logger.log("💾 Processing BACKUP request");
      return handleBackup(sheet, payload);
    } else {
      Logger.log("➕ Processing CREATE request");
      return handleCreate(sheet, payload);
    }
  } catch (error) {
    var errorMsg = error.toString();
    Logger.log("❌ Error in doPost: " + errorMsg);
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: errorMsg })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getBackupSheetName() {
  return "Backup_" + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function createDailyBackupIfNeeded() {
  try {
    var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    var sourceSheet = spreadsheet.getActiveSheet();
    var backupName = getBackupSheetName();
    var existingBackup = spreadsheet.getSheetByName(backupName);

    if (existingBackup) {
      return { status: "exists", sheetName: backupName };
    }

    var backupSheet = spreadsheet.insertSheet(backupName, spreadsheet.getNumSheets());
    var data = sourceSheet.getDataRange().getValues();
    if (data.length > 0 && data[0].length > 0) {
      backupSheet.getRange(1, 1, data.length, data[0].length).setValues(data);
    }

    Logger.log("✅ Created daily backup: " + backupName);
    return { status: "created", sheetName: backupName };
  } catch (err) {
    Logger.log("⚠️ Backup creation failed: " + err.toString());
    return { status: "error", message: err.toString() };
  }
}

function createDailyBackupTrigger() {
  var existingTriggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < existingTriggers.length; i++) {
    if (existingTriggers[i].getHandlerFunction() === "createDailyBackup") {
      return { status: "exists" };
    }
  }

  ScriptApp.newTrigger("createDailyBackup")
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(0)
    .create();

  return { status: "created" };
}

function createDailyBackup() {
  createDailyBackupIfNeeded();
}

function handleBackup(sheet, payload) {
  var backupResult = createDailyBackupIfNeeded();
  return ContentService.createTextOutput(JSON.stringify({
    status: "success",
    backup: backupResult
  })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * THÊM MỚI 1 ĐĂNG KÝ + LƯU EMAIL GOOGLE + GỬI EMAIL XÁC NHẬN
 * + TỰ ĐỘNG GHÉP NỐI & GỬI THÔNG TIN CHO ĐỐI PHƯƠNG CÙNG KHU VỰC
 */
function handleCreate(sheet, payload) {
  var time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  payload.time = time; // gắn lại thời gian để đưa vào email

  sheet.appendRow([
    time,
    payload.role || "",
    "",
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

  // 1) Gửi email xác nhận đăng ký thành công cho chính người vừa đăng ký
  if (payload.userEmail) {
    try {
      sendConfirmationEmail(payload);
    } catch (mailErr) {
      Logger.log("Lỗi gửi email xác nhận: " + mailErr.toString());
    }
  }

  // 2) Tự động tìm những hồ sơ đối phương cùng khu vực và gửi email ghép nối
  try {
    matchAndNotify(sheet, payload);
  } catch (matchErr) {
    Logger.log("Lỗi ghép nối / gửi email đối phương: " + matchErr.toString());
  }

  createDailyBackupIfNeeded();
  return ContentService.createTextOutput(JSON.stringify({ status: "success" })).setMimeType(ContentService.MimeType.JSON);
}

/**
 * XÓA CÁC DÒNG (dùng cho trang quản trị)
 */
function handleDelete(sheet, payload) {
  try {
    var rowsToDelete = typeof payload.rows === "string" ? JSON.parse(payload.rows) : payload.rows;
    
    if (!Array.isArray(rowsToDelete) || rowsToDelete.length === 0) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "No rows to delete" })).setMimeType(ContentService.MimeType.JSON);
    }

    // Sắp xếp theo thứ tự giảm dần để xóa từ dưới lên trên (tránh lỗi index)
    rowsToDelete.sort(function (a, b) { return b - a; });
    
    var deletedRows = [];
    for (var i = 0; i < rowsToDelete.length; i++) {
      var r = parseInt(rowsToDelete[i]);
      if (r > 1) { // Chỉ xóa dòng > 1 (tránh xóa header)
        sheet.deleteRow(r);
        deletedRows.push(r);
      }
    }
    
    Logger.log("✅ Deleted rows: " + deletedRows.join(", "));
    createDailyBackupIfNeeded();
    return ContentService.createTextOutput(JSON.stringify({ status: "success", deletedCount: deletedRows.length, deletedRows: deletedRows })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log("❌ Delete error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

function handleUpdate(sheet, payload) {
  try {
    var rowData = typeof payload.rowData === "string" ? JSON.parse(payload.rowData) : payload.rowData;
    
    // TỐI ƯU LOGIC: Chấp nhận cả sheetRow hoặc row để tránh lỗi bất đồng bộ client-server
    var row = parseInt(rowData.sheetRow || rowData.row);

    // Kiểm tra số dòng hợp lệ (phải > 1 để tránh ghi đè header)
    if (row <= 1 || isNaN(row)) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid row number: " + (rowData.sheetRow || rowData.row) })).setMimeType(ContentService.MimeType.JSON);
    }

    // Chuẩn hóa dữ liệu trước khi cập nhật theo đúng cấu trúc 14 cột
    if (rowData.name !== undefined)         sheet.getRange(row, 4).setValue(String(rowData.name).trim());
    if (rowData.phone !== undefined)        sheet.getRange(row, 5).setValue(String(rowData.phone).trim());
    if (rowData.address !== undefined)      sheet.getRange(row, 6).setValue(String(rowData.address).trim());
    if (rowData.age !== undefined)          sheet.getRange(row, 7).setValue(String(rowData.age).trim());
    if (rowData.avatar !== undefined)       sheet.getRange(row, 8).setValue(String(rowData.avatar).trim());
    if (rowData.childAge !== undefined)     sheet.getRange(row, 9).setValue(String(rowData.childAge).trim());
    if (rowData.bmAgeReq !== undefined)     sheet.getRange(row, 10).setValue(String(rowData.bmAgeReq).trim());
    if (rowData.workTime !== undefined)     sheet.getRange(row, 11).setValue(String(rowData.workTime).trim());
    if (rowData.careNeonatal !== undefined) sheet.getRange(row, 12).setValue(String(rowData.careNeonatal).trim());
    if (rowData.detail !== undefined)       sheet.getRange(row, 13).setValue(String(rowData.detail).trim());

    Logger.log("✅ Update row " + row + " successfully");
    createDailyBackupIfNeeded();
    return ContentService.createTextOutput(JSON.stringify({ status: "success", updatedRow: row })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    Logger.log("❌ Update error: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() })).setMimeType(ContentService.MimeType.JSON);
  }
}

/* =====================================================================
 *  KHU VỰC LOGIC GHÉP NỐI (MATCHING) THEO ĐỊA CHỈ & GỬI EMAIL THÔNG BÁO
 * ===================================================================== */

function normalizeText(str) {
  if (!str) return "";
  return str.toString().toLowerCase().trim();
}

function extractRegion(address) {
  if (!address) return "";
  var parts = address.toString().split(",");
  var last = parts[parts.length - 1] || address;
  return normalizeText(last);
}

function getOppositeRole(role) {
  var r = normalizeText(role);
  if (r === "bảo mẫu") return "Phụ huynh";
  if (r === "phụ huynh") return "Bảo mẫu";
  return null;
}

function matchAndNotify(sheet, newEntry) {
  if (!newEntry.address) return;

  var oppositeRole = getOppositeRole(newEntry.role);
  if (!oppositeRole) return; 

  var newRegion = extractRegion(newEntry.address);
  if (!newRegion) return;

  var data = sheet.getDataRange().getValues();
  var headers = data[0];
  var colIndex = {};
  for (var c = 0; c < headers.length; c++) colIndex[headers[c]] = c;

  var matches = [];
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var rowRole = row[colIndex["role"]];
    var rowAddress = row[colIndex["address"]];
    var rowEmail = row[colIndex["userEmail"]];

    if (!rowRole || !rowAddress) continue;
    if (normalizeText(rowRole) !== normalizeText(oppositeRole)) continue;
    if (extractRegion(rowAddress) !== newRegion) continue;
    if (rowEmail && newEntry.userEmail && rowEmail === newEntry.userEmail) continue; 

    matches.push({
      role: rowRole,
      name: row[colIndex["name"]],
      phone: row[colIndex["phone"]],
      address: rowAddress,
      age: row[colIndex["age"]],
      childAge: row[colIndex["childAge"]],
      bmAgeReq: row[colIndex["bmAgeReq"]],
      workTime: row[colIndex["workTime"]],
      careNeonatal: row[colIndex["careNeonatal"]],
      detail: row[colIndex["detail"]],
      userEmail: rowEmail
    });

    if (matches.length >= MAX_MATCHES) break;
  }

  if (matches.length === 0) return;

  // Gửi cho người vừa đăng ký danh sách đối tác phù hợp
  if (newEntry.userEmail) {
    sendMatchListEmail(newEntry, matches);
  }

  // Gửi thông tin của người mới cho từng đối tác cũ đã lưu trên hệ thống
  matches.forEach(function (m) {
    if (m.userEmail) {
      sendMatchListEmail(m, [newEntry]);
    }
  });
}

/**
 * HOÀN THIỆN: Khôi phục đoạn mã HTML tạo bảng thông tin cấu trúc hồ sơ trong email
 */
function buildProfileRowsHtml(profile) {
  var isBaomau = normalizeText(profile.role) === "bảo mẫu";
  var rows = [];
  rows.push(["Vai trò đối tác", profile.role]);
  rows.push(["Họ và tên liên hệ", profile.name]);
  rows.push(["Số điện thoại / Zalo", profile.phone]);
  rows.push(["Địa chỉ khu vực", profile.address]);

  if (isBaomau) {
    if (profile.age)          rows.push(["Tuổi bảo mẫu", profile.age]);
    if (profile.workTime)     rows.push(["Khung giờ làm việc mong muốn", profile.workTime]);
    if (profile.careNeonatal) rows.push(["Khả năng nhận trẻ sơ sinh", profile.careNeonatal]);
  } else {
    if (profile.childAge)     rows.push(["Độ tuổi hiện tại của bé", profile.childAge]);
    if (profile.bmAgeReq)     rows.push(["Yêu cầu độ tuổi bảo mẫu", profile.bmAgeReq]);
    if (profile.workTime)     rows.push(["Khung giờ cần bảo mẫu", profile.workTime]);
  }
  if (profile.detail)         rows.push(["Mô tả / Yêu cầu chi tiết", profile.detail]);

  var html = '<table border="1" style="border-collapse: collapse; width: 100%; max-width: 600px; margin-top: 10px;">';
  rows.forEach(function(r) {
    html += '<tr>' +
              '<td style="padding: 10px; background-color: #f8fafc; font-weight: bold; width: 200px;">' + r[0] + '</td>' +
              '<td style="padding: 10px;">' + r[1] + '</td>' +
            '</tr>';
  });
  html += '</table>';
  return html;
}

/**
 * HOÀN THIỆN: Gửi email xác nhận gửi đơn thành công cho đối tác
 */
function sendConfirmationEmail(payload) {
  var subject = "🎉 [BaomauConnect] Hồ sơ đăng ký kết nối trực tuyến thành công";
  var htmlBody = '<div style="font-family: Arial, sans-serif; color: #1e293b; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px;">' +
                   '<h2 style="color: #4A90E2; margin-bottom: 6px;">Xin chào ' + payload.name + ',</h2>' +
                   '<p>Hệ thống đám mây <b>BaomauConnect</b> đã lưu trữ thành công thông tin đăng ký của bạn vào cơ sở dữ liệu nội bộ trực tuyến theo thời gian thực.</p>' +
                   '<p>Dưới đây là thông tin chi tiết hồ sơ của bạn:</p>' +
                   buildProfileRowsHtml(payload) +
                   '<p style="margin-top: 20px; font-size: 0.9rem; color: #64748b;">Hệ thống đang tiến hành rà soát dữ liệu khu vực để tự động ghép nối bạn với đối tác phù hợp nhất. Xin vui lòng kiểm tra email thường xuyên.</p>' +
                 '</div>';
                 
  MailApp.sendEmail({
    to: payload.userEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * HOÀN THIỆN: Gửi danh sách ghép nối tự động
 */
function sendMatchListEmail(recipient, matches) {
  var isRecipientBm = normalizeText(recipient.role) === "bảo mẫu";
  var partnerType = isRecipientBm ? "PHỤ HUYNH" : "BẢO MẪU";
  
  var subject = "🔔 [BaomauConnect] Phát hiện đối tác (" + partnerType + ") phù hợp tại khu vực của bạn";
  var htmlBody = '<div style="font-family: Arial, sans-serif; color: #1e293b; line-height: 1.6; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e2e8f0; border-radius: 12px;">' +
                   '<h3 style="color: #2ecc71;">Xin chào ' + recipient.name + ',</h3>' +
                   '<p>Dựa trên bộ lọc địa chính thời gian thực, hệ thống đã tìm thấy <b>' + matches.length + ' hồ sơ đối tác mới phù hợp</b> tại khu vực sinh sống/làm việc của bạn.</p>' +
                   '<p>Bạn có thể chủ động liên hệ trao đổi công việc theo danh sách dưới đây:</p>';

  matches.forEach(function (m, idx) {
    htmlBody += '<div style="margin-top: 25px; padding-top: 15px; border-top: 2px dashed #e2e8f0;">' +
                  '<h4 style="color: #0f172a; margin-bottom: 8px;">🎯 Đối tác phù hợp #' + (idx + 1) + '</h4>' +
                  buildProfileRowsHtml(m) +
                '</div>';
  });

  htmlBody += '<p style="margin-top: 25px; font-size: 0.85rem; color: #94a3b8; border-top: 1px solid #e2e8f0; padding-top: 10px;">Trân trọng cảm ơn bạn đã đồng hành cùng BaomauConnect - Hệ thống kết nối tự động 24/7.</p>' +
              '</div>';

  MailApp.sendEmail({
    to: recipient.userEmail,
    subject: subject,
    htmlBody: htmlBody
  });
}