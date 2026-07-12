/**
 * ============================================================
 *  BaomauConnect - Google Apps Script Backend (Code.gs)
 * ============================================================
 *  Cấu trúc cột trên Google Sheet (đúng theo thứ tự ghi dữ liệu):
 *  1.time | 2.role | 3.category | 4.name | 5.phone | 6.address
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

/**
 * CẬP NHẬT 1 DÒNG (dùng cho trang quản trị)
 */
function handleUpdate(sheet, payload) {
  try {
    var rowData = typeof payload.rowData === "string" ? JSON.parse(payload.rowData) : payload.rowData;
    var row = parseInt(rowData.sheetRow);

    // Kiểm tra số dòng hợp lệ (phải > 1 để tránh xóa header)
    if (row <= 1 || isNaN(row)) {
      return ContentService.createTextOutput(JSON.stringify({ status: "error", message: "Invalid row number" })).setMimeType(ContentService.MimeType.JSON);
    }

    // Chuẩn hóa dữ liệu trước khi cập nhật
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
 *  KHU VỰC LOGIC GHÉP NỐI (MATCHING) THEO ĐỊA CHỈ
 * ===================================================================== */

/**
 * Chuẩn hóa chuỗi để so sánh: chỉ cần chuyển chữ thường + trim.
 * (Không cần xử lý bỏ dấu vì địa chỉ được người dùng CHỌN từ dropdown
 * tỉnh/thành có sẵn trên form, nên luôn đồng nhất chính tả.)
 */
function normalizeText(str) {
  if (!str) return "";
  return str.toString().toLowerCase().trim();
}

/**
 * Trích "khu vực" (thường là Tỉnh/Thành phố) từ chuỗi địa chỉ đầy đủ.
 * Địa chỉ được ghép dạng: "Số nhà, Phường/Xã, Quận/Huyện, Tỉnh/Thành phố"
 * => lấy đoạn cuối cùng sau dấu phẩy làm khu vực để so khớp.
 * Nếu địa chỉ không có dấu phẩy, dùng luôn toàn bộ chuỗi.
 */
function extractRegion(address) {
  if (!address) return "";
  var parts = address.toString().split(",");
  var last = parts[parts.length - 1] || address;
  return normalizeText(last);
}

/**
 * Xác định vai trò đối lập để tìm ghép nối:
 * "Bảo mẫu" <-> "Phụ huynh"
 * (So khớp chính xác vì giá trị role là cố định trong code, không phải
 * do người dùng gõ tự do, nên không cần xử lý sai lệch chính tả.)
 */
function getOppositeRole(role) {
  var r = normalizeText(role);
  if (r === "bảo mẫu") return "Phụ huynh";
  if (r === "phụ huynh") return "Bảo mẫu";
  return null;
}

/**
 * Tìm các hồ sơ đối phương (vai trò ngược lại) có cùng khu vực địa chỉ
 * với hồ sơ vừa đăng ký, rồi gửi email thông báo ghép nối cho cả 2 phía.
 */
function matchAndNotify(sheet, newEntry) {
  if (!newEntry.address) return;

  var oppositeRole = getOppositeRole(newEntry.role);
  if (!oppositeRole) return; // Không xác định được vai trò thì bỏ qua

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
    if (rowEmail && newEntry.userEmail && rowEmail === newEntry.userEmail) continue; // tránh tự ghép với chính mình

    matches.push({
      role: rowRole,
      category: row[colIndex["category"]],
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

  // (a) Gửi cho người VỪA đăng ký: danh sách đối phương phù hợp cùng khu vực
  if (newEntry.userEmail) {
    sendMatchListEmail(newEntry, matches);
  }

  // (b) Gửi cho TỪNG đối phương đã có sẵn: thông báo có người mới phù hợp
  matches.forEach(function (m) {
    if (m.userEmail) {
      sendMatchListEmail(m, [newEntry]);
    }
  });
}

/**
 * Tạo bảng HTML hiển thị thông tin 1 hồ sơ (dùng chung cho email ghép nối)
 */
function buildProfileRowsHtml(profile) {
  var isBaomau = normalizeText(profile.role) === "bảo mẫu";
  var rows = [];
  rows.push(["Vai trò", profile.role]);
  rows.push(["Họ và tên", profile.name]);
  rows.push(["Số điện thoại", profile.phone]);
  rows.push(["Địa chỉ", profile.address]);

  if (isBaomau) {
    if (profile.age)          rows.push(["Tuổi", profile.age]);
    if (profile.workTime)     rows.push(["Thời gian có thể làm việc", profile.workTime]);
    if (profile.careNeonatal) rows.push(["Nhận chăm bé sơ sinh", profile.careNeonatal]);
  } else {
    if (profile.childAge)     rows.push(["Độ tuổi của bé", profile.childAge]);
    if (profile.bmAgeReq)     rows.push(["Yêu cầu độ tuổi bảo mẫu", profile.bmAgeReq]);
    if (profile.workTime)     rows.push(["Khung giờ cần bảo mẫu", profile.workTime]);
  }
  if (profile.detail) rows.push(["Nội dung / Ghi chú thêm", profile.detail]);

  return rows
    .filter(function (r) { return r[1] !== undefined && r[1] !== null && r[1] !== ""; })
    .map(function (r) {
      return '<tr>' +
        '<td style="padding:8px 12px;border:1px solid #e2e8f0;background:#f8fafc;font-weight:600;color:#334155;width:40%;">' + r[0] + '</td>' +
        '<td style="padding:8px 12px;border:1px solid #e2e8f0;color:#1e293b;">' + r[1] + '</td>' +
        '</tr>';
    })
    .join('');
}

/**
 * Gửi email cho `recipient` (phải có userEmail), liệt kê danh sách
 * `profiles` là các hồ sơ đối phương phù hợp cùng khu vực.
 */
function sendMatchListEmail(recipient, profiles) {
  if (!recipient.userEmail) return;

  var region = extractRegionDisplay(recipient.address);
  var subject = "BaomauConnect - Có " + profiles.length + " hồ sơ phù hợp gần bạn (" + region + ")";

  var blocksHtml = profiles.map(function (p) {
    return '<div style="margin-bottom:18px;border:1px solid #e2e8f0;border-radius:8px;overflow:hidden;">' +
      '<table style="border-collapse:collapse;width:100%;font-size:14px;">' +
        buildProfileRowsHtml(p) +
      '</table>' +
    '</div>';
  }).join('');

  var htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">' +
      '<div style="background:linear-gradient(135deg,#22c55e,#16a34a);padding:24px 30px;border-radius:10px 10px 0 0;">' +
        '<h2 style="color:#ffffff;margin:0;font-size:20px;">Tìm thấy hồ sơ phù hợp!</h2>' +
      '</div>' +
      '<div style="padding:24px 30px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px;">' +
        '<p style="color:#334155;font-size:15px;">Chào <b>' + (recipient.name || '') + '</b>,</p>' +
        '<p style="color:#334155;font-size:15px;">Hệ thống BaomauConnect vừa tìm thấy <b>' + profiles.length + '</b> hồ sơ cùng khu vực <b>' + region + '</b> có thể phù hợp với bạn:</p>' +
        blocksHtml +
        '<p style="color:#334155;font-size:15px;">Bạn có thể chủ động liên hệ trực tiếp qua số điện thoại ở trên để trao đổi thêm chi tiết.</p>' +
        '<p style="color:#94a3b8;font-size:13px;margin-top:24px;">Đây là email tự động, vui lòng không trả lời trực tiếp email này.</p>' +
        '<p style="color:#334155;font-size:15px;">Trân trọng,<br><b>Đội ngũ BaomauConnect.asia</b></p>' +
      '</div>' +
    '</div>';

  var plainBody = profiles.map(function (p, idx) {
    return "Hồ sơ " + (idx + 1) + ":\n" +
      "- Họ tên: " + (p.name || "") + "\n" +
      "- SĐT: " + (p.phone || "") + "\n" +
      "- Địa chỉ: " + (p.address || "") + "\n";
  }).join("\n");

  GmailApp.sendEmail(
    recipient.userEmail,
    subject,
    "Chào " + (recipient.name || '') + ",\n\nHệ thống tìm thấy các hồ sơ phù hợp cùng khu vực " + region + ":\n\n" + plainBody +
    "\nTrân trọng,\nĐội ngũ BaomauConnect.asia",
    {
      htmlBody: htmlBody,
      name: "BaomauConnect"
    }
  );
}

/**
 * Lấy tên khu vực để hiển thị (giữ nguyên chữ gốc, không bỏ dấu)
 * dùng cho tiêu đề / nội dung email cho thân thiện.
 */
function extractRegionDisplay(address) {
  if (!address) return "";
  var parts = address.toString().split(",");
  return (parts[parts.length - 1] || "").toString().trim();
}

/* =====================================================================
 *  EMAIL XÁC NHẬN ĐĂNG KÝ THÀNH CÔNG (gửi cho chính người đăng ký)
 * ===================================================================== */

function sendConfirmationEmail(data) {
  var isBaomau = (data.role || "").toString().trim() === "Bảo mẫu";
  var subject = "BaomauConnect - Đăng ký thành công (" + (data.category || "") + ")";

  var rows = [];
  rows.push(["Vai trò", data.role]);
  rows.push(["Loại đăng ký", data.category]);
  rows.push(["Họ và tên", data.name]);
  rows.push(["Số điện thoại", data.phone]);
  rows.push(["Địa chỉ", data.address]);

  if (isBaomau) {
    if (data.age)           rows.push(["Tuổi", data.age]);
    if (data.workTime)      rows.push(["Thời gian có thể làm việc", data.workTime]);
    if (data.careNeonatal)  rows.push(["Nhận chăm bé sơ sinh", data.careNeonatal]);
  } else {
    if (data.childAge)      rows.push(["Độ tuổi của bé", data.childAge]);
    if (data.bmAgeReq)      rows.push(["Yêu cầu độ tuổi bảo mẫu", data.bmAgeReq]);
    if (data.workTime)      rows.push(["Khung giờ cần bảo mẫu", data.workTime]);
  }

  if (data.detail) rows.push(["Nội dung / Ghi chú thêm", data.detail]);
  rows.push(["Thời gian đăng ký", data.time]);

  var tableRowsHtml = rows
    .filter(function (r) { return r[1] !== undefined && r[1] !== null && r[1] !== ""; })
    .map(function (r) {
      return '<tr>' +
        '<td style="padding:10px 14px;border:1px solid #e2e8f0;background:#f8fafc;font-weight:600;color:#334155;width:40%;">' + r[0] + '</td>' +
        '<td style="padding:10px 14px;border:1px solid #e2e8f0;color:#1e293b;">' + r[1] + '</td>' +
        '</tr>';
    })
    .join('');

  var htmlBody =
    '<div style="font-family:Arial,Helvetica,sans-serif;max-width:600px;margin:0 auto;background:#ffffff;">' +
      '<div style="background:linear-gradient(135deg,#4A90E2,#357ABD);padding:24px 30px;border-radius:10px 10px 0 0;">' +
        '<h2 style="color:#ffffff;margin:0;font-size:20px;">Đăng ký thành công!</h2>' +
      '</div>' +
      '<div style="padding:24px 30px;border:1px solid #e2e8f0;border-top:none;border-radius:0 0 10px 10px;">' +
        '<p style="color:#334155;font-size:15px;">Chào <b>' + (data.name || '') + '</b>,</p>' +
        '<p style="color:#334155;font-size:15px;">Hệ thống <b>BaomauConnect</b> đã nhận được hồ sơ đăng ký của bạn với thông tin chi tiết như sau:</p>' +
        '<table style="border-collapse:collapse;width:100%;margin:16px 0;font-size:14px;">' +
          tableRowsHtml +
        '</table>' +
        '<p style="color:#334155;font-size:15px;">Đội ngũ BaomauConnect sẽ kiểm tra và liên hệ lại với bạn trong thời gian sớm nhất qua số điện thoại đã đăng ký.</p>' +
        '<p style="color:#94a3b8;font-size:13px;margin-top:24px;">Đây là email tự động, vui lòng không trả lời trực tiếp email này.</p>' +
        '<p style="color:#334155;font-size:15px;">Trân trọng,<br><b>Đội ngũ BaomauConnect.asia</b></p>' +
      '</div>' +
    '</div>';

  var plainBody = rows
    .filter(function (r) { return r[1] !== undefined && r[1] !== null && r[1] !== ""; })
    .map(function (r) { return "- " + r[0] + ": " + r[1]; })
    .join("\n");

  GmailApp.sendEmail(data.userEmail, subject,
    "Chào " + (data.name || '') + ",\n\nHệ thống BaomauConnect đã nhận được hồ sơ đăng ký của bạn:\n\n" +
    plainBody +
    "\n\nĐội ngũ BaomauConnect sẽ liên hệ lại sớm nhất.\n\nTrân trọng,\nĐội ngũ BaomauConnect.asia",
    {
      htmlBody: htmlBody,
      name: "BaomauConnect"
    }
  );
}