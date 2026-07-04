/**
 * Hàm xử lý dữ liệu từ form và tự động gửi email xác nhận
 */
function doPost(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var payload = JSON.parse(e.postData.contents);
    
    // 1. Kiểm tra hành động (Xóa/Cập nhật/Thêm mới)
    if (payload.action === "delete") {
      return handleDelete(sheet, payload);
    } else if (payload.action === "update") {
      return handleUpdate(sheet, payload);
    } else {
      // 2. Thêm mới dữ liệu (Case mặc định)
      var time = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
      sheet.appendRow([
        time, payload.role, payload.category, payload.name, payload.phone, 
        payload.address, payload.age, payload.avatar, payload.childAge, 
        payload.bmAgeReq, payload.workTime, payload.careNeonatal, payload.detail,
        payload.userEmail // Lưu email người dùng để đối soát
      ]);

      // 3. TỰ ĐỘNG GỬI EMAIL THÔNG BÁO CHO NGƯỜI DÙNG
      sendConfirmationEmail(payload);

      return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Hàm gửi email sử dụng GmailApp (đơn giản, hiệu quả thay cho Java)
 */
function sendConfirmationEmail(data) {
  var subject = "Thông báo: " + data.category + " đã được ghi nhận";
  var body = "Chào " + data.name + ",\n\n" +
             "Hệ thống BaomauConnect đã nhận được thông tin của bạn.\n\n" +
             "Chi tiết yêu cầu:\n" +
             "- Vai trò: " + data.role + "\n" +
             "- Địa chỉ: " + data.address + "\n" +
             "- Nội dung: " + data.detail + "\n\n" +
             "Chúng tôi sẽ kiểm tra và phản hồi lại sớm nhất.\n" +
             "Trân trọng,\nĐội ngũ BaomauConnect.";

  // Gửi email tới email mà người dùng đã đăng nhập (lấy từ payload)
  GmailApp.sendEmail(data.userEmail, subject, body);
}

// Giữ nguyên các hàm bổ trợ handleDelete, handleUpdate... (đảm bảo logic của bạn không bị gãy)