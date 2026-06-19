## 📝 Hệ Thống Kết Nối Bảo Mẫu Tại Gia - BaomauConnect

Hệ thống **BaomauConnect** là một nền tảng web single-page (SPA) gọn nhẹ, giúp tối ưu hóa quy trình kết nối thời gian thực (Real-time) giữa **Bảo mẫu** có nhu cầu tìm việc và **Phụ huynh** có nhu cầu thuê người chăm sóc trẻ. Dữ liệu được đồng bộ trực tiếp với hệ thống đám mây thông qua Google Sheets (sử dụng Google Apps Script làm API backend).

## 🛠️ Danh Sách Chức Năng Chi Tiết

### 1. Giao diện Phía Người Dùng (`index.html`)

Giúp người dùng dễ dàng khai báo và gửi thông tin lên hệ thống dựa theo 2 vai trò cốt lõi:

* **Lựa chọn vai trò trực quan:** Sử dụng thẻ điều hướng động để chuyển đổi nhanh giữa biểu mẫu cho Bảo mẫu và Phụ huynh mà không cần tải lại trang.
* **Hồ sơ Ứng tuyển Bảo mẫu:**
* Thu thập thông tin cá nhân cơ bản: Họ tên (bắt buộc nhập đầy đủ họ và tên), tuổi tác (giới hạn từ 18 - 70 tuổi), số điện thoại, ảnh chân dung cá nhân.
* **Đồng bộ địa chính động:** Gọi API từ dịch vụ `provinces.open-api.vn` để tự động tải danh sách Tỉnh/Thành phố, Quận/Huyện, Xã/Phường tương ứng theo thời gian thực.
* **Linh hoạt khung giờ làm việc:** Lựa chọn làm việc *Toàn thời gian* hoặc *Bán thời gian* (Tự động mở cấu trúc chọn khung giờ cụ thể từ `00:00` đến `24:00`).
* **Xử lý định dạng lương tự động:** Cho phép nhập số tiền mong muốn và chọn đơn vị tính (`đ/tháng` hoặc `đ/giờ`). Hệ thống tự động định dạng hiển thị dạng chuỗi phân cách hàng nghìn (Ví dụ: `9.000.000đ/tháng`).
* Khai báo chuyên sâu: Đánh dấu khả năng nhận chăm sóc trẻ sơ sinh (Newborn) và bộ lọc kỹ năng/kinh nghiệm tích lũy.


* **Phiếu Yêu cầu từ Phụ huynh:**
* Thu thập thông tin người đại diện liên hệ, số điện thoại, địa chỉ cần bảo mẫu đến làm việc (áp dụng đồng bộ địa chính động).
* Khai báo chi tiết về trẻ: Độ tuổi hiện tại của bé và yêu cầu mong muốn về độ tuổi của bảo mẫu.
* Cài đặt ngân sách chi trả (áp dụng bộ xử lý định dạng tiền tệ tương tự biểu mẫu bảo mẫu) cùng các mô tả công việc, lưu ý riêng biệt.


* **Ràng buộc & Kiểm tra dữ liệu (Validation):** Khóa nút gửi khi đang xử lý nhằm tránh spam dữ liệu, tự động kiểm tra định dạng số điện thoại Việt Nam (phải có đúng 10 chữ số và bắt đầu bằng số 0).

---

### 2. Bảng Điều Khiển Quản Trị Đám Mây (`admin.html`)

Hệ thống backend mini phục vụ cho điều phối viên quản lý danh sách dữ liệu đổ về từ biểu mẫu:

* **Xác thực bảo mật cơ bản (Cloud Auth Layer):**
* Tích hợp lớp mã hóa mật khẩu bằng thuật toán **MD5** thông qua thư viện `CryptoJS`.
* Yêu cầu mật khẩu để ẩn/hiển thị nội dung quản trị viên (Hệ thống so khớp chuỗi hash tĩnh có sẵn nhằm đảm bảo quyền truy cập thô).


* **Bảng Thống kê Tổng quan (Live Stats Dashboard):** hiển thị trực quan tổng số lượng đơn hiện có trên hệ thống đám mây, bóc tách chi tiết số lượng hồ sơ từ Bảo mẫu và yêu cầu từ Phụ huynh.
* **Bộ lọc trạng thái dữ liệu (Toolbar Filter):** Cho phép lọc nhanh danh sách hiển thị theo nhu cầu: *Tất cả đơn*, *Chỉ Bảo mẫu*, hoặc *Chỉ Phụ huynh*.
* **Bảng dữ liệu tương tác động (Interactive Data Table):**
* Hiển thị thông tin trực quan theo từng cột (Thời gian nhận, Vai trò, Họ tên & SĐT, Khung giờ, Địa chỉ, Yêu cầu).
* Tự động hiển thị nhãn `Nhận sơ sinh` nổi bật đối với các hồ sơ bảo mẫu có khả năng chăm trẻ Newborn.
* Xem ảnh đại diện thu nhỏ (Avatar preview) trực tiếp trên dòng nếu bảo mẫu có cung cấp liên kết ảnh hợp lệ.


* **Chức năng Cập nhật (Update - CRUD):** Tích hợp cửa sổ biểu mẫu thông minh (`SweetAlert2`) tự động bắt dữ liệu theo dòng trên Google Sheets, cho phép sửa đổi bất kỳ trường thông tin nào và lưu đồng bộ ngược lại Cloud.
* **Chức năng Xóa dữ liệu (Delete - CRUD):**
* Xóa đơn lẻ từng hồ sơ.
* Xóa hàng loạt (Bulk Delete): Cho phép tích chọn hộp kiểm (`Checkbox`) nhiều dòng cùng lúc để thực hiện một lệnh xóa diện rộng, giúp dọn dẹp hòm thư dữ liệu nhanh chóng.



---

## 🌐 Ngăn Xếp Công Nghệ Sử Dụng (Tech Stack)

* **Frontend:** HTML5, CSS3 (Google Fonts Inter, Flexbox, Grid Layout), JavaScript thuần (Vanilla JS).
* **Thư viện hỗ trợ UI/UX:**
* `Font-Awesome 6.4.0` (Hệ thống biểu tượng trực quan).
* `SweetAlert2` (Hộp thoại thông báo và biểu mẫu pop-up hiện đại).


* **Thư viện Bảo mật:** `CryptoJS 4.1.1` (Mã hóa MD5).
* **Cloud API Backend:** Kết nối bất đồng bộ (`fetch API`) tới **Google Apps Script URL** hỗ trợ 2 phương thức chính:
* `GET`: Tải toàn bộ cấu trúc mảng JSON từ Google Sheets.
* `POST`: Truyền tham số payload `action: 'update'` hoặc `action: 'delete'` dạng chuỗi dữ liệu (URL-encoded) để ghi đè hoặc loại bỏ hàng trong trang tính.
