# 1. Di chuyển vào thư mục code trên máy của bạn
cd ~/Documents/baomau-connect

# 2. Xóa bỏ liên kết cũ kết nối với tài khoản cá nhân
git remote remove origin

# 3. Kết nối với kho chứa mới của tổ chức bằng đường SSH siêu tốc
git remote add origin git@github.com:baomauconnect/baomauconnect.github.io.git

# 4. Đẩy toàn bộ code lên mạng
git push -u origin main
