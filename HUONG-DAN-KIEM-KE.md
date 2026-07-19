# Hướng dẫn bật màn Kiểm kê kho (thêm 18/07/2026)

Đã thêm màn **"Kiểm kê"** vào app kho (index.html) + backend (backend-apps-script.gs). Backup bản cũ: `index_backup_20260718b.html`, `backend-apps-script_backup_20260718.gs`.

## Bước 1 — Cập nhật Backend (bắt buộc, làm trước)
1. Mở Sheet **kho-ion** → menu **Tiện ích mở rộng (Extensions)** → **Apps Script**.
2. Mở file `backend-apps-script.gs`, **thay toàn bộ nội dung** bằng bản mới (file em vừa gửi). (Thực chất chỉ thêm 2 chỗ: khai báo `HEAD.KiemKe` và 1 dòng `if (p.kiemke ...)` trong `doPost`.)
3. **Lưu** (Ctrl+S).
4. **Deploy lại**: nút **Deploy** → **Manage deployments** → bấm bút chì (Edit) ở deployment đang chạy → mục **Version** chọn **New version** → **Deploy**.
   - ⚠️ Phải tạo **version mới** thì `doPost` mới nhận được dữ liệu kiểm kê. Không redeploy = tab KiemKe không được ghi.
5. Tab **"KiemKe"** trên Sheet sẽ **tự tạo** khi đồng bộ phiếu kiểm kê đầu tiên — anh không cần tạo tay.

## Bước 2 — Cập nhật App
- Đẩy `index.html` mới lên GitHub Pages như mọi lần (chạy `push-app.ps1` hoặc `day-github.ps1`).
- Mở app trên điện thoại → refresh (nếu là app đã cài, có thể cần đóng mở lại). Sẽ thấy nút **"Kiểm kê"** mới ở thanh dưới.

## Bước 3 — Dùng hàng tuần
1. Vào tab **Kiểm kê** → chọn kho (**Hưng Yên** hoặc **Hà Nội – Huy**).
2. Chọn ngày, nhập tên người đếm.
3. Gõ **SL đếm** từng mã. Cột **Tồn sổ** (số app đang tính) hiện sẵn để đối chiếu; cột **Lệch** tự nhảy (xanh dư, đỏ thiếu).
4. Bấm **Lưu phiếu kiểm kê** → tự đồng bộ lên tab KiemKe.
5. Kho **Hà Nội** không có số app nên cột Tồn sổ để trống — Huy chỉ đếm tay rồi lưu.

## Kết nối với agent giám sát
- Agent giám sát kho + tài chính DA05 **tự đọc tab KiemKe** mỗi lần chạy, lấy dòng ngày mới nhất theo từng (Kho, Mã) làm **mốc tồn chuẩn**.
- Kho Hưng Yên: tồn = kiểm kê tuần gần nhất + nhập-xuất sau đó. Kho Hà Nội: tồn = số Huy đếm tuần gần nhất.
- Nhờ đó bảng tồn luôn sát thực tế, và mỗi tuần lộ ra chênh lệch (như Lux500 hụt 12 thùng đợt kiểm kê 13/07) để anh truy.

*Ghi chú: app không dùng service worker cache mạnh nên refresh là thấy bản mới. Nếu vẫn thấy bản cũ, xóa app đã cài rồi cài lại từ trình duyệt.*
