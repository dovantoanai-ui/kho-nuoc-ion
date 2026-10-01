# Hướng dẫn: Giao VIBA — kho VIBA, công giao, báo cáo (01/10/2026)

App hoá đơn → tab **Giao VIBA** có 4 tab con: **Đơn giao · Kho VIBA · Báo cáo · Cài đặt**.

## Triển khai (làm 1 lần, đúng thứ tự)
1. `hoadon-backend.gs` → dán đè vào project Apps Script hoá đơn → Deploy → Manage deployments → Edit → **New version** (giữ URL /exec).
2. `backend-hn-apps-script.gs` → dán đè vào Apps Script của Sheet **ION Kho HN** → Deploy → New version. Chạy tay hàm `testKhoViba` để kiểm tra (xem Logger).
3. Đẩy `hoadon.html` lên GitHub (`git add -A && git commit && git push`).
   Đẩy app SAU khi đã deploy backend — nếu không, thông tin phiếu VIBA lúc xác nhận giao sẽ không được lưu.

## Bắt đầu dùng
0. Ảnh lưu Drive: nếu báo `LOI ANH` thì chạy tay hàm `capQuyen` trong Apps Script hoá đơn 1 lần.
1. **Cài đặt**: kiểm tra đơn giá (mặc định theo bảng VIBA đề xuất), nhập **kg/đơn vị** từng mã (để tính bốc xếp), VAT mặt bằng nếu có, ngày bắt đầu dùng kho VIBA (mặc định 01/10/2026) → Lưu.
2. **Kho VIBA** → phiếu **Tồn đầu kỳ**: số hàng + vỏ đang nằm ở kho VIBA tại ngày bắt đầu.
3. Mọi phiếu kho VIBA đều **bắt buộc kèm ảnh** (lưu Drive, folder `Anh kho VIBA ION FUJI`). Kho VIBA **độc lập**: phiếu **Nhập kho** chỉ cộng kho VIBA (hàng từ HN, Hưng Yên hay Hạ Long đều được). Hàng rời kho nào thì lập phiếu xuất ở app kho đó như thường. VIBA trả hàng/vỏ: phiếu **Trả về ION**. Sửa tồn sau kiểm kho: phiếu **Điều chỉnh tồn**.
   Đơn VIBA giao từ ngày bắt đầu KHÔNG trừ kho HN nữa (xuất từ kho VIBA).

## Hằng ngày — nhập đơn VIBA (SAU KHI VIBA đã giao xong)
- Tab **Giao VIBA → Nhập đơn**: form giống hoá đơn thường (khách, chiết khấu, sản phẩm, VAT, thanh toán, cọc, ghi chú) + khối **Phiếu giao VIBA**: số phiếu, ngày giao, NV giao, địa chỉ, vỏ thu về, lên tầng / ngoại thành, người nhập.
- Bấm **Lưu đơn VIBA đã giao** → ghi thẳng HoaDon (ngày = ngày giao): doanh thu, công nợ, trừ kho VIBA. Xong bấm **Chia sẻ / Lưu ảnh** gửi hoá đơn cho khách.
- App tự tính công giao và báo đỏ nếu khách giữ vỏ mà chưa thu cọc.
- Tab Hoá đơn **không còn** tích "Giao qua VIBA". **Không còn trạng thái "Chờ giao"**: đơn VIBA chỉ nhập khi đã giao thành công. Kho VIBA trừ hàng từ ngày có phiếu kho VIBA đầu tiên (18/09/2026).
- Đơn đã giao thiếu thông tin: **Đã giao** → Xem "Đã giao, thiếu thông tin VIBA" → **Bổ sung thông tin VIBA**.

## Kiểm kho VIBA (biên bản đối chiếu)
Kho VIBA → **Kiểm kho VIBA**: chọn ngày, người kiểm, người chứng kiến bên VIBA, nhập số đếm thực tế từng mã (để trống = không kiểm), **chụp ảnh biên bản (bắt buộc)**.
App hiện tồn sổ sách tính đến hết ngày kiểm + chênh lệch. Lưu = biên bản (tab `KiemKhoVIBA`), **không sửa tồn**. Muốn sửa thì lập phiếu "Điều chỉnh tồn".

## Cuối tháng
**Báo cáo** → chọn tháng → xem Nhập–xuất–tồn, Công theo NV, Tiền phải trả VIBA → **Xuất Excel theo mẫu VIBA**
(Bảng Nhập Xuất Tồn · Tính Công Giao Hàng · Thanh toán VIBA · Chi tiết đơn · Kiểm kho nếu tháng có biên bản).

## Quy tắc tính
- Công: Bình/thùng 12.000đ/SP nội thành, 15.000đ ngoại thành; can 5L 6.000đ. Điểm giao < 10 SP: trả cố định 100.000đ (thay cho tính theo SP). Lên tầng +50.000đ/điểm.
- Hoa hồng 6% × tiền hàng chưa VAT (không tính cọc) của đơn VIBA đã giao trong tháng.
- Mặt bằng 5.000.000đ/tháng (chỉ tính từ tháng bắt đầu dùng kho VIBA). Bốc xếp = tấn hàng "Nhập kho" trong tháng × đơn giá/tấn.
- Báo cáo tính lại theo đơn giá HIỆN TẠI trong Cài đặt.

## Dữ liệu (Sheet DA05)
- `GiaoVIBA`: thêm 7 cột (So phieu VIBA, NV giao, So SP, Vo thu, Len tang, Ngoai thanh, Cong giao) — backend tự thêm.
- `KhoVIBA`: phiếu kho VIBA (mỗi dòng 1 mã hàng). `KiemKhoVIBA`: biên bản kiểm kho. `CaiDatVIBA`: cài đặt (1 dòng JSON).
- Backup trước khi sửa: `*_backup_20261001_viba.*` cùng thư mục.

## Đối chiếu kiểm kê (01/10/2026)
- Kho VIBA → thẻ **Đối chiếu kiểm kê**: chọn biên bản + thời gian sổ sách (Tất cả thời gian / Từ lần kiểm trước / Tháng của biên bản / Từ ngày…).
- Bảng Sổ – Đếm – Lệch. Bấm từng mã để xem sổ tính từ đâu (phiếu nhập, từng phiếu giao).
- Mã thiếu: liệt kê hoá đơn kho cùng mặt hàng KHÔNG có phiếu VIBA trong kỳ (chỉ là gợi ý), đánh dấu đơn/ngày trùng đúng số lệch.
- Nút **Xuất Excel đối chiếu**.
- Báo cáo VIBA: ô **Kỳ báo cáo** có thêm "Tất cả thời gian" (mặt bằng tính theo số tháng từ ngày bắt đầu).

## Sheet riêng cho nhân viên VIBA (01/10/2026)
- File "VIBA - Kho ION FUJI Ha Noi" (backend tự tạo lần đầu, nằm cạnh sheet DA05). Link ở Giao VIBA → Cài đặt → "Sheet cho nhân viên VIBA".
- 5 tab: Tổng quan · Nhập xuất tồn · Kiểm kho · Tiền giao hàng · Doanh số. Không có hoa hồng %, mặt bằng, bốc xếp.
- Tự cập nhật sau mỗi lần lưu đơn VIBA / phiếu kho / kiểm kho / cài đặt. Nút "Cập nhật sheet VIBA ngay" để bấm tay.
- Sheet bị ghi đè mỗi lần cập nhật → KHÔNG sửa tay trong sheet VIBA. Share quyền Người xem.
