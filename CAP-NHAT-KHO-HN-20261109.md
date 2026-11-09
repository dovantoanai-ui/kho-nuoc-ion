# Cập nhật 09/11/2026 — Kho Hà Nội trên app Hóa đơn (V-261109-07)

Backup trước khi sửa: `hoadon_backup_20261109.html`, `hoadon-backend_backup_20261109.gs`.

## Cái gì mới (màn Hóa đơn, dưới "+ Thêm sản phẩm")
| Nút | Làm gì | Ghi vào đâu (Sheet duan-nuoc-ion) |
|---|---|---|
| 📥 Nhập kho | Phiếu nhập kho HÀ NỘI (không phải kho VIBA). Dòng "Vỏ 19L" = thu vỏ lẻ | tab mới `PhieuKhoHN` |
| 📋 Kiểm kê kho | Đếm thực tế kho HN → thành MỐC tồn | tab mới `KiemKeHN` |
| 🔄 Thu vỏ bình | Bấm +/− số vỏ, gắn vào hóa đơn đang lập; cộng vào tồn vỏ HN khi lưu; ghi chú hóa đơn thêm "Thu vỏ n bình" | `PhieuKhoHN` (Loại "Thu vỏ") |
| 📊 Tồn kho | Bảng tồn kho HN | (chỉ tính, không ghi) |

Cọc bình: có Bình 19L → báo "cần cọc SL × 50.000 đ", nút "Điền cọc". Lưu hóa đơn khi thiếu cọc bị chặn trừ khi tích "Không cọc" + ghi chú ≥ 5 ký tự (ghi chú tự gắn `[KHÔNG CỌC]` / `[CỌC THIẾU x/y]`).

## Cách tính tồn kho HN
Tồn = mốc kiểm kê gần nhất + nhập/thu vỏ SAU ngày mốc − bán (hóa đơn Kho = "Hà Nội", TRỪ đơn giao VIBA). Chưa kiểm kê → tính từ 0, nên kiểm kê trước.

## Giả định em đã chọn (sửa được)
- Cọc 50k áp mọi khách (hằng `COC_BINH` trong hoadon.html).
- Thu vỏ chỉ ghi số vỏ, không trừ/hoàn cọc.
- "Bình 19L" = mọi sản phẩm có "19L" trong tên (không phải vỏ).
- Nhập kho HN chưa có ảnh phiếu.
- Mã hàng dùng chung SP01–SP07 như kho VIBA (khớp theo tên hàng).

## BẮT BUỘC để chạy
1. Dán `hoadon-backend.gs` mới vào Apps Script của Sheet duan-nuoc-ion → Deploy → New version.
2. Đẩy `hoadon.html` lên GitHub (như `_day-len-20261105.bat`).
Chưa deploy backend thì nút vẫn mở nhưng lưu sẽ không vào Sheet (app báo "Đã lưu" vì gửi no-cors).
