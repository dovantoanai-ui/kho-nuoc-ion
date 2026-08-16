# Tự động trừ tồn kho theo hoá đơn (16/08/2026)

Nhập bán hàng ở **app Hoá đơn** → tồn kho ở **app kho tương ứng** (Hà Nội `hn.html` / Hưng Yên `index.html`) **tự trừ**, không phải nhập 2 lần.

## Cách hoạt động

```
App Hoá đơn ──ghi──▶ Sheet DA05, tab HoaDon (cột Kho: Hà Nội / Hưng Yên)
                             │  backend kho ĐỌC thêm tab này khi tính tồn
              ┌──────────────┴──────────────┐
   Kho = "Hưng Yên"                Kho = "Hà Nội"
   backend kho HY                  backend kho HN
              ▼                             ▼
Tồn = Tồn đầu + Nhập − Xuất(kho) − Bán(hoá đơn) − Tặng + Thu hồi (− Trả HY)
```

- Chỉ hoá đơn **từ 16/08/2026** trở đi mới trừ tồn (hoá đơn cũ có thể đã nhập phiếu Xuất tay → tránh trừ đúp). Muốn đổi mốc: sửa `HOADON_TU_NGAY` trong 2 file backend kho.
- Khớp sản phẩm theo **tên hàng** (đã chuẩn hoá bỏ dấu/hoa thường). Kho HN tra tên qua cả Data kho HY — nguồn sinh ra tên trên hoá đơn — nên khớp chính xác.
- Không ghi thêm dòng nào vào Sheet kho: sửa/xoá hoá đơn trên Sheet DA05 thì tồn tự đúng lại.
- 3 file HTML **giữ nguyên**, chỉ deploy lại 3 backend.

## Deploy (3 backend, mỗi cái ~2 phút)

**1. Backend kho Hưng Yên** — mở Sheet `kho-ion` → Extensions → Apps Script:
   1. Dán toàn bộ `backend-apps-script.gs` mới đè lên code cũ → Lưu.
   2. Chọn hàm **`testDocHoaDon`** → Run (lần đầu Google hỏi quyền → Cho phép). Xem Log: phải báo đọc được hoá đơn, không cảnh báo thiếu Mã.
   3. Deploy → **Manage deployments** → Edit (bút chì) → Version: **New version** → Deploy. URL `/exec` giữ nguyên.

**2. Backend kho Hà Nội** — mở Sheet `ION Kho HN` → Extensions → Apps Script:
   1. Dán toàn bộ `backend-hn-apps-script.gs` mới đè lên → Lưu.
   2. Chạy tay hàm **`capNhatDanhMucTuHY`** 1 lần → đồng bộ tab Data theo kho HY (đủ 13 SKU, tên có dấu, giữ nguyên Tồn đầu). Anh mở tab Data điền lại **Ton dau** thực tế kho HN nếu cần.
   3. Chạy **`testDocHoaDon`** → xem Log như trên.
   4. Deploy → Manage deployments → New version.
   5. Mở app `hn.html` → Danh mục → **Tải lại danh mục từ Google Sheet**.

**3. Backend Hoá đơn** — mở project Apps Script của Sheet DA05 (hoặc script.google.com):
   1. Dán toàn bộ `hoadon-backend.gs` mới đè lên → Lưu.
   2. Deploy → Manage deployments → New version.
   - Thêm mới: bấm "đồng bộ" hoá đơn nhiều lần **không còn ghi trùng dòng** (bỏ qua ID đã có).

## Quy tắc vận hành từ 16/08/2026 (quan trọng)

1. **Đơn nào đã nhập ở app Hoá đơn thì KHÔNG nhập phiếu Xuất ở app kho nữa** — kể cả đơn đại lý. Nhập cả 2 nơi = trừ tồn 2 lần.
2. App kho chỉ còn dùng cho: Nhập kho, điều chuyển HY↔HN, Tặng/Thu hồi không hoá đơn, Trả về HY, Kiểm kê.
3. **Không đổi tên hàng** trong tab Data kho HY tuỳ tiện — tên là khoá khớp giữa hoá đơn và kho. Muốn đổi tên: đổi xong chạy lại `capNhatDanhMucTuHY` bên HN, và hiểu là hoá đơn cũ mang tên cũ sẽ không khớp nữa.
4. Sản phẩm bán trên hoá đơn mà không có trong Data kho → không trừ được tồn (vẫn tính doanh thu). Chạy `testDocHoaDon` định kỳ để phát hiện dòng lệch.

## Lưu ý còn lại

- Báo cáo email/tab BaoCao (`email-baocao.gs`) và agent giám sát kho–tài chính **chưa** gộp hoá đơn theo cách này — số "Ban" ở đó vẫn là số cũ. Cần thì làm bước sau.
- File backup cùng thư mục: `*_backup_20260816.gs`. Lỗi thì dán lại file backup + Deploy New version là về như cũ.
