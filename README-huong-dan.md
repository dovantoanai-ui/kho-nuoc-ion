# App Kho Nước Ion Kiềm — DA05

Hai file:
- `kho-nuoc-ion.html` — app (mở trực tiếp bằng trình duyệt, chạy điện thoại + máy tính).
- `backend-apps-script.gs` — code đổ dữ liệu vào Google Sheet.

## Cài đặt (làm 1 lần, ~10 phút)

**B1. Tạo Google Sheet kho MỚI** (riêng, không dùng chung Sheet hạch toán chi phí DA05).

**B2. Gắn backend:** trong Sheet đó → menu **Tiện ích mở rộng (Extensions) → Apps Script** → xóa code mẫu → dán toàn bộ `backend-apps-script.gs` → Lưu.

**B3. Tạo danh mục:** trong Apps Script, chọn hàm **`setupDataTab`** → **Run** (lần đầu Google hỏi quyền → Cho phép). Xong sẽ có tab **Data** + 8 sản phẩm mẫu.

**B4. Deploy Web App:** **Deploy → New deployment → chọn Web app** → Execute as: **Me**, Who has access: **Anyone** → Deploy → **copy URL kết thúc bằng `/exec`**.

**B5. Nối app với Sheet:** mở app → tab **Danh mục** (dưới cùng) → dán **URL Web App (`/exec`)** → **Lưu kết nối + tải danh mục**. App tự lấy danh mục + giá **100% từ tab Data** (qua JSONP, chạy cả khi Sheet để riêng tư). Không cần Sheet ID nữa.

> ⚠️ **QUAN TRỌNG — để Đồng bộ (ghi) chạy được:** app phải mở qua **link web (https)**, KHÔNG mở bằng cách double-click file (`file://`). Trình duyệt chặn gửi dữ liệu ra ngoài khi chạy từ file cục bộ. Cách nhanh nhất: deploy lên Netlify/Cloudflare (hoặc ghép vào nemthanh-hub đã có sẵn deploy) để có link https + cài được iPhone. Việc **đọc danh mục** thì chạy được cả từ file.

## Dùng hằng ngày

- **Nhập / Xuất ĐL / Bán lẻ / Thu tiền:** nhập xong bấm Lưu — dữ liệu lưu trên máy trước.
- Cuối buổi bấm **Đồng bộ** (góc trên phải) → dữ liệu ngày đó nhảy vào Google Sheet (tab NhapKho / XuatKho / ThuTien). Đồng bộ lại nhiều lần **không** bị trùng dòng (khử theo LineID).
- **Tồn kho** = tồn đầu kỳ (8 SKU nhập lần 1) + tổng nhập − tổng xuất, tính trên máy.
- **Báo cáo:** doanh thu, tổng nhập, công nợ đại lý, giá trị tồn theo Hôm nay / Tháng / Tất cả.

## Cài lên iPhone (PWA)

Mở file (sau khi deploy web lên Netlify/Cloudflare) bằng **Safari → nút Chia sẻ → Thêm vào Màn hình chính**. Chạy như app.

## Còn cần anh

- Điền **giá bán 4 bậc (NPP/C1/C2/C3) + giá lẻ** trong tab Data (hoặc trong app → Danh mục). Giá vốn đang tạm lấy từ giá tồn MISA (có mã thấp bất thường — anh soát lại).
- Muốn em **deploy lên web** (Netlify/Cloudflare) để dùng link + cài iPhone, hay **ghép vào nemthanh-hub** cạnh kho-hn/kho-th thì báo em.

## Kỹ thuật (theo chuẩn skill html-business-app)

1 file HTML đơn · localStorage (JSON nhất quán) · ngày local (không toISOString) · fetch no-cors · Apps Script upsert theo LineID · validate JS pass. Chưa gắn PWA manifest/service worker (dùng "Thêm vào màn hình chính" là đủ) — cần bản cài offline đầy đủ thì em bổ sung.
