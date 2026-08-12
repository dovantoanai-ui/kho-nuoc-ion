# App Kho Nước Ion — Hà Nội (hn.html)

App kho riêng cho kho Hà Nội, chạy song song với app kho Hưng Yên (`index.html`).
Hai app **dùng hai Google Sheet khác nhau**, dữ liệu tách bạch hoàn toàn.

| | Kho Hưng Yên (kho tổng) | Kho Hà Nội |
|---|---|---|
| App | `index.html` | `hn.html` |
| Link | `.../kho-nuoc-ion/` | `.../kho-nuoc-ion/hn.html` |
| Sheet | `kho-ion` (đang chạy) | Sheet **mới**, anh tạo ở Bước 1 |
| Backend | `backend-apps-script.gs` | `backend-hn-apps-script.gs` |
| Khoá localStorage | `kda_...` | `khn_...` |

> Vì hai app cùng tên miền `dovantoanai-ui.github.io`, chúng dùng chung bộ nhớ trình duyệt. Prefix `khn_` là thứ giữ cho dữ liệu kho HN không đè lên kho Hưng Yên trên cùng một máy/điện thoại. **Không đổi prefix này.**

---

## Cài đặt (làm 1 lần, ~10 phút)

**B1. Tạo Google Sheet MỚI** — đặt tên ví dụ `kho-ion-hanoi`. Không dùng lại Sheet `kho-ion` của Hưng Yên.

**B2. Gắn backend:** trong Sheet vừa tạo → **Tiện ích mở rộng (Extensions) → Apps Script** → xoá code mẫu → dán toàn bộ nội dung `backend-hn-apps-script.gs` → Lưu.

**B3. Tạo cấu trúc:** trong Apps Script chọn hàm **`setupDataTab`** → **Run** (lần đầu Google hỏi quyền → Cho phép).
Xong sẽ có 5 tab: **Data** (8 SKU mẫu), **NhapKho**, **XuatKho**, **TraVeHY**, **KiemKe**.

**B4. Deploy Web App:** **Deploy → New deployment → Web app** → Execute as: **Me** — Who has access: **Anyone** → Deploy → copy URL kết thúc bằng `/exec`.

**B5. Nối app với Sheet:** mở `https://dovantoanai-ui.github.io/kho-nuoc-ion/hn.html` → tab **Danh mục** (dưới cùng) → dán URL `/exec` → **Lưu kết nối + tải danh mục**.

**B6. Điền tồn đầu kỳ:** trong tab **Data**, cột **Ton dau** — điền số lượng thực tế đang có ở kho Hà Nội tại ngày bắt đầu dùng app. Để 0 nghĩa là kho bắt đầu từ rỗng.

> Sau khi sửa tab Data, bấm **Tải lại danh mục từ Google Sheet** trong app để app lấy số mới.

---

## Dùng hằng ngày

**Nhập kho** — có 2 nguồn:

- *Từ kho Hưng Yên*: nhận hàng điều chuyển từ kho tổng. Giá vốn tự lấy theo danh mục (chỉ để tính giá trị tồn, không phải mua hàng).
- *NCC trực tiếp*: kho HN tự đặt hàng nhà cung cấp — gõ tên NCC và giá vốn thực tế.

**Xuất kho** — giao đại lý / khách / nhân viên. Ba loại dòng:

| Loại | Ý nghĩa | Tác động tồn |
|---|---|---|
| Xuất | giao hàng bình thường | trừ tồn |
| Tặng | hàng tặng, không thu tiền | trừ tồn |
| Thu hồi | khách trả hàng về kho | cộng lại tồn |

**Trả về HY** — hàng Hà Nội trả ngược lên kho tổng Hưng Yên (cận date, thừa, đổi mẫu). Phiếu này **trừ tồn kho HN**.

**Kiểm kê** — gõ số đếm thực tế, app hiện cột Lệch so với tồn sổ (xanh = dư, đỏ = thiếu).

**Đồng bộ** — bấm nút góc trên phải sau khi lưu phiếu. Đồng bộ lại nhiều lần **không** trùng dòng (khử theo LineID). App tự đồng bộ ngay sau mỗi lần lưu nếu đã nối Sheet.

**Bán lẻ / đơn lẻ**: nhập ở app Hoá đơn, không nhập ở app này.

---

## Công thức tồn kho Hà Nội

```
Tồn = Tồn đầu (tab Data) + Nhập − Xuất − Tặng + Thu hồi − Trả về HY
```

Đã kiểm bằng dữ liệu giả (2 SKU, 3 loại phiếu): bảng Tồn kho, bảng Tổng hợp mã hàng, bảng Nhập–Xuất–Tồn theo ngày và cột Tồn sổ ở Kiểm kê đều ra cùng một con số.

---

## Đối soát hai kho (làm cuối tuần / cuối tháng)

Hai kho là hai sổ độc lập, nên mỗi lần chuyển hàng phải có **hai phiếu khớp nhau**:

| Chiều hàng đi | Kho Hưng Yên ghi | Kho Hà Nội ghi |
|---|---|---|
| HY → HN | tab `XuatKho`, đối tượng nhận = Huy / kho HN | tab `NhapKho`, Nguồn = `Kho Hung Yen` |
| HN → HY | tab `NhapKho` | tab `TraVeHY` |

Cách soát: lọc tab `NhapKho` của Sheet HN theo cột **Nguon = Kho Hung Yen**, cộng số lượng theo mã, đem so với phần xuất cho HN bên Sheet Hưng Yên trong cùng kỳ. Lệch thì tìm phiếu thiếu ở một trong hai bên.

> Lưu ý nghiệp vụ: từ trước tới nay "Xuất kho HN (Huy)" bên Hưng Yên được coi gần như đã bán. Khi kho HN có sổ riêng thì hàng ở HN vẫn là **tồn của công ty**, chưa phải doanh thu — doanh thu vẫn chỉ ghi nhận ở app Hoá đơn khi bán thật. Anh cân nhắc chốt lại cách ghi nhận này với agent giám sát kho + tài chính DA05 để hai bên không đếm trùng.

---

## Cài lên iPhone

Mở link `hn.html` bằng **Safari → nút Chia sẻ → Thêm vào Màn hình chính**. Nên đặt tên icon là **Kho HN** để không nhầm với app kho Hưng Yên.

---

## Kỹ thuật

1 file HTML đơn · vanilla JS · localStorage prefix `khn_` · hàm ngày local (không `toISOString`) · onclick gắn bằng DOM closure cho các dòng sản phẩm sinh động · đọc Sheet qua JSONP (chạy được cả khi Sheet để riêng tư) · ghi Sheet qua `fetch no-cors` + upsert theo LineID · pass acorn `ecmaVersion 2020`.
