# Cách đẩy app lên GitHub để có link web (3 bước ~2 phút)

> Em không đăng nhập GitHub thay anh từ đây được, nên anh chạy trên máy anh (đã có git đăng nhập sẵn). Repo này đã sẵn sàng, chỉ cần đẩy.

## Bước 1 — Tạo repo rỗng trên GitHub
Vào https://github.com/new → **Repository name:** `kho-nuoc-ion` → chọn **Public** → **KHÔNG** tích "Add a README" → **Create repository**.

## Bước 2 — Đẩy code (mở Git Bash / CMD tại thư mục này)
Mở terminal ngay trong `E:\Sieu du an\opc-nemthanh\du-an\DA05-nuoc-ion-kiem\web-kho-nuoc-ion` rồi dán:

```
git init
git add .
git commit -m "App kho nuoc ion kiem DA05"
git branch -M main
git remote add origin https://github.com/dovantoanai-ui/kho-nuoc-ion.git
git push -u origin main
```

## Bước 3 — Bật GitHub Pages
Vào repo trên GitHub → **Settings** → **Pages** → mục "Build and deployment":
- Source: **Deploy from a branch**
- Branch: **main** — folder: **/ (root)** → **Save**.

Đợi ~1 phút, link web app của anh là:

**https://dovantoanai-ui.github.io/kho-nuoc-ion/**

## Sau khi có link
- Mở link đó bằng trình duyệt → tab **Danh mục** → dán **URL Apps Script Web App** (`/exec`) → Lưu. Từ giờ **Đồng bộ chạy được** (vì đã là https, không còn `file://`).
- Trên iPhone: mở link bằng **Safari → Chia sẻ → Thêm vào Màn hình chính** để cài như app.

## Lần sau sửa app
Sửa `index.html` xong, tại thư mục này chạy:
```
git add . && git commit -m "cap nhat" && git push
```
GitHub Pages tự cập nhật lại link.
