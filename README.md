# SAPOCHE — Báo cáo KPI nhân sự

Web app Apps Script đọc dữ liệu KPI từ Google Sheet (nguồn BigQuery qua Connected Sheets), hiển thị dashboard 2 trang: Tổng quan công ty và Chi tiết nhân sự.

```
src/
├── appsscript.json   manifest: scope, advanced service Sheets, cấu hình web app
├── Code.js           tầng dữ liệu — clasp đẩy lên thành Code.gs
└── index.html        toàn bộ giao diện + biểu đồ (SVG vẽ tay, không thư viện ngoài)
```

---

## 1. Đưa lên GitHub lần đầu

### Chuẩn bị

**Bật Apps Script API** (chỉ làm 1 lần cho tài khoản): vào https://script.google.com/home/usersettings → bật **Google Apps Script API**. Không bật thì `clasp push` báo lỗi 403.

**Lấy Script ID**: trình soạn Apps Script → ⚙️ Project Settings → mục *IDs* → copy **Script ID**.

### Các lệnh

```bash
npm install -g @google/clasp
clasp login                       # mở trình duyệt, đăng nhập tài khoản có quyền sửa script

git clone https://github.com/<user>/sapoche-kpi.git
cd sapoche-kpi

cp .clasp.json.example .clasp.json
# mở .clasp.json, dán Script ID vào

clasp pull                        # kéo code hiện có trên Apps Script về src/
git add -A && git commit -m "Đưa Apps Script lên repo" && git push
```

> `clasp pull` sẽ **ghi đè** `src/`. Nếu code trên Apps Script đang cũ hơn repo thì bỏ qua bước này, dùng `clasp push` để đẩy code repo lên.

---

## 2. Vòng lặp làm việc hằng ngày

```bash
# sửa code trong src/ bằng editor quen thuộc
clasp push                        # đẩy lên Apps Script
git commit -am "Sửa filter theo tuần" && git push
```

**Quan trọng:** `clasp push` chỉ cập nhật code, **URL web app vẫn chạy bản deploy cũ**. Muốn người dùng thấy thay đổi phải tạo version mới:

```bash
clasp deployments                             # xem danh sách, copy Deployment ID
clasp deploy -i <DEPLOYMENT_ID> -d "Sửa filter tuần"
```

Dùng `-i` để giữ **nguyên URL** đang gửi cho khách. Chạy `clasp deploy` không có `-i` sẽ tạo URL mới.

---

## 3. Tự động deploy khi push lên main (tuỳ chọn)

File `.github/workflows/deploy.yml` đã có sẵn. Cần thêm 3 secret trong repo (Settings → Secrets and variables → Actions):

| Secret | Lấy ở đâu |
|---|---|
| `CLASPRC_JSON` | Nội dung file `~/.clasprc.json` sau khi `clasp login` (Windows: `C:\Users\<tên>\.clasprc.json`) |
| `SCRIPT_ID` | Project Settings trong trình soạn Apps Script |
| `DEPLOYMENT_ID` | Kết quả lệnh `clasp deployments` |

Từ đó mỗi lần push lên `main`, Actions tự `clasp push` rồi cập nhật đúng bản deploy đang dùng.
Chưa khai `DEPLOYMENT_ID` thì workflow chỉ push code, bỏ qua bước tạo version — không lỗi đỏ.

> `CLASPRC_JSON` chứa refresh token của tài khoản Google. Repo phải để **private**. Token hết hạn thì `clasp login` lại rồi cập nhật secret.

---

## 4. Những thứ KHÔNG được commit

`.gitignore` đã chặn sẵn:

- **`.clasprc.json`** — chứa token đăng nhập Google. Lộ file này là mất quyền vào toàn bộ Apps Script của tài khoản.
- **`.clasp.json`** — chứa Script ID, gắn với từng môi trường. Mỗi người tự tạo từ `.clasp.json.example`.

---

## 5. Vài chỗ hay vấp

**`clasp push` báo "User has not enabled the Apps Script API"**
Chưa bật API ở bước chuẩn bị.

**Push xong mà web app không đổi**
Chưa tạo version deploy mới. Xem mục 2.

**File `.gs` thành `.js` dưới local**
Đúng như thiết kế. clasp tự đổi `.js` ↔ `.gs` khi đẩy/kéo. Đừng đổi tên `Code.js` thành `Code.gs` trong repo.

**`clasp pull` xoá mất file**
`clasp pull` đồng bộ một chiều từ Apps Script về. File chỉ có ở local mà chưa push sẽ bị xoá. Luôn `clasp push` trước khi `pull`.

**Sửa `appsscript.json` xong bị mất quyền**
Đổi `oauthScopes` khiến Google yêu cầu cấp quyền lại. Chạy tay `kiemTra()` một lần trong trình soạn để cấp lại.

---

## 6. Cấu hình trong `Code.js`

| Khoá | Ý nghĩa |
|---|---|
| `CONFIG.TAB` | Tên tab dữ liệu. Để trống = tự dò tab có đủ cột `ten_cv`, `nguoi_phu_trach`, `trang_thai`, `diem_task` |
| `CONFIG.BO_TRANG_THAI` | Trạng thái loại khỏi báo cáo, mặc định `['Đã huỷ']` theo đúng DAX `So_CV` |
| `CONFIG.FIELD_LY_DO_LUI` | Tên field lý do lùi trong `form_json`. Để trống = tự dò |
| `CONFIG.CACHE_PHUT` | `0` = luôn đọc mới. Đặt `10` nếu tải chậm |

Hàm chạy tay trong trình soạn:

- **`chanDoan()`** — in ra từng tab: kiểu GRID hay DATASOURCE, header thật, cột nào thiếu
- **`kiemTra()`** — số dòng, phân bố trạng thái, phân bố tháng, `Diem_TB` toàn bộ. Đối chiếu với Power BI cùng kỳ
- **`lamMoiDuLieu()`** — làm mới Connected Sheets từ BigQuery. Gắn trigger Time-driven chạy sáng sớm
- **`lietKeTab()`**, **`xoaCache()`**

---

## 7. Ba điểm gãy dữ liệu ở tầng BigQuery

Không sửa được ở repo này, phải sửa view `kpi_nhansu`:

1. `so_lan_doi_han` đang là `CAST(NULL AS INT64)` → cột "Số lần lùi" luôn trắng
2. `Lý do lùi` không có trong view, `Code.js` phải bóc từ `form_json`
3. `diem_TN` và `diem_NL` hardcode `10` → điểm lãnh đạo chấm chưa vào, mọi người đều 2.5 + 1.5
