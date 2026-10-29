# RULES MASTER — Hướng dẫn sử dụng

Cài đặt (1 lần/máy): xem `HUONG-DAN.md`. Master: `D:\COSTING EXCEL\_MASTER\RULES_MASTER.xlsx` · macro: `PERSONAL.XLSB`.

## A. Bắt đầu dự án mới (1 lần/dự án)

| Bước | Thao tác | Ghi chú |
|---|---|---|
| 1 | Tạo file dự án (chép từ template hoặc file dự án cũ), đặt tên theo dự án | Chép từ file đã Setup → có sẵn cột + query |
| 2 | Mở file → **Enable Content** nếu Excel hỏi | Cần bật để Save đẩy được quy tắc |
| 3 | Sheet `_RULES`: kiểm tra `tblRules` có cột `ID · Phạm vi · Dự án gốc` và có sheet `_RULES_MASTER` | |
| 4 | Thiếu → Alt+F8 → `PERSONAL.XLSB!SetupProjectRules` → Run → bấm **OK** | 1 lần/file |
| 5 | Chép từ dự án cũ → rà dòng `RIENG` trong `tblRules`: dòng chỉ đúng cho dự án cũ thì xoá hoặc đặt Trạng thái `Obsolete` | Dòng `CHUNG` đã có ID: giữ nguyên |
| 6 | Data → Refresh All → xem `_RULES_MASTER` có quy tắc chung → Ctrl+S | |

> Mẹo: chạy `SetupProjectRules` 1 lần trên **file template** → mọi dự án tạo từ template có sẵn, bỏ qua bước 4.

## B. Khi dự án có quy tắc mới → tự cập nhật về quy tắc chung

```
Dự án A: thêm dòng tblRules, Phạm vi = CHUNG, ID trống
        │  Ctrl+S
        ▼
Macro (tự động): cấp ID R000x → ghi vào RULES_MASTER.xlsx (Active, v1)
                 → ghi ID + Dự án gốc ngược về dự án A → làm mới _RULES_MASTER của A
        │
        ▼
Dự án B, C…: mở file (hoặc Refresh All) → _RULES_MASTER tự có quy tắc mới
```

| Việc | Ai làm | Tự động? |
|---|---|---|
| Quyết định quy tắc là CHUNG hay RIENG | Bạn (Claude hỏi khi chốt quy tắc) | Không — bắt buộc chọn, vì macro không đoán được |
| Cấp ID, ghi Master, ghi ngược ID | Macro khi Save | ✅ |
| Phát hiện trùng nội dung với quy tắc đã có ở Master | Macro | ✅ dùng lại ID cũ, không thêm dòng |
| Các dự án khác nhận quy tắc mới | Query khi mở file / Refresh All | ✅ |
| Claude dùng quy tắc mới ở dự án khác | Claude đọc `_RULES_MASTER` đầu phiên | ✅ |

Chỉ cần nhớ 1 việc: **quy tắc dùng cho mọi dự án → chọn Phạm vi = `CHUNG` rồi Save.**

Lưu ý:
- Dòng đã có `ID` **không** được đẩy lại. Sửa nội dung ở dự án sẽ không cập nhật Master → muốn sửa quy tắc chung, sửa ở `RULES_MASTER.xlsx` (xem mục C).
- Master đang mở chưa Save → dự án khác chưa thấy dòng mới (query đọc bản đã lưu).
- Dòng cũ đánh dấu bằng `Nguồn = CHUNG-DỰÁN` (cách cũ) **không** tự lên Master — đổi Phạm vi sang `CHUNG` rồi Save.

## C. Làm việc hằng ngày

| Muốn | Làm |
|---|---|
| Quy tắc **chỉ dự án này** | Thêm dòng `tblRules`, cột **Quy tắc**, **Phạm vi = RIENG** |
| Quy tắc **mọi dự án** | Như trên, **Phạm vi = CHUNG**, để trống **ID** → Ctrl+S → ID hiện `R000x` |
| Làm cùng Claude | Claude đọc `_RULES` → `_RULES_MASTER` → skill; khi chốt quy tắc mới, Claude ghi `tblRules` và hỏi CHUNG/RIENG — bạn chỉ cần Save |
| Xem quy tắc chung mới từ dự án khác | Mở lại file, hoặc Data → Refresh All |
| **Sửa** quy tắc chung | Mở `RULES_MASTER.xlsx` → sửa nội dung, tăng **Phiên bản** +1 → Save. Không sửa trong `_RULES_MASTER` (bị ghi đè khi làm mới) |
| **Huỷ** quy tắc chung | Trong `RULES_MASTER.xlsx` đặt **Trạng thái = Obsolete** → Save. Các dự án bỏ nó ở lần mở tiếp theo |
| Lỡ đánh dấu CHUNG | Đặt `Obsolete` ở Master, đổi Phạm vi trong `tblRules` về RIENG |

## D. Sự cố

| Hiện tượng | Xử lý |
|---|---|
| Save xong ô ID vẫn trống | Kiểm tra Phạm vi = `CHUNG` (chữ hoa) và Quy tắc có nội dung. Vẫn không được → Alt+F8 → `PERSONAL.XLSB!StartRulesSync` → Ctrl+S lại |
| Hộp "Lỗi đồng bộ Master: … read-only" | `RULES_MASTER.xlsx` đang mở ở máy/cửa sổ khác → đóng nó → Ctrl+S lại (ID trống sẽ tự đẩy lần Save sau) |
| `_RULES_MASTER` chưa thấy dòng vừa thêm ở Master | Master chưa Save → Save Master rồi Refresh All |
| Excel báo PERSONAL.XLSB read-only | Đang mở 2 phiên Excel → tắt hết Excel, mở lại |
| File thiếu `_RULES_MASTER` | Chạy `SetupProjectRules` (mục A bước 4) |
