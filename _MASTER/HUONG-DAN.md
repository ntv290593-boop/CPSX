# RULES MASTER — Hướng dẫn cài (1 lần)

Chốt: Master `D:\COSTING EXCEL\_MASTER\RULES_MASTER.xlsx` · quy tắc CHUNG vào Master với trạng thái **Active** (không qua duyệt) · macro đặt trong **PERSONAL.XLSB**.

| # | Việc | Thao tác |
|---|---|---|
| 1 | Master | Chép `RULES_MASTER.xlsx` vào `D:\COSTING EXCEL\_MASTER\` (hoặc bỏ qua — macro tự tạo nếu chưa có) |
| 2 | Trusted Location | File > Options > Trust Center > Trusted Locations > Add `D:\COSTING EXCEL\` (tick *Subfolders*) |
| 3 | Mở PERSONAL.XLSB | Nếu chưa có: View > Macros > Record Macro > *Store in: Personal Macro Workbook* > Stop |
| 4 | Import VBA | Alt+F11 > chọn `VBAProject (PERSONAL.XLSB)` > File > Import: `modRulesSync.bas`, `clsAppEvents.cls` |
| 5 | Dán ThisWorkbook | Mở `ThisWorkbook` của PERSONAL.XLSB, dán nội dung `ThisWorkbook_PERSONAL.txt` > Save > tắt/mở lại Excel |
| 6 | Setup file dự án | Mở `D:\260903_DA_CPSX_RE00.xlsm` > Alt+F8 > `PERSONAL.XLSB!SetupProjectRules` |
| 7 | Đánh dấu | Trong `_RULES` cột `Phạm vi`: chọn `CHUNG` cho quy tắc muốn dùng mọi dự án (mặc định `RIENG`) |
| 8 | Save | Ctrl+S → dòng CHUNG được cấp ID `R0001…` và ghi vào Master; `_RULES_MASTER` tự làm mới |
| 9 | Skill | Dán `skill-cpsx-master-section0.md` vào cuối §0 của skill cpsx-master |
| 10 | Chạy thử | File A: thêm 1 quy tắc CHUNG > Save. File B: chạy Setup (bước 6) hoặc mở lại → thấy quy tắc trong `_RULES_MASTER` |

## Macro có sẵn

| Macro | Dùng khi |
|---|---|
| `SetupProjectRules` | Chạy 1 lần/file: thêm cột `ID · Phạm vi · Dự án gốc` vào `tblRules` (tạo `_RULES` nếu thiếu), tạo query + sheet `_RULES_MASTER` |
| `PushRulesNow` | Đẩy tay (có báo số dòng). Bình thường không cần — Save tự đẩy |
| `CreateMasterFile` | Tạo Master trống nếu chưa có |
| `StartRulesSync` | Bật lại bắt sự kiện Save nếu vừa sửa code VBA (reset project làm mất sự kiện) |

## Hành vi

| Tình huống | Kết quả |
|---|---|
| Dòng `CHUNG`, `ID` trống, có nội dung | Ghi vào Master (Active, Phiên bản 1), ghi ID ngược về file dự án |
| Nội dung trùng 1 quy tắc đã có ở Master | Không thêm dòng mới, dùng lại ID cũ |
| Dòng đã có ID | Không đẩy lại (sửa nội dung → sửa ở Master) |
| Master đang mở ở máy/phiên Excel khác (read-only) | Báo lỗi, giữ ID trống → tự đẩy lại ở lần Save sau |
| Quy tắc sai | Đặt `Obsolete` ở Master → các dự án bỏ nó ở lần mở file tiếp theo |
| Template file mới | Chạy Setup trên template 1 lần → dự án tạo từ template có sẵn query |

## Giả định cần kiểm

- `tblRules` có cột nội dung tên đúng **`Quy tắc`** (và tuỳ chọn `Ngày`, `Chủ đề`). Khác tên → đổi tên cột, hoặc báo để sửa macro.
- Query đọc bản **đã lưu** của Master; Master đang mở chưa Save thì chưa thấy dòng mới.
