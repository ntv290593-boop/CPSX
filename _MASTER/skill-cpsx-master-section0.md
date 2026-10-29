### §0b. Quy tắc 3 tầng (thêm vào cuối §0 của skill cpsx-master)

Trước khi làm bất kỳ việc gì trên workbook costing, đọc đủ 3 nguồn theo thứ tự ưu tiên (cao → thấp):

| Ưu tiên | Nguồn | Vị trí | Ghi chú |
|---|---|---|---|
| 1 | Quy tắc dự án | sheet `_RULES`, bảng `tblRules` | Cột `Phạm vi`: `RIENG` = chỉ dự án này · `CHUNG` = đã/sẽ đẩy lên Master |
| 2 | Quy tắc chung | sheet `_RULES_MASTER`, bảng `tblRulesMasterView` (query chỉ đọc từ `D:\COSTING EXCEL\_MASTER\RULES_MASTER.xlsx`) | Chỉ áp dụng dòng `Trạng thái = Active`; `Obsolete` đã bị lọc |
| 3 | Skill | file skill này | Mặc định khi 2 tầng trên không nói gì |

- Xung đột: tầng ưu tiên cao thắng. Trùng `ID` giữa `_RULES` và `_RULES_MASTER` → lấy nội dung ở Master (bản mới nhất), trừ khi dòng `_RULES` ghi `RIENG`.
- Không sửa sheet `_RULES_MASTER` (sẽ bị ghi đè khi làm mới). Muốn sửa/huỷ quy tắc chung → sửa ở file Master.
- Khi chốt quy tắc mới với user: ghi vào `tblRules` của file dự án, hỏi user `CHUNG` hay `RIENG`. Để trống `ID` — macro tự cấp ID khi Save.
- Nếu workbook thiếu `_RULES_MASTER`: báo user chạy macro `SetupProjectRules` (PERSONAL.XLSB).
