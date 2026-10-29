# HANDOFF — Bộ skill CPSX (Cowork ↔ Claude trong Excel)

Cập nhật: 2026-10-21 (chiều: xong P4) · Đọc file này trước khi làm tiếp. Người dùng: thích bảng, ngắn gọn, **duyệt plan trước khi làm**.

## 1. Mục tiêu
Cowork đọc spec + bản vẽ (PDF CAD/scan; PDF/Word/Excel; Việt/Anh/song ngữ) → bóc tách → Claude trong Excel chiết tính trên file CPSX (.xlsm). Dự án **hỗn hợp** (cửa, tủ, bếp…). Train dần: bóc tới đâu, ghi rule tới đó.

## 2. Kiến trúc đã chốt

| Khâu | Ai | Cách |
|---|---|---|
| Đọc spec/bản vẽ, bóc tách | Cowork | Ghi `<dự án>\02_takeoff\AI_INPUT.xlsx` (mẫu: `_MASTER\schemas\AI_INPUT_TEMPLATE.xlsx`, cột: `_MASTER\schemas\README.md`) |
| Nạp vào CPSX | Power Query `qAiInput` → sheet `AI_INPUT` (`tblAiInputView`) | Tự khi mở file (`SyncAiInput`) hoặc Alt+F8 `PERSONAL.XLSB!RefreshAiInput` |
| Chiết tính, link BoQ, mô tả | Claude trong Excel | Skill `cpsx-master` v2 + thư viện `_LIB` (`L.map`, `L.newItem(s)`, `L.edit`, `L.desc`, `L.addRules`…), §1b đọc AI_INPUT |
| **Cowork KHÔNG ghi trực tiếp .xlsm** | — | Python làm mất Power Query/VBA |

### Bộ nhớ chung (thông suốt 2 bên)

| Tầng | Nơi | Ghi bởi |
|---|---|---|
| 1 Rule dự án | `_RULES` / `tblRules` trong từng file | Add-in (`L.addRules`); `Phạm vi`=CHUNG → Ctrl+S macro đẩy lên Master |
| 2 Rule chung | `_MASTER\RULES_MASTER.xlsx` / `tblRulesMaster` → query vào `_RULES_MASTER` | Macro khi Save; **Cowork ghi thẳng** (hỏi user trước, khi file đóng), ID `R00xx` tiếp theo |
| 3 Skill | `cpsx-master` (skill tài khoản, dùng chung Cowork + add-in) | Sửa `skill-src\SKILL.template.md` → `node skill-src\build.js` → zip → user upload |

## 3. Thư mục `D:\COSTING EXCEL\`

| Đường dẫn | Vai trò |
|---|---|
| `_MASTER\` | Hệ thống: RULES_MASTER, macro (.bas/.cls), skill-src, cpsx-master (SKILL.md + lib\cpsx_lib.js v2.1), schemas, backup, HUONG-DAN*.md |
| `260903_DA_CPSX_RE00.xlsm` | **Template chuẩn** (đã có AI_INPUT + qAiInput + qRulesMaster) |
| `01. DOOR\` (02. TU…) | Chỉ mẫu + ví dụ đã duyệt (README) |
| `_ACTIVE\<yymmdd>_<Ma>_<Ten>\` | Dự án đang làm: 00_input, 01_extract, 02_takeoff, 03_costing (CPSX_RE00, RE01…), 04_audit. Xong dự án → đẩy server (quy trình đóng: thiết kế sau) |

PERSONAL.XLSB đã cài: `modRulesSync`, `modCpsxLib`, `modAiInput`, `clsAppEvents` (BeforeSave → PushSharedRules; WorkbookOpen → AutoUpdateLib + SyncAiInput). Compile OK.

## 4. Trạng thái

| Hạng mục | Trạng thái |
|---|---|
| P1 thư mục | ✅ |
| P3 AI_INPUT (macro, template, skill §1b) | ✅ |
| Lib v2.1 (không nhận `Matrix`/`Spec` là BoQ) | ✅ ghi đĩa |
| Skill cpsx-master build 21/10 (§3.3 cửa 2 cánh dùng mẫu 2 cánh, §3.1 #9 đọc rule MAU/TRACUU/FILE-00) | ✅ ghi đĩa |
| **Lib v2.2 + skill §5.0 Tự học (21/10 chiều)**: `_EDITLOG` lưu giá trị Claude ghi (cột F `f:`), `L.map` trả `userEdits` (ô user sửa), `L.userEdits`, `L.learned`; `L.unmark` giữ dòng chưa học; `L.addRules` mặc định CHUNG; skill bắt ghi rule ngay khi bị sửa/nhắc | ✅ ghi đĩa; ⏳ user upload `cpsx-master.zip`, mở file costing (AutoUpdateLib) → kiểm `lib:"2.2"`; dán 2 dòng vào Instructions add-in |
| Tắt 7 skill cũ (cpsx-project-rules, costing-door-*, costing-chiet-tinh-hao-hut, cpsx-ve, boq-melamine-laminate-toggle) | ⏳ user |
| P4 Danh mục mẫu | ✅ 21/10: RULES_MASTER R0075–R0151 (MAU 59, TRACUU 15, DOOR R0134, CHUNG R0150, FILE-00 R0151); dọn 14 rule Obsolete, sửa R0004/19/22/28/33/34, R0074 v2. Nháp: `_MASTER\MAU_DRAFT.md` |
| **P5 Chạy thử dự án hỗn hợp thật** | ▶ **việc tiếp theo**: chờ user bỏ input vào `_ACTIVE` |
| Đóng dự án / đẩy server, đóng gói plugin | Sau |

## 5. Danh mục mẫu — ĐÃ XONG 21/10 (giữ để tham khảo cách làm)

Mục tiêu: model biết **mẫu nào dùng khi nào, mỗi dòng chi tiết là gì, ô nào là ô nhập liệu, biến thể sửa ra sao** — không cần hiểu chi tiết từng công thức.

| Sheet trong RE00 | Số dòng | Ghi chú |
|---|---|---|
| Template | 1606 | 22 block (tủ/loose/mẫu trống 1312–1370) |
| Door | 928 | 9 block (§3.3 skill đã có quy tắc chọn) |
| Kitchen | 1331 | 23 block |
| Chiet tinh | 8 | rỗng (END dòng 8) |
| Material | 1123 | tra mã vật tư A→F |
| Khác | OverView, Spec, Matrix, Dinhmuc, Van; ẩn: PK, KINH, GO, VAN KHAC, Linhtinh, PCCC, VANCHUYEN | |

Block: dòng đầu có số ở cột B, tên ở A; kết thúc `END` ở A; dòng chi tiết được group 1 cấp. Cột chi tiết: A tên · B dày · C rộng · D dài · E SL · F hao hụt · G KL · H đơn giá · I thành tiền · J ĐVT · K mã vật tư · L/N/O hoàn thiện · Q tên vật liệu · AN…AU phụ hao hụt.

Cách làm (đã được user duyệt "đọc từng template"):
1. Đọc template **chỉ đọc**: stage RE00 vào workspace. openpyxl lỗi vì drawing → patch `openpyxl.reader.excel.find_images = lambda *a,**k: ([],[])`; đọc 2 lần (formulas + data_only).
2. Mỗi block → 1 dòng danh mục: Tên · Sheet!dòng đầu–END · Dùng khi (đoán, đánh dấu ❓) · Dòng chi tiết (+k: nhãn A, mã K, SL E) · Ô nhập liệu (ô số cứng/không công thức) · Thông số đổi giá · Mẫu gần giống.
3. Làm **Template trước**, gửi user duyệt định dạng → rồi Door, Kitchen, rồi các sheet tra cứu (Material, PK, Dinhmuc…).
4. User duyệt → ghi `RULES_MASTER.xlsx`, Chủ đề = `MAU` (1 dòng/mẫu) — add-in tự nhận qua `_RULES_MASTER`. Nội dung dài/ổn định có thể đưa vào skill (§3.x) sau.

## 6. Lưu ý môi trường

| Vấn đề | Xử lý |
|---|---|
| Clipboard máy user bị ghi đè `VOB@boho.vn` → gõ ký tự đặc biệt qua computer-use bị dán sai | Gõ chữ thường; chọn file bằng double-click trong hộp thoại; chạy macro bằng con trỏ trong Sub + F5 |
| Excel ở màn hình HP, VBE ở Built-in Display | `switch_display` |
| Hôm 20/10 có phiên khác sửa `_MASTER` song song; một lần commit bị ghi đè bằng bản cũ | Kiểm `device_list_dir` (size/mtime) sau mỗi lần ghi |
| Có file trùng tên `D:\260903_DA_CPSX_RE00.xlsm` (gốc ổ D) | Không dùng; chuẩn là `D:\COSTING EXCEL\…` |
| RULES_MASTER: 151 rule (137 Active: MAU 59, BOQ ~30, CHUNG ~20, TRACUU 15…) | Đọc trước khi thêm rule để tránh trùng; Trạng thái chỉ Active/Obsolete |
| `device_commit_files` cùng stagedPath lần 2 có thể ghi lại bản CŨ (21/10) | Mỗi lần ghi dùng tên file nguồn mới; sau ghi stage lại và đọc kiểm nội dung |
| Backup RULES_MASTER | `_MASTER\backup\RULES_MASTER_bk_20261021_truocMAU.xlsx`, `…_truocDON.xlsx` |
