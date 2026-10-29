# HANDOFF — 10-tender: link bản vẽ/spec + BOQ_TONGHOP (28/10/2026)

Dự án: SOFITEL tender (`_ACTIVE\SOFITEL\10_TENDER`). Skill: `SKILL\10-tender`. Chỉ sửa **tool build** (`td.py`), không chạy lại bước LLM nào, không sửa JSON trong `_data` (trừ `_data\_zip` được giải nén thêm).

## 1. Vấn đề user báo
1. File sau scan tra được gì → đã trả lời (DASHBOARD, sheet BOQ CĐT + cột nội bộ ẩn, SCOPE, TCQ_LOG, MAT, DWG/KHU, VIEC_MO, FILE_INDEX, `02_extract\*.md`).
2. BOQ_TRACK không có link bản vẽ/spec như 01-scan. Đo trên bản build cũ: SCOPE "Bản vẽ hiệu lực" 600 ô có chữ / **7** có link; MAT "Nguồn spec" **0/92**; sheet BOQ CĐT không link nào.
   Nguyên nhân:
   - `link_target` bỏ qua file nằm trong zip — mọi file vai trò "Chính" (284) đều trong zip.
   - Cột nội bộ trong sheet BOQ là công thức `INDEX(SCOPE…)` → chỉ có chữ, mất link.
   - Ô chỉ link được tờ đầu; không có `_LINKS` như 01-scan.
3. Link không nhảy đúng trang PDF. Nguyên nhân: hyperlink thường của Excel không mang được số trang. 01-scan dùng add-in `_MASTER\vba\SRC_LINKS.xlam` (`src\modSrcLinks.bas`): bắt sự kiện bấm link → tra `_LINKS` theo Key `Sheet!Ô` → mở PDF bằng Edge `#page=n`; nhiều dòng cùng Key → menu chọn. Add-in đã cài trên máy (registry `Excel\Options\OPEN`).

## 2. Đã sửa — `SKILL\10-tender\tool\td.py`
| Hàm | Thay đổi |
|---|---|
| `unzip` (mới, ~dòng 271) | giải nén 1 file trong zip ra `_data\_zip\<FID>_<tên>` (ghi `.tmp` rồi đổi tên; lỗi → None). `cmd_zip` dùng lại hàm này |
| `link_item` (thay `link_target`, ~286) | link JSON `[nhãn, FID, trang, ô đích]` → `(nhãn, file tương đối theo p.base, trang, ô đích)`; file trong zip → bản giải nén; URL giữ nguyên |
| `set_links` (mới, ~1110) | ô có link file → hyperlink **trỏ về chính ô** + ghi đích vào `p.links` (để add-in xử lý). 1 đích trong workbook → hyperlink thường `#'Sheet'!A1`. URL → hyperlink thường |
| `put` | cột LINK gọi `set_links` |
| `write_links` (mới, ~1140) | sheet `_LINKS` đúng định dạng add-in: A Key(`Sheet!Ô`) B Sheet C Ô D Nhãn E File(tương đối) F Trang G Ô đích; thêm H "Mở (dự phòng)", I "Về ô". Đích trong chính workbook: E = tên file BOQ_TRACK, G = `'TCQ_LOG'!A88` |
| `add_boq_cols` (~1226) | cột nội bộ kiểu LINK ("Bản vẽ hiệu lực") ghi thẳng giá trị + link; cột khác vẫn là công thức tra SCOPE |
| `TONGHOP`, `boq_tonghop` (mới, ~1248) | sheet **BOQ_TONGHOP** (sau DASHBOARD) — user yêu cầu "BOQ tổng hợp giống scan costing" |
| `cmd_build` (~1400) | đặt `p.base, p.links, p.self_name` riêng cho từng file ra (03_boq / 01_index / 04_out); gọi `boq_tonghop` rồi `write_links` trước khi lưu BOQ_TRACK |

BOQ_TONGHOP: 1 dòng = 1 dòng SCOPE (724). Cột: Dòng BOQ gốc (→ hàng trong sheet BOQ CĐT) · Item · Vùng · Nhóm · Mô tả BOQ · ĐV · KL BOQ · Mã phòng · Trạng thái · Nhóm thực hiện · NTP chi tiết · Bản vẽ hiệu lực (→ PDF đúng trang) · Rev/ngày/đợt · Mã VL (→ MAT) · Spec · VL ("mã: mô tả MAT [trạng thái nếu ≠ Theo spec]" → file spec đúng trang) · Mã spec/FF&E · Mã TCQ (→ TCQ_LOG) · Nội dung cuối (VN) · Việc còn mở · Mô tả NT – Base. Giá trị tĩnh, dựng lại mỗi build; không vào `_snapshot` nên `pull_user` không đọc sheet này.

Tài liệu: `SKILL\10-tender\SKILL.md` mục 3 thêm 3 gạch đầu dòng (BOQ_TONGHOP, `_LINKS` + add-in, link file zip). Memory `skill-10-tender.md` đã cập nhật.

## 3. Kết quả build thật (28/10, 0 LỖI, 0 XEM)
- `03_boq\SSP_BOQ_TRACK_261028.xlsx` (bản trước ở `03_boq\_backup\…_171711`, `…_173126` và 1 bản cùng ngày nữa); `01_index\FILE_INDEX.xlsx`; `04_out\SSP_BOQ_261028_CDT.xlsx` (không đổi nội dung — chỉ cột chính thức, không link).
- Số đo bằng openpyxl:
  - SCOPE "Bản vẽ hiệu lực": 600 có chữ / 590 có link (10 ô còn lại là "KHÔNG THẤY – …").
  - MAT "Nguồn spec" 91/92; DWG "File PDF" 313/313; KHU 32/32; TCQ_LOG "Nguồn" 158/158.
  - BOQ_TONGHOP: Dòng BOQ gốc 724/724, Spec · VL 302/302, Mã VL 302/302, Mã TCQ 341/341.
  - `_LINKS` 12.159 dòng; Key cùng ô liền nhau (add-in quét các dòng liên tiếp); mọi File tồn tại khi tính từ `03_boq`; 10.960 dòng có Trang.
- File BOQ_TRACK ~1,7 MB (trước ~0,76 MB).

## 4. CHƯA kiểm — session sau làm
1. **Chưa bấm thử trong Excel.** Mở BOQ_TRACK (bật macro), bấm thử:
   - BOQ_TONGHOP cột L (1 tờ và nhiều tờ → menu) → Edge mở đúng trang?
   - Cột O (Spec · VL) → Finishes Schedule đúng trang?
   - Cột Q nhiều mã TCQ → menu → nhảy đúng dòng TCQ_LOG? (đích trong chính workbook đi qua `SL_GotoWb`, so `FullName` với `wb.Path\E` — kiểm có khớp không)
   - Sheet BOQ CĐT (vd `2. Finishing-L4-17`, cột "Bản vẽ hiệu lực" đang **ẩn** trong nhóm outline) → bấm được?
   - Link TCQ_LOG "Nguồn" (file xlsx TCQ + ô đích) → mở đúng ô?
   - Excel có báo "found a problem" khi mở không (file có ~12k hyperlink tự trỏ).
2. Không có add-in → bấm ô link file không mở gì; dự phòng = cột H của `_LINKS`. Hỏi user có cần cài add-in cho máy khác trong phòng tender không (`_MASTER\vba\src\install_SRC_LINKS.ps1`).
3. 1 mã MAT không có link spec — tra mã nào và vì sao (có thể "Không thấy spec").
4. Vòng giữ sửa tay user (`pull_user`) chưa thử lại sau sửa: sửa 1 ô SCOPE trong xlsx → `build` → sửa còn giữ, và ô LINK không bị nhận nhầm là sửa tay (chữ trong ô không đổi so với trước, nhưng nên thử).
5. Build với `--ra` sang ổ khác (C:) → đường dẫn tuyệt đối trong cột E → add-in ghép `wb.Path` + đường dẫn tuyệt đối sẽ hỏng. Chỉ ảnh hưởng build thử, không ảnh hưởng build thật.
6. User có thể muốn đổi cột BOQ_TONGHOP (vd tách KL theo mã phòng thành cột riêng) — chưa hỏi.

## 5. Lệnh kiểm nhanh
```
py -I "D:\COSTING EXCEL\SKILL\10-tender\tool\td.py" build "D:\COSTING EXCEL\_ACTIVE\SOFITEL"
```
Dừng nếu `~$SSP_BOQ_TRACK_*.xlsx` tồn tại (file đang mở). Đếm link: đọc bằng openpyxl, đếm `cell.hyperlink` từng cột; `_LINKS` kiểm Key liền nhau + `os.path.exists(cột E)` với cwd = `03_boq`.
