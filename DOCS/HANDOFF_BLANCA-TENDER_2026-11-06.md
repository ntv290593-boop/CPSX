# HANDOFF — Blanca City tender (skill 10-tender), 06/11/2026

Dự án active: Blanca City (ACTIVE.md đã đổi 06/11). GOC = `_ACTIVE\Blanca_City`. Đọc `_ACTIVE\Blanca_City\10_TENDER\README.md` trước (đợt, bẫy, quyết định của user).

## Đã xong (check 0 LỖI từng bước)
- Hồ sơ 06/11: zip lồng zip + rar giải nén vào `00_input\261106_Lan2`; zip gốc ở `00_input\_zip_goc` (`bo_thu_muc`).
- `init --mau cdt_vn` (mẫu mới, BOQ CĐT tiếng Việt, không TCQ), `boq_bo_sheet = ["TH"]`.
- Bước 1 kiemke (773 file), `trang` (16 PDF / 417 trang), 2 boq (404 dòng / 6 sheet), 4 dwg (316 tờ), 5 mat (119 mã), 6 banve (4 nhóm, 404/404 dòng), 7 doichieu (3 nhóm, 404 Base), `phanloai` tool (404 `Chờ LLM`), `docsau` (361 cụm → 46 nhóm, `_data\_docsau.txt`).
- Bước 8: `dmvt` đã ra `03_boq\BLANCA_DMVT_261106.xlsx` (118 dòng, 0 ảnh — tool chỉ cắt ảnh từ trang spec PDF có chữ; ảnh mẫu nằm ở cột "Minh họa" xlsx F0434 và PDF ảnh F0692). Chưa có `_data\_dmvt.json`.
- Bước 3 tcq: không có TCQ. Bước 7b ben: chưa có `_data\_ben.json`.

## Việc tiếp theo (theo thứ tự)
1. Bước 2b phanloai (1 subagent, `buoc/2b-phanloai.md`) → chạy lại `phanloai`.
2. Bước 7c spec (1 subagent).
3. Bước 7d docsau: **user chọn đọc sâu toàn bộ 46 nhóm**. Giao theo đợt ~8 nhóm song song; prompt chỉ subagent tới khối `== Nhóm <n>` trong `_data\_docsau.txt`, ghi `*.docsau-<n>.json`. Sau đó `neo`, gửi NEO LỆCH về từng nhóm (≤ 2 vòng).
4. `tincay` → bước 7e (1 subagent).
5. `build` → báo user theo SKILL.md mục 3.
(User đã ngắt phiên lúc vừa giao 2b/7c/7d nhóm 1 — chưa subagent nào chạy các bước này; có thể user muốn đổi cách làm, hỏi lại trước khi giao.)

## Cập nhật 07/11
- User bảo bỏ 7d. Đã làm 2b (404 dòng LLM, 6 KHÔNG RÕ = picomat che điều hòa), 7c (SCOPE/QUAN_SAT/SPEC/ISSUES.spec.json, S01–S12), `tincay` + 7e (ISSUES 42, RFI_LOG 13 nháp, ASK_USER A01–A05; 361 cụm THAP), `build` → `03_boq\BLANCA_BOQ_TRACK_261107.xlsx`, `BLANCA_DMVT_261107.xlsx`, `04_out\BLANCA_BOQ_261107_CDT.xlsx`.
- 361 LỖI còn lại đều do chưa có 7d (QUAN_SAT/Kiểm chéo Gemini từng cụm).
- Lỗ hổng skill: `buoc/2b-phanloai.md` bảo ghi `VIEC_MO.phanloai.json` nhưng `COT.json` `thu_tu_buoc` không có `phanloai` → check LỖI "bước lạ"; subagent đã xóa file.

## Câu còn chờ user
1. Phụ kiện cửa chào theo bộ hay từng chi tiết (dòng đầu mục và dòng chi tiết đều Base → dễ tính trùng).
2. STUDIO!12 gạch sàn TL2: BV hoàn công ghi "NHÀ THẦU KHÁC" — giữ Base hay đổi Chờ.
3. DMVT: có sửa tool lấy ảnh từ cột "Minh họa" của F0434 không; `_dmvt.json` theo mẫu nào.

## Sửa tool trong phiên (CHƯA COMMIT)
- `SKILL\10-tender\tool\td.py`: mẫu `cot_ma` (mã từ cột MÃ SP), `c1_la_muc` (Vùng/Nhóm/Mục phụ); `_cauhinh` `rx_phong`, `phong_theo_sheet`, `dwg_theo_trang` (hàm `dwg_trang`: sổ DWG theo trang, số tờ ưu tiên ở 700 ký tự cuối lớp chữ); `mat` cờ `KHÁC NGHĨA THEO LOẠI CĂN`.
- `SKILL\10-tender\form\MAU_DU_AN.json`: mẫu `cdt_vn`. SKILL.md mục 0, flowchart, CHANGELOG đã cập nhật.
- Mặc định không đổi; `DOCS\scan-toi-uu\tender-test\kiemthu.py --mau` 75/75 đạt. SOFITEL/CTX không có trên máy nên chưa thử dữ liệu thật.
- Git working tree còn thay đổi `_MASTER\vba\…` và flowchart từ trước phiên này (không phải của phiên này).
