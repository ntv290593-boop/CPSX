# HANDOFF — Tool hóa thêm / kiểm máy skill 01-scan (2026-10-31)

Trạng thái (cập nhật cuối 31/10): **4.1 a (phần số), b, c đã làm, chưa commit, chưa chạy trên hồ sơ thật.** Xem mục 6. Việc tiếp theo là đo 4.2, chờ user chọn dự án.
Căn cứ đã đọc: `DOCS\Flowchart\01-scan_flowchart.md`, `SKILL\01-scan\tool\pd.py`, `SKILL\01-scan\buoc\2-boq.md`, `5-doi-chieu.md`, `DOCS\PLAN_SCAN-TOI-UU_2026-10-28.md`, `DOCS\PLAN_NEO-ANCHOR_2026-10-29.md`.

## 1. Việc tiếp theo

User nói làm tiếp mà không chọn, nên đã làm 4.1 trước theo đề xuất, không qua bước duyệt plan riêng. Việc tiếp: user chọn dự án đo 4.2 (Mục "Chưa rõ"). Kết quả 4.1 ở mục 6.

## 2. Bối cảnh
User hỏi các bước trước bước 5 có dùng máy đối chiếu không, rồi đưa 2 phản hồi bên ngoài về việc thay LLM bằng máy để verify và phản biện. Hai phản hồi và phân tích đầy đủ chỉ nằm trong phiên chat, phần đáng giữ đã chép vào đây.

**Máy đã đối chiếu trước bước 5** (chỉ từng phần):
- Bước 1: file trùng, trang trùng, file tổng.
- Bước 2: `check` so mã BOQ_CDT↔BOQ_TONGHOP, cặp Mã/Khu BOQ_CDT↔BOQ_MAP.
- Bước 3: `neo` N3.2.
- Bước 4: `neo` N3.1 + N3.3; `check` chuỗi `a+b=c`; `gemini --claude`; `check` so tổng KL theo mã, chỉ XEM (`pd.py:812`).

**Vẫn do LLM làm:** spec↔ML↔BoQ; Claude↔Gemini ở trường không phải số; mã hoàn thiện BV↔ML; kích thước BoQ↔BV (`doichieu` chưa có, dù plan B1 có ghi).

## 3. Kết quả verify (khi làm thì dùng, đừng lặp lại sai lầm)

| Hướng / nhận định | Kết luận | Bằng chứng |
|---|---|---|
| Chấm Độ tin cậy 100% bằng script, LLM khỏi rà | Sai | Trùng 91% với LLM, 8 mã tool lạc quan hơn, 5 mã thận trọng hơn (PLAN_SCAN L140). Điểm còn dựa vào dòng LLM viết |
| Scope: LLM chỉ xét khi mơ hồ | Không đủ | Blanca 148/148 gán, 0 "không chắc", lúc đầu rà tay sửa 5 lỗi, chưa có bản LLM để so (PLAN_SCAN L148). Điều kiện mơ hồ gần như không kích hoạt |
| Tìm trang, đọc số từ text: mức Cao | Quá lạc quan | SOFITEL 40% @12, trần 60% (PLAN_SCAN L127). Máy không biết số nào thuộc chiều nào (L87) |
| Đếm tag trên lớp text ra KL | Đã bị bác | DH114 trên 57 trang, KL thật 646 cánh (PLAN_SCAN L88). Chỉ làm tín hiệu kiểm chéo |
| So tổng KL với subtotal | Chưa làm được chắc | BOQ tender SOFITEL (`07-08 SSP-ID package-Tender BOQ.xlsx`) không có dòng subtotal KL ở sheet hạng mục: "Total Quantity" là tiêu đề cột, dòng tổng ở 0.Summary là tiền. Mới thử 1 BOQ. `boq-tach` chỉ in dòng tổng (`pd.py:993`) |
| Đặt mã V01 tự động theo mô tả | Chỉ làm nháp | R/S/C có thể ở cột riêng (`pd.py:1016`), gộp theo mô tả sẽ nhập 2 món khác kích thước mà không báo. Khóa = mô tả + ĐV + R/S/C; LLM duyệt |
| Thêm 3 lớp L1/L2/L3 | Bỏ | Chủ yếu đặt lại tên: `check` đã có LỖI/XEM, `neo` tách riêng, vòng sửa ≤ 2 lần đã có |
| Giảm đáng kể số lần gọi LLM | Chưa có số đo | `do` đếm ô, không đếm lượt gọi/token. Lợi chắc chắn của tool hóa là bớt ô ghi và lỗi chép |
| BOQ_MAP, phần sao chép BOQ_TONGHOP suy từ BOQ_CDT | Đúng, rủi ro thấp | `2-boq.md` mục 3–4 |

## 4. Thứ tự đề xuất

### 4.1. Vá lỗ hổng độ tin cậy (không cần chờ đo)

**a. Tool sinh dòng `Kiểm chéo Gemini`, phần so số.** Hiện LLM tự viết (`4-ban-ve.md` L38) mà dòng này quyết định Độ tin cậy (`pd.py:1163-1166`, `1213-1217`); `check` chỉ kiểm dòng có tồn tại (`pd.py:810`). Dòng không có chữ "lệch:" thì không bị hạ mức. Cần xử lý:
- Kết quả so chỉ in ra, không lưu (`pd.py:562`).
- So bằng `Counter`: không theo vị trí, gộp mọi crop. Đọc đúng số mà gán sai chiều vẫn tính khớp.
- Số Claude do LLM gõ vào `--claude`; nên lấy từ QUAN_SAT.
- Lệnh chưa nhận mã, nên chưa biết crop thuộc mã nào.
- `--quansat`: không lưu, không phân tích (`pd.py:549`, `556`), khóa `GEM_TRUONG` khác tên Trường trong QUAN_SAT, giá trị viết tự do. Để sau, làm phần số trước.

**b. Đưa đủ danh sách đã loại ra `CẦN XEM`.** `_bo_qua` không vào `check` (chỉ dùng ở `pd.py:181`, `209-210`); `du_phong` chỉ kiểm file có tồn tại (`pd.py:779`); mã ngoài scope chỉ in 12 mã đầu (`pd.py:805`). Bỏ nhầm file hoặc mã thì mọi bước sau vẫn qua. User duyệt theo ngoại lệ, không phải máy quyết.

**c. Sửa `buoc/5-doi-chieu.md` L10.** Câu "tool thận trọng: lệch nhau chủ yếu tool THẤP hơn" ngược số đo (8 lạc quan, 5 thận trọng), khiến LLM rà lỏng hơn.

### 4.2. Đo mù trên mẫu
Khoảng 10 mã, cố ý lấy cả mã CAO và mã ngoài scope; user kiểm như đợt `KIEM_MAU` Blanca (57 dòng/10 mã); kích thước chấm được bằng DWG; đo thêm lượt gọi/token. Lý do: user chỉ duyệt ngoại lệ nên không thấy lỗi lọt, không thể tính tỷ lệ bắt lỗi từ review thường.

**Chưa rõ:** đo trên dự án nào. N5 chờ "lần scan 01 thật kế tiếp" (PLAN_NEO L127) mà dự án active đang chạy skill 10-tender, nên có thể treo lâu. Cần hỏi user.

### 4.3. Sau đó
- Tool sinh `BOQ_MAP`, phần sao chép `BOQ_TONGHOP`, issue B (cùng mã khác mô tả/ĐV, dòng trùng).
- `doichieu` so kích thước BoQ↔BV: tận dụng `neo_kt_gt`; thứ tự R×S×C có thể khác nhau.
- Thêm khuôn cho `Mã hoàn thiện trên BV` rồi so với ML.
- Tra ML và bắt callout: cần đo trước, chỉ hiệu quả với hồ sơ có lớp text.

### 4.4. Tạm bỏ
L1/L2/L3; đếm KL bằng tag; so subtotal cho đến khi có BoQ thật có dòng đó.

## 5. Lưu ý khi làm
- Đọc `ACTIVE.md` trước. Dự án active là SOFITEL tender. `_ACTIVE\SOFITEL\01_scan` chỉ được dùng để ĐO, không lấy số liệu sang tender. `_ARCHIVE` bị chặn: Blanca dùng `DOCS\scan-toi-uu\_cache`.
- Sao lưu `pd.py` vào `tool\_backup` trước khi sửa. Đổi tính năng skill thì cập nhật `DOCS\Flowchart\01-scan_flowchart.md` + `CHANGELOG.md` cùng phiên.
- Hồi quy: `DOCS\scan-toi-uu\bench.py`, `DOCS\scan-toi-uu\neo-test\neo01_regress.py`.
- Lõi dùng chung `SKILL\_core\scan_core.py` cho 01-scan và 10-tender: sửa ở đó thì kiểm cả hai.
- Git: nhánh `main-1`; còn 4 file backup chưa track (`pd_2026-10-30_truoc-n32.py`, `truoc-n33.py`, ở `SKILL` và `.claude\skills`). Phiên này không commit gì; file handoff này cũng chưa commit.

## 6. Đã làm 31/10 (mục 4.1, chưa commit)
Sao lưu: `tool\_backup\pd_2026-10-31_truoc-vaGemini.py`. Test: `DOCS\scan-toi-uu\gem_so_test.py`.

**a. Phần số Gemini.** `gemini <crop> --ma <mã> --goc --dot` lưu số Gemini vào `_gemini_so\<mã>.json`, dừng nếu QUAN_SAT chưa có `Kích thước BV` R×S×C (buộc Claude ghi trước). `check`/`doichieu` so với QUAN_SAT lúc chạy nên đổi số là thấy ngay. `doichieu` sinh dòng `Kiểm chéo Gemini số` (`QUAN_SAT.doichieu-tool.json`): số Claude không có trong số Gemini → THAP + issue; chưa chạy / agy lỗi → tối đa TB. Dòng `Kiểm chéo Gemini` do LLM viết vẫn dùng cho trường không phải số. `--claude` giữ nguyên, chỉ để xem. Cache `.gemini.txt` bỏ qua khi crop mới hơn; trả lời rác không cache.

**Còn hở (cố ý, chưa giải quyết):**
- Chưa biết số thuộc chiều nào: đọc đúng số, gán sai chiều vẫn KHỚP.
- Chỉ so chuỗi R×S×C đầu tiên (`neo_kt_gt`), số ≥ 100. Số phổ biến (600, 900) có thể khớp ngẫu nhiên khi crop nhiều số.
- Số so theo từng mã, chưa theo từng crop; chưa kiểm crop thuộc đúng tờ đã link.
- LỆCH có thể oan khi crop không có số đó: thử agy thật với hình tự vẽ, số 450 không nằm trên đường kích thước thì Gemini không liệt kê → `chỉ Claude: 450`. Chiều sâu thường ở mặt cắt, nên phải chạy đủ crop. Tỷ lệ lệch oan chưa đo, là việc của 4.2; nếu cao thì xem lại mức THAP.
- `--quansat` vẫn chưa lưu/so (handoff 4.1a để sau).

**b. Danh sách đã loại.** `kiemke` lưu `_kiemke_bo.json` (file + lý do; dòng `#` trong `_bo_qua.txt` thành lý do). `check` thêm dòng `ĐÃ LOẠI` đủ danh sách: file bỏ, dự phòng, file tổng, trang trùng, mã ngoài scope theo lý do. Terminal cắt ở 400 ký tự, DASHBOARD đủ. Dự án đã `kiemke` bằng bản cũ: báo thiếu file, chạy lại `kiemke` mới có danh sách.

**c.** Sửa `5-doi-chieu.md` L10 theo số đo (8 lạc quan, 5 thận trọng).

**Đã kiểm:** `gem_so_test.py` đạt. Bản sao Blanca: chưa chạy Gemini số thì CAO 9 → 0, TB 12 → 21, THAP giữ 127; Gemini khớp hết thì Độ tin cậy y như bản cũ (khác 0 mã); `check` 100 LỖI giữ nguyên. `bench.py` QUA; `neo01_regress.py` D1d 199/199. `agy` thật qua dòng lệnh chạy được (hình tự vẽ). `build` thật trên bản sao ra DASHBOARD đủ dòng `ĐÃ LOẠI`.
**Chưa kiểm:** cả luồng trên hồ sơ thật với subagent bước 4.
