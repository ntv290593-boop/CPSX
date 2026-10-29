# HANDOFF — Nghiên cứu repo GitHub + thử ý tưởng `neo` cho 01-scan (29/10/2026)

Mục đích file: để session khác **kiểm tra lại tính đúng** của những gì đã làm và đã kết luận. Mỗi kết luận ghi rõ nguồn và độ chắc. Chưa có thay đổi nào ở skill `01-scan` hay `10-tender`.

## 1. Bối cảnh

- User muốn bổ sung cho skill `D:\COSTING EXCEL\SKILL\01-scan` (quét hồ sơ thầu → PROJECT_DATA). Bước đầu: tìm repo GitHub public làm việc tương tự để học thuật toán và cách dùng tool.
- Dự án active (theo `ACTIVE.md`, cập nhật 28/10): **SOFITEL tender**, thư mục `_ACTIVE\SOFITEL`, làm việc ở `10_TENDER`. Dự án này dùng skill `10-tender`, không phải 01-scan. Các dự án costing cũ nằm ngoài phạm vi theo `CLAUDE.md` nên **không có dữ liệu `QUAN_SAT` của 01-scan để chấm**.
- Cách 01-scan hiện chống ảo giác (đọc từ `SKILL.md`, `buoc/4-ban-ve.md`): điều phối giao từng bước cho subagent ngữ cảnh sạch; `pd.py check` kiểm máy; Gemini đọc lần hai; mỗi giá trị phải có link `[nhãn, file, trang]`.

## 2. Việc đã làm

### 2.1 Nghiên cứu repo (WebSearch + WebFetch)
Đọc README/trang mô tả: OpenConstructionERP, matchline (issue #746), floorplan-pipeline, Floorplan-Dimractor, opendataloader-pdf, bas-showcase PR#2, cad-floor-plan-extractor, opentakeoff, drawing-analyzer, bài báo Handoff-H1 (arXiv 2608.15032).

Đọc sâu hơn (cả mã nguồn/tài liệu):
- `drawing-analyzer/src/drawing_analyzer/anchor.py` và `tiling.py` — lấy qua `raw.githubusercontent.com`, đọc được. Nội dung neo trích dẫn (chuẩn hoá, bậc EXACT/FUZZY/UNANCHORED, ngưỡng 0,85, veto số, nới 8 pt) và `choose_grid` (thử mọi lưới r·c ≤ 99, DPI tối thiểu theo ANSI E, chồng lấn 8%).
- `matchline/docs/pdf_schedules.md` và `ingest/sheet_index.py` — **qua trang GitHub, do mô hình nhỏ của WebFetch tóm tắt**, nên các hằng số (0,85, 1,4 / 2,2 chiều cao dòng, `MIN_INDEX_ROWS = 3`, regex số tờ) cần kiểm lại bằng mở file gốc.
- `matchline/ingest/pdf_schedules.py` (39 KB): URL raw trả 404, **chưa đọc**.
- Mọi repo khác: chỉ README. Cách `cad-floor-plan-extractor` nhận diện cửa **chưa xác định** (README không nói).

Kết luận nghiên cứu (độ chắc: cao với anchor.py/tiling.py, vừa với phần còn lại): không có repo public nào làm đúng bài toán quét hồ sơ thầu ra bảng dữ liệu; các kỹ thuật dùng được là (a) neo trích dẫn về toạ độ text, (b) đọc bảng schedule bằng lưới vector, (c) lọc lớp OCG, (d) phân tầng model.

### 2.2 Đối chiếu với code hiện có
Đã đọc: `SKILL\01-scan\SKILL.md`, `JSON.md`, `buoc\4-ban-ve.md`, `pd.py` dòng 739–829 (`run_checks`) và danh sách hàm, `scan_core.py` dòng 1–60.
- `run_checks` kiểm: cột bắt buộc, kiểu số/LIST, link có file và trang nằm trong file, BOQ khớp, `QUAN_SAT` nguồn HÌNH/TEXT phải có tờ/chi tiết, chuỗi cộng, trang không có lớp text. **Không thấy** kiểm giá trị đã ghi có thật trên trang được link.
- **Giới hạn của kết luận này:** tôi chỉ đọc `run_checks`. Các lệnh `doichieu` (pd.py:1164) và `do` (1253) chưa đọc; nếu có kiểm nội dung thì kết luận "chưa có" sai.

### 2.3 Các bài test (chỉ đọc, không sửa dự án)
Script ở `DOCS\scan-toi-uu\neo-test\neo_test.py … neo_test7.py` (đã sửa đường dẫn; chạy bằng `py -I -X utf8 <file>`; `neo_test2…7` đọc `neo_test.py`/`neo_test2.py`… theo đường dẫn tuyệt đối trong DOCS). Dữ liệu: `10_TENDER\_data\MAT.mat.json` (91/92 dòng có link), `FILE_INDEX.kiemke-tool.json` (ID → đường dẫn + tên file trong zip), 3 PDF spec F0042/F0043/F0044 trong `OneDrive_1_6-10-2026.zip`.

| # | Bài test | Kết quả đo được | Ghi chú / giới hạn |
|---|---|---|---|
| A | Lớp text 3 PDF spec | 110 trang đều có text | — |
| B | Giá trị thật vs trang link, bản chuẩn hoá 1 | 60 KHỚP / 31 LECH | 31 LECH là **báo oan** do tokenization (dấu phẩy, `69 -2393`, đơn vị mm); đã mở text từng ca để xác nhận |
| B2/B3 | Sau khi sửa (bỏ ký tự không chữ-số, so chuỗi liền; kích thước so từng số) | 91/91 KHỚP | **Tôi chỉnh thuật toán trên chính bộ này → lạc quan**, cần bộ khác để kiểm |
| C | Đổi 1 chữ số trong mã/kích thước (giả lập LLM ghi sai) | bắt 700/702 (99,7%) | Lọt 2 ca `5–10mm`→`2–10mm`/`1–10mm`. Là lỗi **giả lập**, không phải lỗi thật của subagent |
| D | Gắn link sang trang khác | bắt 91/91; gợi ý trang đúng duy nhất 89/91 (2 ca nhiều ứng viên) | Mỗi file chỉ 21–62 trang |
| E | Mã VL tình cờ có trên trang ngẫu nhiên | 0/91 | Mã VL (ví dụ `ST-02`) đủ đặc hiệu |
| F | Lớp text bản vẽ ngoài zip (101 file <80MB, mẫu 184 trang) | 76% trang không có text | Toàn bản as-built 2012 (ảnh scan) — **mẫu lệch** |
| G | Lớp text bản vẽ trong zip (32 file / 104 trang, bộ 2026) | 2% không text; 46% trang ít số | Đã xem tay: trang ít số là trang chữ (preamble, spec, form lighting). Mẫu lấy trang 0, n/3, 2n/3, n−1 → nghiêng về trang bìa |
| H | Chuỗi kích thước (1.021 chuỗi rút từ text 401 trang), đột biến "khó" | Kiểm theo **tập số**: bắt 212/3063 (7%). Kiểm **theo thứ tự**: 2632/3063 (86%). Đổi chỗ 2 chiều: tập 0/37, thứ tự 36/37 | **15% chuỗi thật (157/1021) không tự khớp** vì cách ghép từ của script sai; một phần "chuỗi" là rác (số điện thoại, `#1-02`). Con số 86% chưa đáng tin tuyệt đối |

## 3. Kết luận hiện tại (cần session khác kiểm)

1. Kiểm "mã vật liệu/model có nằm trên trang được link" hoạt động tốt trên bảng spec có lớp text (B3, C, D, E). **Độ chắc: vừa** (một bộ dữ liệu, tự tinh chỉnh).
2. Kiểm kích thước bằng tập số vô dụng; cần kiểm theo thứ tự (H). **Độ chắc: vừa**, script còn lỗi ghép từ.
3. Không cần thêm cột "Trích nguyên văn" vào form: `Mô tả` là tiếng Việt diễn đạt lại nên khớp cả câu sẽ hỏng; kiểm mã/số đã đủ (B, C). **Độ chắc: vừa.**
4. LECH phải là cảnh báo "XEM", không phải lỗi chặn `check` (B: 31 báo oan ở bản đầu).
5. Bản vẽ ảnh scan (as-built 2012) → chỉ ra "KHÔNG KIỂM ĐƯỢC"; bản vẽ 2026 trong zip có text (G). **Độ chắc: vừa**, mẫu lệch.
6. **Chưa kiểm chứng:** hiệu quả trên `QUAN_SAT` thật của 01-scan (kích thước tủ/cabinet từ ảnh crop); trang chứa nhiều hạng mục; bản vẽ có số vẽ bằng nét (trường hợp SOFITEL costing cũ có, ghi trong `modules` — chưa gặp trong mẫu G).

## 4. Việc sẽ làm (chưa làm, chờ user chọn)

User chưa chọn hướng. Hai lựa chọn đã nêu:
- (a) User chỉ định một dự án + đợt có `QUAN_SAT` thật để chấm báo oan/bỏ sót với dữ liệu thật. Theo `CLAUDE.md`, cần user cho phép mở dự án đó.
- (b) Làm luôn bản đầu, chạy chế độ cảnh báo, đo trong lần scan thật tiếp theo.

Thiết kế bản đầu đã đề xuất (chưa viết code):
- Lệnh mới `neo` trong `pd.py` (chưa đụng `scan_core.py`; file đó yêu cầu chạy `DOCS\scan-toi-uu\bench.py` trước/sau khi sửa).
- Với mỗi dòng `QUAN_SAT` (và `Kích thước BV` ở `BOQ_TONGHOP`) có link PDF: lấy mã/số trong `Giá trị`, mở đúng trang link, so với text trang (bỏ ký tự không chữ-số, so chuỗi liền; kích thước so theo **thứ tự** các số trong cửa sổ ≤ 3 số).
- Ba trạng thái: KHỚP · LECH (cảnh báo XEM, không chặn) · KHÔNG KIỂM ĐƯỢC (trang không có text). Dòng `Nguồn = BOQ` bỏ qua.
- Việc để sau: đọc bảng schedule bằng lưới vector (matchline); lọc lớp OCG; phân tầng model theo bước.
- Không sửa form `PROJECT_DATA_FORM.xlsx` (quy tắc: chỉ khi user yêu cầu).

## 5. Danh sách để session khác kiểm

- [ ] Mở trực tiếp `ingest/pdf_schedules.py` và `ingest/sheet_index.py` của matchline, đối chiếu các hằng số ở mục 2.1.
- [ ] Đọc `cmd_doichieu` và `cmd_do` trong `pd.py` xem có kiểm nội dung giá trị với text trang chưa (mục 2.2).
- [ ] Sửa bug ghép từ trong `neo_test7.py` (dùng `get_text("words")` sắp theo block/line/word nhưng 15% chuỗi không tự khớp) rồi chạy lại bài H.
- [ ] Chạy lại bài B–E trên một bộ MAT/spec **khác** (không dùng để chỉnh thuật toán).
- [ ] Lấy mẫu trang chi tiết (không chỉ bìa) của bản vẽ 2026 để đo tỉ lệ số vẽ bằng nét.
- [ ] Kiểm lại kết quả bằng lệnh, không tin số trong file này: `py -I -X utf8 "D:\COSTING EXCEL\DOCS\scan-toi-uu\neo-test\neo_test3.py"` (kỳ vọng 91 KHOP) và `neo_test4.py` (kỳ vọng ~99,7% bắt đột biến; có random seed).

## 5b. Kết quả kiểm lại (session khác, 29/10/2026)

Chạy lại script và mở file gốc; không sửa skill, không mở dự án khác.

| Mục | Cách kiểm | Kết quả |
|---|---|---|
| B3 | `neo_test3.py` | `{'KHOP': 91}` ✔ |
| C3 | `neo_test4.py` (file này chạy **cả C3 và F**, không chỉ C) | 700/702, lọt 2 ca `GL-02 5–10mm` ✔ |
| F | `neo_test4.py` | 139/184 trang không text (76%) ✔ |
| H | `neo_test7.py` | thứ tự 2637/3063 (lần trước 2632), tập 212/3063, đổi chiều 36/37. Lệch nhỏ dù có `random.seed(5)` — nghi thứ tự duyệt `set` đổi theo hash ngẫu nhiên; muốn tái lập: `PYTHONHASHSEED=0` hoặc sort trước khi random. Mẫu lọt có `['1','02']` từ địa chỉ `#1-02` → xác nhận có chuỗi rác |
| 2.2 | đọc `pd.py:1164–1265` | `cmd_doichieu` chỉ chấm độ tin cậy từ trạng thái QUAN_SAT/MAP/Gemini; `cmd_do` chỉ đếm ô. **Không** kiểm giá trị với text trang → kết luận "chưa có" đúng. (`pd.py` chỉ gọi `get_text` ở dòng 236 kiểm lớp text, `cmd_chu`, `cmd_so`) |
| matchline | tải raw nhánh **`develop`** (nhánh mặc định; URL `main` → 404, đó là lý do `pdf_schedules.py` 404 trước đây) | `sheet_index.py`: `TITLE_MATCH_RATIO = 0.85`, `MIN_INDEX_ROWS = 3` ✔. `pdf_schedules.py` (39 KB, đã đọc được): `UNRULED_CONT_GAP = 1.4`, `UNRULED_ROW_GAP = 2.2` ✔; thêm: bảng có khung `CONFIDENCE = 0.9`, bảng không khung `CONFIDENCE_UNRULED = 0.75`, `UNRULED_WORD_GAP = 0.6`, `MIN_ROW_LINES = MIN_COL_LINES = 3`, `MAX_ROW_PT = 72` |
| drawing-analyzer | tải raw | `anchor.py`: `_PAD_PT = 8.0`, `_FUZZY_WINDOW_MIN_OVERLAP = 0.85`, EXACT/FUZZY/UNANCHORED ✔. `tiling.py`: ≤99 ô, `DEFAULT_OVERLAP_FRAC = 0.08`, sàn ANSI E ✔ |

Checklist mục 5: đã xong dòng 1 (matchline), 2 (doichieu/do), 6 (chạy lại lệnh). Còn: sửa ghép từ `neo_test7`, bộ MAT/spec khác, mẫu trang chi tiết bản vẽ 2026.

### Ý kiến của session kiểm (chờ user chọn)
1. **Bộ thử lệch đối tượng:** mọi test chạy trên `MAT` của **10-tender** (bảng spec có text), còn thiết kế `neo` nhắm `QUAN_SAT` của **01-scan** (kích thước từ ảnh crop) — phần này chưa đo. Đề xuất: viết `neo` dạng hàm thuần trong `SKILL\_core\scan_core.py` (chỉ thêm, không sửa hàm cũ; vẫn chạy `bench.py`), áp cho 10-tender trước (dự án active, có dữ liệu đo), 01-scan dùng lại ở chế độ cảnh báo lần scan thật sau (= hướng b).
2. **Bộ dữ liệu khác lấy ngay trong SOFITEL:** spec/schedule ở TCQ/ADD khác F0042–F0044 → kiểm độc lập mà không vi phạm `CLAUDE.md`.
3. **Đổi tên trạng thái "KHỚP" → "CÓ NEO":** chỉ chứng minh mã/số có trên trang link, chưa chứng minh gán đúng hạng mục (trang nhiều hạng mục). Chỉ LECH / KHÔNG KIỂM ĐƯỢC đưa lên user (duyệt theo ngoại lệ).
4. Đồng ý kết luận 3.3, 3.4, 3.5.
5. Chưa tin 86% bài H trước khi sửa ghép từ (15% chuỗi không tự khớp) và lọc rác (số điện thoại, `#x-xx`).
6. Khi làm đọc bảng schedule: tham khảo cách gom bảng không khung trong `matchline/ingest/pdf_schedules.py` (nhánh `develop`).

## 6. Nguồn tham khảo

- https://github.com/Abe-Borg/drawing-analyzer (anchor.py, tiling.py)
- https://github.com/anchapin/matchline (docs/pdf_schedules.md, ingest/sheet_index.py, issue #746)
- https://github.com/Kentucky-ai/opentakeoff · https://github.com/shorokhlev-sketch/floorplan-pipeline · https://github.com/opendataloader-project/opendataloader-pdf · https://github.com/datadrivenconstruction/OpenConstructionERP · https://github.com/jasoncobra3/Floorplan-Dimractor · https://github.com/swoods5565/bas-showcase/pull/2 · https://github.com/HMPCGithub/cad-floor-plan-extractor · https://arxiv.org/abs/2608.15032
