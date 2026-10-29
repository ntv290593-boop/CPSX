# PLAN — Kiểm neo `neo` (giá trị đã ghi có thật trên trang được link) — 29/10/2026

Căn cứ: `DOCS\HANDOFF_NEO-ANCHOR_2026-10-29.md` (mục 2–5b). Chưa sửa skill nào. Plan chờ user duyệt các câu hỏi Q1–Q4.

## 1. Tổng hợp handoff

**Vấn đề:** `check` của 01-scan và 10-tender chỉ kiểm hình thức (cột, kiểu, link có file/trang). Không có bước nào kiểm *giá trị LLM ghi có thật trên trang nó link tới*. Đã xác nhận bằng đọc code: `run_checks` (pd.py:739), `cmd_doichieu`, `cmd_do` đều không làm.

**Ý tưởng `neo`** (học từ `drawing-analyzer/anchor.py`): với mỗi ô có link PDF, lấy mã và số trong `Giá trị`, mở đúng trang link, tìm trong lớp text. Kết quả là một trong 3 trạng thái, chỉ cảnh báo, không chặn.

**Đã đo được** (chỉ trên `MAT` của 10-tender SOFITEL, 3 PDF spec F0042–F0044, có lớp text):

| Bài | Kết quả | Mức tin |
|---|---|---|
| Mã/model có trên trang link (B3) | 91/91 | Lạc quan: thuật toán chỉnh trên chính bộ này |
| Đổi 1 chữ số giả lập (C) | bắt 700/702 | Lỗi giả lập, không phải lỗi thật |
| Link sang trang khác (D) | bắt 91/91; gợi ý trang đúng 89/91 | File ngắn (21–62 trang) |
| Mã VL trùng ngẫu nhiên (E) | 0/91 | Tốt: mã VL đủ đặc hiệu |
| Kích thước kiểm theo tập số / theo thứ tự (H) | 7% / 86% | 86% chưa tin: 15% chuỗi thật không tự khớp do ghép từ sai, có chuỗi rác |
| Lớp text bản vẽ | as-built 2012: 76% không text; bộ 2026 trong zip: 2% | Mẫu lệch về trang bìa |

**Kết luận đã thống nhất:** kiểm theo mã/số, không khớp cả câu (`Mô tả` là tiếng Việt diễn đạt lại). LECH là XEM, không phải LỖI. Trang không text thì ra KHÔNG KIỂM ĐƯỢC. Không thêm cột vào form.

**Chưa kiểm chứng:**
- `QUAN_SAT` thật của 01-scan (kích thước đọc từ ảnh crop).
- Trang chứa nhiều hạng mục.
- Bản vẽ có số vẽ bằng nét.

**Lệch đối tượng:** mọi số đo ở trên làm trên bảng spec của 10-tender, trong khi đích là kích thước bản vẽ của 01-scan.

> **Cập nhật 29/10 (user):** tập trung cho **01-scan**. Bỏ P4 (10-tender); tender chỉ dùng làm nguồn dữ liệu thử phụ. Mục 3b bên dưới thay cho P2–P5 ở mục 3.

## 2. Câu hỏi cần user chốt

| # | Câu hỏi | Đề xuất |
|---|---|---|
| Q1 | Bộ thử cho 01-scan: dùng `_ACTIVE\SOFITEL\01_scan` (PROJECT_DATA v1–v8 + `_dwg_tangA` đáp án DWG của đợt costing trước, theo `ACTIVE.md` cần user cho phép)? | **Cho phép, chỉ để đo thuật toán**, không lấy số liệu sang tender |
| Q2 | Đặt code ở đâu? | Lệnh `neo` + hàm trong **`pd.py`** (đổi so với bản trước). Chuyển sang `scan_core.py` khi tender cần |
| Q3 | Tên trạng thái | **CÓ NEO** · **LỆCH** · **KHÔNG KIỂM ĐƯỢC**. Không gọi là "KHỚP" vì chỉ chứng minh mã/số có trên trang, chưa chứng minh gán đúng hạng mục |
| Q4 | LỆCH có hạ `Độ tin cậy` ở `doichieu` không? | **Chưa**, ở bản 1. Đo báo oan trước, rồi mới nối vào độ tin cậy |

## 3. Kế hoạch

### P1 — Sửa bộ đo trước khi viết code thật (`DOCS\scan-toi-uu\neo-test`)
1. `neo_test7.py`: ghép từ theo `get_text("words")`, sắp theo (block, line, x). Nối các token cùng dòng khi khoảng cách < 0,6 × chiều cao chữ (cùng ý `UNRULED_WORD_GAP` của matchline). Xử lý dạng `69 -2393`, `1,200`, `1.200`, đơn vị `mm`.
2. Lọc rác trước khi coi là chuỗi kích thước: số điện thoại, `#x-xx`, ngày, số tờ (`ID-SK-52-101`), mã revision.
3. Tái lập được: sort trước khi random, chạy với `PYTHONHASHSEED=0`.
4. Đạt khi: ≥ 98% chuỗi thật tự khớp. Chạy lại bài H, ghi số mới vào handoff mục 5b.

### P2 — Đóng băng thuật toán, chạy trên bộ thử độc lập
Không chỉnh thuật toán theo kết quả của các bộ dưới đây. Lệch thì ghi nguyên nhân, sửa ở vòng sau với bộ khác.

| Bộ | Nguồn (SOFITEL 10_TENDER, cùng dự án active) | Đo gì |
|---|---|---|
| S1 | `SCOPE.banve-*.json`: 155 dòng có `Mã VL` + `Bản vẽ hiệu lực` link tới trang **bản vẽ** | Mã VL có trên tờ bản vẽ; tỉ lệ KHÔNG KIỂM ĐƯỢC; trang nhiều hạng mục |
| S2 | Spec/schedule trong TCQ/ADD khác F0042–F0044 (lọc từ `FILE_INDEX`) | Lặp lại B–E trên bộ khác |
| S3 | Mẫu trang **chi tiết** bản vẽ 2026 (không lấy trang bìa), khoảng 40 trang | Tỉ lệ số vẽ bằng nét / không có số trong text |

Với mỗi bộ: đếm 3 trạng thái, **xem tay toàn bộ LỆCH** và 20 dòng CÓ NEO ngẫu nhiên. Ghi bảng báo oan / bỏ sót.

Dừng lại hỏi user nếu: báo oan > 20% số LỆCH, hoặc KHÔNG KIỂM ĐƯỢC > 50% ở S1/S3. Khi đó `neo` cho bản vẽ chưa đáng làm, chỉ giữ phần spec.

### P3 — Hàm lõi trong `scan_core.py` (chỉ thêm)
- `neo_chuan(text)`: bỏ dấu, chữ thường, bỏ ký tự không chữ-số.
- `neo_tach(gia_tri)` → `{ma: [...], so: [[...], ...]}`: mã theo họ mã (`ST-02`, `WD-01`…), chuỗi số kích thước theo thứ tự.
- `neo_trang(words_trang, tach, clip=None)` → `(trạng thái, lý do, vị trí)`:
  - Mã: tìm chuỗi liền trên text đã chuẩn hóa.
  - Số: tìm theo **thứ tự** trong cửa sổ ≤ 3 số liền nhau.
  - **Trang có text nhưng không có số nào** (hoặc vùng clip ít số) → KHÔNG KIỂM ĐƯỢC, không ra LỆCH. Chỗ này chặn báo oan khi số vẽ bằng nét.
  - Trang `lop = KHÔNG` → KHÔNG KIỂM ĐƯỢC.
- `neo_goi_y(ix, tach)`: khi LỆCH thì gợi ý trang khác có đủ neo (dùng lại `chi_muc`). Bài D cho 89/91.
- Trước/sau khi sửa: chạy `DOCS\scan-toi-uu\bench.py` và `td.py ungvien --do`, kết quả phải giữ nguyên mốc (@12 93% / 78%; SOFITEL @12 61%).

### P4 — 10-tender: lệnh `td.py neo <GOC>`
1. Đọc text trang từ cache `_data\_text` (có sẵn từ `td.py trang`, không mở lại PDF hay zip).
2. Duyệt mọi ô có link PDF ở `MAT`, `SCOPE`, `VIEC_MO`. Bỏ qua dòng có `Nguồn = BOQ` và link tới Excel.
3. Ghi `_data\_neo.json` theo dạng `{"<sheet>|<khóa>|<cột>": {"tt", "ly_do", "goi_y"}}`. File bắt đầu bằng `_` nên không ghép vào sheet.
4. `check` đọc `_neo.json`, in `XEM NEO LỆCH …` / `XEM NEO KHÔNG KIỂM ĐƯỢC (n ô)` ở dạng gộp. Không bao giờ ra LỖI.
5. Cập nhật `SKILL\10-tender\SKILL.md` và `JSON.md`: chạy `neo` trước `check` cuối. Chỉ đưa LỆCH lên user (duyệt theo ngoại lệ).

### P5 — 01-scan: lệnh `pd.py neo <GOC> <DOT>`
1. Phạm vi: `QUAN_SAT` có `Nguồn = HÌNH/TEXT` và link PDF; `BOQ_TONGHOP.Kích thước BV`; `Bản vẽ đối chứng`.
2. Dùng `clip` nếu link có vùng; không có thì xét cả trang.
3. Ghi `_data\<DOT>\_neo.json`. `check` in XEM như P4.
4. Sửa `buoc/4-ban-ve.md` ("Trước khi trả kết quả": chạy `neo`, đọc lại ô LỆCH) và `JSON.md` (bảng file phụ trợ).
5. Không đụng `doichieu` (theo Q4).

### P6 — Đo trong lần scan/tender thật kế tiếp
- Ghi: số ô kiểm được, LỆCH, báo oan (user hoặc LLM xác nhận đúng), bỏ sót (lỗi user tìm ra mà `neo` không báo).
- Sau 2 lần đo đạt (báo oan < 10% LỆCH) thì đề xuất Q4: LỆCH chưa giải thích được sẽ hạ `Độ tin cậy` xuống TB.

## 3b. Kế hoạch tập trung 01-scan (thay P2–P5)

**Ô nào kiểm được.** Link của 01-scan chỉ có `[nhãn, file, trang, ô đích]`, không có vùng crop, nên chỉ kiểm được ở mức trang.

| Cột / trường | Kiểm gì | Mức tin dự kiến |
|---|---|---|
| `QUAN_SAT` `Mã hoàn thiện trên BV` (WD/PT/LAM…) | mã có trên trang | Cao (bài E: mã VL đặc hiệu) |
| `QUAN_SAT` `Kích thước BV`, `BOQ_TONGHOP.Kích thước BV`, chuỗi `450+450=900` | các số theo thứ tự (cửa sổ ≤ 3) trên trang | **Chưa biết**: trang bản vẽ có rất nhiều số 450/900, dễ trùng ngẫu nhiên |
| `Revision tờ`, `Bộ tờ đã đọc`, nhãn `Bản vẽ đối chứng` | số tờ/rev có trên khung tên trang link (so `PAGES.Số BV`) | Cao |
| Lắp đặt, Có cánh, Ngăn kéo, Kệ… (đọc bằng hình) | không kiểm, ghi `—` (không tính là KHÔNG KIỂM ĐƯỢC) | — |
| Dòng `Nguồn = BOQ / GEMINI` | bỏ qua | — |

**N1 — Sửa bộ đo** (= P1 cũ): ghép từ, lọc rác, tái lập được.

**N2 — Đo trên dữ liệu 01-scan** (giữ nguyên thuật toán, không chỉnh theo kết quả):
- D1 (cần Q1): `QUAN_SAT` + `Kích thước BV` trong PROJECT_DATA mới nhất của `SOFITEL\01_scan`, đối chiếu đáp án `_dwg_tangA`. Đo: tỉ lệ CÓ NEO / LỆCH / KHÔNG KIỂM ĐƯỢC; báo oan (LỆCH nhưng DWG xác nhận đúng); bỏ sót (DWG nói sai mà ra CÓ NEO).
- D2 — **bài trùng ngẫu nhiên cho kích thước** (mới, quan trọng nhất): lấy chuỗi kích thước thật của mã A, kiểm trên trang của mã B. Tỉ lệ ra CÓ NEO = tỉ lệ lọt. Lọt > 15% thì kiểm kích thước mức trang vô dụng; khi đó chỉ giữ kiểm mã + số tờ, kích thước để sau khi link có vùng crop.
- D3: tỉ lệ trang bản vẽ SOFITEL có số trong lớp text. Bộ này nhiều số vẽ bằng nét, kỳ vọng phần lớn ra KHÔNG KIỂM ĐƯỢC.
- D4 (phụ, không cần Q1): `SCOPE.banve-*.json` của tender (155 dòng Mã VL → trang bản vẽ) để đo kiểm mã trên bản vẽ.
- **Dừng, báo user số đo.**

**N3 — Lệnh `pd.py neo <GOC> <DOT>`:**
1. Đọc text trang từ cache `kiemke` (`_data\<DOT>\_text\F*.json`), không mở lại PDF. Dùng `get_text("words")` chỉ khi cần thứ tự số; cache words theo trang.
2. Chỉ kiểm các cột ở bảng trên mà N2 cho kết quả đạt.
3. Trang `Lớp text = KHÔNG`, hoặc có text nhưng không có số → KHÔNG KIỂM ĐƯỢC.
4. LỆCH thì gợi ý trang khác có đủ neo (dùng `build_index`/`chi_muc` sẵn có).
5. Ghi `_data\<DOT>\_neo.json`. `check` đọc file này, in `XEM NEO LỆCH <Mã> <Trường> …` và 1 dòng tổng KHÔNG KIỂM ĐƯỢC. Không bao giờ là LỖI.
6. Chạy `DOCS\scan-toi-uu\bench.py` trước/sau (không đổi mốc tìm trang).

**N4 — Đưa vào quy trình:**
- `buoc/4-ban-ve.md` "Trước khi trả kết quả": chạy `neo`, ô LỆCH thì mở lại đúng trang, sửa hoặc ghi lý do vào `Giá trị`.
- `SKILL.md` mục 2: sau bước 4, điều phối chạy `neo`, gửi danh sách LỆCH cho subagent (giống vòng sửa LỖI, tối đa 2 vòng).
- `JSON.md`: thêm `_neo.json` vào bảng file phụ trợ.
- Không sửa form, không đụng `doichieu` (Q4).

**N5 — Đo ở lần scan 01 thật kế tiếp** (= P6).

**Để sau:** link có vùng crop (phần tử thứ 5 `clip` trong link; đổi `JSON.md` + `norm_links`) để kiểm kích thước trong vùng. Chỉ làm nếu D2 cho thấy kiểm mức trang lọt nhiều.

Thứ tự: N1 → N2 → **dừng báo user** → N3 → N4 → N5.

## 3c. Kết quả N2 (29/10, user cho phép dùng `SOFITEL\01_scan` để đo)

Script: `DOCS\scan-toi-uu\neo-test\neo01_d1d3.py` (PROJECT_DATA_SOFITEL_v8 + `_dwg_tangA`) và `neo01_d4.py` (SCOPE tender). Chạy `py -I -X utf8 <file>`. Thuật toán giữ nguyên, không chỉnh theo kết quả.

| Bài | Kết quả | Đọc thế nào |
|---|---|---|
| D1a mã đọc hình (v8 cột H) trên trang tờ | 5 CÓ NEO / **407 LỆCH** | **Báo oan toàn bộ.** Lớp text bản vẽ GA Design chỉ có khung tên + ghi chú chung (ví dụ CO-TY-703: 2.870 ký tự, không có `WD`); mã VL, tag vẽ bằng nét |
| D1b trùng ngẫu nhiên mã | 1/1208 | Không có ý nghĩa, vì mã không có trong text |
| D1c đột biến mã | 1236/1236 | Không có ý nghĩa, cùng lý do |
| D1d số tờ trong nhãn link (cột P) trên trang link | **199/199 CÓ NEO**; số tờ A trên trang B: **0/299** | **Dùng được ngay**: bắt link sai trang |
| D1e tên chi tiết (cột Q) trên trang tờ | 13/13 | Mẫu nhỏ, tên chi tiết là text |
| D3 kích thước DWG có trong lớp text PDF | 420/5209 số (8%), 0/218 tờ đạt ≥ 50% | 8% là trùng tình cờ với ghi chú/cao độ; **kích thước không có trong text** |
| D4 Mã VL tender (SCOPE) trên trang bản vẽ 2026 | 0 CÓ NEO / 930 LỆCH | Bộ 2026 cùng kiểu GA Design. Text vùng vẽ chỉ có số chi tiết, tỉ lệ (`01 10`), cao độ (`LEV 2500`, `FFL ±0`) |

**Kết luận:**
1. Luật "trang có lớp text → kiểm được" **sai** với bản vẽ kiểu GA Design. Nếu chạy `neo` như thiết kế N3 cũ thì gần như mọi ô mã/kích thước ra LỆCH oan.
2. Luật "trang có mã dạng XX-NN" cũng sai: số tờ (`ID-SK-52-601` → `SK52`), `FFL0`, `LEV850` đều khớp regex.
3. Kiểm được chắc chắn: **số tờ / số chi tiết / tên chi tiết** (text khung tên + nhãn chi tiết) và **bảng spec/ML có text** (MAT tender: 91/91, đột biến 99,7%).
4. Bản vẽ có kích thước nằm trong text (NORTON theo memory) thì kiểm được, nhưng **chưa có dữ liệu được phép để đo**.

**Sửa thiết kế N3** (thay mục N3 ở 3b):
- **N3.1 (làm được ngay):** kiểm số tờ của nhãn link (`Bản vẽ đối chứng`, link `Tờ · Chi tiết` của QUAN_SAT, `Bộ tờ đã đọc`) có trên khung tên trang link. Lệch thì gợi ý trang có số tờ đó (đã đo: duy nhất, 0/299 trùng). Kèm số/tên chi tiết khi nhãn có `#nn`.
- **N3.2:** link spec/ML (`SPEC`, `Spec · ML`, QUAN_SAT `Nguồn = SPEC`): kiểm mã/model theo cách MAT tender.
- **N3.3 (luật KHÔNG KIỂM ĐƯỢC theo họ):**
  - Mã VL: chỉ kiểm khi text trang có ≥ 1 mã thuộc **họ mã của dự án** (lấy từ FINISHES/ML: WD, MT, PT, ST, GL…), không tính token nằm trong số tờ.
  - Kích thước: chỉ kiểm khi vùng vẽ có ≥ 10 số ≥ 100, sau khi bỏ dòng lặp ở ≥ 30% trang của file (khung tên/ghi chú) và cao độ/tỉ lệ.
  - Ngưỡng 10 **chưa đo** trên bản vẽ có text. SOFITEL sẽ ra KHÔNG KIỂM ĐƯỢC toàn bộ, đúng như mong muốn.
- **Để chọn sau:** nguồn neo kích thước cho bản vẽ vẽ nét: (a) `_MASTER\scan\pdf_dim_decode.py` (chỉ font SOFITEL, đúng 99%); (b) DWG tầng A khi có DWG cùng bản. Cả hai là "lớp text thay thế", có thể cắm vào `neo_trang`.

## 3d. N3.1 + N3.2 đã làm (29/10, user chọn)

Lệnh `pd.py neo <GOC> <DOT>` (khối `B7 neo` trong `pd.py`; không sửa `scan_core.py`, form, `doichieu`). Ghi `_data\<DOT>\_neo.json`; `check`/`build` in `XEM NEO LỆCH …` + 1 dòng tổng (luôn hiện, không bị cắt ở 60 dòng XEM). Link đổi sau lần chạy thì `check` báo chạy lại (so dấu md5 của các link).

- **N3.1 số tờ:** cột `Bản vẽ đối chứng`, `Tờ đếm KL`, `Tờ chi tiết`, `Tờ mặt bằng`, QUAN_SAT `Tờ · Chi tiết` (Nguồn HÌNH/TEXT). Số tờ trong nhãn nhận theo **đuôi số BV trong PAGES** (không dùng regex riêng SOFITEL): `504.4`, `b/504.4`, `HC-1BR-707`, `ID-CO-TY-703` đều nhận ra. CÓ NEO khi: Số BV của trang link thuộc tờ đó, hoặc text trang có đủ số tờ (khớp có biên: `709` ≠ `709A`), hoặc text trang có đúng token nhãn (callout `b/502.1`, tên hình `MB 121`). Có `#nn TÊN` thì kiểm thêm ≥ 80% từ của tên.
- **N3.2 spec/ML:** `Spec · ML` (+ `Thông số riêng`), `SPEC.File · Trang` (+ `Mục`), QUAN_SAT Nguồn SPEC (+ `Giá trị`). Mã = token có cả chữ và số, ≥ 3 ký tự (bỏ `F024`, `p11`, `r16`, `18mm`). Ô có nhiều link: chỉ lấy mã trong nhãn. Đích PDF: text cache; đích Excel: cả dòng của ô link.

**Kết quả** (`DOCS\scan-toi-uu\neo-test\neo01_regress.py`, gọi thẳng hàm `pd.py`):
| Bộ | Gốc | Đột biến trang ±1 |
|---|---|---|
| SOFITEL v8 cột P (D1d) | 199 CÓ NEO / 0 LỆCH / 155 KKĐ (nhãn không có số tờ) | bắt 193/193 link đổi được (6 link ở trang đầu/cuối file không đổi) |
| Bench cache Blanca (mọi cột) | 1201 CÓ NEO / 0 LỆCH / 1181 KKĐ | bắt 951/1201 (79%): nhãn gọn kiểu `p21 504.4` yếu hơn vì trang cạnh nhau in callout của nhau |
| Thử tổng hợp N3.2 (7 ca) | đúng 7/7: sai dòng ML, sai model, sai trang spec đều bắt + gợi ý đúng | — |

Lần đầu chạy có 19 LỆCH oan trên Blanca (callout `b/502.1`, tên hình `121`) → thêm luật "token nhãn có trên trang"; 1 lọt `709`/`709A` → khớp có biên. N3.2 **chưa đo trên ML/spec thật của 01-scan** (Blanca cache không có file Excel ML; spec SOFITEL ngoài phạm vi được phép) — đo ở lần scan thật kế tiếp (N5). `bench.py` qua cổng (93% / 78%).

## 4. Để sau (không làm trong plan này)
- Đọc bảng schedule bằng lưới vector theo `matchline/ingest/pdf_schedules.py` (nhánh `develop`): bảng có khung tin cậy 0,9, không khung 0,75.
- Lọc lớp OCG.
- Chia ô ảnh theo `drawing-analyzer/tiling.py` cho trang lớn.
- Phân tầng model theo bước.

## 5. Rủi ro

| Rủi ro | Chặn bằng |
|---|---|
| CÓ NEO nhưng gán nhầm hạng mục (trang nhiều món) | Tên trạng thái "CÓ NEO"; P2-S1 đo riêng loại trang này; về sau dùng clip/vị trí |
| Số vẽ bằng nét → báo LỆCH oan | Luật "có text nhưng không có số → KHÔNG KIỂM ĐƯỢC" (P3) |
| Tự chỉnh thuật toán theo bộ thử nên số đo lạc quan | Đóng băng ở P2; mỗi vòng sửa dùng bộ mới |
| Sửa lõi làm hỏng tìm trang | Chỉ thêm hàm; chạy 2 phép đo trước/sau (P3) |
| Lấy số liệu nhầm dự án | Mọi bộ thử lấy trong `_ACTIVE\SOFITEL\10_TENDER` |

## 6. Thứ tự và điểm dừng
P1 → P2 → **dừng, báo user số đo** → P3 → P4 → (user dùng thử trên tender) → P5 → P6.
