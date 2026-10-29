# PLAN — Viết lại phần tìm trang của skill 01-scan (2026-10-28) — PHẦN A ĐÃ LÀM 28/10

Căn cứ: `DOCS\HANDOFF_SCAN-TOI-UU_2026-10-28.md` (mục 3–3c) + kết quả user kiểm `KIEM_MAU_DAP_AN_Blanca_v2.xlsx` (28/10).
Phạm vi: chỉ phần **chọn trang ứng viên** (kiemke → boq → `ungvien` → bước 4 dùng ứng viên). Không đụng phần đọc hình, Gemini, build/check, form (trừ khi user duyệt câu hỏi Q1).

## 0. Mốc chính thức (đáp án Blanca đã sửa theo user kiểm)
User kiểm 57 dòng / 10 mã:
- Đáp án cũ (subagent chọn): 24/25 trang đúng → **chính xác nhưng thiếu**: 21 trang đúng chỉ máy tìm ra (danh mục BV, mặt bằng bố trí, tờ chi tiết kế bên).
- Máy (top-5 hiện tại): 10 trang SAI + 2 KHÔNG CHẮC. Các kiểu sai:
  - file **F004 layout dự án** (concept) cho cửa D-02, gối FR15 — 5 trang;
  - tờ **MEP** (mặt bằng đèn trần, công tắc) cho thân khóa PK03 / gối LO03 — 3 trang;
  - từ khóa vị trí: "Gối nằm **@ phòng ngủ**" khớp "CHI TIẾT CỬA **P. NGỦ**" — 3 trang;
  - tờ phù điêu, mặt đứng không có gối (LO03) — 2 trang.
- Đáp án Blanca: 164 → 184 trang (thêm 21, bỏ 1). Script: `DOCS\scan-toi-uu\eval_v2.py`.

| Bộ đo | Hiện tại @12 | Tốt nhất đã thử @12 | Biến thể |
|---|---|---|---|
| Blanca 10 mã đã kiểm (44 trang) | 73% (@5 27%) | **98%** (@5 84%) | mã chuẩn hóa + lọc khu + BM25 + tên tờ |
| Blanca 69 mã (59 mã chưa kiểm) | 62% | 89% | như trên |
| SOFITEL 155 mã | 44% | 51% | BM25 + tên tờ |
| NORTON D01 cửa | 22% | 100% | BM25 + mã + bỏ file tổng + ghim tờ thống kê |
| NORTON vanity | 0/12 tờ phòng tắm | 12/12 (V01, V02) | + mã BV VC1 (giả định có người gán) |

## 1. Thay đổi (theo thứ tự làm)

### P0. Bộ đo cố định trước khi sửa code
- `DOCS\scan-toi-uu\bench.py`: chạy cả 4 bộ đáp án, gọi **thẳng hàm xếp hạng trong `pd.py`** (không chép lại logic như các script prototype), in bảng @5/@12/@40 + số trang SAI đã biết (từ KIEM_MAU) lọt vào top-5.
- Lớp text các bộ đo lưu cố định ở `DOCS\scan-toi-uu\_cache\` (vài MB), để lần sau không phải trích lại 17 phút và không cần mở `_ARCHIVE`.
- Cổng qua: mỗi bước P1–P3 chạy bench; không bộ nào được tụt so với mốc trên.

### P1. `pd.py kiemke` — dữ liệu phụ cho tìm kiếm (máy làm, không tốn token)
1. **L1 trùng ở mức trang:** md5 lớp text từng trang (bỏ khoảng trắng). Trang trùng ghi `trùng với F00x pN`; giữ bản trong **file cụ thể hơn** (file có khu/khối trong tên, ít trang hơn). File mà mọi trang đều trùng file khác (kiểu file tổng NORTON) → in `L1 FILE TỔNG` và không đưa vào tìm kiếm. Vẫn giữ cả file trong FILES/PAGES.
2. **Loại tờ đoán từ tên tờ:** `danh mục` · `mặt bằng bố trí` · `thống kê` (schedule) · `chi tiết` · `mặt đứng` · `MEP` (điện, đèn, công tắc, ổ cắm, cấp thoát nước, PCCC) · `hoàn thiện` (sàn, trần, tường) · `legend`. Bước 4 sửa khi sai (đã có trong bước 4 mục 8).
3. **Bảng mã trên BV (từ legend):** quét lớp text tìm cặp "mô tả / MÃ" kiểu `BÀN RỬA MẶT / VANITY COUNTER VC1`, `CHẬU RỬA MẶT /LAVABO LA1` → danh sách mã BV + mô tả + trang. Dùng ở P3 để nối mã BoQ tự đặt (V01, F001…) với mã BV. **Chưa test** → làm prototype trên vanity + SOFITEL trước, chỉ giữ nếu bench tăng.
4. Lưu tất cả vào `_data\<DOT>\_timkiem.json` (không vào xlsx — xem Q1).

### P2. Bước kiemke (B) + bước boq — phần LLM điền
- **kiemke B:** quyết định DỰ PHÒNG (concept, layout dự án, fit-out) đã có trong hướng dẫn nhưng code không dùng → ghi danh sách file DỰ PHÒNG vào `_timkiem.json`; `ungvien` bỏ các file này. (Chặn kiểu sai F004.)
- **boq** ghi thêm vào `_timkiem.boq.json`:
  - `khu → file`: mỗi Khu của BOQ_MAP ứng với file/thư mục nào (dựa trên tên file trong FILES; LLM làm vì "Hành lang" ≠ "Area 50 - Guestroom Corridor").
  - `mã cha`: phụ kiện → cửa (PK03 → D-01, D-02…), hạng mục con `-HM`/`-DT` → tủ.
  - `từ khóa tìm`: tên món bỏ phần vị trí (`Gối nằm @ phòng ngủ` → `gối nằm`), thêm tên tiếng Anh nếu bản vẽ tiếng Anh.
- `pd.py check --den boq` kiểm `_timkiem.boq.json` (mã có thật, file có thật).

### P3. Viết lại `pd.py ungvien`
Điểm mỗi trang = cộng các tín hiệu, mỗi tín hiệu **tự bật/tắt theo hồ sơ**:
| Tín hiệu | Bật khi |
|---|---|
| Mã chuẩn hóa (bỏ `_ST`/`_1BR`/`-HM`/`a`, tiền tố `ID-`, `.`/`-`, số 0 đầu) | mã có trên ≥ 1 trang BV |
| Mã BV qua bảng legend (P1.3) | mã BoQ không có trên BV |
| BM25 từ khóa tìm + Mô tả BoQ (ngắn) + tên tờ ×3 | ≥ 50% từ của Mô tả có trong lớp text BV (cùng ngôn ngữ) |
| Từ khóa Hạng mục (logic cũ) | luôn, điểm thấp |

Lọc trước khi chấm: bỏ trang trùng (P1.1), file tổng, file DỰ PHÒNG, file khác khu (P2), tờ MEP/hoàn thiện khi món thuộc 5 xưởng.

**Đầu ra chia nhóm theo loại tờ** (thay một danh sách top-12 trộn chung):
```
## D-01 · Cửa phòng ngủ   [khu: 1br, 2br]
   CHI TIẾT     F005 p40  … (top 8, xếp điểm)
   THỐNG KÊ     F00x pN   (tờ thống kê có chứa mã — luôn kèm)
   MẶT BẰNG     F006 p5, F024 p4   (để đếm KL theo khu)
   DANH MỤC     F005 p3
   CỦA MÃ CHA   (chỉ với phụ kiện/hạng mục con: trang chi tiết của mã cha)
```
Lý do: user tính danh mục BV và mặt bằng bố trí là ĐÚNG; NORTON cần ghim tờ thống kê; nhưng ở Blanca ghim mặt bằng lên đầu làm @5 tụt → tách nhóm thì không phải đánh đổi.

### P4. Sửa hướng dẫn
- `SKILL.md` mục 2 bước 4: định dạng ứng viên mới, cách chia nhóm theo `mã cha` (mã con cùng nhóm với cha).
- `buoc/1-kiem-ke.md`: ghi DỰ PHÒNG vào `_timkiem.json`; đọc dòng `L1 FILE TỔNG` / trang trùng.
- `buoc/2-boq.md`: mục mới cho `_timkiem.boq.json`.
- `buoc/4-ban-ve.md`: đọc ứng viên theo nhóm (chi tiết trước, mặt bằng để đếm, thống kê với cửa); vẫn được tự `tim` khi thiếu.
- `JSON.md`: thêm `_timkiem.*.json`.

### P5. Kiểm cuối
- `bench.py` đạt mốc; chạy `ungvien` trên Blanca **ghi ra scratchpad** (thêm tùy chọn `--ra`), không ghi đè `_data` của dự án.
- Không chạy lại cả skill trên Blanca trừ khi user muốn.

## 1B. Phần B — chuyển việc trích / chép / so từ LLM sang tool (bổ sung 28/10, từ phản hồi "3 lớp")
Nguyên tắc: **tool trích → LLM chỉ quyết phần mơ hồ → tool kiểm**. Lệnh chung trong `pd.py`, không script riêng dự án.

**Verify phản hồi (đo trên `_data` Blanca 27/10):**
| Ý | Kết quả kiểm |
|---|---|
| Subagent có dùng tool `pd.py` | Đúng. `scan_spec.py` có thật (`_MASTER\scan\scan_spec.py`, gọi từ `buoc/3-spec.md`). |
| Bước 2 "LLM chép tay từng dòng BoQ" | **Sai một phần**: lần Blanca subagent tự viết script đọc xlsx (handoff mục 2). Nhưng đó là script tự chế mỗi lần chạy → vẫn nên có lệnh chung `boq-tach`. Khối lượng: boq ghi 946 dòng / 6384 ô. |
| Bước 5 LLM so toàn bộ dù quy tắc đã viết rõ | Đúng. `5-doi-chieu.md` mục 4 là luật rõ (CAO/TB/THAP theo KHÔNG THẤY, MÂU THUẪN, Kiểm chéo Gemini, KL khớp). doichieu ghi 672 dòng / 3170 ô. |
| Bước 4 LLM cộng chuỗi kích thước | Đúng: 315 dòng QUAN_SAT có chuỗi `a+b=c`. Tool kiểm được phép cộng **trên chuỗi LLM đã ghi**; không tự đọc được số nào thuộc chiều nào (handoff mục 2). |
| Tool đếm số lần mã trên lớp text ra KL | **Không dùng làm KL.** NORTON: DH114 có trên 57 trang, KL thật 646 cánh (lấy từ bảng thống kê); trang lặp nguyên văn, legend, mặt bằng điển hình × số căn đều làm đếm sai. Chỉ dùng làm tín hiệu kiểm chéo. |
| SQLite làm lớp đọc | **Chưa cần.** Lợi ích thật là "lọc đúng dòng cần" — làm được bằng một lệnh lọc trên JSON đã ghép (`pd.py` đã có `merge` cho build/check). JOIN ở bước 5 sẽ nằm trong `pd.py doichieu` (Python). SQLite thêm một bản sao dữ liệu có thể cũ so với JSON. Chuyện khóa ghi không liên quan vì chỉ đọc. |

**Việc, xếp theo lợi / rủi ro:**
| # | Việc | Lợi | Ghi chú |
|---|---|---|---|
| B1 | `pd.py doichieu`: tách R×S×C, so BoQ↔BV, tổng KL BOQ_CDT↔BOQ_MAP theo khu, mã hoàn thiện BV↔ML; chấm Độ tin cậy đúng mục 4; sinh issue `D…` nháp | ~3170 ô; luật đã có, ít mơ hồ | LLM chỉ viết `Giả định tính giá`, gom RFI, ASK_USER |
| B2 | `pd.py boq-tach`: nhận dòng tiêu đề (Description/Unit/Qty), sinh nháp BOQ_CDT (sheet, dòng, mô tả nguyên văn, ĐV, KL, link) + issue B (KL chữ, cùng mã khác mô tả, lệch subtotal) | ~6384 ô; bỏ script tự chế | LLM đặt Khu, Mã khi BoQ không có mã, Hạng mục tiếng Việt |
| B3 | Phân loại scope bằng luật từ khóa (từ `2-boq.md` mục 4b) | giảm câu LLM phải xét | chỉ gán khi khớp chắc; "lẫn vật tư xưởng" vẫn ra CHƯA RÕ cho LLM/user |
| B4 | `pd.py loc <sheet> <cột>=<giá trị>`: lọc dòng trên JSON đã ghép (vd mã MÂU THUẪN) | giảm token đọc file | thay cho SQLite |
| B5 | Bước 4: `check` tự kiểm phép cộng chuỗi kích thước + so số Claude/Gemini đã ghi | chặn lỗi số học | không thay việc đọc hình |
| B6 | Đo tỉ lệ ô do tool ghi / tổng ô | biết tool gánh bao nhiêu | cần quy ước file nháp tool (vd `*.boq-tool.json`) và thứ tự ưu tiên khi LLM sửa nháp (LLM thắng tool, user thắng tất cả) |

**Thứ tự đề xuất:** Phần A (P0–P5, tìm trang) trước, vì đã có bộ đo và đáp án đã kiểm. Phần B làm sau, bắt đầu B1 → B2 → B4; mỗi lệnh có bài kiểm: chạy trên `_data` Blanca, so với kết quả LLM đã ghi 27/10 (đếm dòng/ô khớp, liệt kê khác biệt cho user duyệt theo ngoại lệ).

## 2. Không làm trong plan này
Đọc callout / số tờ từ hình (trần 64% ở SOFITEL), cache Gemini `--quansat`, `boq-extract`, `doichieu-prep`, lưu tọa độ chữ để sửa `title_guess`, embedding/RAG.

## 3. An toàn
- Sao lưu `pd.py` → `tool\_backup\pd_2026-10-27.py` trước khi sửa. Không đụng `*.user.json`, `_snapshot.json`, `_data` của dự án khi đo.
- Không lấy số liệu dự án nào sang dự án khác: SOFITEL/NORTON chỉ là bộ đo.

## 4. Câu hỏi cần user chốt
- **Q1.** Dữ liệu phụ (khu→file, mã cha, mã BV, loại tờ, trang trùng) để trong `_data\_timkiem*.json` (không hiện trong PROJECT_DATA, không đổi form) — hay thêm cột vào form (PAGES `Loại tờ`, `Trùng với`; BOQ_TONGHOP `Mã cha`, `Khu BV`)? Đề xuất: JSON riêng trước, khi chạy ổn mới đưa lên form.
- **Q2.** SOFITEL đã về `_ARCHIVE` (bị chặn đọc). Cho phép mở 1 lần để lưu lớp text + đáp án vào `DOCS\scan-toi-uu\_cache\`, hay bỏ SOFITEL khỏi bộ đo?
- **Q3.** NORTON: sau khi lưu cache bộ đo thì chuyển lại `_ARCHIVE`?
- **Q4.** Phụ kiện/hạng mục con: chỉ dùng trang của mã cha để **tìm** (bước 4 vẫn đọc riêng từng mã), đúng ý bạn chưa?
- **Q5.** Phần B: làm sau Phần A (đề xuất), hay làm B1 `doichieu` song song?

## 5. Kết quả thực hiện Phần A (28/10) — user bảo "thực hiện plan này"; Q1 = JSON riêng, Q4 = đúng, Q5 = A trước (mặc định đề xuất)
Đã làm: P0 `DOCS\scan-toi-uuench.py` + `_cache\Blanca\_data` · P1 (trang trùng bỏ ngày, file tổng ≥85%, loại tờ → `_timkiem.json`; kiemke in `L1 TRANG TRÙNG/FILE TỔNG`) · P2 (`_timkiem.kiemke.json` du_phong, `_timkiem.boq.json` khu_file/ma_cha/tu_khoa; `check` kiểm) · P3 (`pd.py ungvien` viết lại, chia nhóm, `--ra`) · P4 (SKILL.md, 1-kiem-ke, 2-boq 4c, 4-ban-ve, JSON.md) · P5 (chạy `--ra` ra scratchpad, `_data` Blanca không bị ghi). Sao lưu: `tool\_backup\pd_2026-10-27.py`.
Bench Blanca (F004 giả lập DỰ PHÒNG): 10 mã kiểm @12 73%→93%, @5 27%→75%; 69 mã @12 62%→81%; SAI đã biết lọt top-5: 10→4. 69 mã tụt so với 87% khi chưa loại F004 vì 14 mã chưa kiểm có đáp án F004 p28/p41 (cùng kiểu user đã chấm SAI) — chưa xác nhận.
KHÔNG làm: SOFITEL/NORTON làm bộ đo (Q2/Q3 chưa trả lời, `_ARCHIVE` bị chặn) → mới đo Blanca; bảng mã BV từ legend (P1.3, chưa test, cần prototype); Phần B (B1–B6). SAI còn lại: tờ phù điêu/mặt đứng không có gối (LO03), "chi tiết cửa P. NGỦ" cho gối.

### 5b. Đo SOFITEL + NORTON (28/10, user yêu cầu) — `DOCS\scan-toi-uuench_ext.py`, `old_base.py`; text thô ở `_cache
aw\`
Cùng bộ đo, so logic cũ (backup) và mới; `Hạng mục` = Mô tả[:120] với SOFITEL/NORTON vì PROJECT_DATA cũ không có cột Hạng mục tiếng Việt, nên con số tuyệt đối khác mốc 44% cũ, chỉ so tương đối.
| Bộ | Cũ @5/@12/@40/hết | Mới @5/@12/@40/hết (đầu ra) |
|---|---|---|
| SOFITEL 155 mã/199 trang | 21/27/33/45% | 34/40/49/60% (38%) |
| NORTON D01 cửa 5 mã/23 trang | 0/17/26/87% | 35/43/**100/100%** (100%); 5/5 mã có ≥1 trang đúng @12 |
| NORTON vanity 3 mã/29 trang (đáp án chỉ 1 tờ đại diện/loại căn → số này kém tin) | 7/10/14/14% | 0/0/0/62% nếu không có gì; với `ma_bv` (giả lập LLM gán VC1…): tờ chi tiết phòng tắm trong top-12 = V01 11/12, V02 11/12, V03 0/12 |
Sửa phát sinh khi đo: (1) lọc khu đổi từ loại hẳn sang hạ điểm ×0.4 — file chung kiểu GENERAL của NORTON mới còn; khóa khu ghép "BLK 01" → blk01; (2) sửa lỗi tách đường dẫn (file Blanca không lọc khu được); (3) mã trên >10% số trang (ô lưới F1 của SOFITEL) bị bỏ khỏi tín hiệu mã; (4) thêm `ma_bv` vào `_timkiem.boq.json` (mã BV của món mã tự đặt, chữ thường như GUONG thêm vào từ khóa). Blanca sau các sửa: 10 mã kiểm @12 93%, 69 mã 78%.
Giới hạn còn lại: SOFITEL trần 60% vì 40% trang đáp án chỉ là tờ vẽ không có chữ khớp (chi tiết `Typical Details` chỉ có khung tên) — cần đọc hình, ngoài phạm vi. Vanity không có `ma_bv`/`tu_khoa` thì 0 vì Hạng mục tiếng Việt, BV tiếng Anh; `ma_bv` do bước boq ghi là điều kiện.

## 6. Kết quả thực hiện Phần B (28/10) — user bảo "thực hiện tiếp phần B"
Đã làm trong `pd.py` (sao lưu trước: `tool\_backup\pd_2026-10-28_truocB.py`): **B1** `doichieu`, **B2** `boq-tach`, **B4** `loc`, **B5** kiểm chuỗi `a+b=c` trong `check` (LỖI nếu tổng sai; bỏ qua `×`, tổng chạy `+5+5=`), **B6** `do` + quy ước nháp `*-tool.json` (merge: tool < LLM < user theo từng ô; ID issue `BT…`/`DT…`). Hướng dẫn sửa: SKILL.md, `buoc/2-boq.md` (mục 0), `buoc/5-doi-chieu.md`, `JSON.md`.
Đo trên `_data` Blanca (ghi ra scratchpad, không ghi đè):
| Lệnh | Kết quả so với bản LLM 27/10 |
|---|---|
| `boq-tach` | 393/393 dòng BOQ_CDT trùng; Mô tả, ĐV, KL, Mô tả kỹ thuật, Item, Bill, Link giống 100% (làm tròn KL 6 số; tự bỏ dòng nhóm "Phụ kiện cửa"). Khu, Mã vẫn LLM. |
| `doichieu` BOQ_MAP `Trạng thái` | 373/373 trùng (46 KHỚP, 6 LỆCH, 321 CHƯA ĐỌC) |
| `doichieu` `Độ tin cậy` mã | 135/148 trùng (91%); 13 mã lệch: tool lạc quan hơn LLM 8 mã (4 TB→CAO như LO01, LO02, FR15 vì issue S09 ghi Mã mơ hồ; 4 THẤP→TB như VK1, PK01 vì LLM tính cả khu chưa có BV) và thận trọng hơn 5 mã (TB→THẤP) — vì thế LLM vẫn phải rà theo ngoại lệ, không coi nháp là kết quả |
| `doichieu` cột `Issue` | 80 trùng, 67 tool thiếu (issue ghi "Nhiều mã"/"Tất cả"/D… LLM tự gắn), 0 thừa |
| `check` | không thêm LỖI mới (100 LỖI/355 XEM giống bản trước B — vì Blanca chưa gắn Phạm vi nhà máy); 0 chuỗi cộng sai thật |
Thử rồi bỏ: "mã có issue mở CAO/TB → tối đa TB" (trùng giảm 91%→86%, LLM vẫn chấm CAO khi đã ghi giả định); khớp issue theo Nội dung (thừa 33).
Chưa làm: **B3** phân loại scope bằng luật từ khóa (cần user duyệt danh sách luật; hiện LLM vẫn phân loại ở bước boq 4b); `Mã hoàn thiện` BV↔ML (QUAN_SAT viết tự do, chưa đủ khuôn để máy so); so số Gemini↔Claude trong B5 (chỉ làm phần cộng chuỗi); chưa chạy `build` thật với file `-tool` (mới thử `merge`/`check`/`do` trên bản sao); P1.3 bảng mã BV vẫn chưa.

## 7. Kết quả B3 (28/10) — `pd.py scope`
Sao lưu trước: `tool\_backup\pd_2026-10-28_truocB3.py`. Lệnh `scope <goc> <dot> [--ra DIR]` → `QUAN_SAT.boq-tool.json` (`Phạm vi nhà máy`, Nguồn BOQ, Model `pd.py scope`). Luật theo thứ tự: phụ kiện cửa → picomat → hộp màn có ván ép → len chân/chỉ (MDF/ván = trong; nhựa = ngoài) → gạch → nhóm ngoài scope (từ khóa đầu món; nhóm có thể lẫn vật tư thì ra `CHƯA RÕ – … nhưng có …`) → trong scope (cửa/vách kính, gương = mộc, tủ/vanity/hộp/băng = cabinet, gối/sofa/ghế/giường/bàn/kệ = loose). Không khớp → không ghi, để LLM. `merge`: dòng `Phạm vi nhà máy` của tool nhường dòng LLM/user cùng mã (QUAN_SAT là NOI nên cần).
Đo trên `_data` Blanca (ghi ra scratchpad): 148 mã → 102 TRONG, 46 KHÔNG THUỘC, 0 CHƯA RÕ, 0 không khớp. Blanca chưa có bản phân loại LLM nên KHÔNG có độ trùng; tôi rà tay cả 148 dòng, sửa 5 lỗi lúc đầu (chỉ gỗ vân gỗ, nẹp inox che gạch, băng để vali, ghế bench). Cần user rà các điểm: FO02/FO11 "len/chỉ gỗ" thực ra nhựa → KHÔNG; FO17/46/47 hộp màn → TRONG vì có ván ép chống ẩm (luật user 27/10); FO25 picomat → mộc; nhãn xưởng của bàn console/ghế bench (cabinet) và bàn ăn/kệ TV (loose) là đoán. SOFITEL/NORTON chưa đo (vẫn bị chặn).
Còn lại: so với phân loại LLM khi có dự án khác; chạy `build` thật với file `-tool`.
