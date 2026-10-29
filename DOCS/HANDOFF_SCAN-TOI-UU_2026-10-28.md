# HANDOFF — Tối ưu skill 01-scan (cập nhật cuối 2026-10-28)

Cho session sau **lập plan rồi sửa** `SKILL\01-scan`. Đến nay **chưa sửa file nào** của skill hay dự án; chỉ đo trong scratchpad / `DOCS\scan-toi-uu\`.

## 0. Trạng thái + việc tiếp theo
**Cập nhật 28/10 tối:** user đã kiểm xong file kiểm mẫu → đáp án Blanca sửa (164→184 trang, `eval_v2.py`), mốc mới + plan ở `DOCS\PLAN_SCAN-TOI-UU_2026-10-28.md` — **chờ user duyệt plan và trả lời Q1–Q5** (plan có Phần A tìm trang + Phần B chuyển việc sang tool), chưa sửa code.

~~**Đang chờ user:**~~ (đã xong) điền cột J (ĐÚNG/SAI/KHÔNG CHẮC) trong `DOCS\scan-toi-uu\KIEM_MAU_DAP_AN_Blanca_v2.xlsx` (10 mã, 57 dòng, 40 dòng tô vàng = đáp án và máy bất đồng). Bản `KIEM_MAU_DAP_AN_Blanca.xlsx` (v1) là bản cũ link trực tiếp — bỏ.

**Session sau làm theo thứ tự:**
1. Kiểm `SKILL\01-scan\tool\pd.py` chưa đổi (lần sửa cuối 27/10 16:54) và hỏi user đã điền file kiểm mẫu chưa.
2. Nếu đã điền: đọc kết quả → sửa đáp án Blanca (thêm trang đúng bị thiếu, bỏ trang sai) → chạy lại `proto_rank.py` lấy mốc chính thức.
3. Lập plan theo hướng ở mục 4, trình user duyệt, rồi mới sửa `pd.py`.
4. Khi sửa: sao lưu `pd.py`; không đụng `*.user.json` / `_snapshot.json`; không ghi vào `_data` của dự án khi đo.

**Lưu ý môi trường:**
- 28/10 chiều: `_ACTIVE` đang có `NORTON` (user chuyển ra để test, đã đo xong — mục 3b), không còn `SOFITEL`. Hỏi user có chuyển NORTON lại `_ARCHIVE` không.
- `_ACTIVE\SOFITEL` được user chuyển từ `_ARCHIVE` sang chỉ để test (28/10). `ACTIVE.md` vẫn ghi Blanca là dự án đang làm. SOFITEL vẫn là dự án đã đóng: chỉ dùng làm bộ đo, không lấy số liệu cho dự án khác. Hỏi user có chuyển lại `_ARCHIVE` không.
- Chạy Python với `-X utf8` (console cp1252). `py -I` bỏ qua `PYTHONIOENCODING`.

## 1. Thuật toán scan hiện tại (đã kiểm trong code)
Lọc nhiều lớp từ rẻ đến đắt; không có BM25 / embedding / RAG.
| Lớp | Kỹ thuật |
|---|---|
| L0 bỏ file nhiễu | regex tên/thư mục + `_bo_qua.txt` |
| L1 trùng | md5 trong nhóm cùng dung lượng; gom tên bỏ ngày/REV — **chỉ báo, không loại khỏi tìm kiếm** |
| Lớp text | PyMuPDF `get_text` → `_text\F*.json` (không tọa độ) |
| Số/tên tờ | `title_guess` heuristic (góc phải dưới, gần nhãn DRAWING NO) |
| Ứng viên trang | `cmd_ungvien`: mã regex (bỏ dấu, KHÔNG bỏ `.`/`-`/hậu tố) hoặc ≥2 từ khóa `Hạng mục`; điểm = mã×5 + số từ khóa; không lọc khu |
| Đọc hình | Claude vision + Gemini (agy) đọc độc lập |
| Điều phối | subagent mỗi bước; bước 4 nhóm ≤8 mã song song; JSON → `pd.py build` |

## 2. Số đo đã verify (Blanca, lần chạy 27/10)
Dữ liệu: `_ACTIVE\Blanca_City\01_scan\_data\261027_DV2.3-DV2.4`.
- 148 mã; THAP/TB/CAO = 124/19/5. Trong 124 THAP: ~57 do 2BR thiếu PDF, ~50–53 do chưa quyết phạm vi (ASK Q01), chỉ ~14–17 do đọc bản vẽ.
- QUAN_SAT: 532 dòng KHÔNG THẤY (Tờ chưa đọc 60, Kích thước BV 50, Mã hoàn thiện 50, Kiểm chéo Gemini 42, Có cánh 36, Ngăn kéo 35).
- Lớp text so với Gemini: trong các dòng "Kiểm chéo Gemini khớp", 93–96% số kích thước có sẵn trong lớp text. Nhưng lớp text không cho biết số thuộc chiều nào.
- Gemini `--quansat` / `--hoi` không cache (chỉ chế độ đọc số mới cache).
- **Đính chính bản phân tích đầu:** BoQ không do LLM gõ tay (subagent viết `make.py` đọc xlsx); lỗi LAM-01 → LA.1 là của costing, không phải scan.

## 3. Chạy thử xếp hạng trang ứng viên
Script trong `DOCS\scan-toi-uu\` (chỉ đọc, chạy `py -X utf8 <file>` từ thư mục đó):
- `eval_ungvien.py`: đọc đáp án Blanca + chép logic `ungvien` hiện tại.
- `proto_rank.py`: các biến thể trên Blanca (mã chuẩn hóa, lọc khu, BM25, tên tờ).
- `make_kiem_mau.py`: tạo file kiểm mẫu; link theo chuẩn PROJECT_DATA (ô trỏ về chính nó + `_LINKS` → add-in SRC_LINKS mở Edge `#page=N`).
- `sofitel_cache.py <thư mục ra>`: trích lớp text 19 PDF SOFITEL (78 giây). Cache đang ở scratchpad session cũ, có thể đã bị xóa → chạy lại nếu cần.
- `sofitel_rank.py <thư mục cache>`: các biến thể trên SOFITEL.

**Đáp án:**
- Blanca: "Bản vẽ đối chứng" của bước banve, 69 mã / 164 trang. Do subagent chọn, **chưa kiểm**. 115/164 trang có chứa mã → có thể thiên về cách tìm theo mã. Có thiếu trang (vd KT2a có 503.1 F005 p15 nhưng thiếu 503.3 F005 p17).
- SOFITEL: link cột P "Bản vẽ chi tiết" của BOQ_TONGHOP trong `_ACTIVE\SOFITEL\01_scan\PROJECT_DATA_SOFITEL_v8.xlsx` (bỏ link sang @SPEC), 155 mã / 199 trang. Quy trình cũ dựng từ callout/DETAILS → khách quan hơn.

**Kết quả: % trang đúng nằm trong top-12** (Blanca đã gộp file trùng F023≡F024)
| Biến thể | Blanca @5 | Blanca @12 | SOFITEL @5 | SOFITEL @12 | SOFITEL trong ứng viên |
|---|---|---|---|---|---|
| Hiện tại | 31% | 62% (55% nếu không gộp file trùng) | 40% | 44% | 49% |
| + mã chuẩn hóa (bỏ `_ST`/`_1BR`/`-HM`/`a`, `.`/`-`, số 0 đầu) | 54% | 84% | 43% | 49% | 54% |
| + lọc khu (khu của file = thư mục sâu nhất) | 71% | 88% | không áp được | | |
| + BM25 + tên tờ | 70% | 88% | 39% | 48–51% | 64% |
| ghim tờ mặt bằng khu (Blanca) | 51% | 88% | | | |

**Rút ra:**
- Không có một cách xếp hạng thắng cả hai dự án.
- Blanca: BoQ có mã thật trên bản vẽ → chuẩn hóa mã + lọc khu + bỏ file trùng là đòn bẩy. BM25 không thêm gì (Mô tả BoQ tiếng Việt, ý "song ngữ" không áp dụng).
- SOFITEL: mã là số dòng tự đặt (F001…), 13/170 "khớp" đều nhầm với ô lưới/cao độ → chuẩn hóa mã vô dụng. BM25 có ích vì mô tả và bản vẽ cùng tiếng Anh. Lọc khu không tự làm được (Vị trí tiếng Việt "Hành lang" ≠ file "Area 50 - Guestroom Corridor").
- Trần của mọi cách dựa lớp text ở SOFITEL là 64%: phần trượt dồn vào 14 trang, chủ yếu tờ FF&E plan (…-201) từng loại phòng và Typical Details. Mã đồ rời/callout nằm trên hình; lớp text chỉ có khung tên.
- Blanca còn trượt: phụ kiện cửa PK01–08 (bản vẽ ghi mã NCC → cần gắn PK vào cửa cha), tờ mặt bằng, F004 p41 (file concept ít chữ).

### 3b. NORTON (bổ sung 28/10, dự án đã đóng, `_ACTIVE\NORTON` chỉ làm bộ đo)
Script: `norton_cache.py <thư mục ra> d01|vanity` (vanity 80 giây, d01 **17 phút** vì file tổng 523 MB) → `norton_rank.py <cache d01> <cache vanity>`. Chấm theo **lớp trang trùng nguyên văn** (trúng bản trùng nào cũng tính).

**D01 cửa** — đáp án: `PROJECT_DATA_v8.xlsx` BOQ_TONGHOP cột X (Thống kê BV) + Y (Chi tiết cửa) của 5 dòng Cửa, 23 trang (link sang `_sheets\…\pNNN_*.pdf` quy về file gốc). 7 PDF / 1534 trang, **file tổng `SB_TENDER_PLOT B_2026 10 08.pdf` (803 tr) trùng nguyên văn 731 trang với 6 file con**. Mã cửa có thật trên BV (DH114 trên 57 trang; ID-DH111 chuẩn hóa thành dh111 thì 77 trang, nguyên văn 51).
| Biến thể | @5 | @12 | @40 | mã ≥1 đúng @12 |
|---|---|---|---|---|
| A hiện tại | 13% | 22% | 35% | 2/5 |
| A + bỏ trang trùng | 13% | 26% | 78% | 2/5 |
| C/D mã chuẩn hóa (± lọc khối) | 13% | 26% | 74% | 2/5 |
| G BM25 + tên tờ + mã + lọc khối | 39% | 43% | 57% | 5/5 |
| G2 = G + bỏ file tổng | 43% | 43% | 87% | 5/5 |
| **J = G2 + ghim tờ thống kê cửa có chứa mã** | **87%** | **100%** | 100% | 5/5 |
| J' = mã chuẩn hóa + bỏ file tổng + ghim tờ thống kê (không BM25) | 57% | 65% | 78% | 5/5 |

- Tờ chi tiết cửa (GEN p54/55) BM25 tìm được ngay (Mô tả BoQ tiếng Anh dài, cùng ngôn ngữ BV). Phần trượt duy nhất là **tờ thống kê cửa** (hạng 27–56): bảng nhiều mã, mỗi mã xuất hiện 1 lần, thua mặt bằng nhắc mã nhiều lần.
- Lọc khối vô hiệu khi còn file tổng: trang trùng được quy về bản trong file tổng (không ghi khối) nên lọt qua bộ lọc. **L1 phải gộp trùng ở mức trang và ưu tiên bản trong file cụ thể**, rồi mới lọc khu.
- Chuẩn hóa mã chỉ giúp ID-DH111/112 (thêm ~26 trang), không đổi điểm. Phụ kiện (25 dòng) không có link BV trong đáp án → không đo được hướng #3.
- Đáp án do quy trình cũ dựng từ tờ thống kê + chi tiết → biến thể J hợp với cách dựng đáp án, có thể đánh giá cao hơn thực tế. Chỉ 5 mã.

**Vanity** — đáp án: `PROJECT_DATA_VANITY_v1.xlsx` BOQ_TONGHOP Q + BOQ_MAP G/H, 3 mã V01–V03 (tự đặt, không có trên BV) / 29 trang. 2 PDF / 334 trang (286 lớp). **Đáp án chỉ lấy 1 tờ đại diện cho mỗi loại căn** (Studio/1BR/2BR/3BR), trong khi BV có 77 tờ chi tiết phòng tắm, 21 tờ có VC1 → @k bị đánh giá thấp; đo thêm "số tờ chi tiết phòng tắm trong top-12".
| Biến thể | @12 | @40 | tờ phòng tắm trong top-12 (V01/V02/V03) |
|---|---|---|---|
| A hiện tại (Hạng mục tiếng Việt) | 10% | 14% | 5/0/0 (V03 "Gương MCB" 0 ứng viên: MCB không có trên BV) |
| E BM25 Hạng mục + Mô tả | 0% | 0% | 0/0/0 |
| E3 truy vấn ngắn + bỏ tờ danh mục/legend | 0% | 34% | 2/0/0 |
| H BM25 + mã BV VC1 (giả định bước boq gán) | 7% | 34% | **12/12/0** |
- BM25 thất bại: từ của Mô tả (laminate, finish, frame, mirror) rơi vào tờ danh mục BV, legend hoàn thiện, mặt bằng trần/tường, những tờ lặp lại ở mọi loại căn.
- Tín hiệu duy nhất hiệu quả là **mã của chính BV (VC1)**, không có trong BoQ → cần bước boq/LLM ghi bảng "mã BoQ → mã/nhãn trên BV" (đọc legend). Gương (V03) không có mã riêng trên BV ("GƯƠNG" có trên 52 trang) → phần này phải đọc hình.
- Hồ sơ căn hộ điển hình: kết quả đúng là **một nhóm tờ theo loại căn**, không phải top-k; nên gộp tờ gần trùng (cùng tên tờ khác loại căn) và chọn đại diện theo loại căn.

### 3c. Gộp ba dự án
| Tín hiệu | Blanca | SOFITEL | NORTON D01 | NORTON vanity |
|---|---|---|---|---|
| Mã BoQ có trên BV | có | không (F0xx tự đặt) | có | không (V0x tự đặt) |
| Mô tả BoQ cùng ngôn ngữ BV | không (Việt) | có (Anh) | có (Anh) | có nhưng chung chung |
| Mã chuẩn hóa | +22 điểm @12 | vô dụng | ~0 | vô dụng |
| BM25 | 0 | +2–7 | +17 @12 | âm |
| Bỏ trùng (file/trang) | +7 | — | +43 @40 | cần gộp gần trùng |
| Lọc khu | +17 @5 | không áp được | chỉ có tác dụng sau khi bỏ file tổng | không áp (khu = khối, không phải loại căn) |
| Ghim tờ theo loại (mặt bằng / thống kê) | @5 tụt | — | **+57 @12** | — |
| Mã BV do LLM gán | — | — | — | 0→12/12 tờ phòng tắm |
- Củng cố hướng #1, #2: không có tín hiệu thắng mọi nơi → cần kết hợp và bật/tắt theo điều kiện (mã có trên BV? cùng ngôn ngữ?).
- Bổ sung: **#2 mở rộng thành "bảng loại tờ"** (mặt bằng, thống kê cửa, FF&E, chi tiết điển hình) do bước boq/kiemke ghi. Với cửa, ghim tờ thống kê có chứa mã là đòn bẩy lớn nhất (ngược với Blanca, nơi ghim mặt bằng làm @5 tụt → ghim theo loại hạng mục, không ghim đại trà).
- Bổ sung: **L1 gộp trùng ở mức trang**, ưu tiên bản trong file cụ thể (không chỉ báo).
- Bổ sung: **bảng mã BoQ → mã BV** (VC1…) do LLM điền cho dự án đặt mã tự do (SOFITEL, vanity).

## 4. Hướng cho plan (chờ user duyệt)
| # | Việc | Căn cứ |
|---|---|---|
| 1 | `ungvien` kết hợp tín hiệu: mã chuẩn hóa (chỉ khi mã BoQ có thật trên BV) + BM25 (khi mô tả cùng ngôn ngữ BV) + bỏ trùng L1 **ở mức trang, ưu tiên file cụ thể** | Blanca @12 62→88%; SOFITEL 44→51%; NORTON D01 @40 35→87% |
| 2 | Bảng **khu → file / tờ mặt bằng, FF&E** và **loại tờ** (thống kê cửa, chi tiết điển hình) do bước boq (LLM) điền, `ungvien` dùng để lọc khu và kèm tờ theo loại hạng mục (cửa → tờ thống kê có mã: ghim; tờ mặt bằng: liệt kê riêng, không ghim vào top vì @5 tụt) | lọc khu Blanca @5 54→71%; SOFITEL trượt chủ yếu ở FF&E plan; NORTON D01 @12 43→100% |
| 2b | Bảng **mã BoQ → mã/nhãn trên BV** (VC1…) do LLM điền khi BoQ đặt mã tự do | NORTON vanity: tờ phòng tắm trong top-12 0→12/12 |
| 3 | Gắn hạng mục con / phụ kiện vào hạng mục cha (PK → cửa; `-DT`/`-HM` → tủ) để thừa hưởng trang của cha | PK01–08 trượt ở Blanca |
| 4 | Callout / số tờ đọc từ hình cho phần lớp text không tới | trần 64% ở SOFITEL |
| 5 | Sau đó mới tới: `pd.py boq-extract`; `doichieu-prep` (script chấm độ tin cậy); cache Gemini `--quansat` + kiểm số bằng `so` trên trang có text; kiemke lưu từ kèm tọa độ để sửa `title_guess` | chưa test |

Không làm embedding/RAG: hồ sơ ~300 trang; khớp chủ yếu là mã, số, tên tờ.
