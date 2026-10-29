# HANDOFF — Lệnh `neo` cho 01-scan + rà quy trình 01-scan (29/10/2026)

Cho session sau làm tiếp. Đọc file này trước, rồi `DOCS\PLAN_NEO-ANCHOR_2026-10-29.md` (mục 3b, 3c). Chưa sửa skill, `pd.py`, `scan_core.py` hay form.

## 0. Quyền và phạm vi (bắt buộc giữ)
- `ACTIVE.md`: dự án active là SOFITEL tender (`_ACTIVE\SOFITEL\10_TENDER`).
- **User cho phép (29/10):** dùng `_ACTIVE\SOFITEL\01_scan` (PROJECT_DATA_SOFITEL_v1–v8, `_dwg_tangA`, `_sheets`, `_sheetmap.json`) và PDF gốc `01. INPUT\@DRAW` **chỉ để đo thuật toán**. Không lấy số liệu sang tender. Không mở `_ARCHIVE`, `_TEST`, NORTON, Blanca.
- **User chốt (29/10):** tập trung `neo` cho **01-scan**. 10-tender chỉ dùng làm nguồn dữ liệu thử phụ, không làm lệnh `td.py neo`.

## 1. Việc đã làm trong session này
1. **Rà câu trả lời về JEV** (TypeSafe AI, mô hình "System One"):
   - JEV có thật: early access 15/9/2026, kiểu đầu ra Choice/Score/Noul, hiệu chỉnh xác suất bằng RLCD.
   - Câu trả lời kia sai ở: tốc độ 15–20ms (hãng công bố 70–500ms) và "chuẩn xác tuyệt đối" (hãng đo 67,8%, thấp hơn Claude Opus 5).
   - Đề xuất ④ chấm độ tin cậy và ⑤ guardrail JSON là sai: `pd.py doichieu` và `check` đã làm bằng luật cố định.
   - Kết luận: chưa dùng JEV. Lý do: mới early access, dữ liệu thầu phải gửi ra ngoài, và lợi ích chỉ ở ③ loại tờ / ① scope cho mã tool không khớp.
2. **Rà 01-scan, tìm việc LLM đang làm thay tool** (chưa làm, để user chọn):
   - Bước 2: kiểm subtotal; BOQ_MAP; `Dòng BoQ gốc`; `Kích thước BoQ` từ `_boq_tach`; "cùng mô tả = cùng mã"; issue B trùng/khác ĐV; từ khóa scope tiếng Anh.
   - Bước 1: CAD đã có PDF; REV/ngày.
   - Bước 3: chép nguyên văn spec theo ref; ghép mã VL ↔ ML; so bảng Claude ↔ Gemini.
   - Điều phối: chia nhóm bước 4 theo trang ứng viên chung.
   - Bước 4: tự ghi PAGES khi `anh`/`chu`/`so` mở trang; so `--quansat` theo trường.
   - Bước 5: Mã hoàn thiện BV↔ML; cột `Mã` dạng danh sách.
   - Đề xuất RAG cho tìm trang, theo thứ tự: bộ đáp án kiểm tay → BM25 + embedding đa ngữ chạy trên máy (bge-m3, RRF) → đồ thị callout/Drawing List → bảng legend → reranker.
3. **Tổng hợp `HANDOFF_NEO-ANCHOR_2026-10-29.md`** thành `PLAN_NEO-ANCHOR_2026-10-29.md`, sau đó sửa plan sang 01-scan (mục 3b).
4. **Đo N2** trên dữ liệu 01-scan (mục 3c của plan). Script đo:
   - `DOCS\scan-toi-uu\neo-test\neo01_d1d3.py`: PROJECT_DATA_SOFITEL_v8 + `_dwg_tangA`.
   - `DOCS\scan-toi-uu\neo-test\neo01_d4.py`: `SCOPE.banve-*.json` của tender.
   - Chạy: `py -I -X utf8 "<file>"`. Mỗi file chạy vài phút; nên chạy từ thư mục Temp.

## 2. Kết quả đo (đã kiểm bằng cách mở text trang, không chỉ đọc số)

| Bài | Kết quả | Kết luận |
|---|---|---|
| D1a mã đọc hình (v8 BOQ_TONGHOP cột H) trên trang tờ | 5 / **407 LỆCH** | Báo oan: lớp text bản vẽ GA Design chỉ có khung tên + ghi chú (`_sheets\A50\p071_ID-CO-TY-703.pdf`: 2.870 ký tự, không có `WD`); mã VL vẽ bằng nét |
| D3 số kích thước DWG có trong text trang | 420/5209 (8%), 0/218 tờ ≥ 50% | Kích thước không có trong text; 8% là trùng với ghi chú/cao độ |
| D4 Mã VL tender trên bản vẽ 2026 | 0 / 930 LỆCH | Bộ 2026 cùng kiểu. Vùng vẽ chỉ có số chi tiết + tỉ lệ (`01 10`), cao độ (`LEV 2500`, `FFL ±0`) |
| D1d số tờ của nhãn link (cột P) trên trang link | **199/199**; trùng ngẫu nhiên **0/299** | Dùng được |
| D1e tên chi tiết (cột Q) trên trang tờ | 13/13 | Dùng được, mẫu nhỏ |
| (Session trước) mã/model spec MAT tender | 91/91; đột biến 700/702 | Dùng được cho bảng spec/ML có text |

**Bẫy đã gặp:**
- Luật "trang có text → kiểm được" sai.
- Luật "trang có mã dạng XX-NN" cũng sai, vì số tờ `ID-SK-52-601` → `SK52`, và `FFL0`, `LEV850` đều khớp regex.
- Cắt khung tên theo `x < 0,8·W` không đủ: ghi chú chung (số điện thoại `7482 0054`, `2001`) nằm trong vùng đó.

## 3. Thiết kế đã sửa (chờ user chọn; xem plan mục 3c)
- **N3.1, làm được ngay:** số tờ của nhãn link (`Bản vẽ đối chứng`, link `Tờ · Chi tiết` của QUAN_SAT, `Bộ tờ đã đọc`) phải có trên khung tên trang link. LỆCH thì gợi ý trang có số tờ đó. Nhãn có `#nn` thì kiểm thêm số/tên chi tiết.
- **N3.2:** link spec/ML (`SPEC`, `Spec · ML`, QUAN_SAT `Nguồn = SPEC`): kiểm mã/model như MAT tender.
- **N3.3:** KHÔNG KIỂM ĐƯỢC theo họ.
  - Mã VL: chỉ kiểm khi text trang có mã thuộc họ mã dự án (lấy từ FINISHES/ML), không tính token nằm trong số tờ.
  - Kích thước: chỉ kiểm khi vùng vẽ có ≥ 10 số ≥ 100, sau khi bỏ dòng lặp ở ≥ 30% trang của file và bỏ cao độ/tỉ lệ. **Ngưỡng chưa đo**, vì không có dự án nào có kích thước trong text mà được phép dùng.
- **Để chọn sau:** nguồn neo kích thước cho bản vẽ vẽ nét: `_MASTER\scan\pdf_dim_decode.py` (font SOFITEL, đúng ~99%) hoặc DWG tầng A.
- Giữ nguyên các quyết định cũ:
  - Trạng thái **CÓ NEO / LỆCH / KHÔNG KIỂM ĐƯỢC**.
  - LỆCH chỉ là XEM, không phải LỖI.
  - Không sửa form, không đụng `doichieu` ở bản 1.
  - Code đặt trong `pd.py` (chuyển sang `scan_core.py` khi tender cần). Sửa `scan_core.py` thì chạy `DOCS\scan-toi-uu\bench.py` trước/sau.

## 3b. Cập nhật cuối ngày 29/10 — ĐÃ LÀM N3.1 + N3.2
User chọn N3.1 + N3.2. Đã viết `pd.py neo`, sửa `check`/`report`, cập nhật `buoc/4-ban-ve.md`, `SKILL.md` mục 2, `JSON.md`. Kết quả + luật: plan mục **3d**. Kiểm hồi quy: `DOCS\scan-toi-uu\neo-test\neo01_regress.py` (D1d 199/199; đột biến bắt 193/193 SOFITEL, 79% Blanca). Backup trước khi sửa: `SKILL\01-scan\tool\_backup\pd_2026-10-29_truoc-neo.py`. Chưa commit.

**Còn lại:** N5 — đo `neo` ở lần scan 01 thật kế tiếp (nhất là N3.2 trên ML/spec thật); N3.3 và nguồn neo kích thước (`pdf_dim_decode` / DWG) vẫn chờ user. Mục 4 dưới đây là kế hoạch cũ, bước 1–2 đã xong.

## 4. Việc tiếp theo cho session sau
1. **Hỏi user** (câu đang treo): làm N3.1 + N3.2 trước (`pd.py neo`, chỉ cảnh báo XEM), hay tính hướng kiểm kích thước bằng `pdf_dim_decode` / DWG trước?
2. Nếu làm N3.1 + N3.2:
   - Text trang lấy từ cache `kiemke` (`<GOC>\01_scan\_data\<DOT>\_text\F*.json`), không mở lại PDF.
   - Số tờ lấy từ nhãn link bằng regex số tờ của dự án (`ID-[A-Z0-9]{1,4}-[A-Z0-9]{1,4}-\d{3}` cho SOFITEL; dự án khác phải tổng quát hóa, vd dùng `TOK_SO` trong `scan_core.py`). So sánh sau khi bỏ ký tự không chữ-số.
   - Ghi `_data\<DOT>\_neo.json`. `check` in `XEM NEO LỆCH …` và 1 dòng tổng KHÔNG KIỂM ĐƯỢC.
   - Kiểm hồi quy:
     - Chạy lại `neo01_d1d3.py`, kỳ vọng D1d 199/199.
     - Làm thêm 1 bài đột biến link: đổi trang ±1 trong cùng file, kỳ vọng bắt gần hết.
   - Cập nhật `buoc/4-ban-ve.md` ("Trước khi trả kết quả"), `SKILL.md` mục 2 (điều phối gửi LỆCH cho subagent, tối đa 2 vòng), `JSON.md` (bảng file phụ trợ: `_neo.json`).
3. Việc còn treo từ handoff trước, ưu tiên thấp vì kích thước SOFITEL không có trong text:
   - Sửa ghép từ trong `neo_test7.py`.
   - Đo ngưỡng N3.3 trên một dự án có kích thước trong text. Cần user chỉ định và cho phép.

## 5. Tệp liên quan
- Plan: `DOCS\PLAN_NEO-ANCHOR_2026-10-29.md` (mục 1 tổng hợp, 3b plan 01-scan, 3c kết quả + thiết kế sửa).
- Handoff gốc: `DOCS\HANDOFF_NEO-ANCHOR_2026-10-29.md` (nghiên cứu repo GitHub, bài A–H trên tender).
- Script đo: `DOCS\scan-toi-uu\neo-test\` (`neo_test*.py` của session trước; `neo01_d1d3.py`, `neo01_d4.py` của session này).
- Dữ liệu đo: `_ACTIVE\SOFITEL\01_scan\PROJECT_DATA_SOFITEL_v8.xlsx`:
  - `_LINKS`: cột Key/Sheet/Ô/Nhãn/File/Trang.
  - BOQ_TONGHOP: H = Mã trên BV, P = Bản vẽ chi tiết, Q = Chi tiết.
  - PAGES: Số BV → file/trang.
- Memory: `neo-anchor-01scan-2026-10-29.md`.
