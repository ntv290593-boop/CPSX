# Bước 3 — spec: spec + Material List + bảng hoàn thiện

**Ghi:** `SPEC`, `BOQ_TONGHOP` (`Spec · ML`, `Thông số riêng`), `QUAN_SAT` (Nguồn `SPEC`), `ISSUES` (ID `S01`…), `NOTES` — file `*.spec.json`.
**Không làm:** đọc bản vẽ để đo; chọn bên đúng khi spec lệch BoQ (bước 5 làm). Chỉ ghi nhận.

Danh sách mã: đọc `BOQ_TONGHOP.boq.json`, **bỏ mã có QUAN_SAT `Phạm vi nhà máy` = `KHÔNG THUỘC SCOPE NHÀ MÁY…`** (sơn nước, thạch cao, sàn/gạch/đá, đèn… — không tra spec cho các mã này, kể cả mục spec sơn tường/gạch/trần). Bước 2 chưa xong thì làm phần 1 trước, phần 2 làm sau khi có danh sách mã.

## 1. Spec (MasterFormat PDF)
- Tìm mục liên quan gói: `py -I "<TOOL>" tim <GOC> <DOT> "<từ khóa>" --file "<tên file spec>"`.
- Tách mục: `pdftotext -layout "<spec.pdf>" "<scratchpad>\spec.txt"`, rồi `py -I "D:\COSTING EXCEL\_MASTER\scan\scan_spec.py" "<scratchpad>\spec.txt"` → danh sách mục (ref, title, text, page).
- **SPEC**: chỉ các mục ảnh hưởng giá hoặc cấu tạo (vật liệu, độ dày, hoàn thiện, phụ kiện, thử nghiệm, bảo hành, mẫu/mock-up). Mỗi mục 1 dòng: `Mục`, `Nội dung` (nguyên văn), `Nhãn`, `Mã liên quan`, `File · Trang`.

## 2. Material List / bảng hoàn thiện / legend
- Với mỗi mã: tìm mã vật liệu (WD-, LAM-, PT-, SS-, MT-…) và dòng ML tương ứng; phụ kiện thì Ref ML, hãng, model.
- Mã hoàn thiện có khu áp dụng (vd chỉ một loại phòng) → ghi khu vào QUAN_SAT.
- **BOQ_TONGHOP**: `Spec · ML` = link mục spec chính + dòng ML (nhiều link ▾). Không có → `"KHÔNG CÓ – đã tìm trong <file/mục lục>"`. `Thông số riêng` = EI, Ref ML, hãng…
- **QUAN_SAT** (`Nguồn = SPEC`): các trường `Vật liệu gỗ (mã WD/PT)`, `Mặt đá`, `Gương`, `Kính`, `Kim loại`, `Giấy/vải`, `Phần ngoài phạm vi` theo spec/ML, kèm `Tờ · Chi tiết` = link trang spec/ML.

## 2b. Trang dạng ảnh (bảng hoàn thiện, ML scan, legend trên bản vẽ)
Trang `Lớp text = KHÔNG/ÍT` hoặc bảng nằm trong hình: tự xem ảnh và đọc trước, rồi cho Gemini đọc độc lập cùng ảnh (`gemini "<ảnh>" --hoi "Transcribe this finish/material table as JSON rows {code, description, brand, area}. Mark unreadable cells as ?"`). So từng mã: khớp → ghi; lệch → xem lại chỗ đó; vẫn lệch → ghi cả hai vào QUAN_SAT (`SPEC` và `GEMINI`, trạng thái `MÂU THUẪN`).

## 3. ISSUES (S…)
Spec ↔ BoQ, spec ↔ ML, ML ↔ BoQ khác nhau (độ dày, vật liệu, hãng, cấp chống cháy…). `Nội dung` nêu đủ hai giá trị + mục; `Nguồn` link cả hai chỗ; `Mức` CAO nếu đổi giá rõ.

## Xong khi
Mỗi mã trong BOQ_TONGHOP có `Spec · ML` (link hoặc "KHÔNG CÓ – …"); `check --den spec` không còn LỖI.
