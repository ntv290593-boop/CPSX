# Bước 7c — spec: spec + bảng hoàn thiện cho từng cụm đọc

**Ghi:** `SPEC.spec.json`, `SCOPE.spec.json` (`Spec · ML`, `Thông số riêng`), `QUAN_SAT.spec.json` (Nguồn `SPEC`), `ISSUES.spec.json` (ID `S01`…), file `*.spec.json`.
**Không làm:** đọc bản vẽ để đo; chọn bên đúng khi spec lệch BOQ (bước 7e). Chỉ ghi nhận. Không sửa `MAT` (bước 5 đã ghi trạng thái cuối của mã).

Danh sách cụm: `_data\_docsau.txt`. Mỗi cụm ghi 1 lần ở dòng đại diện (`Khóa` = `Cụm đọc`), dòng cùng cụm tự nhận theo.

## 1. Spec (MasterFormat PDF)
- Tìm mục liên quan: `py -I "<TOOL>" tim <GOC> "<từ khóa>" --file "<tên file spec>"`.
- Tách mục: `pdftotext -layout "<spec.pdf>" "<scratchpad>\spec.txt"`, rồi `py -I "D:\COSTING EXCEL\_MASTER\scan\scan_spec.py" "<scratchpad>\spec.txt"` → danh sách mục (ref, title, text, page). File trong zip: `py -I "<TOOL>" zip <GOC> <FID>` trước.
- **SPEC**: chỉ các mục ảnh hưởng giá hoặc cấu tạo (vật liệu, độ dày, hoàn thiện, phụ kiện, thử nghiệm, bảo hành, mẫu/mock-up). Mỗi mục 1 dòng: `Mục`, `Nội dung` (nguyên văn), `Nhãn`, `Mã liên quan`, `File · Trang`.

## 2. Bảng hoàn thiện / material list / legend
- Với mỗi cụm: lấy mã vật liệu từ `Mã VL`, `Mã spec/FF&E`; tra `loc <GOC> MAT "Mã VL=<mã>" --cot "Mã VL,Mô tả,NCC/Mã SP,Nguồn spec,Trạng thái cuối,Nội dung cuối (VN)"`. Dùng lại `Nguồn spec` của MAT, chỉ mở trang khi cần thông số MAT chưa có. Mã có `Trạng thái cuối` khác `Theo spec` thì theo kết luận của MAT, không đọc lại TCQ.
- Mã hoàn thiện có khu áp dụng (vd chỉ một loại phòng) → ghi khu vào QUAN_SAT.
- **SCOPE**: `Spec · ML` = link mục spec chính + trang bảng hoàn thiện (nhiều link). Không có → `"KHÔNG CÓ – đã tìm trong <file/mục lục>"`. `Thông số riêng` = EI, hãng, model…
- **QUAN_SAT** (`Nguồn = SPEC`): các trường `Vật liệu gỗ (mã WD/PT)`, `Mặt đá`, `Gương`, `Kính`, `Kim loại`, `Giấy/vải`, `Phần ngoài phạm vi` theo spec, kèm `Tờ · Chi tiết` = link trang spec.

## 2b. Trang dạng ảnh (bảng hoàn thiện scan, legend trên bản vẽ)
Trang ít chữ hoặc bảng nằm trong hình: tự xem ảnh và đọc trước, rồi cho Gemini đọc độc lập cùng ảnh (`py -I "<TOOL>" gemini "<ảnh>" --hoi "Transcribe this finish/material table as JSON rows {code, description, brand, area}. Mark unreadable cells as ?"`). So từng mã: khớp → ghi; lệch → xem lại chỗ đó; vẫn lệch → ghi cả hai vào QUAN_SAT (`SPEC` và `GEMINI`, trạng thái `MÂU THUẪN`).

## 3. ISSUES (S…)
Spec ↔ BOQ, spec ↔ bảng hoàn thiện khác nhau (độ dày, vật liệu, hãng, cấp chống cháy…). `Cụm đọc`, `Loại`, `Nội dung` nêu đủ hai giá trị + mục; `Nguồn` link cả hai chỗ; `Mức` CAO nếu đổi giá rõ.

## Trước khi trả kết quả
Chạy `py -I "<TOOL>" neo <GOC>`. Máy kiểm mã/model trong nhãn link và trong `Thông số riêng` / `Mục` / `Giá trị` có trên trang được link. Dòng `NEO LỆCH` của cụm bạn làm: mở lại đúng trang (dùng `gợi ý`), sửa link hoặc giá trị; nếu đúng thì ghi lý do. LỆCH chỉ là XEM.
Mỗi cụm có `Spec · ML` (link hoặc "KHÔNG CÓ – …"); `check --den spec` không còn LỖI ở file của bạn.
