# Bước 5 — mat: sổ vật liệu, trạng thái cuối

**Ghi:** `MAT.mat.json`, `VIEC_MO.mat.json`.
Tool `mat` đã lập: mỗi `Mã VL` gặp trong mô tả BOQ hoặc trong TCQ, số dòng BOQ, dòng mẫu, các TCQ nhắc tới. Thay đổi ở cấp **mã VL** ghi ở đây, không nhồi vào từng dòng BOQ.

## Làm
1. Đọc README. `loc <GOC> MAT --max 300`.
2. Nguồn spec: Finishes Schedule (bản mới nhất theo FILE_INDEX — vd trong zip bộ chính), FF&E spec, spec thảm/rèm của đợt sau. Giải nén file cần: `py -I "<TOOL>" zip <GOC> <FID>`; đọc `py -I "D:\COSTING EXCEL\SKILL\01-scan\tool\pd.py" chu "<pdf>" <trang>` (tìm trang có mã trước, không đọc cả file bằng ảnh).
3. Mỗi mã ghi: `Loại`, `Mô tả` (tên VL + quy cách chính, như spec), `NCC/Mã SP`, `Nguồn spec` (LINK có trang), `Trạng thái cuối`, `TCQ thay đổi`, `Nội dung cuối (VN)`, `Nhóm thực hiện`.
   - `Trạng thái cuối`: Theo spec / Đổi theo TCQ (TCQ đổi mã SP, độ dày, grade, loài gỗ… — ghi `TCQ thay đổi` + `Nội dung cuối`) / Omit / VE / Chờ (TCQ hoãn, "to be confirmed") / Không thấy spec.
   - TCQ của mã: lấy câu **còn hiệu lực** `loc <GOC> TCQ_LOG "Mã VL~<mã>" --cot "Mã TCQ,Bị thay bởi,Nội dung cuối (VN)"`; câu đã bị thay thì không dùng làm kết luận.
4. Mã có trong spec nhưng không có trong BOQ/TCQ: thêm dòng mới (khóa = mã) chỉ khi spec nói dùng cho khu thuộc gói; ghi `Ghi chú`-kiểu vào `Nội dung cuối`.
5. Mã gặp trong BOQ/TCQ nhưng không thấy trong spec → `Trạng thái cuối` = Không thấy spec + VIEC_MO (Cần ai trả lời = CĐT).
6. Mô tả spec mâu thuẫn mẫu/TCQ (vd tên đá khác mẫu duyệt) → giữ kết luận TCQ còn hiệu lực, ghi mâu thuẫn trong `Nội dung cuối`, thêm VIEC_MO nếu TCQ còn "to be confirmed".

## Trước khi trả kết quả
- `check --den mat` không còn LỖI ở MAT.
- Trả về: số mã, số mã đổi theo TCQ, mã không thấy spec, mã chờ.
