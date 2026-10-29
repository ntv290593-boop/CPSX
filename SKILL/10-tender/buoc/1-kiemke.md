# Bước 1 — kiemke: vai trò từng file, đợt hiệu lực

**Ghi:** `FILE_INDEX.kiemke.json` (cột `Vai trò`, `Thay thế cho`, `Ghi chú`), `VIEC_MO.kiemke.json`, và các mục "Đầu vào / Đợt hồ sơ hiệu lực / Bẫy đã biết" trong `<GOC>\10_TENDER\README.md`.
**Không đọc nội dung bản vẽ/spec.** Chỉ tên file, thư mục, mail, trang đầu khi cần.

## Làm
1. Đọc `<GOC>\10_TENDER\README.md`. Tool `kiemke` đã chạy (điều phối chạy). Xem kết quả:
   - `py -I "<TOOL>" loc <GOC> FILE_INDEX "Vai trò=" --cot ID,Đường dẫn,Trong zip --max 400` — file chưa có vai trò.
   - Tool đã đoán: Mail (.msg), Làm rõ (file TCQ), Tham khảo (as-built, schedule, nội quy, annex), Ngoài scope (zip MEP/LIGHTING), Bản cũ/Trùng (DWG trùng PDF cùng ngày issue). Đoán sai → ghi đè.
2. Đọc **mail của từng đợt** (`py -I "D:\COSTING EXCEL\SKILL\01-scan\tool\pd.py" mail "<file.msg>"`) — mail nói đợt đó gửi gì, thay gì, link nào chưa đính kèm.
3. Gán `Vai trò` cho mọi file, ưu tiên dòng `_loc` theo thư mục:
   ```json
   {"_loc": {"Trong zip": "/PDF/"}, "Vai trò": "Chính"}
   {"_loc": {"Đường dẫn": "^6\\.261016_TCQ04/20261015_Carpet"}, "Vai trò": "Bổ sung/Thay thế", "Thay thế cho": "Thảm trong bộ FF&E 20260605", "Ghi chú": "TCQ04"}
   ```
   | Vai trò | Khi nào |
   |---|---|
   | Chính | bộ thiết kế phát hành gốc dùng để chào giá (IDD, FF&E, Finishes Schedule, BOQ CĐT, HSMT) |
   | Bổ sung/Thay thế | đợt sau phát hành bổ sung hoặc thay một phần bộ chính (addendum, bản vẽ/spec cập nhật) → ghi `Thay thế cho` |
   | Làm rõ | TCQ, bản dịch nội bộ |
   | Tham khảo | as-built, VE list, tiến độ, nội quy, logistic, rendering |
   | Ngoài scope | thuộc gói khác (MEP, lighting…), annex hợp đồng mẫu |
   | Bản cũ/Trùng | bản bị đợt sau thay, file trùng, DWG đã có PDF |
   | Mail | .msg |
   | Chưa có (link) | dòng `L01…` do bước tcq thêm (link ngoài chưa tải) |
   PDF ưu tiên; DWG chỉ "Chính" khi tờ đó không có PDF.
4. **Đợt hiệu lực** — thứ tự đợt theo tiền tố số + yymmdd của thư mục. Ghi README mục "Đợt hồ sơ hiệu lực": mỗi đợt 1 dòng (ngày, gửi gì, thay gì, nguồn = mail/file). Thư mục input cũ của dự án (nếu ACTIVE.md nói đã bị thay) ghi rõ "bị thay bởi …".
5. **Bẫy** thấy được từ tên file/thư mục (vd 2 bộ đánh số Area khác nhau cho cùng khu) → README mục "Bẫy đã biết", kèm nguồn.
6. Hạng mục có bộ bản vẽ nhưng BOQ CĐT không có sheet tương ứng (tên sheet BOQ: các khóa trong `"sheet"` của `<GOC>\10_TENDER\_data\_boq_map.json`) → 1 dòng VIEC_MO (`Cần ai trả lời` = Nội bộ), không tự xử lý.

## Trước khi trả kết quả
- `check --den kiemke` không còn LỖI "thiếu Vai trò".
- Trả về: số file theo vai trò, các đợt và quan hệ thay thế (1 dòng/đợt), câu cần hỏi user.
