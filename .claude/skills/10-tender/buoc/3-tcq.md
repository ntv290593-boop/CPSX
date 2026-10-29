# Bước 3 — tcq: phân loại từng câu TCQ, chuỗi "bị thay bởi"

**Ghi:** `TCQ_LOG.tcq-<n>.json`, `FILE_INDEX.tcq-<n>.json` (dòng `L…` cho link chưa tải), `VIEC_MO.tcq-<n>.json`.
**Chỉ làm các đợt TCQ được giao** (vd 04–05). Tool `tcq` đã tách file TCQ **gộp mới nhất**: mỗi câu 1 dòng, EN + bản dịch VN (từ takenote nội bộ, nếu khớp câu), `Tham chiếu` (Refer TCQ # …), `Link ngoài`, `Tình trạng trả lời` (máy đoán), `Nguồn` (ô gốc).

## Làm
1. Đọc README (thứ tự ưu tiên, bẫy). Lấy câu của nhóm: `loc <GOC> TCQ_LOG "Đợt TCQ~^(4|5)$" --max 200`.
2. Lấy **danh sách gọn các câu đợt trước** để biết câu nào bị câu của bạn thay: `loc <GOC> TCQ_LOG "Đợt TCQ~^(1|2|3)$" --cot "Mã TCQ,Câu hỏi (VN),Trả lời (VN)"`.
3. Mỗi câu ghi:
   | Cột | Ghi |
   |---|---|
   | `Chủ đề` | Phạm vi / Vật liệu / Bản vẽ / Khối lượng / MEP / Thương mại/HĐ / Tiến độ/Logistic / Mẫu/Hồ sơ nộp / Khác |
   | `Loại thay đổi` | KL / VL / Bản vẽ / Phạm vi / Không (nhiều: `VL; Bản vẽ`) |
   | `Trạng thái phạm vi` | Bao gồm / Loại trừ / Omit / VE / MEP / Owner / Chờ / Không đổi — chỉ khi câu nói về phạm vi của hạng mục |
   | `Mã VL` | mã vật liệu nêu trong câu/trả lời (WD-01; ST-07) |
   | `Mã phòng/Khu` | mã phòng hoặc khu câu áp dụng (theo **mã phòng**, không theo số Area) |
   | `Dòng BOQ` | khóa SCOPE khi câu chỉ rõ dòng BOQ (tìm: `loc <GOC> SCOPE "Mô tả BOQ~<từ khóa>" --cot Khóa,Mô tả BOQ,Mã phòng`). Không chắc → bỏ, đừng đoán |
   | `Nội dung cuối (VN)` | 1–2 câu: **kết luận cuối cùng** nhà thầu phải làm theo, có mã nguồn: "Veneer gỗ Walnut cho base; loại khác chào ở VE (TCQ06-1)". Câu bị thay → "Bị TCQ05-24 thay: …" |
   | `Câu hỏi (VN)`, `Trả lời (VN)` | chỉ khi tool chưa có bản dịch (dịch sát, giữ mã/số) |
   | `Tình trạng trả lời` | sửa nếu máy đoán sai: Chưa trả lời (trống) / Hoãn/Chờ (deferred, to be confirmed, later) / Tham chiếu (chỉ "Refer TCQ #…") / Đã trả lời |
4. **Bị thay bởi**: câu sau đổi/huỷ kết luận của câu trước → ghi vào **dòng câu trước** `{"Mã TCQ": "TCQ04-22", "Bị thay bởi": "TCQ05-24"}` (dòng của đợt trước chỉ ghi cột này; nhóm đợt trước không ghi cột đó cho câu ấy). Câu chỉ "Refer TCQ #n" thì `Nội dung cuối` chép kết luận của câu được dẫn + mã, không ghi "bị thay".
5. **Link ngoài** (OneDrive/Drive…): không tải. Mỗi link 1 dòng FILE_INDEX: `{"ID": "L01", "Đường dẫn": "<url>", "Vai trò": "Chưa có (link)", "Ghi chú": "TCQ06-21 spec rèm cập nhật"}` — đánh số `L<n>01`, `L<n>02`… theo nhóm để không trùng nhóm khác.
6. **VIEC_MO** cho: câu Chưa trả lời / Hoãn/Chờ còn hiệu lực; câu trả lời mâu thuẫn với câu khác mà không rõ câu nào thay; hạng mục được nói "bao gồm" nhưng không có trong BOQ. `Liên quan` = mã TCQ, `Cần ai trả lời` = CĐT hoặc Nội bộ.
7. Thứ tự ưu tiên tài liệu (vd "Spec → Drawing → BOQ") nếu có câu nói rõ → README mục "Thứ tự ưu tiên", kèm mã TCQ.

## Trước khi trả kết quả
- Mọi câu của nhóm có `Chủ đề` + `Nội dung cuối (VN)`; `check --den tcq` không còn LỖI ở file của bạn (LỖI "thiếu" của nhóm khác thì bỏ qua).
- Trả về: số câu, số câu bị thay, số câu còn mở, link chưa tải, câu cần hỏi user.
