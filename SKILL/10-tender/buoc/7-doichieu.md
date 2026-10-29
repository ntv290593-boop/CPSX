# Bước 7 — doichieu: gắn phạm vi, nhóm thực hiện, thay đổi cuối cho từng dòng BOQ

**Ghi:** `SCOPE.doichieu-<n>.json` (`Nhóm thực hiện`, `NTP chi tiết`, `Trạng thái`, `Mã TCQ`, `Nội dung cuối (VN)`, `Loại thay đổi`, `Việc còn mở`, `Mô tả NT – Base`, `Mã VL`, `Nguồn`), `VIEC_MO.doichieu-<n>.json`.
**Chỉ làm các sheet được giao.** Tool `doichieu` đã nháp: `Mã TCQ` (TCQ còn hiệu lực có `Dòng BOQ` = dòng này, hoặc trùng `Mã VL`) và `Trạng thái` khi TCQ nói rõ (Bao gồm→Base, Loại trừ, Omit, VE, Chờ).

## Làm
1. Đọc README (thứ tự ưu tiên, phạm vi loại trừ, bẫy). Dòng của nhóm: `loc <GOC> SCOPE "Sheet=<sheet>" --cot "Khóa,Vùng,Nhóm,Mô tả BOQ,Mã phòng,Ghi chú BOQ,Mã VL,Mã TCQ,Trạng thái,Bản vẽ hiệu lực" --max 400`.
2. Tham chiếu gọn: TCQ phạm vi còn hiệu lực `loc <GOC> TCQ_LOG "Bị thay bởi=" "Chủ đề~Phạm vi|MEP" --cot "Mã TCQ,Trạng thái phạm vi,Mã phòng/Khu,Dòng BOQ,Nội dung cuối (VN)"`; MAT đổi `loc <GOC> MAT "Trạng thái cuối!=Theo spec" --cot "Mã VL,Trạng thái cuối,Nội dung cuối (VN)"`.
3. Mỗi dòng ghi (dùng `_loc` cho cả nhóm/vùng, dòng cụ thể cho ngoại lệ):
   | Cột | Ghi |
   |---|---|
   | `Nhóm thực hiện` | **Nhà máy** (mộc, sơn đồ gỗ, cửa, cabinet, loose furniture) / **NTP** (đá, thảm, rèm, kính, inox, chống thấm, xây, phá dỡ, sơn nước, trần, sàn gỗ, giấy dán tường… → ghi loại ở `NTP chi tiết`) / **MEP** (gói MEP làm) / **Owner** (CĐT tự làm/cấp) / **Gói khác** / **Chung** (điều kiện chung, prelims, provisional sum) |
   | `Trạng thái` | Base / VE (chỉ có ở phương án VE) / Loại trừ (TCQ loại khỏi gói) / Omit (CĐT bỏ hạng mục) / Chờ (TCQ hoãn, chưa quyết) / Không áp dụng (dòng tổng, ghi chú) |
   | `Mã TCQ` | mọi TCQ **còn hiệu lực** tác động dòng này (giữ nháp tool nếu đúng, bỏ mã không liên quan) |
   | `Nội dung cuối (VN)` | khi có thay đổi so với mô tả BOQ: 1–2 câu kết luận + mã nguồn (TCQ / MAT / Số BV). Không đổi → bỏ |
   | `Loại thay đổi` | KL / VL / Bản vẽ / Phạm vi / Không |
   | `Việc còn mở` | câu hỏi còn treo cho dòng này (TCQ hoãn, lệch BOQ ↔ bản vẽ chưa rõ bên thắng) |
   | `Mô tả NT – Base` | **chỉ** khi TCQ/MAT đã đổi spec base so với mô tả BOQ: viết mô tả nhà thầu chào theo kết luận cuối (tiếng Anh, văn phong BOQ, ghi mã TCQ). Không có thay đổi → bỏ. Đây là cột gửi CĐT — không thêm đề xuất riêng |
   `Mô tả NT – Đề xuất`, `KL – Đề xuất`: **không ghi** (chỉ user).
4. Thứ tự ưu tiên khi mâu thuẫn: theo README (vd Spec → Bản vẽ → BOQ). Áp đúng thứ tự đó cho `Nội dung cuối`; nếu thứ tự không giải được → `Trạng thái` Chờ + VIEC_MO.
5. Phạm vi tender là **toàn gói** — dòng ngoài năng lực nhà máy vẫn là Base của nhà thầu chính (Nhóm = NTP), không phải Loại trừ.

## Trước khi trả kết quả
- Mọi dòng của nhóm có `Nhóm thực hiện` + `Trạng thái`; `check --den doichieu` không còn LỖI ở file của bạn.
- Trả về: đếm theo Trạng thái và Nhóm thực hiện, số dòng có thay đổi, số việc còn mở.
