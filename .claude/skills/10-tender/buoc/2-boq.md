# Bước 2 — boq: kiểm bản đồ BOQ do tool tách

**Ghi:** `SCOPE.boq.json` (chỉ cột `Mã phòng` khi tool tách sai), `_data\_cauhinh.json` (chỉ khóa `boq_bo_sheet`), `VIEC_MO.boq.json`, README mục "Bẫy đã biết".
Tool `boq` đã tách: mỗi dòng có số Item (hoặc có KL + ĐV) = 1 dòng SCOPE, khóa `<sheet>!<dòng>`; `Vùng` (cột sau Item), `Nhóm`, `Mục phụ`, `Mô tả BOQ`, `ĐV`, `KL BOQ`, `Mã phòng` (KL theo từng cột phòng/khu), `Ghi chú BOQ` (Remark Base/VE), `Mã VL`, `Mã spec/FF&E` (regex trong mô tả). Bản đồ cột ở `_data\_boq_map.json`.

## Làm
1. Đọc README. Mở `_boq_map.json`: mỗi sheet có hàng tiêu đề, cột Item, cột KL/ĐV, cột phòng (`phong`), cột Remark, cột cuối.
2. **Kiểm mẫu** mỗi sheet: lấy 3 dòng (đầu, giữa, cuối) bằng `loc <GOC> SCOPE "Sheet=<tên>" --max 400` rồi so với file BOQ (`py -I "<TOOL>" xem "<BOQ>" "<sheet>" A1:AP40`). So: mô tả, ĐV, KL, KL từng mã phòng, Vùng/Nhóm.
3. Sheet không phải hạng mục (Summary, Cover, Room schedule…) → thêm vào `boq_bo_sheet` trong `_cauhinh.json`, chạy lại `py -I "<TOOL>" boq <GOC>`. **Không bỏ** sheet có hạng mục dù ngoài scope nhà máy (tender quét toàn gói).
4. Đối chiếu sheet `Room schedule` (nếu có) với mã phòng trong header: số phòng mỗi mã, mã có trong schedule nhưng không có cột BOQ (vd loại phòng bị loại khỏi gói) → README "Bẫy đã biết" + VIEC_MO nếu chưa rõ.
5. Tool tách sai (mất dòng, gộp nhầm, mã phòng sai cột) → ghi `SCOPE.boq.json` cột `Mã phòng` cho dòng đó, hoặc VIEC_MO mô tả lỗi để sửa tool (không tự viết script).

## BOQ ra bản mới (chạy lại `boq`)
Khóa dòng là số hàng, nên BOQ chèn/xóa hàng thì khóa cũ trỏ sang dòng khác. Lệnh `boq` tự khớp dòng cũ ↔ mới theo nội dung (Vùng, Nhóm, Mục phụ, Mô tả BOQ, ĐV; không xét Item, KL) rồi đổi khóa trong mọi file `_data`. Bản gốc ở `_data\_dot\boq_<giờ>\`.
1. Trước khi chạy thật: `py -I "<TOOL>" boq <GOC> --thu`. Xem số dòng khớp, cũ không còn, dòng mới, nội dung đổi. Nhiều dòng "cũ không còn" mà BOQ chỉ sửa ít → dừng, báo user (có thể đổi mẫu cột, đổi tên nhóm hàng loạt).
2. Chạy `boq`. Đọc `_data\_dot\boq_<giờ>.json`: `moi` = dòng chưa có kết quả, `doi_noi_dung` = dòng giữ kết quả cũ nhưng phải xem lại (cột nào đổi), `xoa` = dòng bị bỏ (kết quả cũ ở `_dong_bo.json`).
3. Dòng mới và dòng đổi nội dung: giao lại `phanloai`, bước 3, 6, 7, `ben`, đọc sâu cho đúng các khóa đó (`thaydoi` cũng liệt kê). `check` báo LỖI ở chỗ còn `(đã xóa: …)` (vd `Dòng BOQ` của TCQ) → sửa tay chỗ đó.

## Trước khi trả kết quả
- Trả về: số dòng mỗi sheet, sheet đã bỏ, mã phòng, lỗi tách (nếu có).
