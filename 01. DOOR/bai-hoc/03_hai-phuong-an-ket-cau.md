# 03 — Hai phương án kết cấu (Theo CĐT / Theo BOHO)

Nguyên tắc: R0159 (không gõ giá trị cứng, dùng IF theo dropdown OverView `KETCAU_CUA`).

| Nhóm ô | Cách làm |
|---|---|
| Ô đổi theo CĐT (J ván nền, M dày khung, dày cánh, FR, lõi, FR/skin khung, chỉ, rebate, SL lớp 6mdf) | `=IF(KETCAU_CUA="Theo CĐT"; <giá trị CĐT>; <công thức gốc Door#1, dịch đúng hàng>)` |
| B rộng, D cao | Công thức đọc "overall size …" từ mô tả BoQ — chung 2 phương án |
| E số lượng | `=SUM(<các dòng BoQ>)` (R0156) |
| C sâu tường, K loại, L mã HT | Theo dự án, chung 2 phương án |
| Q mô tả | Tự đổi "Kết cấu theo yêu cầu CĐT" / "Kết cấu theo BOHO đề xuất" |

Kiểm tra: đổi qua lại 2 lần, giá H mỗi block phải về đúng số cũ; không vòng lặp; bảng gộp tấm không lỗi.
