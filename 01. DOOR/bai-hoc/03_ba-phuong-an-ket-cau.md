# 03 — Ba phương án kết cấu (Theo bản vẽ / Theo BoQ / Theo BOHO)

Nguyên tắc: R0159 (không gõ cứng, IF theo dropdown OverView `KETCAU_CUA`). Mặc định chào giá **Theo bản vẽ** (R0163).

| Phương án | Lấy cấu tạo từ | Ghi chú |
|---|---|---|
| Theo bản vẽ | Chi tiết cấu tạo trên BV, đọc từng lớp bằng HÌNH (R0167) | BV thường đủ lớp hơn BoQ (vd khung bao 2 mặt ván / 2 mặt FR) |
| Theo BoQ | Mô tả BoQ CĐT | Để so khi BoQ ≠ BV, kèm RFI |
| Theo BOHO | Công thức gốc block mẫu | Phương án đề xuất của công ty |

| Nhóm ô | Cách làm |
|---|---|
| Ô đổi theo phương án (J ván nền, M dày khung, dày cánh, FR, lõi, FR/skin khung, chỉ, rebate) | `=IF(KETCAU_CUA="Theo BOHO"; <công thức gốc>; IF(KETCAU_CUA="Theo bản vẽ"; <BV>; <BoQ>))` — **không** gộp BV+BoQ bằng `<>"Theo BOHO"` (R0170) |
| Lớp thêm (ván/FR giáp tường) | **Chèn HÀNG RIÊNG** (copy dòng ván/FR, bỏ hoàn thiện M=0), E = IF(phương án có lớp này; E đố; 0). KHÔNG nhân E dòng có hoàn thiện (M=1) → laminate bị tính 2 mặt. Lõi = M − Σ dày các lớp |
| B, D, E, C, K, L | Chung 3 phương án (C = dày tường đo trên mặt bằng, R0166) |
| Q mô tả | Tự đổi theo phương án |

Kiểm tra (R0179): đổi lần lượt 3 phương án, báo giá trước/sau mọi block (bảng gộp tấm làm cửa khác nhích giá), đếm ngoặc công thức, trả C52 về giá trị cũ;, giá H từng block về đúng số cũ; không vòng lặp; tổng dày các lớp = dày ghi trên BV.
