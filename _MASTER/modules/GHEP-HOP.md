# GHÉP HỘP (tấm / thùng / mê-tay-lưng)

Nguồn: 260903_DA_CPSX_RE00.xlsm, đọc công thức 2026-10-24. Dòng module `K` quyết định cả cấu trúc bên dưới.

## 1. Nơi có
Template: "GHÉP HỘP TẤM" (Template#1…#4,#19,#21, vd hàng 35), "GHÉP HỘP THÙNG" (#8 Tủ gương hộp, hàng 609), "GHÉP HỘP - MÊ NGỒI/TAY/LƯNG" (#15 Bench, hàng 1020/1028/1036), Giường "ghép hộp" (#17), Kitchen "GHÉP HỘP". Vách ghép hộp (#11) dùng K = XƯƠNG SẮT (xem VACH-GHEP-HOP, chưa làm).

## 2. Ô nhập (dòng module)
- `K`: `GHÉP HỘP - XƯƠNG GỖ` hoặc `GHÉP HỘP - XƯƠNG SẮT`.
- `E`: SL (0 = bỏ module). `B/C/D`: kích thước. Nên là công thức link header/module trên (mẫu đã làm: C35 `=+C12`, D35 `=+B12`).
- Ý nghĩa B/C/D: **Tấm** B = dày hộp (40), C = sâu, D = dài. **Thùng** B = rộng, C = sâu, D = cao. **Mê/Tay/Lưng** B = dày hộp, C = sâu, D = dài, E nhập tay.

## 3. Chi tiết (K = XƯƠNG GỖ), công thức theo mẫu
| Chi tiết | Rộng | Dài | SL | Vật liệu / mặt |
|---|---|---|---|---|
| Tấm mặt ×2 (Tấm) | C+20 | D+20 | E | 12MDF, M=1 |
| Tấm viền dọc (Tấm) | B+20 | D+20 | `IF(B<51,0,E×2)` | 12MDF, M=1 |
| Tấm viền ngang (Tấm) | B+20 | C+20 | `IF(B<51,0,E×2)` | 12MDF, M=1 |
| Mặt trước (Thùng) | B+20 | D+20 | E | 17mdf, M=3 |
| Mặt sau (Thùng) | = trước | = trước | `E×0` | tắt |
| Mặt bên (Thùng) | C+20 | D+20 | `IF(B<51,0,E×2)` | 17mdf, M=3 |
| Trên, dưới (Thùng) | B+20 | C+20 | `IF(B<51,0,E×2)` | 17mdf, M=3 |
| Xương gỗ (3 hàng) | dày `IF(B−2×dày mặt>100,30,B−2×dày mặt)`; rộng 30 | 0 / C / D | `ROUNDUP(dài₁/400+1)×ROUNDUP(dài₂/400+1)×E` | THONG1 |
Tên hàng xương đổi thành "Tấm nền" khi K khác XƯƠNG GỖ. Hộp dày ≤100: dài xương đầu = 0 và SL = 0 (chỉ hai hàng sau tính).

## 4. Số cứng trong công thức (tri thức của user)
`+20` (cắt gọt hoàn thiện), `51` (ngưỡng có viền), `100` (ngưỡng đổi tiết diện xương), `400` (bước xương), xương 30×30 / 16×30.

## 5. Biến thể
- Tấm: mặt 12MDF 1 mặt hoàn thiện; Thùng: 17mdf 3 mặt; XƯƠNG SẮT: 6MDF, M=1.
- Mê/Tay/Lưng: dày mặt 18, viền 12 gõ cứng trong mẫu, vật liệu K = J header (PLW), xương 99×50 / 30×30. Cả cụm Mê + 2 Tay + Lưng gộp trong một block ghế.

## 6. Rule liên quan
R0155 (nhận diện mặt/đáy theo KÍCH THƯỚC, đáy tách riêng), R0074 (độ dày ván), R0058 (công thức B từ mã K sai với mã ≥30mm), R0038/R0040 (ghép hộp xương sắt).

## 7. Takeoff cần khai
Dày/rộng/sâu/cao hộp, SL, xương gỗ hay sắt, vật liệu mặt, mã hoàn thiện + số mặt, dán cạnh (S).

## 9. Chưa xác nhận
- Mẫu Template#1 tấm viền chỉ có 2 hàng (dọc + ngang), chưa thấy hàng đáy tách riêng như R0155 mô tả; cần user chỉ block đã sửa.
- Hộp dày ≤100 xương chỉ 2 hàng có SL: có chủ ý không?
