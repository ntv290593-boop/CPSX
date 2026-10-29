# CỬA — Header, công tắc loại cửa và module CÁNH CỬA (Door#1 + các block cửa phẳng)

Nguồn: Door "Cửa phẳng 1 cánh - new" hàng 10–38 (đọc công thức 2026-10-24); #3, #4, #6, #7 dùng cùng logic. Giá theo **bộ**.

## 1. Ô nhập (header)
`B` rộng, `C` **sâu khung bao = dày tường** (R0166 kiểm trên mặt bằng), `D` cao, `E` SL; `J` ván nền (MDF | HDF | PB | PLW | WPB | PIM); `K` **loại cửa** (`Cửa thường` | EI30 | EI45 | EI60 | EI70 | EI90 | EI120); `L` mã hoàn thiện (LA-01…); `M` dày khung bao (EI 35, thường 30; chỉ để mặc định, dự án chỉnh); `O` = 1; `R` = `Khung bao ván ghép` | `khung bao gỗ đặc` (xem CUA-KHUNG-BAO).

## 2. Công tắc theo K (loại cửa), dùng chung cả block
| Cấu phần | Cửa thường | EI30 | EI45/60/70 | EI90/120 |
|---|---|---|---|---|
| Lớp FR (cánh, khung bao, rebate) | không | 5ERON | 10ERON | 10KHS.FA |
| Lớp da ngoài | 6 + J (PIM: 5) | 3 + J (MDF/PLW→3, hdf/PB→6) | = EI30 | = EI30 |
| Lõi cánh | `<dày>Honeycomb` | RW (Rockwool) | RW1 (EI70) | KHS.HF.20 / .24 |
| Dày cánh (B12) | 40 (46 nếu "Cùi chỏ hơi âm") | 50 | 50 | 46 (#3: 64) |
Cột P (test) = `K&O` khi không phải cửa thường: nối vào chi phí test OverView.

## 3. Module "Cánh cửa" (E = số cánh)
| Chi tiết | Rộng | Dài | SL |
|---|---|---|---|
| Cánh cửa (dòng module) | `B−(đố đứng + FR + 3MDF)×2−6` (÷ E nếu nhiều cánh) | `D−(đố đứng+FR+3MDF)−10` | E |
| Lớp da ngoài ×2 mặt | rộng cánh+20 | dài cánh+20 | 2E |
| Lớp FR ×2 mặt (EI) | = da | = da (12KHS.FA: +2×dày) | 2E (thường: 0) |
| Khung xương TRAMGHEPD1 ×2 dọc, ×2 ngang | dày `B cánh−2×da−FR`; rộng `IF(D>2500,110,60)` | D da / `C da−2×rộng` | 2E |
| Xương phụ 25mdf (ngang, giữa, 6 miếng 600 gõ cứng) | 25×dày xương | theo cánh | `ROUNDUP(C/900)×E`, `E×4×2`, **6** |
| Lõi (tổ ong / Rockwool …) | `C cánh−2×rộng xương` | `D da−2×rộng xương` | E |
Cửa 2 cánh: nhân đôi module "Cánh cửa" (xem CUA-BIEN-THE).

## 4. Rule
R0099 (Door#1 mới nhất), R0100 (Door#2 không dùng), R0134 (2 cánh luôn dùng mẫu 2 cánh), R0166–R0179 (kết cấu theo BoQ / bản vẽ / BOHO, ERON mọi cấu phần EI).

## 5. Takeoff cần khai
Rộng, cao, **sâu khung bao theo từng vị trí cửa**, SL, loại EI, ván nền, mã hoàn thiện, số cánh, kiểu cánh (cùi chỏ nổi/âm), có kính/lam chớp.

## 6. Đã chốt (user 2026-10-24)
Template Door **không** có công tắc KETCAU_CUA (Theo bản vẽ / Theo BoQ / Theo BOHO) và **không đưa vào mẫu**. Công tắc chỉ dựng trong block của file dự án khi cần (R0159–R0179), không sửa mẫu Door RE00.
