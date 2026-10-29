# CỬA — Vật tư phụ, gia công ngoài + phụ kiện

Nguồn: Door#1 hàng 63–84; Door#3 hàng 236–257. Mã ở K, tra Material. Thông số bật/tắt ở **OverView** (không sửa trong block).

| Dòng (K) | SL (cột G) | Bật/tắt |
|---|---|---|
| VTPK | `0,06×E` (gói) | luôn |
| Đóng gói `DG.D` (thường) / `DG.D1` (EI) | `((B+D)×2×C+D×B×2)×10⁻⁶×E`, 2 dòng: khung 200×C×D và cánh | luôn |
| `fm` = **chi phí phủ phim bề mặt** (Material: 35.000 đ/mặt, phân loại FM) | EI: Σm² ván PN/PN1 chưa sơn /2,98; thường: 0 | theo K |
| `KEO.MDF` = keo Techbond L560S + xúc tác LAP (đ/tấm) | thường `(da+lớp mặt rebate)/2,98`; EI `Σm² FR/2,98` | R0027: cộng cứng số hàng |
| `KEO.MGO` = keo Cosmo 125.130 hoặc Adino PUR 5890.0 (đ/tấm) | EI: `Σm² FR/2,98`; thường 0 | |
| `TK+VIT` (tắc kê nhựa + vít 10×100, bộ) | `F×5` | OverView!C50 = "Không" → 0 |
| `KEOEPCC` = keo ép cửa thường Kony Bond CH42 (hoặc CU3 cửa chậm cháy) | `*0` | **đã bỏ, thay bằng các keo khác (KEO.MDF/KEO.MGO/fm); công thức để sẵn phòng khi dùng lại** |
| `KEOCANH` | `SUM(BJ)` | theo mã chỉ cạnh cột R hàng "Cánh cửa" |
| `FOAM` (thường) / `FOAMCC` (EI) | thường E>100: 0,75; khác 1 | OverView!C49 |
| `DOORSEAL` (Hafele) / `DOORSEAL1` | `(D rebate đứng×E + D rebate ngang×E + D cánh×(E−1))×1,05/1000` | OverView!C48, D48 (hãng) |
| `RonCC` / `RonCC1` (ron chống cháy) | `ROUNDUP((C+D)×2×E/2100)` (chỉ EI) | OverView!C47, D47 |
| `tipbond` (keo dán nẹp khung bao, chai) | `(D chỉ cửa đứng×E + D chỉ cửa ngang×E)/15×10⁻³` | OverView!C51 |
| `A300` = keo silicon Apollo A300, màu trong (chai), dùng cho cửa kính | `SUMIF("Ô kính",F)/10` (F = hao hụt/diện tích ô kính) | cửa có kính |
| `10KCL` = kính trong cường lực 10 mm (m²) | `C×D×E` của ô kính | cửa thường có kính; cửa EI dùng mã kính EI theo K |
| `SS.01` len inox | `C=100`, `D=(C cánh+B cánh)×2` | mở khi dự án có |

## Nhân công (giữ nguyên, R0150)
PX MỘC 1 người × 2 × 6 ngày ×1,1 (=13,2 công/SP); PX SƠN theo AA (SONPU×D19, khác×D20); CƠ KHÍ `I54/H83`; BỐC XẾP `ΣBBPL×D37`.

## Chi phí test cửa chống cháy
Cột P dòng đầu = `K&O` (vd EI601) → OverView!C21:C23 tra PCCC (R0144).

## Đã xác nhận (user 2026-10-24)
Mã không biết thì tra sheet Material (cột D tên, E ĐVT). `KEOEPCC` đã bỏ, giữ công thức `*0` để dùng lại khi cần. Không còn mục chưa xác nhận.
