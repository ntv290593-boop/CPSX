# Phụ kiện tủ + khung sắt (mục GIA CÔNG NGOÀI+PHỤ KIỆN / Vật tư cơ khí)

Nguồn: Template#1 hàng 56–60, 70–76; Template#2 hàng 161–167. Mã ở K, tra Material; SL (cột G) = công thức **SUMIF/SUMIFS theo TÊN dòng chi tiết** (xem CONG-THUC-DONG).

| Mã K (Material) | SL (G) | Ghi chú |
|---|---|---|
| BASDD: Bas đỡ kệ di động Hafele 282.24.720 (cái) | `SUMIF("Đợt ngang", E)×4` | 4 bas/đợt ngang |
| BLGC: Bản lề Hafele trùm ngoài giảm chấn + đế 315.07.113 (cái) | `SUMPRODUCT(cánh tủ; D≥2000→5, >1500→4, ≥1000→3, còn lại→2; ×E)` | bản lề theo chiều cao cánh |
| CHW700: Bas treo tủ đế gắn tường Hafele 290.00.740 (bộ) | tủ treo: `SUM(SUMIF("Hông tủ treo lộ","Hông tủ treo khuất","Sắn giữa treo"))` | |
| LUOIGAZ: Lưỡi gà treo nhôm chữ Z (thanh) | đổi K72 sang LUOIGAZ: `ΣJ module tủ treo/6000` | thanh dài 6 m |
| PSF: Thanh chống mo Hafele 406.99.120, Hmax 2400 (bộ). Có PSF18/PSF21 (DTC, Hmax 1800/2100) | cánh D>1700: rộng<590→1; 590–1200→2 (×E) | cánh cao, chọn mã theo chiều cao |
| TIPON513: Nêm nhấn Hafele 356.12.513 + đế 356.12.526 (cái) | công thức `*0` (tắt) | tủ không tay nắm |
| TAYNANG (tay nâng tủ bếp trên), TAYNAM (tay nắm, "tạm tính") | công thức `*0` (tắt), đơn giá Material = 0 | **công thức để sẵn**; dự án nào có thì mở `*0` và tính KL; giá chờ báo giá |
| RA<n> (ray, tủ thấp) | `SUMIFS("ngăn kéo", E)`; K = `"RA"&MIN(MROUND(dài hông thùng,50),550)` | Template#2 |
| CHANNHUA: Chân nhựa Hafele 637.45.326 + bas 637.45.906 (bộ) | `IF(K len chân="Chân nhựa", SUMIF(K,"TỦ THẤP", I), 0)` | I của dòng module xem dưới |

**Số chân nhựa (cột I dòng module tủ thấp):** `(4 + CEILING.MATH(MAX(B−790,0),400)/400) × E`. Module rộng ≤790: 4 chân; mỗi 400 mm rộng thêm: +1 chân.

## Khung sắt (Vật tư cơ khí, hàng 57–60)
Dòng đầu nhập B/C/D (kích thước khung, mặc định 50/1000/1000), E = SL. Hàng "Khung xương": K = `MT.30301.2` (tiết diện parse từ mã: `MID(K,4,2)`×`MID(K,6,2)`); SL = `ROUNDUP(dài₁/600+1)×ROUNDUP(dài₂/600+1)×E` (bước 600). Rule: khung sắt gia công ngoài dùng MTGC1 (R0033, R0034): xương cũ nhân `*0`.

## Vật tư phụ (cố định theo mẫu)
`Vật tư phụ khác` (VTPK, B=`kd(…)` là UDF VBA, #NAME? ở Office.js bình thường, R0068), `Vật tư đóng gói` (DG.3, B/C/D = header), KEOCANH, KEOEP. Nhân công giữ như mẫu (R0150).

## Đã xác nhận (user 2026-10-24)
Dòng `*0` là công thức để sẵn: dự án nào có phụ kiện đó mới mở và tính KL. `CHANNHUA` dùng cột I (số chân) của dòng module. Không còn mục chưa xác nhận.
