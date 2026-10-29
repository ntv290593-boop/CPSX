# KITCHEN — Khung tủ lạnh, Tủ cao, Ốp hông, Gia giảm, Len chân, Ốp hậu (Kitchen#18–#24)

Nguồn: Kitchen hàng 986–1312.

## #18 Khung tủ lạnh (mã hệ TL1) — "Tủ khác" hàng 986–1054
Header 800×700×2400, giá **m²**. Module **TỦ TREO TRÊN** (`K=Tủ treo`, 800×700×600) + **HÔNG TỦ CAO** (`K=Tủ Thấp`, 800×700×2400) + **Giằng tủ** (17×200×800) + **GHÉP HỘP** (dày 40 xương gỗ, tấm 9MDF, xương THONG1 22): cơ chế GHEP-HOP. Cánh phần trên `E=ROUNDUP(B/600)×E`, `C=B/E×E`. Spec TL1 không có thùng NK/len chân.

## #19 Tủ Cao (mã hệ TC1) — hàng 1055–1130
Module TỦ (600×600×600, `K=Tủ Thấp`) + **Ngăn kéo** (NGAN-KEO: `C=C−50`, `D=220`, ray theo Spec) + GHÉP HỘP (40×600×600) + LEN CHÂN (LEN-NOC-CHAN). Spec TC1 có cột F thùng ngăn kéo, M len chân. Cánh `E=ROUNDUP(B/600)×E`, `C=B/E×E`.

## #20–#24 Block tấm / thanh, mỗi block một chi tiết
Cùng cấu trúc: header `B`=dày (18), `C`, `D`; module "CABINET CÓ CÁNH CỬA" `E=1`; **Tấm hoàn thiện** 17MDF (`K=VLOOKUP(BL,Spec,cột)`, `N=VLOOKUP(Spec code HT)`, M=3); **Ván liên kết** 17mdf (`C=50`, `D = D tấm`, M=2, N=sl) chỉ ở #21, #22.
| Block | Hệ | Spec cột | Giá theo | Kích thước mặc định |
|---|---|---|---|---|
| #20 Ốp hông (Panel) | TL1 | J (10) | m² | 18×350×750 |
| #21 Gia giảm nóc | BT1 | K (11) | md (theo D) | 18×100×1000 |
| #22 Gia giảm hông | BT1 | L (12) | md (theo D) | 18×50×750 |
| #23 Len chân | BD1 | M (13) | m² | 18×100×1000 |
| #24 Ốp hậu (đảo bếp) | DB1 | N (14) | m² | 18×750×2000 |
Dùng cho các cấu phần Matrix nằm ngoài module (cột Ốp hông, Gia giảm hông/nóc ở Matrix; len chân và ốp hậu khi có).

## Rule
R0127–R0133. Thứ tự cột Spec: KITCHEN-TONG-QUAN mục 2.

## Chưa xác nhận
#18 và #19 chưa đối chiếu từng hàng với Template#6 và #4 (cùng tên module, khác vật liệu/độ dày: giằng 17, ghép hộp 9MDF, xương 22). Mapping cột "Ốp hông / Gia giảm" của Matrix tới #20–#22 chưa rõ (chờ link đơn giá).
