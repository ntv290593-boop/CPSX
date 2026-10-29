# MODUN CÁNH MỞ (tủ treo + tủ thấp)

Nguồn: Template#1 hàng 12–22 (tủ treo), Template#2 hàng 94–104 (tủ thấp); Kitchen "Bếp trên/dưới - Tủ n cánh" cùng cấu trúc.

## 1. Ô nhập (dòng module)
`B` rộng (gõ tay), `C`=C header, `D`=D header, `E` SL (0 = bỏ), `K`: `Tủ treo` | `Tủ Thấp`, `L`: `0|1|2 lộ hông`, `M`: `Giáp tường trần` | `Giáp tường - hở trần`.
Header: `B = SUMIF(K,"TỦ TREO" hoặc "TỦ thấp", F)` (tổng rộng module → giá theo md), C, D, E nhập.

## 2. Chi tiết (B rộng, C sâu, D cao của module)
| Chi tiết | Rộng | Dài | SL | Vật liệu |
|---|---|---|---|---|
| Nóc | C | B | E | 17MDF (tủ thấp: 9CEM); M = 2 nếu hở trần, ngược lại 3 |
| Đáy | C − dày cánh | B − 2×dày hông | E | 17MDF |
| Hông lộ | C − dày cánh | D − dày nóc | `IF(E=0,0,VALUE(LEFT(L,1)))` | 17MDF |
| Hông khuất | = hông lộ | = hông lộ | `E×2 − E hông lộ` | 17MDF |
| Hậu | MIN(B−2×hông, D hông) | MAX(…) | E | 12MDF |
| Đợt ngang | MIN(C−hậu−cánh, B−2×hông) | MAX(…) | E | `IF(MAX(C,D)>800,25,17)&"MDF"` |
| Cánh | `B/E cánh×E` (rộng 1 cánh) | D | `ROUNDUP(B/600)×E` | 17MDF |
| Chặn cửa | 70 (gõ cứng) | B − 2×hông | `IF(E cánh=0,0,E)` | 17MDF |
| Bổ gia cố | 60 (gõ cứng) | = chặn cửa | `IF(RIGHT(K nóc,3)="cem",E×2,0)` | 17MDF |
| Sắn giữa | C − hậu − cánh | D − nóc − đáy | `MAX(0, E cánh − 2×E)` | 17MDF |
Tên hàng đổi: tủ treo "Hông tủ treo lộ/khuất", "Sắn giữa treo"; tủ thấp "Hông tủ lộ/khuất", "Sắn giữa".

## 3. Cột riêng của dòng module
`G = IF(B×D>0,35 m², ROUND(B×D/0,5)×E, E)`; `J` (tủ treo) = `B×E×4×r`, r = ROUND(D/1200) (hoặc ROUND(D/500) nếu 0) → chiều dài thanh treo (mm), dùng ở LUOIGAZ (xem PHU-KIEN-TU).

## 4. Số cứng
600 (1 cánh/600 mm), 800 (đổi đợt ngang 25mm), 70, 60, 0,35 và 0,5 (G).

## 5. Rule / Chưa xác nhận
R0075, R0039, R0032.
- **Nóc tủ thấp = 9CEM (Cemboard CMB 9mm, Material m²): chủ ý của mẫu**, vì tủ bếp là khu ướt (user xác nhận). Khi K nóc kết thúc "cem" thì "Bổ gia cố khung" tự bật (E×2).
- Cột G của dòng module: tạm để nguyên, chưa dùng (user chốt).
- Kitchen góc/hút mùi chưa đối chiếu từng biến thể.
