# NGĂN KÉO (tủ thấp / MODUN CÁNH MỞ - NGĂN KÉO)

Nguồn: Template#2 hàng 105–112, ray ở hàng 163; Kitchen#12–14 (FULL NK, NK + cánh, NK + ô trống), Template#3, #4, #5.

## 1. Ô nhập
`B105` rộng ngăn kéo (gõ 450), `D105` cao mặt ngăn (gõ 220), `C105 = C module − 50`.
SL ngăn: `E105 = ((ROUNDUP(B module/B105 − 1)+1) × ROUND(D module/D105)) × E module` = số cột × số hàng.

## 2. Chi tiết
| Chi tiết | Rộng | Dài | SL | Vật liệu |
|---|---|---|---|---|
| Mặt ngoài NK | D105 | B105 | E105 | 17MDF |
| Mặt + hậu thùng | C mặt ngoài − 50 | D mặt ngoài − 13×2 − dày hông thùng×2 | 2×E105 | 12MDF |
| Hông thùng | = mặt+hậu | C105 − dày mặt ngoài | 2×E105 | 12MDF |
| Đáy NK | MIN(dài mặt+hậu, dài hông) | MAX(…) | E105 | 12MDF |
| Gia cố đáy | 100 | dài hông − 2×dày mặt+hậu | `IF(B105>450, ROUNDUP(B105/600),0)×E105` | 12MDF |
| Chặn NK | 50 | B105 − 2×dày đợt ngang | E105 | 17MDF |
| Sắn giữa | C − hậu − cánh | D − nóc − đáy | `ROUNDUP(B module/B105 − 1)×E module` | 17MDF |

## 3. Phụ kiện (xem PHU-KIEN-TU)
Ray: K = `"RA"&MIN(MROUND(dài hông thùng,50),550)` (RA550 khi hông ≥ 550); SL = tổng E dòng "ngăn kéo".

## 4. Số cứng
50, 13 (khe ray mỗi bên), 100, 450 và 600 (ngưỡng gia cố), 17/12 mm.

## 5. Chưa xác nhận
Kitchen FULL NK: header ghi D cao, E số NK = 3; chưa đối chiếu cách tính khác với Template#2. Tên khóa phải giữ nguyên: "ngăn kéo", "Hông thùng Ngăn kéo".
