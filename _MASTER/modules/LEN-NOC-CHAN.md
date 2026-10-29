# LEN NÓC (tủ trên) / LEN CHÂN (tủ dưới)

Nguồn: Template#1 hàng 44–48 (len nóc), Template#2 hàng 135–139 (len chân), Template#3–#6 (LEN CHÂN), Kitchen#23.

## 1. Ô nhập (dòng module)
Len nóc: `B`=100 (cao len, gõ), `C`=C header, `D`=B header, `E`=SL (nhập, 0 = bỏ), `T`=SONPU. Len chân: `B`=100, `C`=C header, `D`=B header, `E`=1 (gõ), `K`: `Chân nhựa` | khác.

## 2. Chi tiết
| Chi tiết | Rộng | Dài | SL | Vật liệu |
|---|---|---|---|---|
| Mặt len (dài) | B len | D len | `E len` (chân: ×1) | 17MDF, M=3 |
| Mặt len (sâu) | = trên | C len | `E len` (chân: ×2) | 17MDF, M=3 |
| Kè gỗ (nóc) 30×30 | 30 | D len / C len | theo mặt len | THONG |
| Khung xương (chân) | B len | `C − dày mặt×E − 2×dày xương` | `IF(K="chân nhựa",0,ROUND(D/400+1)×E)` | THONG |
| Khung xương ngang | 30 | `D mặt − dày×E` | `IF(K="chân nhựa",0,ROUND(C/600+1)×E)` | THONG |

## 3. Công tắc
`K = "Chân nhựa"` → khung xương E = 0; chân nhựa tính ở phụ kiện `CHANNHUA` (xem PHU-KIEN-TU). Chân gỗ (K khác) → khung xương THONG bước 400/600.

## 4. Số cứng
100 (cao len), 30×30, bước xương 400 (dài) và 600 (sâu).

## 5. Chưa xác nhận
Len nóc: xương/kè gỗ có tính SL ở Template#1 không (E47 = E45). Len chân Kitchen#23 là block riêng (m²) khác cấu trúc này, chưa đọc.
