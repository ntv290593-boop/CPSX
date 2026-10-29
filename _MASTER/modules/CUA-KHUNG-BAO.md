# CỬA — Khung bao, chỉ cửa, rebate

Nguồn: Door#1 hàng 22–38; Door#3 hàng 188–204; Door#7 hàng 572–588.

## 1. Ô nhập
Header: `M` dày khung bao, `C` sâu (dày tường), `R` (xem mục 3). Dòng module "Khung bao": `B,C,D` = header, `E`=1.

## 2. Chi tiết
| Chi tiết | Rộng | Dài | SL | Vật liệu |
|---|---|---|---|---|
| Đố đứng | C | D | 2E | TRAMGHEPD; dày `M − dày da khung − dày FR khung` |
| Lớp FR + lớp da đố đứng | C | D | `IF(R="khung bao gỗ đặc",0,2E)` | FR theo K (ERON/KHS.FA); da 6MDF |
| Đố ngang | C | B | E | TRAMGHEPD |
| Chỉ cửa đứng | 50 | D | 2×số đố đứng | thường 12 + J; EI30 9 + J; khác 3 + J |
| Chỉ cửa ngang | 50 | B+100 | 2×số đố ngang | như trên |
| Rebate đứng | **C − dày cánh** | D cánh + 20 | 2E | 15 + J (vd 15MDF) |
| Rebate ngang | = rebate đứng | `C cánh×E+15×2` | E | như trên |
Mỗi dòng đố/chỉ/rebate có thêm "lớp mặt": **cửa thường → 6mdf (E = E dòng gốc), cửa EI → lớp FR cùng loại (E=0 khi cửa thường)**.

## 3. Công tắc R (header)
`R = "khung bao gỗ đặc"` → mọi lớp phủ, FR, chỉ cửa, rebate có E = 0; chỉ còn đố gỗ đặc. Mặc định `Khung bao ván ghép`.

## 4. Quan hệ kích thước (không gõ cứng)
Dày đố đứng = `M − Σ lớp phủ`; rebate rộng = `C − dày cánh` (cánh 50 → 60; cánh 40 → 70); xương cánh không phụ thuộc khung bao.

## 5. Rule
R0166 (kiểm dày tường trên MB), R0168–R0172 và R0175–R0178 (rebate EI 3 lớp: 3MDF + 6ERON + 9MDF, rộng 80; cửa thường rộng 35), R0027 (keo tính theo vị trí hàng khung bao).

## 6. Đã chốt (user 2026-10-24)
Rộng rebate 35/80 (R0177) chỉ áp dụng ở block dự án (gõ đè ở C hàng "Rebate đứng"); **mẫu giữ công thức `C − dày cánh`**, không sửa.
