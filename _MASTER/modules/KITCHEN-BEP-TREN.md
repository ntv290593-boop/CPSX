# KITCHEN — Bếp trên (Kitchen#1–#7, mã hệ BT1/BT2)

Nguồn: Kitchen hàng 10–407. Chung: MODUN CÁNH MỞ / GÓC như MODUN-CANH-MO, MODUN-GOC; vật liệu từ Spec (KITCHEN-TONG-QUAN). `K` module = `Tủ treo`; header cao 750, sâu 350 (góc L: 600). Giá theo md.

| Block | B | Khác với MODUN CÁNH MỞ |
|---|---|---|
| #1 Tủ 1 cánh | 600 | `E cánh = E module` (1 cánh) |
| #2 Tủ 2 cánh | 900 | `E cánh = 2×E module` (mỗi cánh 450) |
| #3 Tủ 3 cánh | 1200 | `E cánh = 3×E module` (mỗi cánh 400) |
| #4 Tủ hút mùi | 600 | thêm **Thùng fix che máy hút mùi** 200×200 (D = D hông): Tấm mặt dựng trước dưới (17, C=150, D = D đáy), Tấm mặt dựng trước che ống (C=200, `D = D thùng − C tấm dưới`), Tấm mặt dựng hông che ống (C=200, D=733, ×2) |
| #5 Tủ góc chữ I | 600 | **MODUN GÓC**: Tấm FIX (17MDF, `C=C module`, `D=D cánh`); cánh `C=(B−C Tấm FIX)/E×E`, `E=ROUNDUP((B−C FIX)/600)×E` |
| #6 Tủ góc chữ L | 600 (sâu 600) | MODUN CÁNH MỞ có **2 dòng "Cánh tủ"**: cánh trước `C = B − C chặn cửa`, cánh bên `C = C sâu − C chặn cửa`, `E=ROUNDUP(B/600)×E` mỗi dòng |
| #7 Tủ không cánh | 600 | `E cánh = E×0` (tắt), chặn cửa `E=0`; K cánh/thùng lấy cột "Tủ không cánh" của Spec |
Cột I (Spec) = "Tủ không cánh". Tấm "Sắn giữa" tự bật khi số cánh vượt 2 mỗi module (`E = MAX(0, E cánh − 2×E)`), nên chỉ #3 (3 cánh) có.

## Phụ kiện (Kitchen#1)
BASDD `Σ"Đợt ngang"×4`; BLGC theo chiều cao cánh; PK treo `K = Spec col 20` (CHW700 | Luoigaz): CHW700 `Σ(Hông treo lộ/khuất, Sắn giữa treo)`; Luoigaz `ΣJ/6000`; PSF; TIPON513, TAYNANG `×0`; **tay nắm `K = Spec col 21`, SL = 1 nếu K khác rỗng** (mới so với Template). Chi tiết mã: PHU-KIEN-TU.

## Chưa xác nhận
Chưa đọc từng hàng của #4 (che hút mùi) ngoài các dòng trên, #6 và #7. Kitchen#2/#3 header ghi "Tủ 2/3 cánh" nhưng rộng 900/1200 chia 450/400: cánh chia đều theo `B/E`.
