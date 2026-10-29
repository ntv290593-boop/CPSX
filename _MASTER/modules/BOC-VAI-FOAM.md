# Bọc vải + foam (vách, sofa, tấm nền)

Nguồn: Template#10 hàng 746–752, #11 hàng 813–821, #13 hàng 906–914. Cùng một cụm "Vật tư Sofa".

| Dòng | Rộng × Dài | SL | Vật liệu |
|---|---|---|---|
| Foam bọc | dày (B)=20–50, C/D = C/D tấm | = E tấm bọc | K=`fd` (Material, m³, hao hụt 1,1) |
| Vải (tổng kl vải) | `C = C foam + 2×dày foam + 2×40..80`; D tương tự | `MIN(ROUNDUP(D/1370),ROUNDUP(C/1370))×E foam` | K=`FB.01` (md) |
| Vải lót | = vải | = vải | K=`kate` (md) |
Khổ vải 1370 mm. Dư vải mỗi bên: 40 mm (Vách phẳng), 80 mm (Vách ghép hộp, Tấm nền).

## Công tắc
Vách phẳng: foam + vải chỉ bật khi `N = "VÁCH PHẲNG - BỌC VẢI"` (`E foam = SUMIF(N,…,E)`). Vách ghép hộp / Tấm nền: bật bằng `E` của "Ván bọc vải" / "Foam bọc".

## Nhân công liên quan
PX SOFA (vách phẳng) = `SUMPRODUCT(D vải,E vải) × 0,15 × 10⁻³`. Sofa/ghế: xem nhóm sofa (chưa làm).

## Chưa xác nhận
Mã vải `FB.01` là mã vải mặc định; vải khác chọn qua K (R0052(d): L010/L011 không tính tiền vải → giữ "Không bao gồm").
