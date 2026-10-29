# Kim loại gia công ngoài (khung sắt, chân sắt, tấm inox)

Nguồn: Template#10 hàng 771–774, #11 hàng 836–839, #12 hàng 881–882, #13 hàng 924–925; R0033, R0034, R0038, R0073.

## Cấu trúc
Dòng tiêu đề "Kim loại gia công ngoài": `B/C/D` = kích thước vùng khung (link B/C/D module), `E` = SL, `F` = "Độ khó", `M` = "Độ dày", `N` = "Loại". Mỗi dòng chi tiết: `K` mã vật tư, `B×C` tiết diện, `M` dày thép, `N` loại, `D` chiều dài, `E` SL.

## Khối lượng (kg), cột G
`G = v × 0,00000785 × E × F`, trong đó v theo `N`:
- `hộp`: `(B+C)×2×M×D`
- `Tấm`: `C×D×M`
- `tròn đặc`: `π×B²/4×D`
- `Ống`: `π×M×(B−M)×D`
7,85e-6 = tỉ trọng thép (kg/mm³). F = hệ số "Độ khó" (mặc định 1).

## SL khung
Ngang: `ROUNDUP(D'/600+1)×ROUNDUP(D''/600+1)×E` (vách ghép hộp, bước 600), 500 (vách phẳng 20×40). Đứng tương tự, hoán đổi D.

## Mã và giá
`MTGC1` = sắt gia công chi tiết dễ, **80.000 đ/kg** (R0033; không dùng MTGC 120.000). `SSGC` = inox gia công (tấm SSGC: `G = C×D×M×7,85e-6`, R0038). Mã phải có trong Material (R0073); không gõ tay giá vào H.

## Khung cơ khí cũ
`Vật tư cơ khí` (MT.20401.2, MT.30301.2) nhân `*0` ở hạng mục đã chuyển sang gia công ngoài (tránh tính 2 lần).

## Chưa xác nhận
Chi tiết "Độ khó" F: giá trị khác 1 dùng khi nào; chưa thấy rule.
