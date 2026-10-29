# Công thức chung của MỌI dòng chi tiết (Template/Door/Kitchen)

Nguồn: Template hàng 13, 95… (đọc công thức 2026-10-24). Người bóc chỉ sửa K, M, N, S (R0075).

| Cột | Ý nghĩa | Công thức |
|---|---|---|
| B | dày (mm) | `IF(K="",0,IFERROR(--IF(LEFT(K,1)<"3",LEFT(K,2),LEFT(K,1)),0))`; sai với mã ≥30mm (R0058) |
| C, D | rộng, dài | công thức từ dòng module/header (xem từng file) |
| E | SL | công thức từ E module; `0` = tắt dòng |
| F | hao hụt | E=0 → 0; PN + N∈{SL,UV,0} → 1,2; mã có trong Material → Material!I; SOL.S → LET theo kích thước nhỏ nhất |
| G | khối lượng | `LET(p=D×E×F; SWITCH(J: "m3"→p×B×C/1e9, "m2"→p×C/1e6, "SF"→p×C×2/93000, "md"→p/1000))` |
| H | đơn giá | `Material!F` của K + chi phí hoàn thiện (nếu L="pn": tra N, O trong Material) |
| I, J | thành tiền, ĐVT | `I = G×H`; `J = VLOOKUP(K, Material, 5)` |
| L | phân loại | `XLOOKUP(K, Material!A, Material!B)` (PN = ván phủ, SOL.S/SOL.SL = gỗ solid…) |
| T, V, W | sơn PU / lót / số mặt-sơn | theo O, L, N, M (xem R0042, R0043, R0154) |

## Quy ước dòng module (R0075)
`B` rộng module, `C`/`D` lấy header, `E` nhập tay (0 = bỏ), `K` loại, `L` "n lộ hông", `M` giáp tường/trần, `F = B×E`.

## CẢNH BÁO: tên cột A là KHÓA công thức
Phụ kiện dùng `SUMIF/SUMIFS(A:A, "Đợt ngang" | "cánh tủ" | "ngăn kéo" | "Hông tủ treo lộ"…)`. Đổi tên dòng chi tiết → mất phụ kiện. Chỉ đổi theo R0032 (tên khóa của mẫu giữ nguyên).
