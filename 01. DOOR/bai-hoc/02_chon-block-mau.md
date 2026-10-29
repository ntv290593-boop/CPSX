# 02 — Chọn block mẫu Door

| Loại cửa | Block | Ghi chú |
|---|---|---|
| Phẳng 1 cánh, không kính (thường hoặc chống cháy) | Door#1 "Cửa phẳng 1 cánh - new" | Chọn ở ô K header: `Cửa thường` / `EI30` — công thức tự rẽ nhánh skin, lớp FR, lõi (Honeycomb/RW), rebate |
| Có ô kính | "Cửa phẳng 1 cánh - Kính" | R0026; không dùng cho cửa không kính |
| 2 cánh | "Cửa phẳng 2 cánh" (/- Kính) | R0134 |

**Header Door#1:** B rộng · C sâu tường (**bắt buộc check MẶT BẰNG từng vị trí cửa – R0166**) · D cao · E SL · J ván nền (MDF/WPB) · K loại · L mã hoàn thiện · M dày khung bao.

**Offset dòng so với header:**

| Offset | Chi tiết | Offset | Chi tiết |
|---|---|---|---|
| +2 | Cánh (B = dày cánh) | +12 | Khung bao |
| +3 | Skin | +13 | Đố đứng |
| +4, +5 | Lớp FR | +14 | FR khung |
| +6..+10 | Khung xương | +15 | Skin khung |
| +11 | Lõi | +16 | Đố ngang |
| +19 / +20 / +21 | Chỉ đứng / Chỉ FR / Chỉ ngang | +23 / +24 / +25 | Rebate đứng / FR / 6mdf |
| +26 / +28 | Rebate ngang / ngang 6mdf | | |
