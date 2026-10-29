# KITCHEN — Bếp dưới (Kitchen#8–#17, mã hệ BD1)

Nguồn: Kitchen hàng 410–984. Chung: MODUN CÁNH MỞ - NGĂN KÉO (MODUN-CANH-MO + NGAN-KEO), `K`=`Tủ Thấp`, cao 750, sâu 600; vật liệu từ Spec (cột D/E/F/G/H…). Nóc: 9CEM chủ ý (khu ướt). Len chân và chân nhựa: LEN-NOC-CHAN, PHU-KIEN-TU. Giá theo md.

| Block | B | Cánh | Ngăn kéo | Ghi chú |
|---|---|---|---|---|
| #8 Tủ 1 cánh | 600 | `E×1` | không | |
| #9 Tủ 2 cánh | 900 | `E×2` | không | |
| #10 Tủ 3 cánh | 1200 | `E×3` | không | |
| #11 SINK | 800 | `E×2` | không | thùng sink (cột G Spec); cánh mỗi cái 400 |
| #12 FULL Ngăn kéo | 600 | `E×0` | `E=3` (gõ) | thùng NK (cột F Spec); `D=220` mỗi ngăn |
| #13 Ngăn kéo + Cánh | 600 | `ROUNDUP(B/600)×E`, **`D = D module − D ngăn kéo`** (cánh thấp lại) | `E=1` | 1 NK trên + cánh dưới |
| #14 Ngăn kéo + Ô trống | 600 | `×0` | `E=1` | ô hở dưới NK |
| #15 Ô trống + Tấm fix | 600 | `×0` | không | **Đợt ngang tấm fixed**: 2 dòng (đợt = đợt ngang thường; thanh `C=150, D=B`), dành cho ô hở chờ thiết bị theo CĐT (lò, máy rửa chén) |
| #16 Tủ góc chữ I | 800 | `ROUNDUP((B−C FIX)/600)×E` | không | **MODUN GÓC**, Tấm FIX 17MDF2M-ME-01 `C=C module`, `D=D cánh` |
| #17 Tủ góc chữ L | 900 (sâu 900) | 2 dòng cánh (trước/bên) | không | MODUN CÁNH MỞ sâu 900, cánh `C = B − C chặn` và `C = C sâu − C chặn` |
Ngăn kéo (khi có): kích thước và SL như NGAN-KEO (`C=C module−50`, `D=220`, ray `RA<n>` theo Spec cột V).

## Quan hệ kích thước
Cánh thấp lại khi có NK phía trên: `D cánh = D module − D ngăn kéo` (530 = 750 − 220). Ô trống (không cánh): thùng + đợt, E cánh `*0`.

## Phụ kiện
Như PHU-KIEN-TU (BASDD, BLGC, ray RA, CHANNHUA theo len chân, tay nắm theo Spec cột U). Số chân nhựa theo cột I của dòng module: `4 + CEILING.MATH(MAX(B−790,0),400)/400`.

## Đã sửa mẫu (user yêu cầu, 2026-10-24): #11 SINK
Vật liệu **thùng sink lấy từ Spec cột G** (thùng sink). Trước đây các hàng thùng của #11 tra nhầm cột H (cánh). Đã sửa trong `260903_DA_CPSX_RE00.xlsm`, sheet Kitchen, hàng **577 Đáy, 578–579 Hông lộ/khuất, 581 Đợt ngang, 584 Bổ gia cố, 585 Sắn giữa**: `K` index 8→7 (A:T), `N` index 7→6 (B:T). Giữ nguyên: Nóc (cột O, tấm đỡ đá), Cánh và Chặn cửa (cột H). Ô sửa tô vàng #FFFF00 (màu cũ #C5D9F1 của cột K/N). Bản sao trước khi sửa: `_MASTER\backup\260903_DA_CPSX_RE00_truoc-sua-sink_2026-10-24.xlsm`.

**Hậu tủ #11 (hàng 580), user chốt:** vật liệu theo thùng sink (Spec cột G), dày 9 như hậu các thùng khác. Đã sửa `K580 = 9&MID(VLOOKUP(BL580,Spec,7,0),3,20)`, `N580` index 6 → `9MDF2M-ME-01`, hoàn thiện `ME-01` (trùng kết quả hậu của block #9).

## #14, #15 đã đọc (2026-10-24)
- **#14 Ngăn kéo + Ô trống:** thùng (đáy, hông, hậu, đợt ngang, bổ gia cố, sắn giữa) lấy **Spec cột I (Tủ không cánh)** vì ô hở thấy được; nóc cột O; cánh `E×0`; ngăn kéo `E=1`; ngăn kéo cùng bộ chi tiết như #12/#13.
- **#15 Ô trống + Tấm fix:** thùng lấy cột I như #14; nóc cột O; cánh `E×0`; 2 dòng "Đợt ngang tấm fixed": dòng 1 = đợt ngang (`C,D` theo đợt thường, E = E module), dòng 2 = thanh `C=150`, `D=B module`, E = dòng 1 (vật liệu cột F, 12MDF2M).
- Block SINK không có dòng riêng cho lỗ bồn.

## Đối chiếu chỉ số cột Spec của mọi block Kitchen (2026-10-24) — đã sửa mẫu theo user
Chuẩn: thùng/đáy/hông/đợt ngang/bổ gia cố/sắn giữa → cột D (K 4, N 3); hậu → E (K 5, N 4); cánh/chặn cửa → H (K 8, N 7); nóc bếp dưới → O (K 15); thùng ngăn kéo → F (K 6, N 5). Đã sửa 56 ô trong `260903_DA_CPSX_RE00.xlsm` (tô vàng, màu cũ #C5D9F1; bản sao trước khi sửa `_MASTER\backup\260903_DA_CPSX_RE00_truoc-sua-spec-idx_2026-10-24.xlsm`):
1. **Thùng ngăn kéo (K12, K13, K14, K19):** Mặt+hậu thùng, Hông thùng, Đáy, Gia cố đáy NK từ cột G (thùng sink) → **cột F (thùng ngăn kéo)**: ra 12MDF2M-ME-01 như Template; riêng K19 (Tủ cao TC1 không có cột G) trước đó vật liệu = 0, nay có.
2. **Mặt ngoài ngăn kéo và Chặn NK (cùng 4 block):** từ cột F → **cột H (cánh / hoàn thiện ngoài)**: ra 17MDF + LA-01, giống Template#2 (M, dán cạnh giữ nguyên).
3. **Hậu tủ K6, K17 (góc chữ L): giữ nguyên, dùng 17MDF theo cột D** (user chốt, không phải lỗi).
4. **N lệch một cột:** Hậu K1–K5, K16 đã sửa N index 3→4 (hậu); Đợt ngang K18, K19 N index 4→3 (thùng).
Block SINK (K11) đã sửa ở mục trên.

## Chưa xác nhận
K14/K15 dùng cột I (Tủ không cánh) cho thùng: user đã đồng ý (ô hở thấy laminate).
