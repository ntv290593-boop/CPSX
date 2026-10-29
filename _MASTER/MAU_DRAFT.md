# MAU_DRAFT — Danh mục mẫu (chờ duyệt, CHƯA ghi RULES_MASTER)

Nguồn: `D:\COSTING EXCEL\260903_DA_CPSX_RE00.xlsm` (đọc 2026-10-21). Chủ đề = `MAU`, Dự án gốc = `260903_DA_CPSX_RE00`, Trạng thái = Active, Phiên bản = 1. ID nối tiếp ID cuối trong RULES_MASTER khi ghi.
Định dạng cột Quy tắc: `[Sheet#n] Tên (~dòng) | Dùng khi | Giá theo | Module | Ô nhập | PK | Gần giống | Ghi chú`

## 0. Cấu trúc chung

| Mã | Quy tắc |
|---|---|
| MAU-00 | [MAU-00] Cấu trúc block (Template/Door/Kitchen): tìm block theo tên ở cột A, kết thúc ở dòng A="END"; số dòng ghi trong rule chỉ để tham khảo. Header: B rộng, C sâu, D cao, E SL (mặc định 0), M ĐVT giá (md = H/B·1000; m2 = H/F; trống = cái/bộ), G nhóm CB tủ / PA tấm / LS loose / DR, DRF cửa. Dòng module (trong mục VẬT TƯ…): E nhập tay (0 = bỏ module), B/C/D nhập khi khác header, K loại (Tủ treo/Tủ Thấp/GHÉP HỘP - XƯƠNG GỖ|SẮT…), L số hông lộ, M giáp tường/trần, T sơn, V sơn lót. Dòng chi tiết: B/C/D/E là công thức, chỉ sửa K mã ván, M số mặt HT, N mã HT, S dán cạnh (và vài C cố định như chặn cửa 70, gia cố 60). Phụ kiện: mã ở K trong mục GIA CÔNG NGOÀI+PHỤ KIỆN (sắt = MTGC, inox = SSGC). Nhân công: B số SP, C số người, D số ngày. Thứ tự mục cố định: VẬT TƯ MỘC → Sofa → Cơ khí → VT phụ/đóng gói → GIA CÔNG NGOÀI+PK → NHÂN CÔNG → % điện nước/VP/khấu hao/GC → GIÁ GỐC NHÀ MÁY. Mẫu chưa nhập mà ra #N/A hoặc #DIV/0! là bình thường, không tự sửa. Độ dày, mã ván, kích thước trong mẫu chỉ là mặc định: luôn chỉnh theo spec/bản vẽ dự án đang bóc. Nghiệp vụ không rõ thì hỏi người dùng, không đoán. |
| MAU-01 | [MAU-01] Chọn Template hay Kitchen cho bếp: BoQ có nhiều hệ bếp → bóc module ở sheet Matrix/Spec, mỗi module tính bằng mẫu Kitchen tương ứng (tủ rời, lặp lại theo hệ). Trường hợp bình thường (không lặp hệ) → dùng Template#1 Tủ bếp trên / Template#2 Tủ bếp dưới. |

## 1. Sheet Template (22 mẫu)

| # | Quy tắc |
|---|---|
| T1 | [Template#1] Tủ bếp trên (~10–91) \| Dùng khi: tủ bếp treo, tủ treo tường, tính cả dãy theo md \| Giá theo: md \| Module: Cánh mở, Góc, Ghép hộp tấm, Len nóc, Khung sắt \| Ô nhập: rộng nhập ở B dòng module (header B = SUMIF tổng module "Tủ treo"), header C, D, E \| PK: bas, bản lề, tay nâng, tay nắm \| Gần giống: Template#2, #6; Kitchen "Bếp trên - …" |
| T2 | [Template#2] Tủ bếp dưới (~92–182) \| Dùng khi: tủ bếp thấp cả dãy, có ngăn kéo \| Giá theo: md \| Module: Cánh mở - Ngăn kéo, Góc, Ghép hộp, Len chân \| Ô nhập: rộng ở B dòng module (header B = tổng "Tủ thấp"), C, D, E; SL ngăn kéo ở dòng "Ngăn kéo" \| PK: bas, bản lề, ray RA550, chân nhựa \| Gần giống: Template#1, #3; Kitchen "Bếp dưới - …" |
| T3 | [Template#3] Cabinet (~183–259) \| Dùng khi: tủ đứng lẻ / cabinet có ngăn kéo \| Giá theo: m2 \| Module: Tủ (thùng + 4 NK), Ghép hộp, Len chân \| Ô nhập: header B, C, D, E; E module \| PK: bas, bản lề, ray RA550 \| Gần giống: #2, #4 |
| T4 | [Template#4] Tủ áo (~260–349) \| Dùng khi: tủ áo cao, thùng 25MDF \| Giá theo: m2 \| Module: Tủ áo (4 cánh), Tủ ngăn kéo (3 NK), Ghép hộp, Len chân \| Ô nhập: header B, C, D, E; B/D dòng "TỦ NGĂN KÉO" \| PK: ray RA450, thanh treo TTA + bas, lưới, tay nắm \| Gần giống: #5, #7 |
| T5 | [Template#5] Locker - Mailbox (~350–438) \| Dùng khi: locker, tủ thư có khóa \| Giá theo: m2 \| Module: Tủ cánh mở, Tủ ngăn kéo, Ghép hộp, Len chân \| Ô nhập: header B, C, D, E; D module \| PK: khóa LOCK, TTA, bas \| Gần giống: #4 |
| T6 | [Template#6] Tủ cao - Khung tủ lạnh (~439–513) \| Dùng khi: khung bao tủ lạnh, tủ cao có tủ treo phía trên \| Giá theo: m2 \| Module: Tủ treo trên, Hông tủ cao (2 lộ), Ghép hộp, Len chân \| Ô nhập: header B, C, D, E; D tủ treo trên \| PK: như #1 \| Gần giống: #1, #4; Kitchen "Khung tủ lạnh" |
| T7 | [Template#7] Vanity (~515–595) \| Dùng khi: tủ lavabo \| Giá theo: md \| Module: Tủ, Ngăn kéo, Thùng fix xương sắt, Len chân, Khung sắt \| Ô nhập: header B, C, D, E; E từng module \| PK: mặt đá MTGC \| Gần giống: #3 |
| T8 | [Template#8] Tủ gương hộp (~596–677) \| Dùng khi: tủ gương WC có thùng \| Giá theo: cái \| Module: Tủ, Ghép hộp thùng, Ghép hộp tấm \| Ô nhập: header B, C, D, E \| PK: MTGC \| Gần giống: #9 |
| T9 | [Template#9] Gương phẳng (~678–733) \| Dùng khi: gương treo có viền/khung \| Giá theo: cái \| Module: Tấm nền + viền, Khung xương gỗ, Khung sắt \| Ô nhập: header B, C, D, E \| PK: MTGC \| Gần giống: #8 |
| T10 | [Template#10] Vách, trần phẳng (~735–789) \| Dùng khi: ốp vách/trần phẳng \| Giá theo: m2 \| Module: Vách (tấm 18), Khung xương gỗ, Khung xương sắt \| Ô nhập: header B, C, D, E; K/N dòng vách \| PK: keo 10KCL, MTGC1 \| Gần giống: #11, #13 |
| T11 | [Template#11] Vách ghép hộp (~790–858) \| Dùng khi: vách đầu giường, vách tivi \| Giá theo: m2 \| Module: Hộp xương sắt, Ván bọc vải, Ô trống tivi \| Ô nhập: header B, C, D, E; B/D ô trống \| PK: MTGC1, vải kate \| Gần giống: #10 |
| T12 | [Template#12] Len chân, thanh lam gỗ (~859–899) \| Dùng khi: len chân rời, nan lam \| Giá theo: m2 \| Module: Len chân, Nan gỗ \| Ô nhập: header B (dày), C, D, E; B/C nan \| Gần giống: #10 |
| T13 | [Template#13] Tấm nền sơn lót (~900–942) \| Dùng khi: tấm ván sơn lót / bọc foam \| Giá theo: m2 \| Module: Tấm, Foam bọc \| Ô nhập: header B, C (dày), D, E \| PK: MTGC \| Gần giống: #10 |
| T14 | [Template#14] Bàn (~944–1017) \| Dùng khi: bàn các loại \| Giá theo: md \| Module: Mặt ghép hộp / Mặt thường (chọn 1), Chân 4 chân / Trụ chân (chọn 1), Đỡ mặt sắt \| Ô nhập: header B, C, D, E; E module bật/tắt \| PK: MTGC, SSGC \| Gần giống: #15 |
| T15 | [Template#15] Bench, Armchair, Ghế đôn (~1018–1118) \| Dùng khi: ghế bọc nệm, bench \| Giá theo: md \| Module: Mê ngồi, Tay, Lưng (ghép hộp), Chân gỗ, Đệm, Gối lưng, Gối trang trí, Vải chạy sơ đồ \| Ô nhập: header B, C, D, E; B (cao/dày) từng module, khổ vải C \| Gần giống: #16, #18 |
| T16 | [Template#16] Ghế gỗ solid (~1119–1196) \| Dùng khi: ghế khung gỗ tự nhiên \| Giá theo: md \| Module: Khung ghế solid (chân, giằng, tay, lưng, nan OAK), Phần bọc nệm (foam, vải) \| Ô nhập: header B, C, D, E, K loại gỗ; seat height D dòng VẬT TƯ MỘC \| Gần giống: #15 |
| T17 | [Template#17] Giường (~1197–1277) \| Dùng khi: giường \| Giá theo: cái \| Module: Không ghép hộp / Ghép hộp (chọn 1), Đuôi giường, Vạt giường \| Ô nhập: header B, C, D, E \| PK: MTGC \| Gần giống: #11 |
| T18 | [Template#18] GỐI (~1278–1319) \| Dùng khi: gối rời \| Giá theo: cái \| Module: Gòn, vải, kate \| Ô nhập: header B, C, D, E \| Gần giống: #15 |
| T19 | [Template#19] TRỐNG (~1321–1380) \| Dùng khi: không có mẫu phù hợp, tự dựng từ đầu \| Giá theo: tự đặt \| Module: placeholder, Ghép hộp tấm \| Ô nhập: tất cả |
| T20 | [Template#20] FORM BAO BÌ (~1381–1468) \| Dùng khi: tính đóng gói/thùng carton riêng \| Module: ~55 dòng carton, xốp, PE, plywood \| Ô nhập: kích thước thùng ở B/C/D từng dòng |
| T21 | [Template#21] Tiêu hao ván tự động (~1469–1531) \| Dùng khi: nhập tấm theo kích thước, tự ra tiêu hao ván \| Module: 3 dòng tấm nhập C/D/K/M/N, Ghép hộp tấm \| Ô nhập: C, D, K, M, N các dòng tấm \| Gần giống: #19 |
| T22 | [Template#22] RAPBOM (~1532–1580) \| KHÔNG dùng để chiết tính. Phục vụ RAPBOM sản xuất, dùng tool riêng (chờ cung cấp). |

## 2. Sheet Door (11 mẫu) — quy tắc chọn và kích thước xem skill §3.3

Header Door khác tủ: B rộng, **C dày tường/khung bao (mặc định 110)**, D cao, E SL; **K loại** (EI30/EI60/EI90/Cửa thường), L mã HT, M dày khung bao (tự: cửa thường 30, chống cháy 35), G tự DR/DRF.

| # | Quy tắc |
|---|---|
| D1 | [Door#1] Cửa phẳng 1 cánh - new (~10–92) \| Dùng khi: cửa chống cháy (EI30/EI60), mẫu chuẩn mới nhất \| Giá theo: bộ \| Module: Cánh cửa (3MDF + FR 5ERON + tràm ghép), Lõi Rockwool 80kg/m3, Khung bao (đố, chỉ, rebate) \| Ô nhập: header B, C, D, E, K, L; E "Cánh cửa" = số cánh \| PK: foam, doorseal, ron, SS.01 \| Gần giống: Door#2, #5 |
| D2 | [Door#2] Cửa phẳng 1 cánh (~93–175) \| KHÔNG dùng (bản cũ). Luôn dùng Door#1 "Cửa phẳng 1 cánh - new". |
| D3 | [Door#3] Cửa phẳng 1 cánh - Kính (~176–265) \| Dùng khi: cửa thường có ô kính \| Giá theo: bộ \| Module: Cánh (6MDF + lõi tổ ong 28), Khung bao, Ô kính (C/D, chỉ ASH) \| Ô nhập: header; C/D/E "Ô kính" \| PK: kính 10KCL, A300 \| Gần giống: Door#5 |
| D4 | [Door#4] Cửa phẳng 2 cánh (~266–364) \| Dùng khi: mọi cửa 2 cánh (kể cả 2 cánh bằng nhau); KHÔNG lấy mẫu 1 cánh + E=2 \| Module: 2× Cánh + lõi tổ ong, Khung bao \| Gần giống: Door#3 |
| D5 | [Door#5] Cửa phẳng 2 cánh - Kính (~365–472) \| Dùng khi: cửa chống cháy 2 cánh có kính \| Module: 2× Cánh (Rockwool), Khung bao, 2 Ô kính \| PK: kính EI30, 10KCL \| Gần giống: Door#1, #3 |
| D6 | [Door#6] Cửa phẳng - Ô chớp (~473–559) \| Dùng khi: cửa có ô lam chớp (cửa kỹ thuật, WC) \| Module: Cánh + tổ ong, Khung bao, Ô lam chớp (nan ASH, D/E) \| Gần giống: Door#3 |
| D7 | [Door#7] Cửa panel đố gỗ 1 cánh (~560–642) \| Dùng khi: cửa panel đố ASH \| Module: Cánh (ván nền 6MDF ×4, nẹp, đố 50ASH, tổ ong ×2), Khung bao \| Ô nhập: kích thước đố theo §3.3 \| Gần giống: Door#8 |
| D8 | [Door#8] Cửa gỗ tự nhiên 1 cánh (Có kính) (~644–708) \| Theo §3.3: KHÔNG dùng cho cửa kính thông thường (gỗ Ash, không có bảng hao hụt tấm); chỉ khi cửa gỗ tự nhiên nguyên khối \| Module: Cụm ô kính (C/D/E), đố, chỉ, rebate 25/50ASH |
| D9 | [Door#9] Cửa gỗ tự nhiên 2 cánh (Có kính) (~709–782) \| như Door#8, 2 cánh; header L/M = tỉ lệ 2 cánh (0,667) \| Module: Ô kính cánh 1, Ô kính cánh 2, Cánh cửa |
| D10 | [Door#10] FX.001 (~785–842) \| Dùng khi: ô fix trên cửa, tấm gỗ \| Giá theo: cái \| Module: Tấm hoàn thiện + nền 9MDF, Khung xương thông/tràm, chỉ nẹp 15MDF |
| D11 | [Door#11] FX.002 (~843–903) \| Dùng khi: ô fix trên cửa, kính khung gỗ ASH \| Module: đố 38Ash, rebate, chỉ nẹp; header K/L = số ô kính |

## 3. Sheet Kitchen (24 mẫu) — từng tủ rời, header K = mã nhóm (BT1 bếp trên, BD1 bếp dưới, TL1, TC1, DB1)

| # | Quy tắc |
|---|---|
| K1 | [Kitchen#1] Bếp trên - Tủ 1 cánh (~10–65) \| Dùng khi: tủ bếp trên lẻ 1 cánh \| Giá theo: md \| Module: Cánh mở (Tủ treo), Khung sắt \| Ô nhập: header B, C, D, E \| PK: bas, bản lề, CHW700, tay nâng \| Gần giống: Kitchen#2, #3; Template#1 |
| K2 | [Kitchen#2] Bếp trên - Tủ 2 cánh (~66–121) \| như K1, rộng 900 \| Gần giống: K1, K3 |
| K3 | [Kitchen#3] Bếp trên - Tủ 3 cánh (~122–177) \| như K1, rộng 1200 \| Gần giống: K2 |
| K4 | [Kitchen#4] Bếp trên - Tủ hút mùi (~178–237) \| Dùng khi: tủ trên có thùng che máy hút mùi \| Module: Cánh mở, Thùng fix che hút mùi (B/C), Khung sắt \| Gần giống: K1 |
| K5 | [Kitchen#5] Bếp trên - Tủ góc chữ I (~238–294) \| Dùng khi: tủ góc trên thẳng hàng \| Module: Modun góc (có tấm FIX) \| Gần giống: K6 |
| K6 | [Kitchen#6] Bếp trên - Tủ góc chữ L (~295–352) \| Dùng khi: tủ góc trên chữ L, sâu 600 \| Gần giống: K5 |
| K7 | [Kitchen#7] Bếp trên - Tủ không cánh (~353–407) \| Dùng khi: tủ trên hở, không cánh \| Gần giống: K1 |
| K8 | [Kitchen#8] Bếp dưới - Tủ 1 cánh (~411–464) \| Dùng khi: tủ bếp dưới lẻ 1 cánh \| Giá theo: md \| Module: Cánh mở - Ngăn kéo (Tủ Thấp) \| PK: bas, bản lề, chân nhựa \| Gần giống: K9, K10; Template#2 |
| K9 | [Kitchen#9] Bếp dưới - Tủ 2 cánh (~465–518) \| như K8, rộng 900 |
| K10 | [Kitchen#10] Bếp dưới - Tủ 3 cánh (~519–572) \| như K8, rộng 1200 |
| K11 | [Kitchen#11] Bếp dưới - SINK (~573–626) \| Dùng khi: tủ chậu rửa \| như K8 \| Gần giống: K9 |
| K12 | [Kitchen#12] Bếp dưới - Tủ FULL Ngăn kéo (~628–690) \| Dùng khi: tủ toàn ngăn kéo \| Module: thùng + Ngăn kéo (D cao, E số NK = 3) \| PK: ray RA550 \| Gần giống: K13 |
| K13 | [Kitchen#13] Bếp dưới - Tủ Ngăn kéo + Cánh (~691–753) \| Dùng khi: 1 NK trên + cánh dưới \| Module: thùng + Ngăn kéo (E=1) \| Gần giống: K12, K14 |
| K14 | [Kitchen#14] Bếp dưới - Tủ Ngăn kéo + Ô trống (~754–816) \| Dùng khi: NK + ô hở \| Gần giống: K13 |
| K15 | [Kitchen#15] Bếp dưới - Tủ ô trống + Tấm fix (~817–869) \| Dùng khi: ô hở chờ thiết bị theo thiết kế CĐT (lò nướng, máy rửa chén…) + đợt fix \| Module: thùng, Đợt ngang tấm fixed (E) \| Gần giống: K14 |
| K16 | [Kitchen#16] Bếp dưới - Tủ góc chữ I (~871–927) \| Module: Modun góc, Khung sắt \| Gần giống: K17 |
| K17 | [Kitchen#17] Bếp dưới - Tủ góc chữ L (~928–984) \| sâu 900×900 \| Gần giống: K16 |
| K18 | [Kitchen#18] Khung tủ lạnh (~987–1054) \| Dùng khi: khung tủ cao bao tủ lạnh trong dãy bếp (hoặc tủ khô theo thiết kế CĐT) \| Giá theo: m2 \| Module: Tủ treo trên (D), Hông tủ cao, Ghép hộp \| Gần giống: Template#6 |
| K19 | [Kitchen#19] Tủ Cao (~1055–1130) \| Dùng khi: tủ cao theo thiết kế CĐT (tủ khô, tủ lò…) \| Giá theo: m2 \| Module: Tủ, Ngăn kéo, Ghép hộp, Len chân \| Gần giống: K18, Template#4 |
| K20 | [Kitchen#20] Ốp hông (~1133–1168) \| Dùng khi: tấm ốp hông tủ \| Giá theo: m2 \| Ô nhập: header B dày, C, D, E |
| K21 | [Kitchen#21] Gia giảm nóc (~1169–1204) \| Dùng khi: tấm gia giảm/che khe nóc \| Giá theo: md (theo D) |
| K22 | [Kitchen#22] Gia giảm hông (~1205–1240) \| Dùng khi: tấm gia giảm khe hông \| Giá theo: md (theo D) |
| K23 | [Kitchen#23] Len chân (~1241–1276) \| Dùng khi: len chân dãy bếp \| Giá theo: m2 |
| K24 | [Kitchen#24] Ốp hậu (đối với đảo bếp) (~1277–1312) \| Dùng khi: tấm ốp mặt sau đảo bếp \| Giá theo: m2 |

## 4. Sheet tra cứu / thông số (Chủ đề đề xuất = `TRACUU`)

| # | Quy tắc |
|---|---|
| L1a | [TRACUU] Material – cấu trúc: A mã (khớp cột K của block; chỉ cạnh khớp cột R, sơn khớp T/V), B phân loại, D tên, E ĐVT, F đơn giá, G tổng KL (SUMIF từ Chiet tinh), H thành tiền, I tỉ lệ hao hụt/hệ số. Nhóm theo thứ tự: Vật tư chính (gỗ solid, giá từ GO) → VÁN TRƠN → TẤM HT → BOHO GIA CÔNG NGOÀI → Sơn HT → Sơn lót → VT phụ + bao bì → Gia công ngoài + PK (ray, PK khác, kính GL, gia công OS) → RAPBOM (BOM-1…8) → Chi phí nhà máy (CPXM, CPSOFA, CPSĐG, CPCK, CPBX, CPKE = công; CPDN, CPVPNM, CPKH, CPGC, CPVC = % từ OverView). Không gõ giá cứng vào F khi F đang là công thức. Mã không có trong Material thì block báo #N/A. |
| L1b | [TRACUU] Material – VÁN TRƠN (bảng MDF, PB, WPB, PLW, PB1 xương): mỗi dòng nhập P = độ dày; mã A = dày + loại (17MDF, 9PLW…); khổ M×N 1220×2440; tiêu chuẩn Q và chống ẩm R lấy từ OverView C39/C40; giá S = giá thấp nhất trong Van khớp mô tả × hệ số U (1,02); W = NCC. Cần độ dày mới: thêm dòng trong đúng bảng loại ván. |
| L1c | [TRACUU] Material – TẤM HT: mỗi mã hoàn thiện (WD-xx veneer, LA-xx laminate, ME-xx melamine, PT-xx sơn, FM phủ mặt, VNTAP veneer tạp, BK backer) là 1 cụm: dòng tấm (giá O lấy từ OverView bảng danh mục L33:Q63 theo option giá) + dòng chỉ cạnh cùng màu (mã <dày chỉ><B>, vd 21LA.1, 23VN.1; MOQ, chi phí/pcs). Mã HT mới: khai trong OverView L33:Q63 trước, rồi chép 1 cụm trong Material và đổi mã. |
| L1d | [TRACUU] Material – BOHO GIA CÔNG NGOÀI: ván phủ sẵn, mã <dày><ván><số mặt>M-<mã HT> (vd 17MDF2M-ME-01), giá = giá ván trơn + giá phủ × số mặt. Thêm màu: chép cụm mẫu (ME-01/ME-02), đổi mã (skill §3.6). |
| L1e | [TRACUU] Material – Q1 = ĐẠI TRÀ / HÀNG LẺ / MẪU đổi hệ số hao hụt (Q2 ván 1,25/1,3; Q3 gỗ 1,8/2; foam 1,1/1,2) và giá bao bì; Q1 và ô K chế độ ở dòng nhân công của block (HÀNG LẺ/ĐẠI TRÀ) để nguyên như mẫu, không tự đổi. PK: tên + giá tra theo hãng đã chọn ở OverView D45; cột I = 0,55 khi hãng Hafele là chiết khấu Hafele. Sơn, keo, bao bì có giá công thức theo khối lượng (bậc theo G). |
| L2 | [TRACUU] Nguồn giá (sheet ẩn, Power Query từ SharePoint FactoryCostingShared): GO = gỗ tự nhiên (mã <dày><loại>, giá m3), Van = ván công nghiệp (MDF/HDF/PLW…, khổ 1220×2440, giá tấm), VAN KHAC = cemboard, MgO (ERON), honeycomb…, KINH = kính theo nhà cung cấp, PK = phụ kiện theo hãng (HAFELE/IMUNDEX/BLUM/ROXAN/DTC…). Không sửa tay các sheet này; cập nhật giá bằng Refresh. |
| L3a | [TRACUU] OverView – báo cáo A7:I24 (chỉ đọc): cơ cấu CPSX từ Material (vật tư, nhân công, gia công ngoài, điện nước, VP, khấu hao, vận chuyển, test cửa CC) → Tổng CPSX chưa VAT D24. G = giá mục tiêu (nhập để so sánh). Bảng L7:U14 tổng hợp theo nhóm G của block (CB md, PA m2, LS cái, DR/DRF bộ). |
| L3b | [TRACUU] OverView – thông số dự án (nhập ở đây, không sửa trong block): C29–C32 % điện nước, văn phòng, khấu hao, gia công ngoài; E27 phương án giá hoàn thiện (BASE/OPTION-1…3: cùng 1 spec, vd WD-01, có nhiều phương án giá); C39 tiêu chuẩn formaldehyde (E1), C40 chống ẩm, C41 hệ sơn (PU), C42 tỉ trọng mút (D40), C43 địa điểm giao + E43 loại xe (tra VANCHUYEN, L28 hệ số sử dụng xe), C44 đóng gói (MODULE…; hệ số đóng gói để theo công thức, không sửa), D45 hãng PK tủ, C47–C51 ron chống cháy, ron cách âm, foam nở, vít + tắc kê, keo nẹp khung bao (Có/Không). |
| L3c | [TRACUU] OverView – bảng danh mục vật tư HT L33:Q63: mỗi mã HT của dự án (LA-xx, WD-xx…) một dòng, giá theo cột BASE / OPTION-1 / OPTION-2 / OPTION-3. Material TẤM HT lấy giá ở đây. Đổi giá HT: sửa ở bảng này. |
| L3d | [TRACUU] OverView – CHI PHÍ TEST CỬA CC C21:C23: nhập mã <cấp EI><loại cửa> (= cột P dòng đầu block Door, vd EI601); tự lấy chi phí test + ra giấy từ PCCC và giá cửa mẫu (max đơn giá cửa đó trong Chiet tinh). A65+: nhật ký REV điều chỉnh giá (nội dung, ngày, trước/sau, chênh lệch). |
| L4 | [TRACUU] Spec: mỗi hệ bếp một nhóm 7 dòng, mã hệ ở B (BT1, BT2 bếp trên; BD1 bếp dưới; TL1 khung tủ lạnh; TC1 tủ cao; DB1 đảo bếp). Cột = cấu phần (thùng, hậu, thùng NK, thùng sink, cánh, tủ không cánh, ốp hông, gia giảm, len chân, ốp hậu, tấm đỡ đá, PK treo, tay nắm, ray). Nhập: ván nền, hoàn thiện, code HT, độ dày, số đợt ngang; dòng cuối nhóm tự sinh mã vật tư. Block Kitchen lấy mã theo header K = mã hệ. Thêm hệ mới: chép 1 nhóm 7 dòng, đổi mã hệ. |
| L5 | [TRACUU] Matrix: bóc module theo hệ bếp. Mỗi bảng (Tủ bếp trên / dưới / Khung tủ lạnh) có dòng rộng module (vd 300/400/450/600/800), dòng tổng SL, dòng đơn giá module; mỗi dòng sau là 1 hệ (B mã, C SL, D chiều dài, E số cabinet, F thành tiền, G đơn giá/md, các cột = số module theo rộng). Không phải BoQ CĐT. Dùng khi BoQ có nhiều hệ bếp (MAU-01). Dòng đơn giá module hiện là số tạm, sẽ link giá sau; F/G đang #REF!/#VALUE! là chờ link, không tự sửa. |
| L6 | [TRACUU] Dinhmuc: định mức nhân công/gia công: cabinet (công/m2 theo số tấm), vách phẳng/ghép hộp, nan gỗ, sơn (PU, pigment, lót), cửa thường/chống cháy/gỗ solid (công/bộ theo MOQ), ô kính, lam chớp, đóng gói. Block tra định mức từ đây; không sửa trong block. |
| L7 | [TRACUU] PCCC (ẩn): chi phí test + ra giấy cửa chống cháy theo cấp EI30–EI120 (cộng ở OverView mục CHI PHÍ TEST CỬA CC). VANCHUYEN (ẩn): cước theo điểm giao và loại xe, OverView C43/E43 chọn. |
| L8 | [TRACUU] Sheet hệ thống, không sửa: Linhtinh (danh sách tên block cho dropdown + thông báo VBA), _LIB, _RULES, _RULES_MASTER, AI_INPUT (Power Query nạp), sheet backup/snapshot (Door_BK_truocsua, Material_Snapshot, Auto_Snapshot_*). |

## 5. Rule sửa đổi đi kèm

| # | Chủ đề | Quy tắc |
|---|---|---|
| X1 | DOOR | Cửa 2 cánh (kể cả 2 cánh bằng nhau) luôn dùng mẫu Door "Cửa phẳng 2 cánh" (hoặc "Cửa phẳng 2 cánh - Kính" khi có kính/chống cháy). Không lấy mẫu 1 cánh + E "Cánh cửa"=2. Thay thế skill §3.3a (cần sửa skill-src lần build sau). |
| X1-ok | | (X1 đã ghi R0134; X2 bỏ vì trùng Door#2) |
| X3 | CHUNG | Nhân công trong block (B số SP, C số người, D số ngày, mã K) giữ nguyên như mẫu, không tự chỉnh. Việc của model khi bóc/chiết tính là đảm bảo KHỐI LƯỢNG VẬT TƯ đúng (kích thước, SL, mã vật tư, số mặt hoàn thiện, dán cạnh). |
