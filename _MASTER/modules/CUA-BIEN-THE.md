# CỬA — Các biến thể block (khác Door#1)

Nguồn: Door hàng 176–265 (#3), 266–364 (#4), 473–559 (#6), 560–642 (#7), 785–842 (#10), 843–903 (#11). Chung: header, khung bao, phụ trợ như CUA-CANH, CUA-KHUNG-BAO, CUA-PHU-TRO.

## Door#4 Cửa phẳng 2 cánh (B=1200)
Nhân đôi module "Cánh cửa" (hàng 268 và 278): cánh 1 rộng `(B−(đố+FR+3MDF)×2−6)/2`, cánh 2 = `… − rộng cánh 1`. **Không** dùng mẫu 1 cánh + E=2 (R0134). Khung bao, rebate, chỉ cửa như #1.

## Door#3 Cửa phẳng 1 cánh - Kính (thường)
Thêm "Ô kính": `B`=dày cánh, `C,D` gõ (mặc định 300×900), `E`=số ô; chỉ đứng/ngang `25ASH` 20×40 (`D+30`, 2 mỗi cạnh). Kính `10KCL` (thường) hoặc mã kính EI (K248 = K loại, khi EI) = C×D×E; `A300`: `SUMIF("Ô kính",F)/10`. **Lớp da, lõi, khung xương KHÔNG trừ diện tích ô kính** (E lớp da = `2×E cánh`): user chốt theo công thức mẫu, giữ nguyên (không tự thêm hàng trừ).

## Door#5 Cửa phẳng 2 cánh - Kính (EI30, B=1200)
Cấu trúc = #4 (hai module "Cánh cửa", cánh 1 `…/2`, cánh 2 `… − rộng cánh 1`) + **2 ô kính, mỗi cánh một ô**: mỗi ô `C=150`, `D=1500` (gõ), `E=1`, kèm chỉ đứng/ngang 25ASH 20×50 (`D+30`, 2 mỗi cạnh). Kính: `10KCL` (cửa thường) hoặc mã kính theo K (vd EI30, khi không phải thường), SL = số ô. Khác #1 ở **rebate ngang**: `D = C cánh 2 + C cánh 1 + 30` (tổng rộng 2 cánh) và `C rebate = C khung bao − dày cánh` (110). Doorseal: `(D rebate đứng×E + D rebate ngang×E + D cánh×(E cánh 1 + E cánh 2 − 1))×1,05/1000`. Ron chống cháy cộng chu vi của cả hai cánh `/2100`.

## Door#6 Cửa phẳng - Ô chớp
Cánh + khung bao như #1; thêm **Ô LAM CHỚP**: `B = C cánh − 200`, `C` = dày cánh, `D`=450 (gõ), `E`. Nan 15×40 25ASH `E = ROUND((D − 2×khung)/30)` (bước 30); khung ngoài đứng/ngang 15×40 25ASH.

## Door#7 Cửa panel đố gỗ
Cánh gồm **panel ván nền 6mdf ×4**, **nẹp chặn panel** 10×15 25ASH (×8 mỗi hướng), **đố đứng bản lề** (C=145), **đố đứng khóa** (C=140), **đố ngang ×3** (trên/giữa/dưới), tất cả 50ASH dày = dày cánh; lõi tổ ong ×2. Khung bao như #1. Theo §3.3 dùng cho cửa panel đố ASH.

## Door#8, #9 Cửa gỗ tự nhiên 1 / 2 cánh (có kính, gỗ ASH)
Chỉ dùng khi cửa **gỗ nguyên khối** (R0106). Không có lớp da/lõi/FR: toàn bộ là gỗ ASH, nhóm `G = DR`. Mặc định B=900/1200, C=230, D=2400.
- **Cụm ô kính cánh** (gõ `C`, `D`; #8: 430×1935, #9: cánh 1 400×1935, cánh 2 170×1935), `E`=số ô (số ô kính tính theo cột đố giữa). Nẹp kính đứng ×2, ngang ×2: 25ash 15×15.
- **Cánh cửa:** dày 45; rộng `B − 2×dày đố đứng khung bao` (35); cao `D − dày đố ngang khung bao`. #9: cánh 1 rộng `(B−2×35)×M`, cánh 2 `(B−2×35)×(1−M)`, **M = tỉ lệ 2 cánh (header, 0,667)**, hiện "Check tỉ lệ cửa" nếu M không hợp lệ.
- **Đố đứng cánh** 50ash: `C = (C cánh − C ô kính)/2`, ×2; **Đố ngang** `C = (D cánh − D ô kính×E − C đố giữa×E giữa)/2`, ×2; **Đố ngang giữa** C=150, `E = số ô − 1`.
- **Khung bao (cụm):** Đố đứng 35 (50ash, C=sâu, ×2), đố ngang; Chỉ cửa đứng 15×50 25ash (`D = D khung + C chỉ ngang`, ×4), chỉ cửa ngang (`D = B + 2×C chỉ`, ×2); Rebate đứng/ngang 15 25ash, rộng `C khung bao − dày cánh` (185), rebate ngang `D = B − 2×35`.
- Chưa có phụ kiện keo/ron riêng trong phần đã đọc; theo mục KHÁC và Gia công ngoài của block.

## Door#10 FX.001 (ô fix tấm gỗ)
Tấm hoàn thiện 9MDF ×2 mặt, tấm nền 9MDF; khung xương gỗ: xương thông 30×30 bước **300** (ngang/dọc/đứng); khung gỗ giáp tường 40 tràm ghép (ngang ×1, dọc ×2); chỉ nẹp 15mdf rộng 50 (ngang ×2, dọc ×4).

## Door#11 FX.002 (ô fix kính khung ASH)
Header `K` = "Số ô kính" (`L` = số ô). Đố đứng/ngang 38Ash 30; rebate 25Ash 15 (`C−50`); **đố giữa** `E = L−1`; cụm ô kính `C=(B−2×đố−đố giữa×E)/L`; chỉ nẹp quanh ô 25Ash 15×15 ×2; chỉ nẹp khung bao 15×50 ×4.

## Chưa xác nhận
Door#8/#9: phần Gia công ngoài + phụ kiện (kính, keo, ron) chưa đọc; hạng mục cửa gỗ tự nhiên nên đối chiếu thêm khi có ca thật.
