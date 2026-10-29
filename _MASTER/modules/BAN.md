# BÀN (Template#14, "LOOSE" / Bàn)

Nguồn: Template hàng 943–1017. Giá theo **md** (`N = H/B×10³`, B = rộng bàn mm).

## 1. Ô nhập
Header: `B` rộng, `C` sâu, `D` cao, `E` SL. Mỗi module có **E nhập tay: 0 = bỏ, 1 = bật**. Chọn **1 trong 2** nhóm mặt và **1 trong 2** nhóm chân (bật cả hai là tính trùng).

## 2. Module mặt bàn (chọn 1)
- **MẶT BÀN GHÉP HỘP:** K = XƯƠNG GỖ, dày 40: dùng đúng cơ chế GHEP-HOP (tấm mặt 12MDF ×2, viền khi dày ≥ 51, xương THONG1). Thêm **Viền gỗ hiển thị**: 50oak, `C=100`, `D=(C mặt+D mặt)×2`, E = E module.
- **MẶT BÀN KHÔNG GHÉP HỘP:** Tấm mặt 25MDF `C×D = C×B` của header, E = E module.

## 3. Module chân bàn (chọn 1)
- **4 CHÂN:** Chân bàn 50×50 50oak (`D=cao`, `E=4×E`); Giằng dọc 30×50 38oak (`D=B`, `E=2×E`); Giằng ngang 30×50 (`D=C−2×dày giằng`, `E=2×E`); Xương ngang 30×50 ashc (`E=ROUND(B/600−1)×E`); Ke góc 30×120×120 ashc (`E = E xương + E giằng ngang`).
- **TRỤ CHÂN:** Chân bàn gỗ đặc 50oak (B×B×cao) hoặc **Chân bàn Trụ - ghép ván** (ván nền 25MDF `C=B+20`, `D=cao+20`, `E=ROUNDUP(C/dày)×E`; tấm hoàn thiện 3mdf `C=B×3,14/E+100`, `E=ROUNDUP(B×3,14/1220)×E`).

## 4. Phụ trợ (gia công ngoài + cơ khí)
Tấm đỡ: **Đỡ mặt bàn** (`MT.TAM5`, 700×700), **Tăng nặng** (`MT.TAM10`). Kim loại gia công ngoài (MTGC/SSGC, xem KIM-LOAI-GIA-CONG): Đỡ mặt bàn `0,8×B × 0,8×C`, Tăng nặng `0,7×B × 0,7×C`, Trụ nối tấm đỡ + đế (Ống Ø60, M=1,4), Tấm/Viền ốp đế tròn giả dày (SSGC, "nếu có"). Kính `10KCL`, `SS.01`, `TI.01` (tiện) để sẵn.

## 5. Rule
R0072 (trụ ghép-tiện: khi E trụ > 1 đổi C = `B×3,14/ROUNDUP(B×3,14/1220)+100`), R0073 (mã gia công/tiện phải có trong Material: TI.03, TI.04, CNC.01), R0074.

## 6. Takeoff cần khai
Rộng × sâu × cao, mặt ghép hộp hay không (dày, vật liệu), kiểu chân (4 chân / trụ / gỗ đặc), loại gỗ, có đỡ mặt kim loại / kính / inox, số lượng.

## 7. Chưa xác nhận
Chưa thấy công thức chặn bật đồng thời hai module cùng nhóm (nhập E đúng là việc của người bóc).
