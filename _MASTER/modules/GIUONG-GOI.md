# GIƯỜNG (Template#17) và GỐI (Template#18)

Nguồn: Template hàng 1197–1277 (giường), 1278–1319 (gối). Giá theo **cái**.

## Giường
Header: `B` rộng 1800, `C` dài 2200, `D` cao 350 (đầu giường), `E` SL.
**Chọn 1 trong 2** module (E nhập tay: 1 = bật, 0 = bỏ):
- **Giường không ghép hộp:** Đầu giường, Đuôi giường (25mdf, `C=cao`, `D=B`, E=E), Vai giường (`D=C`, `E=2E`).
- **Giường - ghép hộp:** Đầu / Đuôi / Vai là 3 module **GHÉP HỘP** (dày 40, xương gỗ, đúng cơ chế GHEP-HOP, kích thước 1820/2220 +20): đầu và đuôi `D=B`, vai `D=C`.
Chung cho cả hai: **Vạt giường** 17MDF (`C=B`, `D=C`, E=1); **Bọ đỡ vạt** thông 30×30 chu vi `(B+C)×2` và 25×95×80 (`E=2×E thanh dọc`).
**Vật tư cơ khí:** Thanh dọc đỡ vạt `MT.40401.2` (`D=C`, `E=ROUNDUP(B/850)−1`), Thanh ngang đỡ vạt (`*0` tắt), Chân đứng `MT.30301.2` (`D=cao`, `E=2×số thanh dọc`). **Kim loại gia công ngoài:** Chân sắt MTGC 30×30×1,2 hộp (xem KIM-LOAI-GIA-CONG).
Vật tư phụ: VTPK `B=kd(…)` (UDF, R0068), đóng gói DG.3 (`KT thùng KD`), KEOCANH. Bọc nệm (foam/vải) để sẵn.

## Gối (rời)
Header `B,C,D` = 600×150×450, K = `GOI`. Module "CABINET CÓ CÁNH CỬA" (E=1): **Gòn tơi** `GONTOI` (`B,C,D` = header, E); Vải `C=B+100`, `D=D+100`, `E = 2×E gòn`; vải lót `KATE`; đóng gói `DG.V`; KEOCANH.

## Rule
R0093 (Giường), R0094 (Gối), R0074.

## Takeoff cần khai
Giường: kích thước (R × D × cao đầu), ghép hộp hay không (và dày), vật liệu đầu giường (25mdf hay ghép hộp), số thanh đỡ/chân sắt, bọc nệm (nếu có), mã hoàn thiện. Gối: kích thước, ruột (gòn tơi/foam), vải.

## Chưa xác nhận
Giường: block không có sẵn dòng đầu giường bọc nệm (foam/vải để trống); hạng mục giường bọc phải bổ sung dòng theo R0039. Cần user cho ví dụ giường bọc nếu có.
