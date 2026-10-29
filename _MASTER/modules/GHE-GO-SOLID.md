# GHẾ GỖ SOLID (Template#16)

Nguồn: Template "Ghế gỗ solid" hàng 1119–1196. Giá theo **md**. Hai module: KHUNG GHẾ - SOLID (gỗ tự nhiên) và PHẦN BỌC NỆM (ván nền + foam + vải).

## 1. Ô nhập
Header: `B,C,D` = rộng, sâu, cao tổng; `J` = ván nền (PLW); `K` = **loại gỗ** (OAK, ASH, …); `D` hàng 1120 = **SEAT HEIGHT** (450). Mỗi module `E` = 1.

## 2. Khung ghế (chi tiết, mặc định)
Chân trước/sau 45×45 (`D = seat height − dày đệm`, `E=2E`); Khung giằng trước/sau 30×50 (`D = B − 2×chân`); giằng hông (`D = C − 2×chân`, `E=2E`); Ke góc ashc 30×100×100; Tay vịn 30×50 (`D=C`); Lưng tựa 20×200; Nan 20×25 (`D = cao − seat height`); **Mê ngồi hiển thị** 17MDF (`C = C−dày giằng×2`, `D = B−2×dày giằng hông`).

## 3. CƠ CHẾ TỰ CHỌN MÃ GỖ (cột K của từng chi tiết gỗ)
`K = LET(Go = UPPER(loại gỗ), Day = MIN(B,C,D), tìm các mã trong Material kết thúc bằng tên gỗ (vd 38OAK, 50OAK), Max = dày lớn nhất có sẵn; nếu không có → trả tên gỗ gốc; còn lại: Can = Day + 5 (phôi dư 5 mm); nếu Can > Max → số lớp ghép Lop = ROUNDUP(Can/Max); Tar = Day/Lop + 5; mã = MIN(dày ≥ Tar) & tên gỗ)`.
Ví dụ: chân 45 mm gỗ OAK → 50OAK. **Đổi loại gỗ chỉ cần sửa K header**; chi tiết tự chọn dày phôi từ Material.

## 4. Phần bọc nệm
Nền mê ngồi 18 (K = J header), Cụm lưng/Cụm tay 18; **Nền lưng / Nền tay** ván 3 mm ép cong: `E = dày cụm / 3 × E cụm` (18/3 = 6 lớp). Foam: đệm 70, nền lưng/tay 50; vải = foam +100; gối lưng rời (`fd` 120×250×250), gối trang trí; vải chạy sơ đồ khổ 1370 + kate (BOC-VAI-FOAM). Mê fixed / rời như SOFA-BENCH.

## 5. Rule
R0092 (mẫu), R0065 (cut list gỗ ASH: dày phôi = mã phách), R0074, R0150.

## 6. Chưa xác nhận
Cột `AN` (tên gỗ) lấy từ K header qua công thức chung; chưa đọc AN của từng dòng. Mã gỗ cần có trong Material theo dạng `<dày><TÊN>` (vd 38OAK); thiếu → trả tên gỗ trơn và báo #N/A.
