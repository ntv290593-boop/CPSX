# KITCHEN — Khi nào dùng, Spec, Matrix (đọc trước các file KITCHEN-*)

Nguồn: sheet Kitchen, Spec, Matrix; R0076, R0110–R0133, R0145, R0146. Đọc công thức 2026-10-24.

## 1. Quy tắc dùng (user chốt)
**Chỉ dùng sheet Kitchen khi làm MA TRẬN ở sheet Matrix** (BoQ có nhiều hệ bếp lặp lại, mỗi module tính bằng mẫu Kitchen). Bếp bình thường (không lặp hệ) → dùng Template#1 Tủ bếp trên / #2 Tủ bếp dưới (MODUN-CANH-MO…). Sheet `Matrix`, `Spec` **không phải BoQ CĐT**.

## 2. Spec: cấu hình vật liệu theo hệ (mỗi hệ một nhóm 7 dòng)
Hệ: **BT1, BT2** (bếp trên), **BD1** (bếp dưới), **TL1** (khung tủ lạnh), **TC1** (tủ cao), **DB1** (đảo bếp). Nhóm gồm: tên cấu phần / **Ván nền** / **Hoàn thiện** (Laminate | Melamine | Acrylic) / **Code HT** / **Độ dày** / **Số đợt ngang** (chỉ thùng) / **dòng khóa** (mã hệ ở cột A).
Dòng khóa sinh mã vật tư: `Melamine → dày&ván&"2M-"&Code` (vd 17MDF2M-ME-01); `Acrylic → dày&ván&"1M-"&Code&"BK"`; còn lại (Laminate/…) `dày&ván` (17MDF), hoàn thiện đi bằng cột N (Code HT).
Cột Spec (đánh số từ A=1): D thùng tủ, E hậu tủ, F thùng ngăn kéo, G thùng sink, H cánh, I tủ không cánh, J ốp hông, K gia giảm nóc, L gia giảm hông, M len chân, N ốp hậu, O tấm đỡ đá, **T PK treo tủ** (CHW700/Luoigaz), **U tay nắm** (TN.01/không), **V ray** (RA/…). Thêm hệ mới: chép 1 nhóm 7 dòng, đổi mã hệ (R0145).

## 3. Cách block Kitchen lấy dữ liệu từ Spec
Ô `K` (mã vật tư) của mỗi chi tiết = `VLOOKUP(BL<hàng>, Spec!A2:T330, <chỉ số cột>, 0)`; `N` (code HT) = `VLOOKUP(BL, Spec!B2:T330, …)`; số đợt ngang `E = E module × VLOOKUP(…,3,0)`; PK treo tủ `K = VLOOKUP(BL, Spec!A2:Z330, 20)`, tay nắm `…21`. **`BL` = mã hệ**, đặt ở `K` của dòng tiêu đề block (BT1…) và chuyền xuống mọi hàng bằng `INDIRECT("BL"&ROW−1)` (giữ nguyên, R0069). → Đổi vật liệu/hoàn thiện cả dự án **chỉ sửa ở Spec**, không sửa block.

## 4. Matrix: bóc module theo hệ bếp (R0146)
Ba bảng: **Tủ bếp trên**, **Tủ bếp dưới**, **Khung tủ lạnh**. Mỗi bảng:
- Hàng 2: bề rộng module theo cột (300/400/450/600/800; ốp hông 20, gia giảm hông 50, gia giảm nóc "md").
- Hàng 3: **tổng SL toàn dự án** = `SUMPRODUCT(cột, SL hệ)`.
- Hàng 4: **đơn giá một module** (đang là số tạm 2.000.000, H4 `#REF!`; sẽ link `'Chiet tinh'!H<block Kitchen>`).
- Mỗi dòng hệ (mã kiểu `KT-1612-C` của BoQ): `C` SL, `D = SUMPRODUCT(rộng, số module)` (chiều dài mm), `E = SUBTOTAL(số cabinet)`, `F = SUMPRODUCT(đơn giá module, số module)`, `G = F/D×10³` (đơn giá/md), cột cuối = tổng md gia giảm nóc.
Block Kitchen chính là **bảng giá của một module ở một bề rộng**; Matrix nhân số lượng module của từng hệ.

## 5. Điểm riêng của Kitchen so với Template
Dùng lại logic module (MODUN-CANH-MO, MODUN-GOC, NGAN-KEO, LEN-NOC-CHAN, PHU-KIEN-TU) nhưng: **mã vật liệu/hoàn thiện lấy từ Spec**; số cánh/ngăn kéo **cố định theo từng block** (không để E tự tính); cao tủ trên 750 (Template 1000). Header `M`: md | m2; nhóm `G = CB`.

## 6. Đã chốt (user 2026-10-24)
- Đơn giá module ở Matrix hàng 4 là **số tạm**; sau này lấy **theo giá của hệ, cột H của Chiet tinh** (`'Chiet tinh'!H<dòng đầu block>`), không gõ tay.
- Thùng sink lấy vật liệu theo **Spec cột G** (đã sửa mẫu, xem KITCHEN-BEP-DUOI).

## 7. Chưa xác nhận
- Matrix: nhiều bề rộng 300/400/450 dùng chung một block hay copy nhiều block (mỗi block một bề rộng) khi link giá.
- Block Kitchen chưa đối chiếu chi tiết từng hàng: Tủ cao, Khung tủ lạnh (xem KITCHEN-TU-CAO-PHU).
