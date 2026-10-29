# 01 — Quy trình tính giá gói cửa

| Bước | Việc | Rule |
|---|---|---|
| 1 | Quét hồ sơ: BoQ CĐT làm xương sống, gắn BV/Spec/ML từng dòng, ghi mã VL/HT trên từng hạng mục | — |
| 2 | Đọc `_RULES_MASTER` + sheet Door, hiểu block mẫu TRƯỚC khi chia BV | R0099–R0109 |
| 2b | **Check mặt bằng** từng vị trí cửa: dày tường (CD.66 + CD.06 + CD.00.07) → ghi sheet DAY_TUONG file scan | R0166 |
| 2c | **Đọc ghi chú loại cửa + HÌNH chi tiết cấu tạo** (render trang, crop, không chỉ text): lập bảng lớp cánh/khung bao/nẹp chặn/chỉ – mỗi lớp vật liệu + dày + số mặt; so BV ↔ BoQ ↔ block | R0167 |
| 3 | Map từng loại cửa → block; lập phiếu bóc (ô cần đổi, giá trị, lý do, link BV) | `02_`, `03_` |
| 4 | Tạo file CPSX, import BoQ CĐT (sheet bên trái), chép block, link đơn giá + mô tả | R0164, R0158, R0023, R0160 |
| 5 | Vận chuyển OverView C43/E43; test cửa chống cháy C21 | R0144 |
| 6 | Kiểm Material G4/H4 | R0048, R0161 |
| 7 | Lệch hồ sơ → Issue → RFI (bỏ cột nội bộ khi gửi CĐT) | R0163, `05_` |
| 8 | User duyệt → bỏ tô vàng | R0024, R0051 |

Phạm vi mặc định: chỉ cánh + khung + nẹp; **không** phụ kiện trừ khi được yêu cầu. Giá v1 = giá xuất xưởng; giữ nhân công mẫu (R0150).
