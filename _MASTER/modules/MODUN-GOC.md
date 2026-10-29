# MODUN GÓC

Nguồn: Template#1 hàng 23–34 (tủ treo), Template#2 hàng 113–125; Kitchen "Tủ góc chữ I/L" (BT/BD).

## 1. Ô nhập
Như MODUN-CANH-MO: `B` rộng (gõ), `C`, `D` theo header, `E` mặc định **0** (bật khi hạng mục có góc), `K`, `L`, `M`.

## 2. Khác với MODUN CÁNH MỞ
Các dòng Nóc, Đáy, Hông lộ/khuất, Hậu, Đợt ngang, Chặn cửa, Bổ gia cố, Sắn giữa: **cùng công thức**, chỉ đổi tham chiếu sang dòng module góc.
- **Tấm FIX** (thêm): rộng = C, dài = D cánh, SL = E, 17MDF. Là tấm che phần góc, cũng là khóa tính cánh.
- **Cánh:** trừ độ rộng tấm FIX: rộng 1 cánh = `(B − C Tấm FIX)/E cánh×E`, SL = `ROUNDUP((B − C Tấm FIX)/600)×E`.
- **Sắn giữa:** D = `D − dày nóc − dày đáy` (tủ treo) như cánh mở.

## 3. Công thức chung và phụ kiện
Xem CONG-THUC-DONG. Phụ kiện góc tính cùng bảng với cánh mở (SUMIF theo tên "cánh tủ", "Đợt ngang"…): xem PHU-KIEN-TU.

## 4. Chưa xác nhận
Kitchen chữ L (sâu 600/900) và chữ I có thể đổi cách tính Tấm FIX và cánh; chưa đối chiếu công thức từng block. Cần đọc Kitchen#5, #6, #16, #17 trước khi dùng cho takeoff.
