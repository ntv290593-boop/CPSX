# Nhận dạng mẫu — loại đồ vật → mẫu ghép (NHÁP 2026-10-24, từ SOFITEL RE00)

Chọn mẫu gồm 2 bước: (1) nhận LOẠI đồ vật từ bản vẽ + dòng BoQ; (2) tra quy ước "loại này ghép bằng mẫu nào". Nhãn Mau trong file đã làm là lựa chọn của người bóc, không thuần theo hình dạng. **Sheet Kitchen chỉ có hiệu lực khi user cho phép dùng Matrix (R0076): không bao giờ gán Kitchen cho hạng mục thường.** (Bản nháp đầu viết tủ áo/minibar/bedside = Kitchen#19: SAI, do script trích đoán mẫu theo chuỗi nhãn và Kitchen thắng Template!Cabinet với biên 0.01–0.03; đã sửa `extract_ai_input.py` loại Kitchen mặc định.) Quy ước dưới đây lấy từ SOFITEL RE00; dự án khác kiểm lại.

> **CẢNH BÁO (user phát hiện 2026-10-24): cột "Mẫu đã dùng" ở bảng §1 KHÔNG ĐÁNG TIN.** Nó lấy từ script đoán mẫu theo chuỗi nhãn, mà nhãn module ("VÁCH, TRẦN", "VẬT TƯ MỘC", "Cabinet có cánh") dùng chung giữa nhiều mẫu và user hay đổi tên. Bằng chứng: len chân F012 (nhãn `VÁCH, TRẦN` → `Len chân` → `Lưỡi gà`) bị gán "Vách, trần phẳng" (0.863) thay vì "Len chân, thanh lam gỗ" (0.727); trim F011 và nan gỗ F009 user nói là TRỐNG nhưng bị gán Vách phẳng/Gương phẳng. Chỉ dùng các quy tắc do user xác nhận ở §0.

## 0. Quy tắc user đã xác nhận (nguồn: cột cmt, BO_SUNG_RE00_v2, 2026-10-24)
| Loại | Mẫu | Ghi chú |
|---|---|---|
| Len chân (WD-01/WD-03 80, 100) | Template "Len chân, thanh lam gỗ" | F007, F008 chép từ block len chân đã làm (F012, F020) |
| Chỉ gỗ / trim / nan gỗ / thanh nan | Template TRỐNG | F023, F026.1; nguồn chép F011 |
| Vách nan gỗ che điều hòa | Template TRỐNG | F024, chép từ F009 |
| Cửa (kể cả cửa swing) | Sheet Door | không dùng Template vách (F005) |
| Tủ áo | Template "Tủ áo" | không dùng vách (F065) |
| Hạng mục chưa có bản vẽ đã đọc | không gợi ý mẫu | F028: phải mở bản vẽ trước |

### 0b. User điền theo loại (`_ACTIVE\SOFITEL\02_takeoff\LOAI_DO_VAT_SOFITEL.xlsx`, 2026-10-24)
Loại do Claude gán theo tên BoQ + ĐVT; mẫu do user điền. Đây là đáp án loại → mẫu đầu tiên đáng tin.

| Loại đồ vật | Mẫu |
|---|---|
| Ốp veneer gỗ lên tường (m2) · Ốp trang trí cột · Ván nền dán giấy · Ván nền ốp gương/kim loại | Template "Vách, trần phẳng" |
| Vách sơn / vách CNC · Gương ốp có cánh mở · Vách góc (corner screen) · Vách gỗ đứng/fins portal · Vanity · Gương rời | Template "TRỐNG" |
| Vách TV / TV panel có hốc · Đầu giường · Đầu giường + vách góc | Template "Vách ghép hộp ( Vách đầu giường, vách tivi..)" |
| Hốc ủi · Minibar cao (1900) · Minibar thấp (900) · Bedside liền tường · Tủ phòng khách + kệ | Template "Cabinet" |
| Bàn rời | Template "Bàn" |
| Ghế bành/lounge/armchair · Đôn/stool/ottoman · Băng ghế | Template "Bench, Armchair, Ghế đôn" |
| Hốc tường nhà tắm · Khung bồn tắm/ghế tắm · Cửa có sẵn làm lại mặt | (user để trống — chờ phạm vi) |

## 1. Bảng quy ước (số lần dùng: số liệu REV01 cũ, còn kiểm lại trên RE00)
| Loại đồ vật (tên BoQ) | Mẫu đã dùng | Ghi chú |
|---|---|---|
| Len chân, trim, nẹp, vách/trần phẳng, ván nền dán giấy/gương | Template#10 | cả m, m² |
| TV panel, đầu giường (headboard), vách có ô tivi/hộp âm | Template#11 (5/5) | chia hộp (VACH-GHEP-HOP) |
| Corner screen | #10 (6) hoặc #11 (2) | #11 khi có hộp |
| Gương, bàn phụ nhỏ (D<600) | Template#9 | bàn phụ nhỏ dùng tấm phẳng |
| Bàn, console, coffee table | Template#14 | console L014 dùng #18 GỐI = nghi sai |
| Sofa, bench, ottoman | Template#15 | L-sofa, sofa bed |
| Ghế, pedestal | Template#16 | pedestal vẫn dùng ghế solid |
| Tủ áo, bedside, dressing table, minibar tủ cao | Template!Cabinet (Template#3) | trích lại không Kitchen: khớp 0.78–0.93, ít ô đè số hơn |
| Minibar thấp F041 | Template!Tủ bếp dưới (Template#2) | khớp 0.74, chưa chắc, cần kiểm |
| Vanity counter, riser door, bedside L022–L025 | Template#19 TRỐNG | tự ghép |
| Cửa các loại (1 cánh, 2 cánh, khung ô mở, cửa thông phòng) | "Cửa lùa gỗ tràm 1 cánh" (10/11) | thư viện REV01 ≠ RE00 |

## 2. Cặp dễ nhầm và cách phân biệt
- **Wardrobe có/không ngăn ủi**: block REV01 giống hệt nhau (F040 = F066). Phân biệt bằng mặt đứng: ngách ủi rộng 250 ốp đá + 4 ngăn 520 (PS02A) so với hai cột 645 có ngăn két (PS02).
- **Kệ WD-04 + hộp đèn trên (cao 290)** của tủ áo Prestige: là tấm ngang dày 40 ở cao độ 1810; mẫu không có dòng sẵn.
- **Minibar type 2 (1900) vs type 3 (900)**: cùng tên, khác bộ tờ (713/714 so với 715/716).
- **Nhận loại bằng số tờ ở khung tên**, không bằng kích thước (F040 và F066 từng bị nhầm).

## 3. Cấu tạo không module nào có cơ chế → cách bóc của user
- Hông bo tròn R80 (minibar type 3): tính **GHÉP HỘP ×2** (hai đầu 105), không dùng hệ số "cong".
- Cụm ngăn chia nhỏ trong hộc: tính qua **ngăn kéo thêm** trong thùng, không thêm dòng tấm lẻ.
- Chia thùng: thùng chứa 3 ngăn kéo + thùng 3 ngăn kéo + thùng tủ 2 cánh (thêm ngăn kéo bên trong) + len chân gỗ.
- Ca thật F074 (REV01, hàng 5626; mẫu thật nhiều khả năng Template!Cabinet, không phải Kitchen): TỦ 900 có 2 cánh + 4 module ngăn kéo + TỦ 250 cánh kính + GHÉP HỘP B=80 + LEN CHÂN 50 (khung xương gỗ). Đã bóc sẵn bằng tay, user chốt GIỮ NGUYÊN (2026-10-24). Chú ý: `AI_MODULE` chỉ liệt kê ô nhập; dòng thêm/bớt nằm ở `NGOAI_EDITS`, nên không kết luận "thiếu" từ AI_MODULE.

## 4. Phạm vi (theo dự án, ghi ở PROJECT)
Mặc định tính mọi vật liệu mình gia công, kể cả kim loại. SOFITEL bỏ kim loại toàn dự án, đá ST, giấy WC, OS&E, đèn, kính GL; tính gỗ ván + ván nền + sơn trên ván.

## 5. Cần xác nhận
Template#19 TRỐNG cho vanity/riser door (chưa xác nhận lý do); quy ước dùng chung hay theo dự án (chưa kiểm NORTON); nhãn nhiễu cần lọc trước khi làm đáp án huấn luyện (bedside, console, corner screen).
