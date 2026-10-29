# Scan: nhận loại đồ vật bằng quan sát có cấu trúc (thiết kế, 2026-10-24)

Viết cho session Claude làm bước scan/takeoff. Bối cảnh lỗi: `DOCS\HANDOFF_SOFITEL_BO-SUNG_2026-10-24.md` §4, memory `template-recognition-convention`.

## 1. Nguyên tắc
1. **Một lần đọc, hai lớp kết quả**: model đọc bộ tờ → (a) QUAN SÁT theo trường cố định, mỗi giá trị kèm nguồn; (b) LOẠI đề xuất từ danh sách đóng (`02_takeoff\LOAI_DO_VAT_<DỰ ÁN>.xlsx` sheet LOAI) hoặc `KHÔNG RÕ`, kèm lý do.
2. **Bước sau không phân loại lại**: luật máy kiểm quan sát ↔ loại ↔ ĐVT; lệch → cờ, đưa user (duyệt theo ngoại lệ).
3. **Mẫu không do model chọn**: loại → mẫu tra bảng user điền.
4. **Không mất dữ liệu**: chỉ THÊM dòng, không ghi đè; mỗi giá trị có nguồn; phân biệt `KHÔNG CÓ` với `KHÔNG THẤY`; giữ mọi giá trị khi nguồn mâu thuẫn.

## 2. Nơi lưu (PROJECT_DATA)
Sheet mới **QUAN_SAT** (dạng dài, 1 dòng = 1 quan sát):

| Cột | Ý nghĩa |
|---|---|
| Mã | mã BoQ |
| Thành phần | `1`, `2`… khi 1 dòng BoQ gồm nhiều đồ vật (đầu giường + vách góc; vanity + gương; bench + side table) |
| Trường | xem §3 |
| Giá trị | text/số; số ghi kèm đơn vị |
| Trạng thái | `CÓ` · `KHÔNG CÓ` (thấy rõ là không có) · `KHÔNG THẤY` (chưa đọc được/không có tờ) · `MÂU THUẪN` |
| Tờ · Chi tiết | số BV + số chi tiết/hình (không chấp nhận chỉ số tờ cho trường cấu tạo) |
| Nguồn | `HÌNH` · `TEXT` · `SPEC` · `BOQ` · `USER` |
| RunId · Model · Bản DS loại | để biết dòng nào lỗi thời khi đổi danh sách loại |

BOQ_TONGHOP thêm cột: **Loại đề xuất**, **Lý do loại**, **Độ chắc**, **Cờ kiểm**, **Loại chốt** (chỉ user/luật ghi). Giá trị hiện hành của 1 trường = dòng mới nhất, ưu tiên `USER`. Chạy lại không bao giờ xóa dòng `USER`.

## 3. Trường quan sát (lấy từ điểm phân biệt giữa các loại)
Lắp đặt (gắn tường/đứng rời/âm tường) · Kích thước BV (R×S×C, từng nhánh nếu chữ L) · Kích thước BoQ · Có cánh (số, mở/lùa) · Ngăn kéo (số) · Thanh treo · Kệ/đợt · Hốc/ô tivi · Mặt đá · Gương · Kính · Kim loại · Giấy/vải · Vật liệu gỗ (mã WD/PT) · Ván nền ẩn sau lớp hoàn thiện · Khung xương · Cong/bo (R) · Phần ngoài phạm vi · Khu (Phòng/JS/PS/Hành lang…) · Bộ tờ đã đọc · Tờ chưa đọc · Revision tờ.

Thứ tự: model điền quan sát từ HÌNH **trước**, rồi mới đọc mô tả BoQ và điền loại (giảm neo theo tên).

## 4. 20 kịch bản sai thông tin và cách chặn

| # | Kịch bản | Ví dụ thật | Chặn bằng |
|---|---|---|---|
| 1 | Đọc thiếu tờ trong bộ | F066 chỉ đọc 711, bỏ 710/712 | `Tờ chưa đọc` ≠ rỗng → Độ chắc tối đa THẤP, không chốt loại |
| 2 | Đọc nhầm tờ của type khác cùng tên | minibar type 2 (713/714) vs type 3 (715/716); F040 vs F066 | đối chiếu số tờ khung tên ↔ `Bản vẽ (CĐT ghi)`; kích thước BV vs BoQ lệch >10% → cờ |
| 3 | Callout vector không có text → không đi tới tờ chi tiết | "18 / ID-PS-TY-712" | model đọc callout từ HÌNH ghi vào CALLOUT; tờ đích chưa đọc → #1 |
| 4 | Trống bị hiểu là "không có" | không thấy ngăn kéo vì chưa xem mặt cắt | trạng thái bắt buộc `KHÔNG CÓ`/`KHÔNG THẤY`; ô trống = lỗi |
| 5 | BoQ và BV khác kích thước, giữ 1 mất 1 | F054 BoQ 1500 / BV 1620; F053 sâu 270 / ~150 | ghi cả 2 dòng (nguồn BOQ, HÌNH), trạng thái `MÂU THUẪN` → ISSUES/RFI |
| 6 | Đọc sai chữ số/đơn vị/tỷ lệ | decoder vector 99% | tổng các đoạn = kích thước tổng; so DWG khi có; số lẻ bất thường → cờ |
| 7 | 1 dòng BoQ nhiều đồ vật bị ép 1 loại | F062 đầu giường + vách góc; F075 vanity + gương; L043 | cột Thành phần; tên BoQ có "&"/"with"/"and" mà chỉ 1 thành phần → cờ |
| 8 | Nhiều dòng BoQ cùng 1 đồ vật / trùng | F054 trùng kích thước F030; F026.1 đã có block | trùng (tờ + chi tiết) giữa 2 mã → cờ TRÙNG |
| 9 | Bỏ phần gỗ ẩn sau lớp ngoài phạm vi | ván nền dưới giấy WC/gương | trường `Ván nền ẩn` bắt buộc khi có giấy/gương/kim loại |
| 10 | Lẫn phần ngoài phạm vi vào tính giá | đá ST, kim loại MT, OS&E | trường `Phần ngoài phạm vi` so dòng PROJECT "Phạm vi tính giá" |
| 11 | Mã hoàn thiện sai khu/phase | WC-07 chỉ JS; WD-04 phase 2 vs WD-01 | kiểm FINISHES (Khu áp dụng) ↔ trường Khu |
| 12 | Trang không có chữ bị bỏ lặng lẽ | 26 trang shop drawing cũ | trang không text → đọc HÌNH hoặc ghi `KHÔNG THẤY` + lý do, không bỏ |
| 13 | Model bịa chi tiết không có trên hình | — | bằng chứng phải có số chi tiết; mẫu ngẫu nhiên 10% đối chiếu lại bằng lượt đọc thứ 2 |
| 14 | Neo theo tên BoQ (thấy điều mình chờ) | "Wardrobe" → ép thành tủ áo | quan sát trước, tên sau (§3); luật: loại tủ áo cần thanh treo `CÓ` |
| 15 | Đồ mới không thuộc loại nào bị ép loại gần nhất | fins portal, khung bồn tắm | cho `KHÔNG RÕ`; luật quan sát ↔ loại sai → cờ; thêm loại mới do user duyệt |
| 16 | Đổi danh sách loại, nhãn cũ mất nghĩa | gộp/tách loại | lưu `Bản DS loại` mỗi dòng; đổi bản → phân loại lại từ QUAN_SAT, không đọc lại BV |
| 17 | Chạy lại ghi đè sửa tay của user | — | chỉ thêm dòng; không xóa dòng `USER`; đếm dòng trước/sau |
| 18 | Ghi file làm mất dữ liệu | PermissionError khi file mở; openpyxl bỏ Conditional Formatting/Data Validation extension (RE00 có); chuỗi "=" thành công thức | kiểm khóa `~$` trước khi ghi; sao lưu; đếm hyperlink/dòng trước-sau; KHÔNG lưu file CPSX bằng openpyxl; `data_type='s'` |
| 19 | Đọc revision cũ của tờ | tờ thay thế trong bộ hồ sơ | trường Revision; 2 file cùng số BV → lấy rev cao nhất, cờ nếu khác |
| 20 | Gán chéo khu: cùng tên, khác thiết kế theo khu | F098 gán nhầm tờ PS; D20 nhầm tờ GM | Khu của tờ = Khu có KL của dòng BoQ (S5) |

## 5. Luật kiểm quan sát ↔ loại (khởi đầu, bổ sung dần)
- Len chân / chỉ gỗ: ĐVT `m`; Vách/ốp/ván nền: `m2`; tủ, bàn, ghế, cửa: `Nr`.
- Mã `D*` → loại cửa; loại cửa → có cánh `CÓ`.
- Tủ áo → thanh treo `CÓ`; hốc ủi → không có cánh hoặc có cánh + kích thước rộng ≤ 400.
- Vanity → mặt đá `CÓ` (ngoài phạm vi SOFITEL) + phần gỗ dưới.
- Loại có Độ chắc CAO chỉ khi: không có `KHÔNG THẤY` ở trường phân biệt của loại đó, không có `MÂU THUẪN` chưa xử lý.

## 6. Chấm trước khi dùng
Tập đáp án: mã RE00 đã có AH (80) + loại user xác nhận ở `LOAI_DO_VAT_SOFITEL.xlsx`. Chạy không cho xem đáp án. Đạt: ≥95% đúng trên mục model trả lời; `KHÔNG RÕ` ≤15%; mọi cờ của 20 kịch bản được kích ít nhất 1 lần trong bộ thử (để biết luật chạy thật).
