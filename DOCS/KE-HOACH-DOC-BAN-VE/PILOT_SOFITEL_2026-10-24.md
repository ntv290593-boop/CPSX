# Pilot SOFITEL: Gemini (agy) đọc hình vs session trước, tự xem lại crop

Nguồn: 5 tờ PDF trong `01_scan\_sheets` (GR-TY-702, PS-TY-701, PS-TY-702, PS-TY-736, CO-TY-701). Chưa ghi gì vào v8.

| Dòng | Nội dung | Session trước | Gemini | Tự xem crop | Kết luận |
|---|---|---|---|---|---|
| F007/F008/F020/F021 | Len gỗ cao | 80 (75+5) | 80 (75+5) | 80 | Khớp 3 bên, độ tin cậy cao |
| F013/F027 | Len sơn cao | 100, dải sơn 40 | 100, 40 | (không xem lại) | Khớp 2 bên |
| F025/F026 | Trim 20-40x40 | khớp một phần | trim 20x20, khe 5 | (không xem lại) | Giữ "khớp một phần" |
| F023 | 4 thanh 40x140, cao 2100 | 4 thanh, khe 105, tổng 470; chưa thấy 2100 | giống | 4x40 + 3x105 = 475, không phải 470 | **Đã sửa sau GĐ3 (DWG):** khe thật 103,33, tổng 470 đúng; 105 là do DIMRND=5 làm tròn, không phải lỗi bản vẽ. 2100 không có trên tờ |
| F024 | Slatting ở air grille | không thấy, THIẾU | không thấy | (không xem lại) | Khớp, giữ THIẾU/RFI-04 |
| F018/F019 | Cột PT-03 630 x 2350 | "7 cột x 0,63 x 2,5 = 11,025 khớp" | "7 cột lỗ" | Tờ chỉ có 1 cột, mặt đứng/chi tiết 06 có 5 cột lỗ | Số lượng 7 không có trên tờ, chỉ suy từ BoQ (vòng tròn). Kích thước 1 cột đúng (630, 510 mặt, 2350/2500). F019 (7 m²) không chia hết cho diện tích 1 cột |
| GR702 ct03 | Len kim loại | "cao 80" | 40 (2+38) | 40 (8+32), dày 20 | Session trước sai (80); Gemini đúng tổng, sai phân tách. Ngoài phạm vi (loại kim loại) |
| PS736 ct05 | Chuỗi kích thước | - | thiếu 2 cặp 31/40 (tổng 419) | 43+40+(31+40)x4+43 = 410, +10x10 = 510 | Kiểm tổng chuỗi bắt được lỗi đọc của Gemini |

## Bài học
1. Trên tờ Sofitel, số kích thước là vector (15-40 nghìn nét), không có trong text layer. Không đối chiếu text layer như NORTON được.
2. Kiểm tổng chuỗi kích thước bắt được cả lỗi của Gemini (419 != 510) lẫn lỗi của bản vẽ (475 != 470).
3. "Khớp BoQ" bằng phép chia diện tích BoQ cho kích thước 1 đơn vị chỉ chứng minh kích thước, không chứng minh số lượng.
4. Gemini đọc đúng ghi chú và tổng, hay sai khi liệt kê chuỗi dài và khi đếm.
