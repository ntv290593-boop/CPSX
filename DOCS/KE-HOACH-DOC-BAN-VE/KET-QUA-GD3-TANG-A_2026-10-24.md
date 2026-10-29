# Kết quả GĐ3: thử tầng A (DWG → DXF → ezdxf) trên 5 tờ pilot SOFITEL

Ngày: 2026-10-24 · Phạm vi: 5 tờ (ID-PS-TY-736, -702, -701, ID-GR-TY-702, ID-CO-TY-701). **Chưa chạy trên cả 271 DWG.**

## Kết luận ngắn

Tiêu chí thoát của GĐ3 đạt ở mức "có mặt": **41/41 giá trị đã xác nhận bằng mắt đều xuất hiện trong dữ liệu DWG**, và việc gắn kích thước vào đúng chi tiết chạy được trên tờ thử (PS-TY-702). 5/5 tờ khớp phiên bản với PDF. DWG còn sửa được hai chỗ sai: một của Gemini, một của chính tôi trong pilot.

## Cách làm (lặp lại được)

1. Copy DWG sang thư mục tạm (không đụng file dự án). ODA File Converter 27.9.0:
   `"C:\Program Files\ODA\ODAFileConverter 27.9.0\ODAFileConverter.exe" <in> <out> ACAD2018 DXF 0 1 "*.dwg"`
   Lệnh trả về ngay nhưng file DXF xuất hiện sau vài giây (chạy nền); 5 file mất khoảng 10 giây.
2. `ezdxf` đọc DXF. Kích thước **nằm ở paper space** (layout `A1 PLAN`, 47–94 DIMENSION/tờ), model space chỉ có hình học và text rời.
3. **Giá trị thật = `get_measurement()` × DIMLFAC** của dimstyle (1:5 → 5, 1:10 → 10, 1:2 → 2). Ví dụ 126,0 × 5 = 630; 102 × 5 = 510; 8,6 × 5 = 43.
4. Gắn vào chi tiết: tiêu đề chi tiết nằm trong thuộc tính INSERT (`TITLE`, `Nº`, `SCALE`); mỗi chi tiết là một viewport; kích thước thuộc chi tiết có viewport chứa nó (hoặc gần nhất).
5. Script: `dwg_extract.py` (trích kích thước, text, phát hiện gõ đè) và `dwg_details.py` (gắn chi tiết), mức spike, lưu ở `_MASTER\scan\` và `_SCRIPTS\scan\`.

## Kết quả đo

| Kiểm tra | Kết quả |
|---|---|
| Giá trị đã xác nhận có trong DWG | 41/41 (PS-736: 14, PS-702: 7, GR-702: 7, CO-701: 4, PS-701: 9). Một giá trị (30 ở PS-702) chỉ khớp khi cho dung sai làm tròn. Lưu ý: 9 giá trị của PS-701 là số Gemini đọc, tôi chưa tự xem; DWG xác nhận cả 9 |
| Gắn vào chi tiết (PS-702) | Chi tiết 01/02 → 5, 5, 10, 10, 75, 80; chi tiết 08 → 5, 4×40, 3×103,3, 470; chi tiết 03 → 2, 8, 20, 32, 40. Một số chi tiết (06, 04, vp18, vp28) gắn chưa chắc |
| Phiên bản DWG ↔ PDF | 5/5 khớp số bản vẽ, ngày phát hành và chữ phiên bản (`-`, `C`, `C`, `A`, `B`) |
| Gõ đè text kích thước | PS-702: 5 chỗ (4 lệch); GR-702: 4 chỗ (4 lệch); 3 tờ còn lại: 0. Ví dụ text "900" trong khi đo thật 530, "VARIES", "<>/220". Đây là chủ ý của người vẽ (kích thước đoạn cắt), không phải lỗi |

## Hai điều chỉnh về kết luận trước đó

1. **F023 (4 thanh 40×140): bản vẽ không tự mâu thuẫn.** Khe thật đo ra **103,33** (3 khe), 4×40 + 3×103,33 = **470**, khớp con số tổng. Số 105 hiển thị là do dimstyle `1-1-PS-5` có **DIMRND = 5** (làm tròn tới bội số 5). Pilot ghi "475 ≠ 470, báo RFI" là sai; đã sửa trong pilot và plan v2.
2. **Phép kiểm cộng chuỗi cần dung sai:** khi dimstyle có DIMRND, dung sai = số đoạn × DIMRND/2 (ví dụ 3 khe × 2,5 = 7,5 mm). Nếu không có dung sai, phép kiểm sẽ báo lỗi giả.
3. **DWG sửa lỗi Gemini:** PS-702 chi tiết 03 (len kim loại): Gemini đọc 30+10, DWG cho 8+32 (+2), cùng giá trị với GR-702 chi tiết 03 mà tôi xem bằng mắt.

## Giới hạn và rủi ro còn lại

- Mới chạy 5 tờ. 271 DWG có thể có tờ lỗi, dimstyle khác, xref hoặc text gõ đè nhiều hơn.
- "Có mặt" khác "gắn đúng". Phép so 41 giá trị chỉ kiểm giá trị có xuất hiện trong dữ liệu, không kiểm gắn đúng chi tiết; việc gắn mới thử trên một tờ.
- Kích thước bị gõ đè: không tự nhận giá trị nào; đánh dấu "xem lại" và chuyển người quyết (giá trị hiển thị là ý đồ thiết kế, giá trị đo là hình học).
- Kích thước hiển thị và giá trị thật có thể khác nhau (làm tròn). Cần quy định dùng giá trị nào cho chiết tính: đề xuất lấy giá trị hiển thị (ý đồ thiết kế) để so BoQ và giá trị thật để kiểm chuỗi.
- Đếm số lượng vẫn phải lấy từ mặt bằng/mặt đứng (DWG có thể đếm INSERT/block, chưa thử).
- Ghi chú, nhãn vật liệu (MTEXT, leader) chưa được trích và gắn vào chi tiết.

## Đề xuất bước tiếp

1. Chạy tầng A trên cả 271 DWG SOFITEL (ước tính 10–15 phút chuyển đổi), kèm báo cáo: tờ lỗi, số gõ đè, tờ không có dimstyle chuẩn.
2. Hoàn thiện gắn chi tiết (kiểm các viewport chưa gắn) và trích ghi chú/leader.
3. Dùng dữ liệu DWG làm đáp án để chấm Gemini và Claude trên PDF (GĐ3b/GĐ2), chỉ nhận kích thước không bị gõ đè.
4. Áp kết quả pilot vào bản sao v9 (GĐ1) với F023 đã sửa.
