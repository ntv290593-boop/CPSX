# Điều chỉnh plan: PDF là đường làm việc chính, DWG là "đáp án" để đo

Ngày: 2026-10-24 · Nguồn: người dùng cho biết công việc chính làm trên PDF, không phải lúc nào cũng có DWG. Bổ sung cho `KE-HOACH-TONG-THE_v2.md` (chưa sửa v2; cần duyệt trước khi gộp).

## Thay đổi vai trò

| Trước | Sau |
|---|---|
| Tầng A (DWG) là nguồn chính cho kích thước | Tầng A chỉ dùng **khi dự án có DWG** và làm **đáp án để chấm các phương pháp đọc PDF** |
| GĐ3b: người dùng duyệt đáp án 10–15 giá trị | Đáp án sinh tự động từ DWG (SOFITEL: 242 tờ có mã, 14.013 kích thước có vị trí), người dùng chỉ duyệt mẫu nhỏ |
| Pipeline chuẩn xoay quanh DWG | Pipeline chuẩn xoay quanh PDF (tầng B, C, D) và có bước **phân loại loại PDF** ở đầu |

## Giá trị của DWG khi PDF là đường chính

DWG của SOFITEL cho biết giá trị đúng của hàng nghìn kích thước, kèm vị trí và chi tiết. PDF của cùng tờ đó là đầu vào thật của người dùng. Ghép hai thứ lại cho phép đo phương pháp đọc PDF trên quy mô lớn mà không cần ai ngồi gán nhãn. Điều kiện: chỉ lấy kích thước không bị gõ đè, và tờ đã khớp phiên bản (192 tờ khớp, 4 lệch, 75 chưa kiểm).

## Các cách đọc số kích thước từ PDF, cần thử và đo

Trên tờ SOFITEL, số kích thước là đường vector, không có trong text layer (đã kiểm). Các hướng sau là **giả thuyết cần kiểm chứng**, chưa thử:

| # | Cách | Ý tưởng | Rủi ro |
|---|---|---|---|
| 1 | **Giải mã glyph vector** | Chữ số được vẽ thành nét vector từ một bộ font hữu hạn. Gom nhóm các hình theo dạng để suy ra chữ số, không dùng model nhìn ảnh, kết quả lặp lại được. Dùng nhãn từ DWG để kiểm | Nét có thể bị cắt/ghép, chữ nhỏ; phụ thuộc font của từng bộ hồ sơ |
| 2 | **Đo hình học đường kích thước** | Chiều dài giữa hai gạch chéo/đường gióng × tỉ lệ của chi tiết = giá trị thật, không phụ thuộc text. Tỉ lệ lấy từ nhãn "Scale 1:5" (có trong text layer) | Phải ghép đúng đường kích thước với chi tiết; PDF in co thì sai tỉ lệ (các tờ đã kiểm đều A1 1:1) |
| 3 | **Model nhìn crop ≥200 dpi** (Gemini/Claude) | Đã dùng ở pilot | Sai ở chuỗi dài và đếm; cần kiểm cộng chuỗi và đọc kép |
| 4 | **OCR chuyên dụng** (ví dụ PaddleOCR) trên crop | Rẻ, chạy cục bộ | Chưa thử; chữ số trong bản vẽ có thể bị nhận sai |

Đo trên bộ nhãn DWG: mỗi cách đúng bao nhiêu phần trăm, sai kiểu gì, lỗi nào bị phép cộng chuỗi bắt được.

## PDF có nhiều loại, không phải tờ nào cũng có vector

Bộ hồ sơ khác có thể là PDF vector kèm text (như NORTON), PDF vector với chữ đã chuyển thành nét (như SOFITEL), hoặc PDF ảnh quét. Mỗi loại cần cách đọc khác nhau. Đề xuất thêm bước đầu pipeline: **phân loại từng tờ** (có text layer không, có vector không, có ảnh nhúng không) và chọn tầng tương ứng. Trường hợp PDF ảnh quét thì chỉ còn tầng D.

## Điều chỉnh lộ trình

| GĐ | Thay đổi |
|---|---|
| 3 (tầng A) | Giữ làm công cụ tạo đáp án; không phát triển thêm cho sản xuất |
| 3b | Chuyển thành "sinh đáp án từ DWG" (gần như tự động) |
| 2 (benchmark) | Chấm các cách đọc PDF (mục trên) bằng đáp án DWG; chọn cấu hình mặc định theo ngưỡng đặt trước |
| 4 (script chuẩn) | Xoay quanh PDF: phân loại tờ → trích B/C → đọc D khi cần → đối soát |
| Mới | Bước phân loại loại PDF ở đầu bước scan |

## Điều cần người dùng quyết

1. Các hồ sơ thật sẽ nhận phần lớn là PDF vector có text, PDF vector với chữ thành nét, hay PDF ảnh quét? Có thể cho vài ví dụ khác SOFITEL và NORTON.
2. Có đồng ý thử cách 1 và 2 (giải mã glyph vector, đo hình học) bằng đáp án DWG của SOFITEL trước khi đầu tư vào pipeline không.
