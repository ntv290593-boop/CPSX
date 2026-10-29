# Kết quả: kiểm chéo số kích thước bằng đo hình học trong PDF ("cách 2")

Ngày: 2026-10-24 · Nối tiếp `KET-QUA-GIAI-MA-CHU-SO-VECTOR-PDF_2026-10-24.md`. Phạm vi: SOFITEL, 36 tờ ngẫu nhiên (2.225 số đã giải mã và nằm đúng vị trí kích thước DWG). Chưa thử dự án khác.

## Ý tưởng

Số trên đường kích thước phải bằng chiều dài hình học của đường đó nhân tỉ lệ chi tiết: `giá trị = chiều dài (pt) ÷ 2,8346 × tỉ lệ`. Hai đầu đường kích thước là gạch chéo ngắn (nét xiên khoảng 5,7 pt) nên tìm được trực tiếp trong PDF. Nguồn đo này độc lập với việc đọc chữ số.

## Kiểm chứng nền tảng

- Trên PS-TY-736, **180/180** đầu đường kích thước (tính từ DWG) đều có gạch chéo đúng vị trí trong PDF.
- Kích thước tuyến tính chỉ có góc 0° và 90° (kích thước xiên chỉ 3,4%).
- Số nằm giữa hai đầu đường kích thước ở 73% trường hợp ngang và 83% trường hợp dọc; phần còn lại nằm ngoài (đoạn ngắn có chữ đặt bên ngoài).

## Kết quả (dùng đúng tỉ lệ thật của kích thước; dung sai ±3 mm vì DIMRND = 5)

| Số chữ số | Số lượng | Xác nhận được số ĐÚNG | Số SAI cố ý (đổi 1 chữ số ngẫu nhiên) vẫn được xác nhận |
|---|---|---|---|
| 1 | 135 | 5,9% | 6,7% |
| 2 | 1.196 | 62,8% | 17,4% |
| 3 | 686 | **98,0%** | 8,0% |
| 4 trở lên | 208 | **95,2%** | 4,3% |

Theo dimstyle: `1-1-PS-10` 80%, `1-1-PS-5` 79%, `1-1-PS-2` 82%, `1-1-PS-25` 53% (mặt bằng 1:25 nhiều kích thước ngắn). Theo hướng: dọc 82%, ngang 65% (do kích thước ngắn nhiều ở hướng ngang).

Khi lấy tỉ lệ từ nhãn "Scale 1:N" gần nhất trong text layer (cách chạy thật, không có DWG): tỉ lệ xác nhận số đúng giảm khoảng 4–5 điểm (60,2% so với 64,7% khi dùng tỉ lệ thật, tính chung mọi số chữ số).

## Giải thích

- **Số lớn (3 chữ số trở lên) kiểm chéo rất tốt:** 95–98% số đúng được xác nhận; số bị MISMATCH rất ít (khoảng 2%) nên nếu bị báo lệch thì đáng nghi.
- **Số nhỏ (1–2 chữ số) kiểm chéo yếu:** đoạn kích thước rất ngắn (hai gạch chéo gần như chồng nhau, bộ lọc hiện loại đoạn ngắn dưới 8 pt), chữ thường đặt bên ngoài đoạn, và dung sai ±3 mm xấp xỉ bằng chính số đó nên không phân biệt được. Đây là các khe, độ dày cần chú ý.
- Thử cho phép cặp gạch chéo không liền kề: không cải thiện số đúng, chỉ làm tăng số sai lọt. Thử chọn đường theo khoảng cách điển hình (3–9 pt): không cải thiện.
- Đổi 1 chữ số ngẫu nhiên là phép thử nghiêm hơn lỗi thật của bộ giải mã (lỗi thật thường là nhầm cụm kế bên hoặc cặp chữ giống nhau như 0 và 5), nên tỉ lệ sai lọt thật có thể thấp hơn các số trên. Chưa đo trực tiếp.

## Quy tắc quyết định đề xuất (kết hợp giải mã + hình học)

| Điều kiện | Quyết định |
|---|---|
| Giải mã được chấp nhận và hình học khớp | **Nhận** (hai nguồn độc lập) |
| Giải mã được chấp nhận, số từ 3 chữ số trở lên, hình học báo lệch | **Xem lại** (hiếm, khoảng 2%) |
| Giải mã được chấp nhận, không tìm được đoạn kích thước (số nhỏ, chữ ngoài đoạn) | **Xem lại nhẹ**: cộng chuỗi kích thước hoặc đọc kép bằng Gemini/Claude |
| Giải mã bị từ chối hoặc không tìm thấy cụm số | **Chuyển người đọc ảnh** (Gemini/Claude trên crop ≥200 dpi) |

## Giới hạn

- Chỉ kiểm trên SOFITEL; chưa biết tỉ lệ tương tự với bộ hồ sơ khác.
- Tỉ lệ chi tiết lấy bằng nhãn "Scale 1:N" gần nhất có thể gán sai; chưa gắn chi tiết bằng tiêu đề và viewport như DWG.
- Chưa xử lý đoạn ngắn, kích thước xiên, chữ xoay xiên.
- Chưa kết hợp thực sự vào một quy trình (mới là các module riêng).
- Chưa đo sai lọt cho lỗi thật của bộ giải mã.

## File

`_MASTER\scan\` (bản git `_SCRIPTS\scan\`, chưa commit): `pdf_dim_geometry.py` (tìm gạch chéo, đoạn kích thước, nhãn tỉ lệ, kiểm tra), `pdf_dim_geometry_eval.py` (đánh giá với đáp án DWG).

## Việc tiếp theo (đề xuất)

1. Xử lý đoạn ngắn: hạ ngưỡng 8 pt, xử lý gạch chéo chồng nhau và chữ đặt ngoài đoạn bằng đường dẫn (leader).
2. Chấm Gemini và Claude trên cùng bộ dữ liệu để biết cách nào bổ sung tốt nhất cho phần "xem lại".
3. Cộng chuỗi kích thước (kiểm chéo thứ ba) cho các số nhỏ.
4. Thử trên hồ sơ khác để biết mẫu chữ số và kiểm chéo có dùng lại được không.
5. Ghép giải mã + hình học + gắn chi tiết vào một bước scan dùng thật.
