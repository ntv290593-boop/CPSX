# So sánh ba cách đọc số kích thước trên PDF: bộ giải mã vector, Gemini, Claude

Ngày: 2026-10-24 · Nối tiếp `KET-QUA-GIAI-MA-CHU-SO-VECTOR-PDF` và `KET-QUA-KIEM-CHEO-HINH-HOC-PDF`.

## Thiết kế

- **Bộ dữ liệu:** 35 crop chi tiết ở 200 dpi, lấy từ 14 tờ SOFITEL **không dùng để huấn luyện mẫu chữ số** (mỗi crop là một viewport mở rộng 30 pt, có từ 8 đến 45 kích thước). Tổng 540 số kỳ vọng.
- **Đáp án:** chuỗi hiển thị của kích thước trong DWG (giá trị đo × DIMLFAC, làm tròn theo DIMRND/DIMDEC kể cả ghi đè từng kích thước; text gõ đè bằng số thì dùng chính số đó; bỏ kích thước có chữ như VARIES và crop có số thập phân).
- **Ba nguồn:** (1) bộ giải mã vector (`pdf_dim_decode.py`, mẫu chỉ từ các tờ train); (2) Gemini 3.8 Flash mức low; (3) Claude Sonnet 4.6, cả hai gọi qua `agy` bằng cùng một prompt, chỉ yêu cầu liệt kê mọi số kích thước thấy được, ghi "?" nếu không đọc được, không đoán, đầu ra là mảng JSON.
- **Cách chấm:** so tập hợp số (có đếm trùng) trong từng crop với đáp án; không kiểm số đó gắn với kích thước nào. Độ phủ = số khớp / số kỳ vọng; độ chính xác = số khớp / số model báo.
- **"Claude" ở đây là Claude Sonnet 4.6 chạy qua `agy`, không phải phiên bản đang làm việc trong phiên này.**

## Kết quả

| Nguồn | Độ phủ | Độ chính xác | Crop khớp hoàn toàn | Số thừa | Số thiếu |
|---|---|---|---|---|---|
| Bộ giải mã vector | 90,2% | **97,8%** | 11/35 | 11 | 53 |
| Gemini 3.8 Flash (low) | **94,8%** | 96,2% | **16/35** | 20 | 28 |
| Claude Sonnet 4.6 (qua agy) | 94,6% | 95,0% | 13/35 | 27 | 29 |

**Nhãn DWG có nhiễu:** 27/540 số kỳ vọng (5%) không nguồn nào trong ba nguồn thấy, nhiều nhất là "40" (9 lần), "5", "30": nhiều khả năng đó là kích thước có trong DWG nhưng không in ra trong PDF. Loại phần này khỏi đáp án (nhãn hiệu chỉnh, 513 số):

| Nguồn | Độ phủ (nhãn hiệu chỉnh) | Độ chính xác |
|---|---|---|
| Bộ giải mã vector | 94,9% | 97,8% |
| Gemini low | **99,8%** | 96,2% |
| Claude Sonnet 4.6 | **99,6%** | 95,0% |
| **Đồng thuận 3/3** (số cả ba nguồn đều báo) | 94,7% | **98,8%** |
| **Đa số ≥2/3** | **99,6%** | 97,1% |

**Số thừa của hai model:** Gemini có 7 mục không phải số thuần (nhãn bán kính R410, R10, R20…, một góc 58°) và 13 số thuần; Claude có 7 nhãn bán kính và 20 số thuần. Bỏ nhãn bán kính ra, độ chính xác của Gemini khoảng 97,5% và của Claude khoảng 96,2%. Một phần "số thừa thuần" có thể là số thật không thuộc đối tượng DIMENSION (mức trung thực: chưa phân loại).

**Bất đồng:** 58/540 giá trị (10,7%) có số lượng khác nhau giữa ba nguồn, đó là phần cần xem lại.

## Tốc độ và chi phí

| Nguồn | Thời gian | Ghi chú |
|---|---|---|
| Bộ giải mã vector | vài giây/tờ | miễn phí, lặp lại được |
| Gemini low | khoảng 13 giây/crop | chạy 35 crop trong vài phút (4 song song) |
| Gemini medium | khoảng 93 giây/crop | trong 15 lượt thử đầu có 7 lượt quá 280 giây và không có kết quả: không dùng được hàng loạt |
| Gemini high | hơn 5 phút/crop | một lượt thử quá thời gian |
| Claude Sonnet 4.6 qua agy | khoảng 26 giây/crop | 35 crop trong 90 giây (4 song song) |

## Nhận định

1. **Hai model nhìn ảnh đọc tốt hơn tôi kỳ vọng khi chỉ hỏi "liệt kê số" trên crop 200 dpi:** độ phủ gần 100% sau khi bỏ nhiễu nhãn. Các lỗi ở pilot (chuỗi dài, đếm) xuất hiện khi hỏi câu phức tạp hơn trên crop lớn. Mức suy luận cao của Gemini không giúp mà còn làm chậm.
2. **Bộ giải mã chính xác hơn nhưng thiếu nhiều hơn:** hợp với vai trò "người đọc thứ hai" độc lập, rẻ, có thể chạy trước.
3. **Kết hợp có lợi:** số cả ba nguồn đồng ý có độ chính xác 98,8%; đa số 2/3 cho độ phủ 99,6% và độ chính xác 97,1%. Phần bất đồng khoảng 11% là phần dồn cho người xem (hoặc cho kiểm chéo hình học).
4. **Không nguồn nào đạt 100%.** Ngay cả khi trừ nhiễu nhãn, vẫn còn khoảng 2–3% số báo không khớp đáp án; chưa tách được phần nào là đọc sai thật và phần nào là số thật ngoài tập DIMENSION.

## Giới hạn

- Chỉ SOFITEL; 35 crop, 14 tờ; mỗi nguồn chạy một lần (không đo mức dao động giữa các lần chạy của model).
- Chấm theo tập hợp số, **không kiểm số gắn đúng kích thước nào hay đúng chi tiết nào**: một model có thể liệt kê đủ số nhưng gán sai vị trí.
- Chỉ so Claude Sonnet 4.6 qua `agy`; chưa thử Claude Opus 4.6 thinking, chưa thử bản Claude của phiên này.
- Chưa so Gemini high trên toàn bộ vì quá chậm.
- Prompt hỏi liệt kê, chưa thử hỏi ngược hoặc đọc kép với crop khác độ phân giải.
- Mọi kết quả dựa trên đáp án DWG đã có nhiễu (5% kích thước không in ra).

## Đề xuất cho quy trình

1. Chạy bộ giải mã vector trước (rẻ, lặp lại được); cắt crop theo chi tiết ≥200 dpi; cho **Gemini low** đọc liệt kê số (nhanh nhất, kết quả tốt nhất trong ba nguồn); dùng Claude Sonnet 4.6 qua `agy` làm người đọc thứ ba khi cần.
2. Nhận số khi các nguồn đồng thuận; số bất đồng đi qua kiểm chéo hình học (kích thước lớn) hoặc cộng chuỗi, còn lại chuyển người xem.
3. Bước kế tiếp nên đo thêm: số có gắn đúng kích thước không (so vị trí), mức dao động giữa các lần chạy, và hồ sơ khác SOFITEL.

## File

Dữ liệu và kết quả thô ở thư mục tạm của phiên (không lưu vào dự án): crop, `bench.json`, `out\<model>\<crop>.txt`, `score.py`, `score2.py`. Có thể tái tạo bằng `prep_bench.py` và `run_bench.ps1`; muốn giữ lại thì cần chép vào `_MASTER\scan\bench\`.
