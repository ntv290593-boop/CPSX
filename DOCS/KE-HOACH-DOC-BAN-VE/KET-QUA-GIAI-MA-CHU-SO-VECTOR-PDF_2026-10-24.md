# Kết quả: giải mã số kích thước vector trực tiếp từ PDF (không dùng model nhìn ảnh, không cần DWG)

Ngày: 2026-10-24 · Thử "cách 1" trong `DIEU-CHINH_PDF-LA-DUONG-CHINH_2026-10-24.md`. Phạm vi: **SOFITEL** (bản vẽ GA Design, cùng một bộ font), 189 tờ có DWG khớp phiên bản. Chưa thử dự án khác.

## Ý tưởng

Trên PDF SOFITEL, số kích thước là các đường nét vector (mỗi chữ số một đường nét, độ dày nét 0,48, cao khoảng 5,6 pt), không có trong text layer. Mỗi chữ số có số đoạn nét rất ổn định (0: 16, 1: 3, 2: 13, 3: 14, 4: 3, 5: 16, 6: 22, 7: 2, 8: 28, 9: 22). Script gom các nét thành cụm số, so từng chữ số với thư viện mẫu, kết quả lặp lại được và không tốn token.

DWG chỉ dùng để gắn nhãn (giá trị đúng, theo chuỗi hiển thị AutoCAD: giá trị đo × DIMLFAC, làm tròn theo DIMRND/DIMDEC đọc cả từ ghi đè trong XDATA của từng kích thước, làm tròn lên). Khi chạy thật không cần DWG.

Phép chuyển tọa độ DWG → PDF chính xác tuyệt đối: `pdf_x = 2,8346 × x`, `pdf_y = H − 2,8346 × y` (khổ A1 1:1, không dịch gốc).

## Kết quả đo (chia tờ huấn luyện/kiểm tra theo tờ, 63 tờ kiểm tra, 3 cách chia ngẫu nhiên)

Quần thể: kích thước có nhãn DWG và không bị gõ đè text (khoảng 3.900–4.300 kích thước mỗi lần chia).

| Kết quả | Cách chia 1 | Cách chia 2 | Cách chia 3 |
|---|---|---|---|
| Chấp nhận và đúng | 90,6% | 94,2% | 95,0% |
| Chấp nhận nhưng sai | 0,9% | 0,6% | 0,3% |
| Từ chối (glyph lạ) | 3,0% | 2,9% | 3,0% |
| Không tìm thấy cụm số | 4,5% | 1,6% | 1,5% |
| **Chính xác trong số được chấp nhận** | **98,98%** | **99,32%** | **99,70%** |

Trong số "chấp nhận nhưng sai": chỉ 0,1–0,2% (3–10 ca/lần chia) là cụm có số chữ số đúng nhưng đọc khác nhãn (nghi đọc sai thật); phần còn lại là cụm bị ghép với kích thước kế bên hoặc nhãn DWG ghép nhầm cụm (đọc đúng số của cụm bên cạnh).

Ba sửa làm kết quả tốt lên: (1) lọc mẫu huấn luyện theo số đoạn nét chủ yếu của từng chữ số (nhiễu nhãn: có chữ "0" bị gắn nhầm thành "8") và buộc ứng viên cùng số đoạn; (2) tính chuỗi hiển thị nhãn đúng cách AutoCAD (làm tròn lên; ghi đè DIMRND theo từng kích thước); (3) ngưỡng ghép chữ theo độ rộng chữ ("1" hẹp cần khoảng cách lớn hơn).

## Dương tính giả (không phải chữ số kích thước)

Trên 36 tờ ngẫu nhiên: 2.951 cụm được chấp nhận, 2.295 cụm nằm trong 9 pt quanh vị trí kích thước DWG (77,8%). 574 cụm "xa" đều bắt đầu bằng số 0, là rãnh CNC/hình bầu dục có nét giống chữ "0". Quy tắc **loại cụm bắt đầu bằng 0** đã được thêm vào `pdf_dim_decode.py`. Còn 82 cụm khác (2,8%): số thật không thuộc đối tượng DIMENSION (ví dụ 2100, 1000, 14); chưa phân loại.

Ví dụ tờ `ID-PS-TY-736`: bộ giải mã cho ra 90 giá trị (630, 510, 2350, 190, 300×12, 75, 43, 31, 35, 80, 60, 40, 25, 30×13, 10×23), khớp DWG; thiếu 4 số "20".

## Giới hạn

- **Chỉ kiểm trên SOFITEL**, cùng bộ font và cùng đơn vị thiết kế. Với dự án khác (font khác, cỡ chữ khác) phải xây lại mẫu. Hiện mẫu được gắn nhãn từ DWG, nên dự án mới không có DWG thì chưa có cách gắn nhãn tự động (cần thử xem mẫu có dùng lại được không).
- **Chữ xoay theo góc xiên** chưa xử lý (nghi là phần lớn của 3% từ chối và 1,5–4,5% không tìm thấy). Chữ cỡ khác 5,6 pt cũng chưa xử lý.
- Chưa kiểm số thập phân (DIMDEC > 0), ký hiệu ±, chữ kèm tiền tố/hậu tố.
- Chưa gắn số vào đúng chi tiết (cần kết hợp tiêu đề chi tiết trong text layer và vị trí).
- Còn 0,3–0,9% "chấp nhận nhưng sai" so với DWG, nên **chưa đủ để dùng một mình cho giá**. Cần kiểm chéo (cộng chuỗi kích thước, đo hình học).
- Chưa so sánh với Gemini/Claude trên cùng bộ dữ liệu; pilot trước đó chỉ trên 5 tờ nên chưa so được.

## Việc tiếp theo (đề xuất)

1. **Kiểm chéo bằng đo hình học (cách 2):** đo khoảng cách giữa hai gạch chéo/đường gióng của kích thước × tỉ lệ chi tiết, so với số giải mã. Số khớp hình học thì độ tin cậy cao; không khớp thì chuyển người xem. Đây cũng là cách phát hiện cụm bị ghép nhầm.
2. **Xử lý chữ xoay xiên** (xoay glyph về thẳng trước khi so mẫu).
3. **Chấm Gemini và Claude trên cùng bộ dữ liệu** (cùng nhãn DWG) để so với bộ giải mã.
4. **Thử trên bộ hồ sơ khác** (NORTON hoặc PDF khác bạn có) để biết mẫu có dùng lại được không.
5. Gắn số vào đúng chi tiết và gộp vào pipeline scan (kèm cộng chuỗi).

## File

`_MASTER\scan\` (bản git: `_SCRIPTS\scan\`, chưa commit): `pdf_glyphs.py`, `pdf_dim_decode.py` (giải mã một trang, không cần DWG), `glyph_dataset.py` (tạo dữ liệu có nhãn từ cặp DWG–PDF), `build_glyph_templates.py`, `glyph_templates_sofitel.npz` (547 mẫu, 15 KB).
