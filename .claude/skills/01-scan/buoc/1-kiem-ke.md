# Bước 1 — kiemke: kiểm kê hồ sơ + thông tin dự án

**Ghi:** `FILES`, `PAGES` (script ghi), `PROJECT`, `NOTES` — file `*.kiemke.json`.
**Không làm:** đọc BoQ chi tiết, đọc bản vẽ, đánh giá spec (bước sau làm).

## Làm — thứ tự: PHẠM VI → model LƯỢC BỎ file → mới QUÉT
**A. Phạm vi trước (chưa quét gì):**
   - `py -I "<TOOL>" kiemke <GOC> <DOT> --ds` → liệt kê gọn theo thư mục (tên, đuôi, MB, số trang PDF; ~1 màn hình, không đọc nội dung, không ghi file).
   - Đọc mail yêu cầu báo giá (`py -I "<TOOL>" mail "<file.msg>"`) và **chỉ phần tổng của BoQ** (`xem <BoQ.xlsx>` lấy tên sheet + vài dòng đầu / sheet tổng; không đọc từng hạng mục). Rút ra: công ty được yêu cầu tính **gì** (loại căn, khu, nhóm hạng mục), **không** tính gì, hạn, phương án.
**B. Model tự lược bỏ file** (suy luận từ phạm vi A + danh sách thư mục, không chờ luật cố định): với từng thư mục / nhóm file quyết **ĐỌC / DỰ PHÒNG / BỎ** kèm lý do một dòng. Hỏi: *thư mục này có phục vụ hạng mục nằm trong phạm vi không? có bản PDF thay thế không? có phải bản cũ/trùng không? chỉ để tham khảo (ảnh hiện trạng, căn khác loại) không?* Mặc định: PDF > CAD; ảnh chụp hiện trạng bỏ; bộ bản vẽ của loại căn không nằm trong BoQ bỏ; bảng vật liệu / chi tiết chung / spec luôn ĐỌC. Không chắc → hỏi user 1 câu gộp, không đoán.
   Bài học (Blanca 27/10 — đã quét thừa những loại sau, lần sau xếp ngay ở B):
   - **Bộ FIT-OUT / hoàn thiện** (MB sàn, trần, tường, sơn, gạch) phục vụ hạng mục sơn/thạch cao/lát — ngoài scope nhà máy → DỰ PHÒNG, không trích text; chỉ mở khi cần mặt bằng vị trí đồ gỗ mà bộ BUILT-IN thiếu.
   - **PDF concept / đề xuất layout** (không khung tên thi công, ngày cũ hơn bộ chính) → DỰ PHÒNG, không trích text, không cho bước 4 tìm mã trong đó.
   - **File trùng** (cùng số trang, cùng danh mục, khác ngày) → chỉ trích text bản mới; bản cũ ghi BỎ.
   - **Loại căn có trong BoQ nhưng chỉ có CAD** → không trích, không cho bước 4 tìm; NOTES ghi "thiếu PDF <loại căn>" để điều phối hỏi user/CĐT ngay, không giao subagent dò cả bộ để ra KHÔNG THẤY.
   Ghi quyết định BỎ vào `<GOC>\01_scan\_bo_qua.txt` (mỗi dòng 1 regex theo đường dẫn, dòng `#` kèm lý do); DỰ PHÒNG/ĐỌC ghi 1 dòng vào NOTES. **Danh sách file DỰ PHÒNG ghi thêm vào `<GOC>_scan\_data\<DOT>\_timkiem.kiemke.json`** dạng `{"du_phong": ["F004", …]}` (ID trong FILES) — `ungvien` bỏ các file này (chặn kiểu sai: trang concept/layout dự án chen vào ứng viên của cửa, gối). Sau khi chạy `kiemke`, đọc thêm các dòng `L1 TRANG TRÙNG` / `L1 FILE TỔNG` trong kết quả: tool đã tự bỏ trang trùng (bản 29/09 vs 30/09 chỉ khác ngày) và file tổng khỏi tìm kiếm, ghi vào `_timkiem.json` — chỉ cần ghi NOTES nếu chọn bản khác.
**C. Rồi mới quét:**
1. Nếu chưa có `FILES.kiemke.json`: `py -I "<TOOL>" kiemke <GOC> <DOT>` (vài chục giây / vài trăm trang). Script tự ghi FILES, PAGES, lưu lớp text từng trang cho lệnh `tim`, và **bỏ các file thuộc `_bo_qua.txt`**.
1b. **Đọc output `kiemke` trước khi làm tiếp (lọc lớp L0/L1, không tốn token):**
   - `L0 bỏ N file nhiễu` — frame/xref/Archive/log đã tự loại, không cần xử lý.
   - Luật riêng dự án nằm ở `<GOC>\01_scan\_bo_qua.txt` (mỗi dòng 1 regex theo đường dẫn; user chốt thư mục/bộ nào không đọc thì ghi vào đây, không sửa `pd.py`). Script in `L0 luật riêng dự án` nếu có file.
   - `L1 TRÙNG NỘI DUNG` — chỉ giữ 1 bản trong số việc đọc; ghi NOTES.
   - `L1 CÙNG TÊN KHÁC NGÀY/REV` — mở đầu mỗi file xem ngày/REV, **chỉ đọc bản mới nhất**; ghi NOTES "bản cũ <tên> không dùng" (không xóa file).
   - **PDF ưu tiên; CAD (`.dwg`) chỉ là dự phòng.** Với mỗi nhóm CAD, nếu cùng thư mục cha / cùng tên đã có PDF in ra → ghi NOTES "CAD <nhóm> đã có PDF, không quét". Chỉ khi một hạng mục/tờ **không có PDF** mới liệt kê CAD đó là "cần PDF hoặc quét CAD" trong NOTES để điều phối hỏi user.
2. Mở `FILES.kiemke.json`, rà cột `Loại` (máy đoán theo tên file). Sai thì sửa đúng ô đó trong file. Loại hợp lệ: `pd.py cot --buoc kiemke`.
3. Mail và sheet tổng BoQ đã đọc ở mục A — dùng lại, không đọc lại. Không có mail thì đọc trang bìa / Drawing List.
4. Ghi `PROJECT.kiemke.json`, mỗi trường 1 dòng, cột `Nguồn` là link tới chỗ đọc được:

| Trường | Ghi gì |
|---|---|
| Tên dự án | như trên bìa / khung tên |
| CĐT | |
| Tư vấn thiết kế | kiến trúc / nội thất |
| Gói thầu | phạm vi hồ sơ mời (BoQ / mail) |
| Đợt hồ sơ | `<DOT>`: các nhóm file chính + số tờ |
| Người yêu cầu | tên, ngày mail |
| Hạn trả giá | ngày, giờ (nếu có) |
| Yêu cầu báo giá | phương án, option, ghi chú đặc biệt trong mail — chép sát ý |
| Phạm vi tính giá | phần công ty mình làm / không làm, nếu mail hoặc BoQ nói rõ |

Không thấy thông tin → `Giá trị` = `chưa có`, không suy ra.
5. **Làm lười (không làm ở bước này):** trang PDF `Lớp text = KHÔNG` cứ để trống `Số BV`/`Tên BV`. Chỉ đọc khung tên bằng Gemini cho **trang ứng viên** của các mã (bước 4 làm, sau `ungvien`), không đọc hết. Nếu vẫn muốn làm ngay (hồ sơ nhỏ, < 20 trang scan), cách làm: xuất ảnh góc khung tên (`anh … --dpi 150 --clip 0.6,0.6,1,1`), rồi cho Gemini đọc một loạt (`gemini a.png b.png … --hoi "Read the drawing number, drawing title and revision in the title block. Output JSON {number, title, rev}."`). Ghi vào `PAGES.kiemke.json` (cột Số BV, Tên BV) và thêm vào NOTES dòng `"Số/tên tờ trang scan do Gemini đọc"`. Hơn 60 trang → chỉ làm các file thuộc loại Bản vẽ, ghi số trang còn lại vào NOTES.
6. Điều đáng chú ý (file hỏng, file trùng tên khác bản, bộ bản vẽ thiếu, file > 200 MB chưa đọc) → `NOTES.kiemke.json`.

## Xong khi
`py -I "<TOOL>" check <GOC> <DOT> --den kiemke` không còn LỖI.
