# Bước 4 — banve: đọc bản vẽ PDF (bước dễ sai nhất — làm chậm, có bằng chứng)

**Ghi:** `QUAN_SAT`, `BOQ_TONGHOP` (`Bản vẽ đối chứng`, `Kích thước BV`), `BOQ_MAP` (KL theo BV, tờ), `DAY_TUONG` (chỉ cửa), `PAGES` (các trang đã dùng), `ISSUES` (ID `V<nhóm>-01`…) — file `*.banve-<nhóm>.json`.
**Chỉ làm các mã được giao.** Không phân loại mẫu Excel (bước takeoff làm), không tính giá.

## Nguyên tắc
1. **Nhìn hình trước, đọc mô tả BoQ sau.** Chỉ dùng `Mã` + `Hạng mục` để tìm tờ. Điền quan sát từ hình xong mới mở `Mô tả BoQ`, để không "thấy điều mình chờ".
2. **Mỗi giá trị có tờ + số chi tiết/hình.** Không có số chi tiết → chưa đủ bằng chứng cho trường cấu tạo.
3. **Không có ≠ không thấy.** `KHÔNG CÓ` = nhìn rõ là không có. `KHÔNG THẤY` = chưa có tờ / chưa đọc được. Không bỏ trống, không đoán.
4. **Lệch thì giữ cả hai.** BoQ một số, bản vẽ một số → 2 dòng QUAN_SAT (Nguồn BOQ, HÌNH), trạng thái `MÂU THUẪN`, cộng 1 issue. Không tự chọn bên thắng.
5. **Khối lượng đếm trên bản vẽ** (bảng thống kê, mặt bằng từng khu). Không lấy KL BoQ chia, không lấy diện tích chia kích thước.
6. Đọc kỹ 20 kịch bản sai và cách chặn: `D:\COSTING EXCEL\_MASTER\modules\SCAN-NHAN-LOAI.md` mục 4 (chỉ đọc mục 4).

## Công cụ (đừng tự viết script đọc PDF)
| Việc | Lệnh `py -I "<TOOL>" …` |
|---|---|
| Tìm trang có mã / từ khóa (regex) | `tim <GOC> <DOT> "VANITY" [--file "MASS"]` |
| Xem trang (ảnh) | `anh "<pdf>" <trang> "<scratchpad>\p13.png" --dpi 100` rồi mở PNG bằng Read |
| Phóng một vùng | `anh … --dpi 220 --clip 0.2,0.1,0.6,0.5` (tỉ lệ 0–1 theo chiều nhìn) |
| Chữ trong vùng | `chu "<pdf>" <trang> --clip …` |
| Số kích thước + vị trí | `so "<pdf>" <trang> --clip …` (chỉ khi `Lớp text` của trang = CÓ) |

| Gemini đọc dãy số kích thước | `gemini "<crop.png>" … --claude "900,450,450"` (in KHỚP / chỉ Claude / chỉ Gemini) |
| Gemini quan sát đủ bộ trường | `gemini "<crop.png>" --quansat "Tủ lavabo dưới chậu LA1 (V01)"` |
| Gemini hỏi riêng | `gemini "<ảnh.png>" --hoi "How many door leaves does the cabinet have?"` |

Trang `Lớp text = KHÔNG` hoặc số vẽ bằng nét → đọc bằng hình, phóng to từng vùng, rồi kiểm chuỗi kích thước (bên dưới). Bộ hồ sơ kiểu SOFITEL (số vẽ bằng nét vector) có bộ giải mã `D:\COSTING EXCEL\_MASTER\scan\pdf_dim_decode.py`; chỉ dùng khi điều phối yêu cầu.

## Gemini = người đọc thứ hai (bắt buộc với mọi crop dùng làm bằng chứng)
Gemini (qua agy) nhìn bản vẽ tốt: nhận đúng cấu kiện, nhãn, đợt/kệ, phần ngoài phạm vi, dãy số in trên hình. Nhưng đã thấy nó **gán nhầm số vào chiều** (lấy sâu 550 làm rộng) và **tự trừ ra số không in** (V01: "cao 300"); đếm số lượng dài cũng hay sót. Vì vậy:
1. **Độc lập:** Claude đọc crop và ghi giá trị của mình **trước**, rồi mới chạy Gemini trên **đúng các crop đó**. Không đưa số hay nhận xét của Claude vào câu hỏi cho Gemini (lệnh `--quansat` chỉ nhận tên món).
2. **Mỗi mã:** chạy `--quansat` trên crop mặt đứng/chi tiết chính, và chạy chế độ số (`--claude "<các số Claude đọc>"`) trên mọi crop có dãy kích thước. Nhiều crop thì truyền một lần, script tự chạy song song.
3. **So từng trường:**
   - Khớp → giữ dòng của Claude.
   - Lệch → Claude phóng to đúng chỗ đó đọc lại. Vẫn lệch → giữ 2 dòng (Nguồn `HÌNH` và `GEMINI`), trạng thái `MÂU THUẪN`, cộng 1 issue.
   - Gemini thấy món mà Claude ghi `KHÔNG THẤY` → đọc lại theo `where` Gemini chỉ.
4. **Không nhận** số Gemini tự tính ra (không in trên hình), và không dùng Gemini làm nguồn duy nhất cho kích thước hay số lượng.
5. Mỗi mã 1 dòng tóm tắt: `Trường = Kiểm chéo Gemini`, `Nguồn = GEMINI`, `Giá trị` = `"khớp: <các trường>; lệch: <các trường>; crop: <tờ/chi tiết>"`. Bước 5 dùng dòng này để chấm độ tin cậy.
6. agy lỗi / không có → ghi `Kiểm chéo Gemini = KHÔNG THẤY – <lỗi>`, làm tiếp; độ tin cậy mã đó không được CAO.

## Tiết kiệm token — lọc trước khi mở ảnh (làm đúng thứ tự)
1. **PDF trước, CAD sau cùng.** Chỉ đụng `.dwg` khi mã đó không có tờ nào trong PDF (sau khi đã `tim` + Drawing List). Dùng CAD thì ghi `Tờ chưa đọc`/issue nêu rõ "chỉ có CAD".
2. **Bắt đầu từ `Ứng viên` điều phối đưa** (kết quả `ungvien`, chia nhóm: đọc `CHI TIẾT` trước; `MẶT BẰNG` để đếm KL theo khu; `THỐNG KÊ` kèm với cửa; `DANH MỤC` để biết tờ; `CỦA MÃ CHA` chỉ để tìm — vẫn đọc riêng từng mã) — như user đọc BoQ rồi mới tìm tờ. Chỉ khi mã `KHÔNG CÓ ỨNG VIÊN` hoặc ứng viên sai tờ mới tự `tim` thêm. Trang scan (`Lớp text = KHÔNG`) trong ứng viên thì đọc khung tên crop nhỏ để xác nhận số tờ và ghi vào `PAGES`.
   **Chỉ mở trang đã khoanh:** lấy danh sách trang từ `tim` (mã, tên tiếng Việt/Anh, mã vật liệu) + Drawing List; tên file/tờ kiểu `501_…@TỦ GIÀY` cho biết tờ ứng mã nào. **Không mở cả file / cả bộ.** Loại trang bìa, mặt bằng kiến trúc chung, MEP, trần/điện khi mã không cần.
3. **Nhóm Loose Furniture (LF-xx) nằm trong phạm vi** — vẫn làm theo mã BoQ, nhưng chỉ mở tờ của chính mã đó.
4. **Ảnh từ rẻ tới đắt:** khung tên/crop nhỏ 100 dpi để xác nhận tờ → phóng vùng chi tiết (150–220 dpi, `--clip`) chỉ khi cần số. Trang `Lớp text = CÓ` thì dùng `chu`/`so` trước, ảnh sau. Không xuất cả trang ở dpi cao.
5. Trang đã đọc ghi vào `PAGES` (bước 8) — lần chạy lại dùng danh sách này, không đọc lại.

## Làm, cho từng mã
1. **Tìm bộ tờ:** `tim` theo mã + tên; theo dấu callout (vd `18 / ID-PS-TY-712`) tới tờ chi tiết; Drawing List. Đủ bộ thường gồm: mặt bằng (vị trí, số lượng), mặt đứng, mặt cắt/chi tiết, bảng thống kê, legend vật liệu. Có tờ trong bộ chưa đọc → ghi trường `Tờ chưa đọc`.
2. **Đúng tờ:** số tờ trên khung tên khớp với tờ được dẫn tới; cùng tên nhưng khác type/khu (type 2 ≠ type 3) thì không dùng lẫn. Khu của tờ phải là khu có KL của dòng BoQ.
3. **QUAN_SAT** — một dòng mỗi trường, `Nguồn = HÌNH` (đọc trên hình) hoặc `TEXT` (lớp text). Các trường:
   `Lắp đặt` (gắn tường / đứng rời / âm tường) · `Kích thước BV` (R×S×C; chữ L thì từng nhánh) · `Có cánh` (số, mở/lùa) · `Ngăn kéo` (số) · `Thanh treo` · `Kệ/đợt` · `Hốc/ô tivi` · `Mặt đá` · `Gương` · `Kính` · `Kim loại` · `Giấy/vải` · `Mã hoàn thiện trên BV` (nhãn WD/PT/LAM… đọc trên hình từng tờ) · `Ván nền ẩn` (bắt buộc khi có giấy/gương/kim loại phủ ngoài) · `Khung xương` · `Cong/bo` · `Phần ngoài phạm vi` · `Khu` · `Bộ tờ đã đọc` · `Tờ chưa đọc` · `Revision tờ`.
   Một dòng BoQ gồm nhiều đồ vật (vd vanity + gương) → `Thành phần` 1, 2…
4. **Kiểm chuỗi kích thước:** tổng các đoạn = kích thước tổng; ghi kết quả vào `Giá trị` (vd `"450+450=900 ✓"`). Không khớp → `MÂU THUẪN` + issue.
4b. **Kiểm chéo Gemini** theo mục trên, cho mọi crop đã dùng ở bước 3–4.
5. **Cửa** (`Nhóm = Cửa`): DAY_TUONG — dày tường tại **từng vị trí thật** của mã (mặt bằng tường xây / ghi chú loại tường), không lấy một tờ điển hình cho mọi vị trí.
6. **BOQ_MAP:** `KL theo BV` từng khu + `Tờ đếm KL`; `Kích thước BV`; `Tờ chi tiết`. Không đếm được → bỏ trống KL + issue nêu lý do. Đếm trên mặt bằng (không có bảng thống kê) → nhờ Gemini đếm cùng tờ (`--hoi "Count every <món> symbol/tag <mã> on this plan; list each with its room"`). Hai số khác nhau → đếm lại theo danh sách phòng; vẫn khác → issue.
7. **BOQ_TONGHOP:** `Bản vẽ đối chứng` = tờ thể hiện rõ nhất (mặt đứng/chi tiết) đứng đầu, sau đó các tờ khác; `Kích thước BV` = R×S×C (chưa rõ chiều nào → `?`).
8. **PAGES:** với mọi trang đã mở: `Loại tờ`, `Mã xuất hiện`, `Đã đọc bằng` (HÌNH/TEXT), `Rev`; sửa `Số BV`/`Tên BV` nếu máy đoán sai.
9. **ISSUES (V…):** thiếu tờ, BV ↔ BoQ lệch (kích thước, số lượng, cấu tạo), mã hoàn thiện BV ≠ ML, tờ khác revision.

## Trước khi trả kết quả
- Mỗi mã có ≥ 1 dòng QUAN_SAT nguồn HÌNH, có `Bản vẽ đối chứng`, có `Tờ chưa đọc` (ghi `không còn` nếu đã đủ), có dòng `Kiểm chéo Gemini`.
- `check --den banve` không còn LỖI ở file của bạn.
- Chạy `py -I "<TOOL>" neo <GOC> <DOT>`: máy kiểm số tờ ghi trong nhãn link (`Bản vẽ đối chứng`, `Tờ · Chi tiết`, `Tờ đếm KL`…) có trên trang được link, mã/model ở link SPEC/ML có trên trang/dòng được link, và QUAN_SAT `Mã hoàn thiện trên BV` / `Kích thước BV` (cả `Kích thước BV` ở BOQ_TONGHOP, BOQ_MAP) có trong lớp text các trang được link. Mã/kích thước chỉ kiểm khi trang có loại đó ở dạng text (bản vẽ vẽ nét ra KHÔNG KIỂM ĐƯỢC); kích thước chỉ lấy chuỗi R×S×C **đầu tiên** của giá trị, nên ghi chuỗi đó trước, diễn giải sau. Số tự cộng/suy ra (khung bao ngoài, tổng các đoạn) không có trên bản vẽ sẽ ra LỆCH: kiểm lại phép tính và ghi rõ vào `Giá trị` (vd `330 = 50 chân + 280 thành`). Dòng `NEO LỆCH` của mã bạn làm: mở lại đúng trang (dùng `gợi ý`), sửa link hoặc giá trị; nếu đúng thì ghi lý do vào `Giá trị`. LỆCH chỉ là XEM, không phải LỖI. Ghi số tờ trong nhãn (vd `F005 p43 604.1 …` hoặc đủ `ID-CO-TY-703`) để máy kiểm được; nhãn chỉ có `p30` ra KHÔNG KIỂM ĐƯỢC.
- Trả về: số mã đã làm, các mã còn KHÔNG THẤY ở trường quyết định giá (kích thước, số cánh/ngăn kéo, vật liệu).
