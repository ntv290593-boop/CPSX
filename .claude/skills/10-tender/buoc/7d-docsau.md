# Bước 7d — docsau: đọc bản vẽ PDF theo cụm (bước dễ sai nhất — làm chậm, có bằng chứng)

**Ghi (file `*.docsau-<nhóm>.json`):** `QUAN_SAT`, `SCOPE` (`Bản vẽ đối chứng`, `Kích thước BV`, `Kích thước BoQ`, `Phạm vi nhà máy`), `BOQ_MAP` (`KL theo BV`, `Kích thước BV`, `Tờ đếm KL`, `Tờ chi tiết`), `DAY_TUONG` (chỉ cửa), `PAGES` (trang đã mở), `ISSUES` (ID `V<nhóm>-01`…).
**Chỉ làm phần được giao** (khối `Lượt 0 · Nhóm <n>` hoặc `Nhóm <n>` trong `_data\_docsau.txt`). Không tính giá. Không sửa `Bản vẽ hiệu lực`, `Mã VL`: thấy sai thì ghi issue.

Cụm = các dòng BOQ cùng mô tả, mã VL, mã FF&E và bộ tờ. Đọc 1 lần, khóa = `Cụm đọc` = `Khóa` của dòng đầu tiên. `SCOPE` ghi ở dòng đại diện, dòng cùng cụm tự nhận. `BOQ_MAP` thì ghi từng khu của từng dòng, vì khối lượng khác nhau: xem bằng `loc <GOC> BOQ_MAP "Khóa=<khóa>"`.

**Toàn gói:** cụm nào cũng đọc, kể cả sơn, thạch cao, sàn, MEP. Trường không áp dụng cho loại đó (sơn không có `Có cánh`) thì bỏ, không ghi `KHÔNG CÓ` cho có. Tối thiểu mọi cụm: `Kích thước BV` hoặc phạm vi/diện tích thể hiện trên tờ nếu có, `Mã hoàn thiện trên BV`, `Khu`, `Bộ tờ đã đọc`, `Tờ chưa đọc`, `Kiểm chéo Gemini`.

## Ghi file (đủ cho 7d — không cần đọc JSON.md)
- Nơi ghi: `<GOC>\10_TENDER\_data\<SHEET>.docsau-<nhóm>.json`, chỉ file của nhóm mình. Không sửa `*-tool.json`, `*.user.json`, file của nhóm/bước khác. Không tự ghi xlsx.
- Dạng: `{"rows": [ {...}, ... ]}`. Cột đúng tên (có dấu) như danh sách "Cột được ghi" trong prompt (`cot --buoc docsau`). Không có giá trị → **bỏ cột**, không ghi `""`, `0`, `"?"`. Xóa giá trị nháp sai của tool → `"(xóa)"`.
- **Nhóm dòng chung** (QUAN_SAT, TRANG_QS nhiều trường cùng cụm/món/tờ): `{"_chung": {"Cụm đọc": "...", "Nguồn": "HÌNH", "Tờ · Chi tiết": [["FUR-3315 p5 D13", "F0006", 5]]}, "dong": [{"Trường": "Kích thước BV", "Giá trị": "...", "Trạng thái": "CÓ"}, ...]}` → mỗi dòng = chung + riêng (riêng thắng). Viết thẳng bằng Write; **không viết script Python sinh JSON**. Sửa bằng Edit.
- SCOPE, BOQ_MAP: sheet có khóa (`Khóa`, `Khóa KL`), chỉ ghi khóa đã có (xem bằng `xemcum`); không thêm dòng mới. PAGES khóa `Khóa trang`, ISSUES khóa `ID` (`V<nhóm>-01`…), DAY_TUONG khóa `Khóa DT`.
- Link: `["nhãn", "FID", trang]`, nhiều link = list các link; nhãn ngắn có số tờ + chi tiết (`"FUR-3315 p5 D13"`). Không có nguồn → chuỗi `"KHÔNG THẤY – đã tìm ở …"`.
- DS đúng 1 giá trị trong danh sách; số ghi dạng số; nhiều giá trị trong 1 ô cách `; `.

## Hai lượt
**Lượt 0 — trang dùng chung** (khối `Lượt 0 · Nhóm <n>`: các `TRANG <FID>|<trang>` và danh sách món cần tìm). Trang này có ≥ 3 cụm dùng nên đọc **1 lần cho mọi cụm**, trước lượt 1.
- Mỗi trang: `xemtrang <GOC> <FID> <trang>` (đường dẫn PDF, khung tên, lớp chữ, các cụm dùng trang kèm mô tả) → ảnh cả trang 1 lần → phóng từng chi tiết có món trong danh sách. Ghi `TRANG_QS.docsau-<n>.json`: 1 dòng = 1 trường của 1 món, `Khóa trang`, `Chi tiết` (số chi tiết/hình), `Món` (nhãn/tên trên hình), `Cụm đọc` (các cụm khớp món, cách `; `; chưa chắc thì để trống), `Trường`, `Giá trị`, `Trạng thái`, `Tờ · Chi tiết`, `Nguồn` HÌNH/TEXT, `Crop` (tham số `anh` đã dùng: `<trang> --clip … --dpi …`). Trường như mục "Làm" bước 3–4, kể cả kiểm chuỗi kích thước.
- Gemini: chạy `gemini <crop> --goc <GOC>` (chế độ số, không `--cum`) và `--quansat "<Món>" --goc <GOC>` trên crop đã dùng, để cache sẵn cho lượt 1. Lệch thì ghi cả 2 dòng (`HÌNH`, `GEMINI`) `MÂU THUẪN`.
- Ghi `PAGES.docsau-<n>.json` cho mọi trang đã mở. Không ghi QUAN_SAT, SCOPE, BOQ_MAP (lượt 1 làm). Trang không có món nào trong danh sách: 1 dòng `Món = (không có)`, `Trạng thái = KHÔNG THẤY`, lý do ở `Giá trị`.

**Lượt 1 — theo cụm** (khối `Nhóm <n>`). `Tờ:` chỉ còn tờ riêng; `Lượt 0 đã đọc:` là các trang đã có quan sát.
- Mở đầu mỗi cụm: `xemcum <GOC> <cụm>` (1 lệnh: mô tả, dòng + KL, MAT, tờ + đường dẫn PDF, TRANG_QS các trang lượt 0, những gì đã ghi). Trang lượt 0: **không mở lại ảnh**, dùng TRANG_QS `xemcum` in ra (cả trang, vì `Cụm đọc` có thể trống; `*` = dòng đã gắn cụm này), chép giá trị đúng món vào QUAN_SAT của cụm với cùng `Tờ · Chi tiết`, `Nguồn`. Chỉ mở lại khi TRANG_QS thiếu đúng trường cụm cần, hoặc món không rõ thuộc cụm nào; mở thì cắt lại đúng `Crop` để trúng cache Gemini.
- `gemini --cum <cụm> --goc <GOC>` vẫn chạy cho mọi crop dùng làm bằng chứng của cụm (kể cả crop lượt 0: cắt lại đúng `Crop` rồi đưa thẳng cho `gemini`, **không mở PNG bằng Read**): trúng cache, không gọi agy, nhưng ghi `_gemini_so\<cụm>.json` cho bước 7e.

## Nguyên tắc
1. **Nhìn hình trước, đọc mô tả BOQ sau.** Điền quan sát từ hình xong mới so với `Mô tả BOQ`, để không "thấy điều mình chờ".
2. **Mỗi giá trị có tờ + số chi tiết/hình.** Không có số chi tiết → chưa đủ bằng chứng cho trường cấu tạo.
3. **Không có ≠ không thấy.** `KHÔNG CÓ` = nhìn rõ là không có. `KHÔNG THẤY` = chưa có tờ / chưa đọc được. Không bỏ trống trường quyết định giá, không đoán.
4. **Lệch thì giữ cả hai.** BOQ một số, bản vẽ một số → 2 dòng QUAN_SAT (Nguồn BOQ, HÌNH), trạng thái `MÂU THUẪN`, cộng 1 issue. Không tự chọn bên thắng.
5. **Khối lượng đếm trên bản vẽ** (bảng thống kê, mặt bằng từng khu). Không lấy KL BOQ chia, không lấy diện tích chia kích thước. Hạng mục m2/md không có bảng thống kê: bỏ trống `KL theo BV`, không tự đo (trạng thái CHƯA ĐỌC, độ tin cậy tối đa TB), không cần issue riêng cho từng dòng.
6. Đọc kỹ 20 kịch bản sai và cách chặn: `D:\COSTING EXCEL\_MASTER\modules\SCAN-NHAN-LOAI.md` mục 4 (chỉ đọc mục 4).

## Công cụ — chỉ dùng các lệnh này
Không tự viết script (Python, sed, grep, find, cat, Select-String) để đọc `_data`, PDF, log Gemini hay mã tool: mỗi lượt tự viết đọc lại ~110k token ngữ cảnh. Thiếu lệnh thì ghi vào kết quả trả về. Gemini chạy thường (đợi tới ~5 phút), không chạy nền rồi `sleep`.
| Việc | Lệnh `py -I "<TOOL>" …` |
|---|---|
| Mọi thứ của 1 cụm | `xemcum <GOC> <cụm hoặc khóa dòng>` |
| Mọi thứ của 1 trang | `xemtrang <GOC> <FID> <trang> [--chu]` (`--chu` = cả lớp chữ) |
| Lọc dữ liệu đã ghép | `loc <GOC> <SHEET> "cột=gt" "cột~regex" --cot a,b` |
| Tìm trang có mã / từ khóa | `tim <GOC> --ma WD-01 [--file s]`, `tim <GOC> "<regex>"` |
| Trang ứng viên của dòng | `ungvien <GOC> --khoa "<Khóa>"` |
| Giải nén file trong zip | `zip <GOC> <FID>` |
| Xem trang (ảnh) | `anh "<pdf>" <trang> "<thư mục ảnh trong prompt>\p13.png" --dpi 100` rồi mở PNG bằng Read (đường dẫn PDF lấy từ `xemcum`/`xemtrang`) |
| Phóng một vùng | `anh … --dpi 220 --clip 0.2,0.1,0.6,0.5` (tỉ lệ 0–1 theo chiều nhìn). Tool tự hạ dpi: cả trang ≤ 1000 px cạnh dài, vùng cắt ≤ 1,15 MP. In `(giới hạn ảnh …)` = vùng quá to so với dpi xin: cắt vùng nhỏ hơn, không xin dpi cao hơn |
| Chữ trong vùng | `chu "<pdf>" <trang> --clip …` |
| Số kích thước + vị trí | `so "<pdf>" <trang> --clip …` (chỉ khi trang có lớp chữ) |
| Gemini đọc dãy số của 1 cụm | `gemini "<crop.png>" … --cum "<Cụm đọc>" --goc <GOC>` (lưu `_gemini_so\<cụm>.json`, in khớp / chỉ Claude; số Claude lấy từ QUAN_SAT) |
| Gemini quan sát đủ bộ trường | `gemini "<crop.png>" --quansat "Tủ lavabo dưới chậu LA1" --goc <GOC>` |
| Gemini hỏi riêng | `gemini "<ảnh.png>" --hoi "How many door leaves does the cabinet have?" --goc <GOC>` |

Trang không có lớp chữ hoặc số vẽ bằng nét → đọc bằng hình, phóng to từng vùng, rồi kiểm chuỗi kích thước.

## Gemini = người đọc thứ hai (bắt buộc với mọi crop dùng làm bằng chứng)
Gemini nhìn bản vẽ tốt: nhận đúng cấu kiện, nhãn, đợt/kệ, phần ngoài phạm vi, dãy số in trên hình. Nhưng đã thấy nó **gán nhầm số vào chiều** và **tự trừ ra số không in**; đếm số lượng dài cũng hay sót. Vì vậy:
1. **Độc lập:** Claude đọc crop và ghi giá trị của mình vào `QUAN_SAT.docsau-<nhóm>.json` **trước** (`--cum` dừng nếu cụm chưa có `Kích thước BV` dạng R×S×C), rồi mới chạy Gemini trên **đúng các crop đó**. Không đưa số hay nhận xét của Claude vào câu hỏi cho Gemini.
2. **Mỗi cụm:** chạy `--quansat` trên crop mặt đứng/chi tiết chính, và chế độ số (`--cum`) trên mọi crop có dãy kích thước; chạy lại khi thêm crop thì kết quả được cộng dồn. Nhiều crop thì truyền một lần. Luôn kèm `--goc <GOC>`: kết quả Gemini lưu ở `_data\_gemini_cache` theo **nội dung ảnh + câu hỏi**, dùng chung mọi nhóm. Crop dùng chung nhiều cụm: vẫn chạy `--cum` cho từng cụm (để ghi `_gemini_so\<cụm>.json`), lần sau trúng cache, không gọi agy lại. Chỉ trúng khi ảnh giống hệt byte (cùng trang, cùng `--clip`, cùng dpi): agent khác muốn dùng lại thì truyền đúng file PNG đã cắt, hoặc cắt lại đúng tham số đó. `--quansat` cùng món + cùng ảnh cũng dùng lại; đổi tên món là câu hỏi khác, gọi lại.
3. **So từng trường:** khớp → giữ dòng của Claude. Lệch → phóng to đúng chỗ đó đọc lại; vẫn lệch → giữ 2 dòng (Nguồn `HÌNH` và `GEMINI`), `MÂU THUẪN`, cộng 1 issue. Gemini thấy món mà Claude ghi `KHÔNG THẤY` → đọc lại theo `where` Gemini chỉ.
4. **Không nhận** số Gemini tự tính ra (không in trên hình), và không dùng Gemini làm nguồn duy nhất cho kích thước hay số lượng.
5. Mỗi cụm 1 dòng: `Trường = Kiểm chéo Gemini`, `Nguồn = GEMINI`, `Giá trị` = `"khớp: <các trường>; lệch: <các trường hoặc không>; crop: <tờ/chi tiết>"`, cho **các trường không phải số**. Phần số do tool so (bước 7e tự sinh `Kiểm chéo Gemini số`), bạn không ghi. Tool chỉ so theo tập số, nên việc gán đúng chiều vẫn phải tự kiểm ở mục 3.
6. agy lỗi / không có → ghi `Kiểm chéo Gemini` trạng thái `KHÔNG THẤY`, giá trị `KHÔNG THẤY – <lỗi>`, làm tiếp; cụm đó không được CAO.

## Tiết kiệm token — lọc trước khi mở ảnh
1. **PDF trước, CAD sau cùng.** Chỉ đụng `.dwg` khi cụm không có tờ nào trong PDF; ghi `Tờ chưa đọc` nêu rõ "chỉ có CAD".
2. **Bắt đầu từ `Bản vẽ hiệu lực`** bước 6 đã gán. Chỉ khi trống hoặc sai tờ mới `ungvien`, `tim` thêm. Không mở cả file / cả bộ.
3. **Ảnh từ rẻ tới đắt:** cả trang (tool giữ ≤ 1000 px) để xác nhận tờ, vị trí chi tiết → phóng vùng chi tiết (150–220 dpi, `--clip` đúng khung chi tiết) chỉ khi cần số. Mỗi ảnh nằm lại trong ngữ cảnh tới hết agent và bị tính lại ở mọi lượt sau: ít ảnh, vùng vừa đủ. Trang có lớp chữ thì `chu`/`so` trước, ảnh sau.
4. Trang đã đọc ghi vào `PAGES` — lần chạy lại dùng danh sách này, không đọc lại.

## Làm, cho từng cụm
1. **Bộ tờ:** từ `Bản vẽ hiệu lực`, theo callout (vd `18 / ID-PS-TY-712`) tới tờ chi tiết. Đủ bộ thường gồm: mặt bằng, mặt đứng, mặt cắt/chi tiết, bảng thống kê, legend vật liệu. Có tờ chưa đọc → trường `Tờ chưa đọc` (ghi `không còn` nếu đủ).
2. **Đúng tờ:** số tờ trên khung tên khớp tờ được dẫn tới; cùng tên nhưng khác type/khu thì không dùng lẫn.
3. **QUAN_SAT** — một dòng mỗi trường, `Nguồn = HÌNH` hoặc `TEXT`, có `Tờ · Chi tiết`: `Lắp đặt` · `Kích thước BV` (R×S×C; chữ L thì từng nhánh) · `Có cánh` · `Ngăn kéo` · `Thanh treo` · `Kệ/đợt` · `Hốc/ô tivi` · `Mặt đá` · `Gương` · `Kính` · `Kim loại` · `Giấy/vải` · `Mã hoàn thiện trên BV` · `Ván nền ẩn` (khi có giấy/gương/kim loại phủ ngoài) · `Khung xương` · `Cong/bo` · `Phần ngoài phạm vi` · `Khu` · `Bộ tờ đã đọc` · `Tờ chưa đọc` · `Revision tờ`. Một dòng BOQ gồm nhiều đồ vật → `Thành phần` 1, 2…
4. **Kiểm chuỗi kích thước:** tổng các đoạn = kích thước tổng; ghi vào `Giá trị` (vd `"450+450=900 ✓"`). Không khớp → `MÂU THUẪN` + issue. `check` báo LỖI nếu phép cộng sai.
4b. **Kiểm chéo Gemini** theo mục trên, cho mọi crop đã dùng ở bước 3–4.
5. **Cửa:** `DAY_TUONG` (`Khóa DT` = `<cụm>|<vị trí>`): dày tường tại **từng vị trí thật** (mặt bằng / ghi chú loại tường), không lấy một tờ điển hình cho mọi vị trí.
6. **BOQ_MAP:** `KL theo BV` từng khu + `Tờ đếm KL`; `Tờ chi tiết`. Đếm trên mặt bằng (không có bảng thống kê) → nhờ Gemini đếm cùng tờ (`--hoi "Count every <món> symbol/tag <mã> on this plan; list each with its room"`). Hai số khác nhau → đếm lại theo danh sách phòng; vẫn khác → issue.
7. **SCOPE (dòng đại diện):** `Bản vẽ đối chứng` = tờ thể hiện rõ nhất đứng đầu; `Kích thước BV` = R×S×C (chưa rõ chiều nào → `?`); `Kích thước BoQ` nếu mô tả BOQ có ghi; `Phạm vi nhà máy` khi tool chưa gắn: `TRONG SCOPE – <mộc|sơn đồ gỗ|cửa|cabinet|loose>` / `KHÔNG THUỘC SCOPE NHÀ MÁY – <nhóm>` / `CHƯA RÕ – <lý do>`. Đây chỉ là nhãn, không bỏ dòng nào, không đổi `Trạng thái`.
8. **PAGES** (`Khóa trang` = `<FID>|<trang>`, kèm `File ID`, `Trang`): `Loại tờ`, `Mã xuất hiện`, `Đã đọc bằng`, `Rev`, `Số BV`, `Tên BV`, `Lớp text`.
9. **ISSUES (V…):** thiếu tờ, BV ↔ BOQ lệch (kích thước, số lượng, cấu tạo), mã hoàn thiện BV ≠ MAT, tờ khác revision.

## Trước khi trả kết quả
- Mỗi cụm có ≥ 1 dòng QUAN_SAT nguồn HÌNH/TEXT (hoặc `KHÔNG THẤY` có lý do), `Bản vẽ đối chứng`, `Tờ chưa đọc`, `Kiểm chéo Gemini`.
- `check --den docsau` không còn LỖI ở file và cụm của bạn.
- Chạy `py -I "<TOOL>" neo <GOC>`: máy kiểm số tờ trong nhãn link, mã/model ở link spec, `Mã hoàn thiện trên BV` / `Kích thước BV` có trong lớp chữ các trang được link. Bản vẽ vẽ nét ra KHÔNG KIỂM ĐƯỢC; kích thước chỉ lấy chuỗi R×S×C **đầu tiên**, nên ghi chuỗi đó trước, diễn giải sau. Số tự cộng không có trên bản vẽ sẽ ra LỆCH: kiểm lại và ghi rõ vào `Giá trị`. Dòng `NEO LỆCH` của cụm bạn: mở lại đúng trang, sửa link hoặc giá trị; nếu đúng thì ghi lý do. LỆCH chỉ là XEM.
- Trả về: số cụm đã làm, các cụm còn KHÔNG THẤY ở trường quyết định giá (kích thước, số cánh/ngăn kéo, vật liệu).
