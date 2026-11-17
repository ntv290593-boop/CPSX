# Bước 7d — docsau: đọc bản vẽ THEO TRANG (bước dễ sai nhất — làm chậm, có bằng chứng)

**Luồng:** `docsau` (tool lập danh sách cần kiểm) → **nhóm Trang / Tìm tờ** (LLM đọc ảnh, mỗi trang 1 lần, ghi `TRANG_QS`) → `ghep` (tool ghép về cụm: `QUAN_SAT`, `SCOPE`, `BOQ_MAP`, `DAY_TUONG`) → **nhóm Đối chiếu** (LLM so với BOQ/MAT, không mở ảnh trừ chỗ thiếu) → `check --den docsau`.
**Chỉ làm phần được giao** (khối `Trang · Nhóm <n>`, `Tìm tờ · Nhóm <n>` hoặc `Đối chiếu · Nhóm <n>` trong prompt `giao`). Không tính giá. Không sửa `Bản vẽ hiệu lực`, `Mã VL`: thấy sai thì ghi issue.

Cụm = các dòng BOQ cùng mô tả, mã VL, mã FF&E và bộ tờ; khóa = `Cụm đọc` = `Khóa` của dòng đầu. Mỗi cụm có **hợp đồng kiểm**: các trường bắt buộc theo loại món (`form\PROFILE.json`, chọn theo `Mã cấp 3`). Trường bắt buộc chưa có kết luận = `check` báo LỖI, `tincay` hạ Độ tin cậy. **Toàn gói:** cụm nào cũng đọc, kể cả sơn, thạch cao, sàn, MEP.

## Ghi file (đủ cho 7d — không cần đọc JSON.md)
- Nơi ghi: `<GOC>\10_TENDER\_data\<SHEET>.docsau-<nhóm>.json`, chỉ file của nhóm mình (nhóm trang: chỉ qua lệnh `luu`). Không sửa `*-tool.json`, `*.user.json`, file của nhóm/bước khác. Không tự ghi xlsx.
- Dạng: `{"rows": [ {...}, ... ]}`. Cột đúng tên (có dấu) như "Cột được ghi" trong prompt. Không có giá trị → **bỏ cột**, không ghi `""`, `0`, `"?"`. Xóa giá trị nháp sai của tool → `"(xóa)"`.
- **Nhóm dòng chung** (nhiều trường cùng trang/món): `{"_chung": {"Khóa trang": "F0006|5", "Món": "LA1", "Cụm đọc": "...", "Nguồn": "HÌNH", "Tờ · Chi tiết": [["FUR-3315 p5 D13", "F0006", 5]]}, "dong": [{"Trường": "Kích thước BV", "Giá trị": "...", "Trạng thái": "CÓ"}, ...]}` → mỗi dòng = chung + riêng (riêng thắng). Viết thẳng bằng Write; **không viết script Python sinh JSON**. Sửa bằng Edit.
- SCOPE, BOQ_MAP: chỉ ghi khóa đã có (`Khóa`, `Khóa KL`), không thêm dòng. PAGES khóa `Khóa trang`, ISSUES khóa `ID` (`V<nhóm>-01`…), DAY_TUONG khóa `Khóa DT`.
- Link: `["nhãn", "FID", trang]`, nhiều link = list; nhãn ngắn có số tờ + chi tiết (`"FUR-3315 p5 D13"`). Không có nguồn → chuỗi `"KHÔNG THẤY – đã tìm ở …"`.
- DS đúng 1 giá trị trong danh sách; số ghi dạng số; nhiều giá trị trong 1 ô cách `; `.

## Nhóm Trang / Tìm tờ — agent `doc-trang`
Giao bằng **subagent_type `doc-trang`** (`SKILL\_agents\doc-trang.md`, nạp qua junction `.claude\agents`). Agent chỉ có Read + Bash, không nạp CLAUDE.md (`omitClaudeMd`); **luật đầy đủ của nhóm trang nằm trong system prompt của agent** (tóm tắt dưới đây để nhóm Đối chiếu hiểu dữ liệu). Sửa luật nhóm trang thì sửa `doc-trang.md` trước, rồi đồng bộ phần tóm tắt này. Agent mới chỉ có ở phiên Claude Code mở sau khi tạo/sửa file.
- **Ghi:** mỗi trang 1 lần bằng `luu <GOC> <n> "<FID>|<trang>" < JSON` (`{"PAGES": {…}, "TRANG_QS": [...], "ISSUES": [...]}`), không Write/Edit file JSON. Tool kiểm cột như `check` (cột lạ, DS, thiếu Tờ · Chi tiết, phép cộng sai, KL không số) — lỗi thì không ghi gì; tự điền `Khóa trang` và PAGES `File ID`/`Trang`/`Số BV`/`Rev`/`Lớp text`; gửi lại = thay dòng cũ của trang, `--them` = nối; in `XEM` món trong danh sách chưa có dòng.
- Đo 16/11: nền khởi động mỗi agent 41,6k (general-purpose) → 10,8k token (doc-trang, gồm cả hướng dẫn).

Nhóm Tìm tờ: các cụm chưa có trang. Tìm tờ bằng `ungvien --khoa`, `tim` (bắt đầu từ `Bản vẽ hiệu lực` nếu có), rồi đọc từng trang tìm được như dưới; trang mới tự đặt `Khóa trang` `<FID>|<trang>`.

Với **mỗi trang**:
1. `xemtrang <GOC> <FID> <trang>` → **danh sách cần kiểm** của trang: từng món (`Cụm đọc`, loại món, tên đã bỏ kích thước/mã hoàn thiện của BOQ) và các trường `cần:` (bắt buộc, tờ loại này thường thể hiện) / `nếu có:`. Danh sách không có số BOQ: **không mở `xemcum` / mô tả BOQ đầy đủ ở nhóm trang** (nhìn hình trước, đối chiếu sau). Món ghi `(trùng tên, phân biệt: <Mã VL>)` = 2 cụm cùng mô tả khác vật liệu: phân theo nhãn vật liệu đọc trên hình, ghi `Cụm đọc` chắc chắn; mã này chỉ để phân biệt, `Mã hoàn thiện trên BV` vẫn ghi theo nhãn đọc trên hình.
2. Ảnh cả trang 1 lần (xác nhận tờ, vị trí từng món, loại tờ) → phóng từng chi tiết có món trong danh sách. Phóng 1 vùng thì đọc **mọi** món, mọi trường có trong vùng đó luôn, không quay lại vùng đã phóng.
   - **Không crop khung tên** khi `xemtrang` in dòng `Số BV / Rev (từ nhãn + sổ DWG …)`: chép thẳng vào PAGES. Chỉ khi `xemtrang` báo sổ thiếu: đọc khung tên trên ảnh cả trang, crop khi không đọc được.
   - Loại tờ `?` (máy chưa rõ): sau ảnh cả trang, tự xác định loại rồi chạy lại `xemtrang <GOC> <FID> <trang> --loai <MB|MBT|MĐ|MC|CT|BTK, nối bằng +>` để lấy đúng danh sách trường (vd tờ chi tiết tủ có mặt đứng + mặt cắt: `--loai MĐ+MC+CT`). Ghi loại đó vào PAGES `Loại tờ`. Không mở thêm ảnh để phân loại.
3. Ghi `TRANG_QS.docsau-<n>.json`: 1 dòng = 1 trường của 1 món. `Khóa trang`, `Chi tiết` (số chi tiết/hình), `Món` (nhãn/tag trên hình), `Cụm đọc` (cụm khớp **chắc chắn**; nhiều cụm cùng 1 đồ vật cách `; `), `Trường`, `Giá trị`, `Trạng thái`, `Tờ · Chi tiết`, `Nguồn` HÌNH/TEXT, `Crop` (tham số `anh`: `<trang> --clip … --dpi …`). Chưa chắc món thuộc cụm nào → **bỏ trống `Cụm đọc`, ghi `Cụm ứng viên`** (không đoán); nhóm Đối chiếu quyết. Một dòng BOQ gồm nhiều đồ vật → `Thành phần` 1, 2…
4. **Mỗi món trong danh sách của trang phải có ≥ 1 dòng.** Món không có trên trang: 1 dòng `Trường = Món trên tờ`, `KHÔNG THẤY`, lý do ở `Giá trị`. Trường `cần:` mà tờ không thể hiện thì không bắt buộc ghi ở trang này (hợp đồng kiểm tính ở cấp cụm, gộp mọi trang); nhưng tờ đúng loại mà không thấy → `KHÔNG THẤY` + lý do.
5. **Trường** (dùng đúng tên): `Lắp đặt` · `Kích thước BV` (R×S×C; chữ L thì từng nhánh) · `Có cánh` · `Ngăn kéo` · `Thanh treo` · `Kệ/đợt` · `Hốc/ô tivi` · `Mặt đá` · `Gương` · `Kính` · `Kim loại` · `Giấy/vải` · `Mã hoàn thiện trên BV` · `Ván nền ẩn` (khi có giấy/gương/kim loại phủ ngoài) · `Khung xương` · `Cong/bo` · `Phạm vi trên BV` (ốp tường/trần/sàn/sơn: vùng, tường nào, cao tới đâu) · `Cao độ trần` · `Phần ngoài phạm vi` · `Khu` · `Revision tờ` (chỉ ghi khi khung tên trên hình **khác** Rev sổ DWG mà `xemtrang` in, kèm issue). Trường không áp dụng cho loại đó (sơn không có `Có cánh`) thì bỏ, không ghi `KHÔNG CÓ` cho có.
6. **Kiểm chuỗi kích thước:** tổng các đoạn = kích thước tổng; ghi vào `Giá trị` (vd `"450+450=900 ✓"`). Không khớp → `MÂU THUẪN` + issue. `check` báo LỖI nếu phép cộng sai.
7. **Khối lượng** (`Trường = KL theo BV`): đếm trên bảng thống kê / mặt bằng, **mỗi khu 1 dòng**, `Khu` = mã phòng/khu như trên bản vẽ, `Giá trị` = số. Cùng món có ở cả bảng thống kê và mặt bằng → ghi cả hai, tool so (bằng nhau lấy 1 lần, khác nhau báo mâu thuẫn, **không cộng**). m2/md không có bảng thống kê: không đo, không ghi. Đếm dài trên mặt bằng → nhờ Gemini đếm cùng tờ (`--hoi "Count every <món> symbol/tag <mã> on this plan; list each with its room"`); khác nhau → đếm lại theo danh sách phòng.
8. **Cửa:** `Trường = Dày tường` tại **từng vị trí thật** (mặt bằng / ghi chú loại tường), `Khu` = vị trí, `Giá trị` = dày (mm) + loại tường. Không lấy 1 tờ điển hình cho mọi vị trí.
9. **Dẫn chiếu:** callout tới tờ khác (vd `18 / ID-PS-TY-712`) → `Trường = Dẫn chiếu`, `Giá trị` = số tờ. Tool đối chiếu với bộ tờ của cụm; tờ ngoài bộ hiện ở `Tờ chưa đọc` để giao đọc bổ sung. Callout chưa biết của món nào (mặt cắt tường…): vẫn ghi, `Cụm ứng viên` nếu đoán được — `ghep` đưa vào `_ghep.json` (`dan_chieu_khong_cum`) và báo tờ chưa ai đọc.
10. **Kiểm chéo Gemini** (mục dưới) cho các crop đã dùng làm bằng chứng của trang.
11. `PAGES.docsau-<n>.json` cho mọi trang đã mở (`Khóa trang`, `File ID`, `Trang`, `Loại tờ`, `Mã xuất hiện`, `Đã đọc bằng`, `Rev`, `Số BV`, `Tên BV`, `Lớp text`). `Số BV`/`Rev` chép từ dòng `Số BV / Rev` của `xemtrang`; `Tên BV` đọc được trên ảnh cả trang thì ghi, không crop riêng để lấy. Trang không có món nào: vẫn ghi PAGES, và 1 dòng TRANG_QS `Món = (không có)`, `KHÔNG THẤY`, lý do.
- Món thấy trên trang mà **không có trong danh sách** nhưng có vẻ là món BOQ khác: ghi với `Cụm ứng viên` nếu đoán được; không thì bỏ qua.
- **Không ghi** QUAN_SAT, SCOPE, BOQ_MAP, DAY_TUONG ở nhóm trang: `ghep` làm.

## Nhóm Đối chiếu — không mở ảnh trừ chỗ thiếu
Chạy sau khi mọi nhóm Trang/Tìm tờ xong và phiên điều phối đã chạy `ghep` (`check` báo LỖI nếu chưa ghép hoặc TRANG_QS mới hơn lần ghép). Xem cụm theo lô: `xemcum <GOC> <cụm 1> <cụm 2> … --gon` (nhiều cụm 1 lần; tool tự dừng ở ~24.000 ký tự và in lệnh gọi tiếp cho cụm còn lại; `--gon` gom TRANG_QS theo trang/món và ẩn dòng QUAN_SAT do `ghep` chép lại, −36–43% chữ) (mô tả BOQ, MAT, tờ đã đọc chưa, TRANG_QS của cụm, QUAN_SAT đã ghép, **CÒN THIẾU** trường bắt buộc, BOQ_MAP, phần `ghep báo`, mục `TRÙNG SỐ BOQ CỤM KHÁC`).
1. **Không chép lại** quan sát đã ghép. Chỉ ghi phần thêm/sửa.
2. **So với BOQ:** `Kích thước BoQ` (SCOPE, nếu mô tả có ghi). BV ↔ BOQ lệch (kích thước, số lượng, cấu tạo) → thêm dòng QUAN_SAT `Nguồn = BOQ` cạnh dòng HÌNH đã ghép, cả 2 `MÂU THUẪN` (ghi dòng BOQ, dòng HÌNH giữ nguyên) + 1 issue. Không tự chọn bên thắng. `Mã hoàn thiện trên BV` ≠ MAT → issue.
3. **Quan sát chỉ có Cụm ứng viên** (dòng `?` trong `xemcum`): quyết thuộc cụm nào → ghi 1 dòng QUAN_SAT cho đúng cụm (chép `Trường`, `Giá trị`, `Trạng thái`, `Tờ · Chi tiết`, `Nguồn`; `Model` ghi `gán từ TRANG_QS <Khóa trang>`). Không quyết được → issue / ASK_USER.
4. **CÒN THIẾU trường bắt buộc:** trang của cụm đã đọc mà thiếu → mở đúng trang, đúng vùng (cắt lại đúng `Crop` nếu có), đọc bổ sung, ghi QUAN_SAT (nguồn HÌNH, có tờ + chi tiết) hoặc `KHÔNG THẤY` + lý do. Còn tờ `[CHƯA]` / dẫn chiếu ngoài bộ → ghi vào kết quả trả về để phiên điều phối giao tìm tờ; không tự mở cả bộ.
5. **KL:** `ghep báo` KL mâu thuẫn giữa trang / không gán được khu → quyết theo bảng thống kê hoặc danh sách phòng, ghi `BOQ_MAP` (`KL theo BV`, `Tờ đếm KL`); không quyết được → issue.
6. **SCOPE (dòng đại diện):** `Bản vẽ đối chứng`, `Kích thước BV` tool đã điền từ trang đọc, sửa nếu sai thứ tự (tờ thể hiện rõ nhất đứng đầu). `Phạm vi nhà máy` khi tool chưa gắn: `TRONG SCOPE – <mộc|sơn đồ gỗ|cửa|cabinet|loose>` / `KHÔNG THUỘC SCOPE NHÀ MÁY – <nhóm>` / `CHƯA RÕ – <lý do>` (chỉ là nhãn, không bỏ dòng, không đổi `Trạng thái`).
7. **ISSUES (V…):** thiếu tờ, BV ↔ BOQ lệch, mã hoàn thiện BV ≠ MAT, tờ khác revision.
8. **`TRÙNG SỐ BOQ CỤM KHÁC`** (máy dò: số trong TRANG_QS đã gán cụm này = kích thước BOQ của cụm khác cùng sheet BOQ / cùng trang, có ít nhất 1 số lẻ không chia hết 50 và không thuộc kích thước BOQ của chính cụm): xét từng cụm nêu. Cùng 1 món (vd dải "len 150 dưới WP-04 3727" ở E3 ↔ bục BF-02A L3727×H150) → issue tính trùng KL / gán nhầm cụm; chỉ là số tường/khoảng cạnh nhau → bỏ qua, không ghi.

## Nguyên tắc
1. **Nhìn hình trước, đọc mô tả BOQ sau.** Nhóm trang chỉ thấy tên món + tên trường, không thấy số BOQ; nhóm đối chiếu mới so.
2. **Mỗi giá trị có tờ + số chi tiết/hình.** Không có số chi tiết → chưa đủ bằng chứng cho trường cấu tạo.
3. **Không có ≠ không thấy.** `KHÔNG CÓ` = nhìn rõ là không có. `KHÔNG THẤY` = chưa có tờ / chưa đọc được. Không bỏ trống trường quyết định giá, không đoán.
4. **Lệch thì giữ cả hai** (BOQ ↔ BV, Claude ↔ Gemini, trang ↔ trang): 2 dòng `MÂU THUẪN` + 1 issue.
5. **Khối lượng đếm trên bản vẽ**, theo khu. Không lấy KL BOQ chia, không lấy diện tích chia kích thước, không tự đo m2/md.
6. Đọc kỹ 20 kịch bản sai và cách chặn: `D:\COSTING EXCEL\_MASTER\modules\SCAN-NHAN-LOAI.md` mục 4 (chỉ đọc mục 4).

## Công cụ — chỉ dùng các lệnh này
Chạy lệnh `td.py` qua **Bash** (PowerShell 5.1 làm mất tham số đường dẫn PDF có 2 dấu cách liền / dấu tiếng Việt — đo 16/11). Không tự viết script (Python, sed, grep, find, cat, Select-String) để đọc `_data`, PDF, log Gemini hay mã tool: mỗi lượt tự viết đọc lại ~110k token ngữ cảnh. Thiếu lệnh thì ghi vào kết quả trả về. Gemini chạy thường (đợi tới ~5 phút), không chạy nền rồi `sleep`.
| Việc | Lệnh `py -I "<TOOL>" …` |
|---|---|
| Danh sách cần kiểm + mọi thứ của 1 trang (nhóm trang) | `xemtrang <GOC> <FID> <trang> [--chu] [--loai MĐ+MC]` (`--chu` = cả lớp chữ; `--loai` = loại tờ tự xác định khi máy in `?`) |
| Mọi thứ của 1 cụm (nhóm đối chiếu) | `xemcum <GOC> <cụm hoặc khóa dòng> […] --gon` (tự dừng khi output quá dài, in lệnh gọi tiếp) |
| Lọc dữ liệu đã ghép | `loc <GOC> <SHEET> "cột=gt" "cột~regex" --cot a,b` |
| Tìm trang có mã / từ khóa | `tim <GOC> --ma WD-01 [--file s]`, `tim <GOC> "<regex>"` |
| Trang ứng viên của dòng | `ungvien <GOC> --khoa "<Khóa>"` |
| Giải nén file trong zip | `zip <GOC> <FID>` |
| Xem trang (ảnh) | `anh "<pdf>" <trang> "<thư mục ảnh trong prompt>\p13.png" --dpi 100` rồi mở PNG bằng Read (đường dẫn PDF lấy từ `xemtrang`/`xemcum`) |
| Phóng một vùng | `anh … --dpi 220 --clip 0.2,0.1,0.6,0.5` (tỉ lệ 0–1 theo chiều nhìn). Tool tự hạ dpi: cả trang ≤ 1000 px cạnh dài, vùng cắt ≤ 1,15 MP. In `(giới hạn ảnh …)` = vùng quá to so với dpi xin: cắt vùng nhỏ hơn, không xin dpi cao hơn |
| Chữ trong vùng | `chu "<pdf>" <trang> --clip …` |
| Số kích thước + vị trí | `so "<pdf>" <trang> --clip …` (chỉ khi trang có lớp chữ) |
| Gemini đọc dãy số của 1 trang | `gemini "<crop.png>" … --trang "<FID>|<trang>" --goc <GOC>` (lưu `_gemini_so\trang\`, số Claude lấy từ TRANG_QS của trang; `ghep` chia về cụm) |
| Gemini quan sát đủ bộ trường | `gemini "<crop.png>" --quansat "Tủ lavabo dưới chậu LA1" --goc <GOC>` |
| Gemini hỏi riêng | `gemini "<ảnh.png>" --hoi "How many door leaves does the cabinet have?" --goc <GOC>` |

Trang không có lớp chữ hoặc số vẽ bằng nét → đọc bằng hình, phóng to từng vùng, rồi kiểm chuỗi kích thước.

## Gemini = người đọc thứ hai (bắt buộc với mọi crop dùng làm bằng chứng)
Gemini nhìn bản vẽ tốt: nhận đúng cấu kiện, nhãn, đợt/kệ, phần ngoài phạm vi, dãy số in trên hình. Nhưng đã thấy nó **gán nhầm số vào chiều** và **tự trừ ra số không in**; đếm số lượng dài cũng hay sót. Vì vậy:
1. **Độc lập:** Claude ghi giá trị của mình vào `TRANG_QS` **trước** (`--trang` dừng nếu trang chưa có `Kích thước BV` dạng R×S×C), rồi mới chạy Gemini trên **đúng các crop đó**. Không đưa số hay nhận xét của Claude vào câu hỏi cho Gemini.
2. **Mỗi trang:** chế độ số `--trang` 1 lần cho mọi crop có dãy kích thước (truyền nhiều crop một lần; chạy lại khi thêm crop thì cộng dồn), `--quansat "<Món>"` cho crop mặt đứng/chi tiết chính của từng món. Luôn kèm `--goc <GOC>`: kết quả lưu ở `_data\_gemini_cache` theo **nội dung ảnh + câu hỏi**, dùng chung mọi nhóm (chỉ trúng khi ảnh giống hệt byte: cùng trang, cùng `--clip`, cùng dpi).
3. **So từng trường:** khớp → giữ dòng của Claude. Lệch → phóng to đúng chỗ đó đọc lại; vẫn lệch → giữ 2 dòng (Nguồn `HÌNH` và `GEMINI`), `MÂU THUẪN`, cộng 1 issue. Gemini thấy món mà Claude ghi `KHÔNG THẤY` → đọc lại theo `where` Gemini chỉ.
4. **Không nhận** số Gemini tự tính ra (không in trên hình), và không dùng Gemini làm nguồn duy nhất cho kích thước hay số lượng.
5. Mỗi **món trên trang** 1 dòng TRANG_QS: `Trường = Kiểm chéo Gemini`, `Nguồn = GEMINI`, cùng `Cụm đọc`, `Giá trị` = `"khớp: <các trường>; lệch: <các trường hoặc không>; crop: <tờ/chi tiết>"`, cho **các trường không phải số**. `ghep` gộp các trang thành 1 dòng mỗi cụm. Phần số do tool so (bước 7e tự sinh `Kiểm chéo Gemini số`); tool chỉ so theo tập số, nên việc gán đúng chiều vẫn phải tự kiểm ở mục 3.
6. agy lỗi / không có → `Kiểm chéo Gemini` trạng thái `KHÔNG THẤY`, giá trị `KHÔNG THẤY – <lỗi>`, làm tiếp; cụm đó không được CAO.

## Tiết kiệm token — lọc trước khi mở ảnh
1. **PDF trước, CAD sau cùng.** Chỉ đụng `.dwg` khi cụm không có tờ nào trong PDF; ghi `Dẫn chiếu`/issue nêu rõ "chỉ có CAD".
2. **Ảnh từ rẻ tới đắt:** cả trang (tool giữ ≤ 1000 px) để xác nhận tờ, vị trí → phóng vùng chi tiết (150–220 dpi, `--clip` đúng khung chi tiết) chỉ khi cần số. Mỗi ảnh nằm lại trong ngữ cảnh tới hết agent: ít ảnh, vùng vừa đủ, đọc hết mọi món trong vùng trong 1 lần. Trang có lớp chữ thì `chu`/`so` trước, ảnh sau.
3. Trang đã đọc ghi vào `PAGES` — lần chạy lại dùng danh sách này, không đọc lại.

## Trước khi trả kết quả
- **Nhóm Trang / Tìm tờ:** mọi trang được giao có PAGES; mọi món trong danh sách của từng trang có ≥ 1 dòng TRANG_QS; `check --den docsau --nhom <n>` không còn LỖI `TRANG_QS` ở file của mình (thiếu tờ/chi tiết, chuỗi cộng sai, KL không phải số). Trả về: số trang đã đọc, món còn `KHÔNG THẤY` ở trường quyết định giá, dẫn chiếu tới tờ ngoài danh sách.
- **Nhóm Đối chiếu:** `check --den docsau --nhom <n>` không còn LỖI "trường bắt buộc chưa có kết luận" ở cụm của mình (hoặc đã ghi `KHÔNG THẤY` + lý do). Chạy `py -I "<TOOL>" neo <GOC>`: máy kiểm số tờ trong nhãn link, mã/model ở link spec, `Mã hoàn thiện trên BV` / `Kích thước BV` có trong lớp chữ các trang được link. Bản vẽ vẽ nét ra KHÔNG KIỂM ĐƯỢC; kích thước chỉ lấy chuỗi R×S×C **đầu tiên**, nên ghi chuỗi đó trước, diễn giải sau. Dòng `NEO LỆCH` của cụm mình: mở lại đúng trang, sửa hoặc ghi lý do. Trả về: số cụm đã làm, cụm còn KHÔNG THẤY ở trường quyết định giá, tờ cần giao tìm/đọc bổ sung.
