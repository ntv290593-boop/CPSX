---
name: 10-tender
description: Quét hồ sơ mời thầu cho PHÒNG TENDER (toàn gói, không lọc phạm vi nhà máy) của dự án ghi ở ACTIVE.md — theo dõi thay đổi cuối cùng qua các đợt TCQ/Addendum (vật liệu, bản vẽ, phạm vi), ra <GOC>\10_TENDER\03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx (BOQ CĐT + cột nhà thầu + cột nội bộ + sổ TCQ_LOG/MAT/DWG/KHU/SCOPE), <DA>_DMVT_<yymmdd>.xlsx (danh mục vật tư có ảnh) và bản gửi CĐT. Dùng khi user nói "scan tender", "quét HSMT", "tender", "cập nhật TCQ", "có TCQ mới", "làm BOQ tender", "build lại tender", "làm DMVT", "danh mục vật tư".
---

# 10-tender — ĐIỀU PHỐI

Phiên này chỉ **điều phối**: chạy tool, giao việc, đọc kết quả kiểm. **Không tự đọc BOQ / TCQ / bản vẽ / spec ở đây** — mỗi bước giao 1 subagent mới (ngữ cảnh sạch, chỉ đọc file hướng dẫn của bước đó).
Khác `01-scan`: quét **toàn gói**, gắn nhãn chứ không bỏ dòng; `Phạm vi nhà máy` chỉ là nhãn, không dùng để lọc; các đợt nối tiếp, đợt sau đè đợt trước. Có đủ phần đọc sâu của 01-scan (mục 1e), áp cho toàn gói.

## 0. Chuẩn bị (không hỏi user nếu đã rõ)
- `GOC` = thư mục dự án trong `D:\COSTING EXCEL\ACTIVE.md`. `TOOL` = `D:\COSTING EXCEL\SKILL\10-tender\tool\td.py`, chạy `py -I "<TOOL>" <lệnh> <GOC> …` (không tham số = xem lệnh). Riêng `chu`, `anh`, `mail`, `xem` nhận thẳng đường dẫn file, không cần `<GOC>`.
- Chưa có `GOC\10_TENDER\_data\_cauhinh.json` → `init` với thư mục hồ sơ, BOQ CĐT, file TCQ **gộp mới nhất**, bản dịch nội bộ (lấy từ ACTIVE.md; thiếu thì hỏi user 1 câu). Có TCQ gộp mới hơn → chạy lại `init --tcq <file mới>`.
- Thư mục trong hồ sơ không phải hồ sơ thầu (file học, form, dữ liệu dự án khác) → `init --bo-thu-muc "<thư mục>;<thư mục 2>"` (đường dẫn tính từ thư mục hồ sơ, ghi vào khóa `bo_thu_muc` của `_cauhinh.json`, thay cả danh sách cũ). `kiemke` không đi vào các thư mục này và báo XEM nếu tên không khớp thư mục nào. Thư mục đã có trong FILE_INDEX thì lần `kiemke` sau báo file mất.
- **Mẫu dự án:** `form\MAU_DU_AN.json` là thư viện các kiểu dự án. Dự án mới chọn mẫu gần nhất bằng `init --mau <tên>`, không viết lại Python. `cdt_item` (mặc định): có BOQ CĐT, TCQ. `tu_lap_vn`: chưa có BOQ CĐT, tự lập BOQ theo form tiếng Việt, TCQ không bắt buộc (mục 1d). Không khớp mẫu nào thì thêm 1 khóa mới trong file này.
- Dừng và báo user nếu `~$<DA>_BOQ_TRACK_*.xlsx` hoặc `~$FILE_INDEX.xlsx` tồn tại (file đang mở).
- `README.md` của `10_TENDER` là trí nhớ của dự án (đợt hiệu lực, thứ tự ưu tiên, bẫy, việc chưa quyết). Mọi subagent đọc nó trước.

## 1. Các bước
| # | Bước | Tool chạy trước | Hướng dẫn | Ai làm | Sau bước |
|---|---|---|---|---|---|
| 1 | kiemke | `kiemke` | `buoc/1-kiemke.md` | 1 subagent | — (xong bước → chạy `trang`: đọc chữ mọi trang PDF có vai trò dùng được, chỉ file mới/đổi) |
| 2 | boq | `boq` | `buoc/2-boq.md` | 1 subagent | 1 |
| 2b | phanloai | `phanloai` | `buoc/2b-phanloai.md`, chỉ khi còn dòng `Chờ LLM` | 1 subagent; xong chạy lại `phanloai` | 2 |
| 3 | tcq | `tcq` | `buoc/3-tcq.md` | ≤ 3 subagent song song, chia theo đợt TCQ (vd 01–03 / 04–05 / 06–07); file `TCQ_LOG.tcq-<n>.json` | 1, 2 |
| 4 | dwg | `dwg` | `buoc/4-dwg.md` | 1 subagent | 1 (song song với 2, 3) |
| 5 | mat | `mat` | `buoc/5-mat.md` | 1 subagent | 3 |
| 6 | banve | `trang` (nếu chưa) | `buoc/6-banve.md` | mỗi nhóm mã phòng/khu 1 subagent, song song; file `SCOPE.banve-<n>.json` | 4, 5 |
| 7 | doichieu | `doichieu` | `buoc/7-doichieu.md` | mỗi 1–2 sheet BOQ 1 subagent; file `SCOPE.doichieu-<n>.json` | 3, 5, 6 |
| 7b | ben | `ben` (dự án mới, sau khi có `_data\_ben.json`) | (mục 1a) | bạn; tender duyệt | 5 |
| 7c | spec | `docsau` | `buoc/7c-spec.md` | 1 subagent | 5, 7 |
| 7d | docsau | (đã chạy ở 7c) | `buoc/7d-docsau.md` | mỗi nhóm trong `_docsau.txt` (≤ 8 cụm) 1 subagent, song song; file `*.docsau-<n>.json` | 6, 7 |
| 7e | tincay | `tincay` | `buoc/7e-tincay.md` | 1 subagent | 7c, 7d |
| 8 | dmvt | `dmvt <GOC> --chi-boc` | `buoc/8-dmvt.md` | bạn hoặc 1 subagent; ghi `_dmvt_llm.json` | 5 |
| 9 | build | `build` (tự dựng DMVT trước) | (mục 3) | bạn | 7, 8 |
| 10 | goi | `goi <GOC> "<bên>"` | (mục 1a) | bạn, khi tender cần | 8 |

Prompt giao việc (thay `<…>`; không thêm số liệu hay suy đoán của bạn vào prompt):
```
Làm bước <buoc> của skill 10-tender.
Đọc rồi làm đúng theo: D:\COSTING EXCEL\SKILL\10-tender\buoc\<file>.md và D:\COSTING EXCEL\SKILL\10-tender\JSON.md
GOC = <…>   TOOL = D:\COSTING EXCEL\SKILL\10-tender\tool\td.py
[chia nhóm: Nhóm <n>, chỉ làm: <đợt TCQ / mã phòng / sheet>. Ghi file *.<buoc>-<n>.json]
Trả về ≤ 15 dòng: file đã ghi, số dòng, việc chưa làm được và lý do, câu cần hỏi user.
```

**Lõi tìm trang dùng chung với 01-scan** (`SKILL\_core\scan_core.py`: đọc chữ + khung tên, trang trùng, BM25, mã chuẩn hóa). Lệnh: `trang` (cache chữ), `tim --ma <mã>` (trang có mã, mọi cách viết), `tim <regex>` (trang có chữ), `ungvien --phong <mã>` (trang ứng viên mỗi dòng BOQ, chia nhóm loại tờ). Sửa lõi hoặc cách chấm điểm → đo trước/sau: `DOCS\scan-toi-uu\bench.py` (01-scan) và `ungvien <GOC> --do` (độ trùng với tờ LLM đã gán; SOFITEL 29/10: dòng có ≥1 trang trùng @12 61%, @40 80%).

**Đọc 1 file** (`SKILL\_core\doc_file.py`): `chu <pdf> <trang> [--clip]`, `anh <pdf> <trang> <png> [--dpi] [--clip]`, `mail <file.msg>`, `xem <xlsx> [sheet] [A1:Z60]`. 10-tender không gọi tool của 01-scan.

**Code ID** (`SKILL\_core\code_id.py`, dữ liệu `_MASTER\code_id`): `phanloai` gán `Mã cấp 3` theo thứ tự User > Luật > Nhớ mô tả > Nhớ đầu mục > LLM, cho mọi dòng có đơn vị trừ sheet khớp `codeid_bo_sheet` (mặc định GR&Prelims, Provisional sum, Summary). `Cờ phân loại` = chỗ máy không chắc. Dự án mới: dòng theo luật có cờ `Luật chưa duyệt` cho tới khi user xem xong và chạy `phanloai <GOC> --duyet-luat`. Ô user sửa vào `nho_local.json` của máy; **`nho.json` chỉ người quản lý sửa** (`DOCS\scan-toi-uu\tender-test\nap_codeid.py`). Sheet BOQ chi tiết có `Mã cấp 3` (droplist, ô tô màu = có cờ), `Code ID` (công thức nhảy theo droplist, tra sheet `_CODE`), `Bên thực hiện` (droplist), `Mã cấu tạo` (droplist); 4 cột này không ẩn. `build` tự chạy lại `phanloai` sau khi kéo sửa tay. Đo lại: `DOCS\scan-toi-uu\tender-test\do_codeid.py`.

## 1a. Bên thực hiện và gửi nhà thầu phụ
- `_data\_ben.json` riêng từng dự án: `ds` (danh sách bên, ra droplist), `gop` (chữ cũ → tên trong `ds`), `loai` (`Loại` của MAT → bên, dùng để nháp và gắn cờ), `tam` (danh sách chưa chốt, DASHBOARD ghi TẠM). Đổi danh sách: sửa file này rồi `build`. Không sửa `SCOPE.ben.json`, đó là chữ gốc.
- Dự án mới chưa có bên: `ben` nháp từ `Mã VL` → `Loại` → bên, cờ `Nháp`. Không có `Mã VL` thì để trống. **Không suy bên từ Code ID.**
- Dòng nhiều bên: tender Unhide rồi khai sheet `CAU_TAO` (khóa `CT-01#1`, thành phần, mã VL, bên, nguồn tờ) rồi chọn `Mã cấu tạo` ở dòng BOQ. `Bên thực hiện` của dòng đó do máy gộp từ mẫu; sửa ở dòng không có tác dụng và bị báo XEM. Máy chỉ gắn cờ `Nhiều bên?`, không tự tạo mẫu. Khối lượng, giá chỉ nằm ở dòng BOQ.
- `goi` lấy dòng Base, VE, Chờ; workbook mới, không chép từ BOQ_TRACK; bỏ mã TCQ khỏi mô tả; máy soát lại file, thấy Code ID, mã TCQ, link, đường dẫn nội bộ, ô giá đã điền thì xóa file và báo LỖI. Chạy sau `build`, vì đọc từ `_data`. Mô tả nhắc tên bên khác (chữ của BOQ, vd `by MEP`) chỉ báo XEM. **Không tự gửi**; nhắc tender xem DASHBOARD `Dòng Base chưa có bên` trước khi gửi.

## 1b. Sổ xuất hiện vật liệu và Area
- `build` tự chạy `area`, `xuathien`. Chạy riêng khi cần xem số.
- Sheet `AREA`: tiền tố mã → tập phòng, riêng từng dự án, dựng từ `KHU` chỉ với Bộ hồ sơ FF&E, Thảm. Tender sửa cột `Mã phòng` ngay ở sheet, build lại vẫn giữ. Mọi ô Area của DMVT lấy theo bảng này đều gắn cờ (làm ở GĐ5).
- Sheet `VL_XUATHIEN` (ẩn): 1 dòng = mã + nơi xuất hiện + file + trang hoặc ô. Chỉ để xem, máy dựng lại mỗi lần build. Dòng `Bản vẽ` đi qua `Bản vẽ hiệu lực` của dòng BOQ và luôn có cờ, vì tag trên bản vẽ là nét. Cờ `Lệch Area` = phòng ở dòng BOQ không thuộc tiền tố mã.
- Cột `Link vật liệu` (ở sheet BOQ và BOQ_TONGHOP): trỏ thẳng trang gốc, không qua sheet MAT. `MAT.Nguồn spec` nếu có, không thì trang Định nghĩa mới nhất của mã. Mở bằng add-in SRC_LINKS. Dòng FF&E: trang spec ở cột này, trang ảnh sản phẩm ở cột `Bản vẽ hiệu lực` (lấy từ cùng tập với trang Định nghĩa mới nhất).
- DASHBOARD đếm mã chưa có trang spec, dòng có mã chưa có link, cặp lệch Area. `MAT.Nguồn spec` trỏ trang không có chữ mã thì ra dòng `XEM spec`. Đo cột Area: `DOCS\scan-toi-uu\tender-test\do_area.py`.

## 1c. Danh mục vật tư (DMVT)
- **Mỗi lần scan phải ra đủ 2 file: BOQ_TRACK và DMVT.** `build` luôn dựng DMVT trước để `Link vật tư` có số hàng. Chạy `dmvt --chi-boc`, làm bước 8, rồi `build`.
- `_data\_dmvt.json` (tiêu đề, tiền tố FF&E, mẫu mã lấy từ trang SPEC, mục I, II…) không bắt buộc: chưa có thì dùng mặc định và báo XEM; dự án mới nên tạo để có tiêu đề và mục đúng.
- Cần Excel. DMVT không dựng được (không có Excel, file đang mở) thì BOQ_TRACK vẫn ra, nhưng `build` báo `LỖI: chưa đủ 2 file` và thoát lỗi: sửa rồi build lại, không báo xong. Chỉ bỏ DMVT khi user yêu cầu: `build --khong-dmvt`.
- Máy bóc A:I theo loại trang: `MATERIAL REFERENCE` theo từng khối `ITEM`, `DRAWING REFERENCE` theo đường kẻ cột, `SPECIFICATION` cho mã khớp `item_spec` (SOFITEL: rug, phụ kiện rèm). Bản mới nhất **theo từng mã**, không theo cả file. FF&E gộp theo Design + Colour; phụ kiện gộp theo nhà cung cấp; mã khác 1 dòng/mã. Mã có ở MAT mà không có trang spec vẫn có dòng, có cờ.
- **Chỉ áp kết luận từ MAT:** Omit giữ dòng, tô xám; Đổi theo TCQ/VE giữ nội dung spec, ghi kết luận vào cột I, có cờ. Mã FF&E chưa có ở MAT có cờ `Chưa qua MAT`, không suy; có câu TCQ còn hiệu lực nhắc tới thì thêm `TCQ nhắc mã`. Spec ghi thẳng OMITTED cũng tô xám.
- **Ảnh:** ảnh gốc theo xref, thuộc khối khi tâm nằm giữa 2 tiêu đề `ITEM` hoặc trong vùng `Sample Swatch`. Nhiều ảnh: ảnh chụp trước, rồi lớn nhất (`"chon_anh": "lon_nhat"` trong `_dmvt.json` để lấy lớn nhất). Mọi dòng có cờ ảnh vào sheet `_SOI_ANH`.
- **File:** `03_boq\<DA>_DMVT_<yymmdd>.xlsx`, dựng từ `form\DMVT_form.xlsx` bằng Excel COM (không lưu bằng openpyxl). Máy ghi A:I, AG `Spec mới khác`, AH `Cờ`, AI `Link spec`, AJ `Link bản vẽ`. J:AF và từ AK trở đi là của tender. Mỗi ô link tối đa 30 đích.
- **Build lại giữ sửa tay** (so `_dmvt_snapshot.json`): ô A:I, ảnh khác snapshot được giữ; spec mới khác thì ghi `Spec mới khác`, không đè. Dòng không mã đi theo dòng có mã phía trên. Tập mã đổi: có cờ, không tự chuyển sửa tay; dòng cũ có công người nằm ở mục `DÒNG CŨ GIỮ LẠI, CẦN XEM`. Không xóa `_dmvt_snapshot.json`.
- **Trước khi gửi CĐT tender tự xóa:** các cột từ AG, sheet `_LINKS`, `_SOI_ANH`; xem thêm L và AB:AC (giá). Skill nhắc đủ các mục này khi báo kết quả, không gửi thay.
- Đo lại: `DOCS\scan-toi-uu\tender-test\do_dmvt.py`. Đổi form: `lam_form_dmvt.py`.

## 1d. Dự án chưa có BOQ CĐT (mẫu `tu_lap_vn`)
Chỉ làm khi user yêu cầu lập BOQ. Hỏi rõ form, phạm vi, mức bóc khối lượng trước khi chạy.
1. `init --mau tu_lap_vn --input <rel> --du-an <DA> [--dwg-regex …] [--bo-thu-muc …]` → gán vai trò file → `kiemke` → `trang`.
2. Có DWG: `_MASTER\tools\.venv\Scripts\python.exe -I SKILL\_core\dwg_tag.py <thư mục DWG> <GOC>\10_TENDER\_data\_dwg_tag.json` (tag PDF là nét vẽ thì đây là nguồn đếm; so vài tờ với PDF trước khi tin). Tag gán về tờ theo viewport, tag Model nằm ngoài mọi viewport không được đếm; `goc` là gốc khung tên để đổi tọa độ sang ảnh. Spec kiểu `Item Desc / Area`: `tool\spec_boc.py`.
3. **Soát bề mặt từng tag** (`tool\be_mat.py`): máy xếp theo nhãn tag rồi tên khung hình; tag mặt đứng còn lại là `can_soat`. Giá trị: `tuong`, `tran`, `san`, `cua`, `portal` (lỗ mở không cánh, ra dòng `Ốp portal lỗ mở` riêng vì không có dòng cửa nhận), `joinery`, `khac`, `khong_ro`. Tag `san`, `khac` mặc định chỉ ghi ở sheet `Ghi chú & ngoài phạm vi`; muốn có dòng thì khai trong `_lap_boq.json`: `san_tinh` (mã → công tác), `chi_tiet` (mã → tên, đơn vị), `ngoai_tick` (công tác không được tick nhưng user chốt phải tính). `tool\chuan_bi_soat.py <GOC>` vẽ ảnh khoanh đỏ vào `_data\_soat\`, người hoặc LLM ghi `_data\_soat_tag.<n>.json`, kết quả này thắng máy. Không suy bề mặt từ loại tờ. Đo độ tin cậy: `tool\mau_soat.py <GOC> tao` rồi `so`, người soát lại không được đọc `_soat_tag.*`. Xem 1 tờ: `tool\soi_tag.py`.
4. Viết `_data\_lap_boq.json` cho dự án (tên tháp, khu, luật công tác, xem CTX làm mẫu) → `tool\lap_boq.py <GOC>` → `0.BOQ_TU_LAP\<DA>_BOQ_tu_lap_<yymmdd>.xlsx`. Mỗi dòng có nguồn và cờ `[XEM]` ở cột Ghi chú.
5. `init --boq <file>` → `kiemke` → `boq`, `tcq`, `dwg`, `mat` → `tool\lap_boq_buoc.py <GOC>` → `phanloai` → `check` → `tool\kiem_boq.py <GOC>` → `build`.
6. DMVT: `_data\_dmvt.json` ghi `"nguon_spec": "spec_boc"` để `dmvt` lấy nội dung và ảnh từ `_spec_boc.json` thay vì mẫu trang SOFITEL. `spec_boc.py` gộp chữ về 1 dòng; `dmvt.py` tách lại theo nhãn, mục cửa `(D..)`, mã hoàn thiện, kích thước và chỉ bỏ chữ rác từ layout (`bo_ma_rac`, `chu_ncc`). Spec kiểu khác có nhãn mới thì thêm vào `_NHAN_HD`.
`lap_boq.py` tự trỏ `boq` trong `_cauhinh.json` sang file vừa ghi (tên có ngày, sang ngày mới là file mới). Kích thước ở cột E, F lên `BOQ_TONGHOP` qua `Kích thước BoQ`.
Dòng cửa ghi `Vật liệu: <mã>. Phụ kiện: <mã>` nên `Mã VL`, `Link vật liệu`, `Link vật tư` có đủ gỗ, kim loại, kính, phụ kiện (`vl_cua_tu_tag`): tag trên ô cửa của tờ lịch + phần mặt cắt tờ chi tiết, không lấy tag mặt bằng chi tiết vì còn chỉ cả tường hai bên. Bộ HD lấy theo ghi chú lịch, lịch không ghi thì lấy tag mặt đứng. Lệch giữa 2 tờ, mã chưa chốt, xếp cửa gỗ mà tag không có mã gỗ: có cờ `[XEM]`. Máy không phân biệt được tag chỉ vào cánh hay khung; xem hình thấy mã chỉ thuộc khung thì khai `cua_ma_khung` (mã cửa → mã → ghi chú) trong `_lap_boq.json`: mã rời khỏi dòng, ghi chú vào cờ.

`check` chỉ kiểm cấu trúc. Kiểm đúng sai là `kiem_boq.py` (tờ link có chứa mã không, dòng không nguồn, PDF khớp số tờ, sót mã, đếm lại, kích thước, mã trùng); khi báo kết quả phải nói rõ đã chạy cái nào. Cả hai không đọc file xlsx đã build: trước khi giao phải đọc lại `BOQ_TONGHOP` và xem ảnh các dòng có cờ của DMVT.
DMVT: quy tắc "ảnh lớn nhất" hay chọn nhầm phối cảnh cả phòng thay vì ảnh sản phẩm; dòng `Nhiều ảnh, máy chọn` phải xem hình, sai thì ghi `Ảnh chọn` trong `_dmvt_llm.json`.
Giới hạn: `lap_boq.py` không đo diện tích; mã FF&E trùng giữa 2 file LF: BOQ tra đúng theo khu, MAT và DMVT chỉ lấy 1 bản.

## 1e. Đọc sâu (chuyển từ 01-scan, `tool\docsau.py`)
- **Phạm vi:** toàn gói. Mọi dòng có đơn vị, trừ sheet khớp `codeid_bo_sheet` (Prelims, Provisional sum, Summary) và dòng `Trạng thái` = Omit, Loại trừ, Không áp dụng. Vì vậy chạy sau bước 7.
- **Cụm đọc:** `docsau` gom các dòng cùng mô tả, `Mã VL`, `Mã spec/FF&E` và bộ tờ `Bản vẽ hiệu lực` thành 1 cụm, đọc 1 lần; khóa cụm là `Khóa` của dòng đầu. Ứng với "mã" của 01-scan. Chia nhóm ≤ 8 cụm (`--nhom n`) → `_data\_docsau.txt`. **Dán khối `Nhóm <n>` vào prompt bước 7d.** CTX đo 01/11: 378 dòng → 339 cụm → 43 nhóm; SOFITEL 730 dòng nên tốn hơn. Báo số nhóm cho user trước khi giao lần đầu (không cần chờ duyệt).
- **Kết quả:** `Kích thước BV`, `KL theo BV`, `Lệch KL`, `Bản vẽ đối chứng`, `Spec · ML`, `Mã hoàn thiện`, `Độ tin cậy`, `Issue`, `Giả định tính giá`, `Phạm vi nhà máy` ở BOQ_TONGHOP và sheet BOQ CĐT. Sổ ẩn: `QUAN_SAT`, `BOQ_MAP`, `DAY_TUONG`, `SPEC`, `PAGES`. Sheet hiện: `ISSUES`, `RFI_LOG`, `ASK_USER`. Sổ nào không có dữ liệu thì không có sheet. Bản `_CDT` và `goi` không có các cột này.
- **Gemini bắt buộc** làm người đọc thứ 2 (`gemini --cum`), Claude ghi số trước. Máy không có agy: bước 7d ghi `KHÔNG THẤY – <lỗi>`, làm tiếp, cụm đó tối đa TB. Máy người dùng cần cài và đăng nhập `agy` mới có kiểm chéo.
- **Kiểm:** `check --den docsau` đòi mọi cụm có QUAN_SAT và `Kiểm chéo Gemini`; phép cộng `a+b=c` sai là LỖI. `check --den tincay` đòi `Độ tin cậy`. `neo` kiểm cả `Bản vẽ hiệu lực` bước 6 và `Nguồn spec` của MAT (chỉ XEM). Dự án chưa chạy `docsau` không bị báo LỖI, chỉ có dòng nhắc, DASHBOARD ghi `Chưa đọc sâu`.
- **Đừng làm lại ở 01-scan:** 01-scan đóng băng, không dùng chung dữ liệu. `Mã` của 01-scan không map sang `Cụm đọc`.

## 2. Sau mỗi bước
1. `py -I "<TOOL>" check <GOC> --den <buoc>`.
2. Có LỖI thuộc bước vừa làm → gửi nguyên danh sách cho subagent đó (SendMessage) sửa JSON của nó. Tối đa 2 vòng; còn lỗi → ghi vào báo cáo, làm tiếp. (LỖI "thiếu cột" của bước SAU là bình thường.)
3. Sau bước 1–3: gom câu hỏi của subagent cần user quyết (vd hạng mục có bản vẽ nhưng không có trong BOQ, file thiếu) → hỏi user **1 câu gộp** trước khi giao bước 6–7. User trả lời → ghi `README.md` mục "Việc chưa quyết" và (nếu là dữ liệu) vào `*.user.json` tương ứng.
4. Bước 6: chia nhóm theo sheet `KHU`/cột `Mã phòng` (vd SK-01+SK-02 / ST-01 / GK+GT / CR / suite L18–20 / hành lang+sảnh thang+cabin). Lấy danh sách bằng `loc <GOC> KHU --cot Khu,Mã phòng`.
5. Bước 7: chạy `doichieu` (nháp từ TCQ_LOG) **sau** bước 3, 5, 6 rồi mới giao subagent.
6. Bước 7c–7e: chạy `docsau` sau bước 7. Sau bước 7d chạy `neo`, gửi dòng `NEO LỆCH` của từng nhóm cho subagent nhóm đó (tối đa 2 vòng, chạy lại `neo` sau mỗi vòng); còn LỆCH thì để trong CẦN XEM, không chặn build. Câu `ASK_USER` gộp thành 1 câu hỏi user khi báo kết quả.

## 3. Kết thúc
- `py -I "<TOOL>" build <GOC>` → `03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx` (bản cùng ngày tự sao `_backup`), `01_index\FILE_INDEX.xlsx`, `04_out\<DA>_BOQ_<yymmdd>_CDT.xlsx` (BOQ gốc + chỉ 3 cột chính thức).
- BOQ_TRACK có sheet **BOQ_TONGHOP** (sau DASHBOARD): 1 dòng = 1 dòng BOQ CĐT, nhìn ngang mô tả – phạm vi – bản vẽ hiệu lực – Link vật liệu – TCQ; ô nào cũng bấm được (dòng BOQ gốc, tờ PDF, file spec, TCQ_LOG). Chỉ để xem — dựng lại mỗi lần build, sửa ở SCOPE/MAT/TCQ_LOG.
- `Link vật tư` (sheet BOQ, BOQ_TONGHOP): trỏ dòng DMVT của từng mã, dựng lại mỗi lần build, mở bằng add-in SRC_LINKS. Chuyển file: xem cách add-in tìm lại đích ở dòng file trong zip bên dưới.
- Sheet BOQ chi tiết của CĐT có các cột của BOQ_TONGHOP theo cùng thứ tự, không ẩn cột nào, thêm `Loại thay đổi`, `Ngày cập nhật`. Không lặp `Item`, `Mô tả BOQ`, `ĐV`, `KL BOQ`, `Dòng BOQ gốc`, vì đã có trên chính hàng đó. Bản `_CDT` chỉ có 3 cột chính thức.
- **Sheet hiện:** `DASHBOARD`, `BOQ_TONGHOP`, sheet CĐT, `MAT`, `TCQ_LOG`, `VIEC_MO`, `ISSUES`, `RFI_LOG`, `ASK_USER`. **Sheet ẩn** (mỗi lần build ẩn lại, chuột phải tab → Unhide; `DASHBOARD` có mục lục): `SCOPE`, `DWG`, `KHU`, `AREA`, `VL_XUATHIEN`, `QUAN_SAT`, `BOQ_MAP`, `DAY_TUONG`, `SPEC`, `PAGES`, `_LINKS`, `_CODE`; `CAU_TAO`, `PHAN_VIEC` ẩn khi chưa có mẫu cấu tạo. Sửa ở sheet ẩn vẫn được kéo vào khi build. Link tới sheet ẩn không nhảy, nên link trên sheet hiện không được trỏ vào các sheet này.
- Cho nhẹ file: `SCOPE`, `DWG`, `KHU`, `AREA`, `VL_XUATHIEN` không có hyperlink; `VL_XUATHIEN` bỏ cột `Khóa`, `Dòng BOQ`, `Chữ gốc` ở sheet. JSON trong `_data` vẫn giữ đủ.
- Link file đi qua sheet `_LINKS` (cùng định dạng 01-scan) + add-in `_MASTER\vba\SRC_LINKS.xlam`: bấm ô → mở PDF **đúng trang** bằng Edge; ô nhiều link → menu chọn. Không có add-in thì bấm ô không mở gì (`_LINKS` đã bỏ cột mở dự phòng; đường dẫn file vẫn ở cột E để mở tay). Link 1 đích trong workbook (dòng BOQ gốc, 1 mã VL/TCQ) là hyperlink thường.
- File trong zip được link qua bản giải nén ở `_data\_zip` (build tự giải nén file đang được link). Link ghi tương đối theo thư mục file xlsx, kèm `BASE_DIR` (`_LINKS!H2`, thư mục file lúc build). Add-in tìm đích lần lượt: tương đối theo chỗ đặt file → theo `BASE_DIR` → bỏ `..\` rồi thử từng thư mục cha của file. Nên chuyển xlsx đi đâu trên máy này vẫn mở được khi hồ sơ gốc còn nguyên; đổi ổ, đổi máy thì để xlsx trong cây dự án. Máy khác cài add-in bằng cách chép `SRC_LINKS.xlam` vào `%APPDATA%\Microsoft\Excel\XLSTART` (gói sẵn ở `_MASTER\vba\GUI_CHO_NGUOI_KHAC`, kèm `HUONG DAN CAI.txt`; không dùng file .bat tự giải mã vì Cortex XDR chặn). File mở từ SharePoint: add-in đổi sang thư mục đồng bộ trên máy, chưa đồng bộ thì mở thẳng PDF trên web (có thể không nhảy đúng trang). Không đặt xlsx đã build vào thư mục hồ sơ `1. THÔNG TIN`/`00_input`: lần `kiemke` sau sẽ quét chúng như hồ sơ.
- Báo user, ngắn: đường dẫn file (cả DMVT) · số dòng BOQ / TCQ / MAT / DWG / DMVT · số dòng Code ID highlight · số cụm đọc sâu CAO / TB / THAP · LỖI còn lại · **CẦN XEM** (chép từ sheet DASHBOARD, gộp theo loại) · câu hỏi. Không kể lại nội dung từng dòng.
- Bản gửi CĐT: **không tự gửi**. Nhắc user kiểm cột "Mô tả NT – Đề xuất / KL – Đề xuất" (chỉ user điền) trước khi gửi.

## 4. Chạy lại / đợt mới
- Có TCQ/Addendum mới: đặt thư mục đợt vào thư mục hồ sơ → `init --tcq <TCQ gộp mới>` (nếu đổi) → `kiemke --dot-moi` ở **lần đầu** của đợt (ID file cũ giữ nguyên; so chữ ký nội dung với lần trước, kể cả file trong zip → `_thaydoi.json`: file mới / đổi / mất; danh sách đợt trước lưu `_dot\`). Chạy lại `kiemke` trong cùng đợt (vd sau khi sửa lỗi) thì **không** thêm `--dot-moi`: danh sách được cộng dồn, không mất; từng lần lưu `_dot\kiemke_<giờ>.json` → bước 1 chỉ gán vai trò file mới → `trang` (chỉ đọc file mới/đổi) → `tcq` → BOQ ra bản mới: `init --boq <file mới>` rồi `boq --thu`, `boq` (khớp dòng theo nội dung, đổi khóa trong `_data`, xem `buoc\2-boq.md`) → bước 3 **chỉ cho đợt mới** (+ ghi `Bị thay bởi` vào câu cũ) → `dwg`, `mat` → **`thaydoi`**: dòng SCOPE có link tới file đổi/mất hoặc Số BV có bản mới, MAT có nguồn spec đổi **hoặc được TCQ mới/đổi nhắc mã** (`tcq` ghi câu mới/đổi của đợt vào `_thaydoi.json`), TCQ mới/đổi + chưa phân loại, link bản vẽ trỏ bản cũ, đếm theo mã phòng → bước 4, 5 cho phần mới (bước 5: mọi mã MAT `thaydoi` báo, kể cả bước đã có JSON) → bước 6 **chỉ cho mã phòng `thaydoi` báo**, file nhóm mới số lớn hơn mọi nhóm cũ và có `"giao_lai": true` (thay kết quả nhóm cũ ở các dòng đó, không cộng dồn; xem `buoc\6-banve.md`) → `doichieu` (nối cả câu TCQ phạm vi chỉ nói theo phòng/khu; dòng nối với TCQ mới/đổi được thêm vào danh sách ảnh hưởng) → bước 7 cho dòng trong `_thaydoi.json` mục `anh_huong.SCOPE` → `docsau --thaydoi` (cụm vẫn tính lại đủ, chỉ giao cụm có dòng bị ảnh hưởng) → bước 7c, 7d cho các cụm đó → `tincay` → bước 7e → build. `check` báo XEM nếu `Cụm đọc` trong QUAN_SAT không còn (bộ tờ đổi nên cụm đổi): dùng `docsau --thieu` để giao lại cụm chưa có quan sát.
- `check` báo XEM "link … là bản cũ của <Số BV>" khi DWG đổi bản hiệu lực sau khi bước 6 gán link → giao lại bước 6 cho các dòng đó.
- Bước đã có `*.<buoc>*.json` → bỏ qua, trừ khi user bảo chạy lại. Chạy lại bước = xóa JSON LLM của bước đó rồi giao lại.
- **Không bao giờ xóa** `*.user.json`, `_snapshot.json`. User sửa trong xlsx (sheet sổ, hoặc 3 cột chính thức trong sheet BOQ) rồi bảo "build lại" → chỉ chạy `build` (tool tự kéo chỗ sửa vào `*.user.json`, user thắng). Trong các cột nội bộ ở sheet BOQ, chỉ `Mã cấp 3`, `Bên thực hiện`, `Mã cấu tạo` được giữ; các cột khác là công thức tra từ SCOPE, sửa không giữ — sửa ở sheet SCOPE. `Code ID`, `Nguồn phân loại`, `Cờ phân loại`, `Cờ bên` do máy tính, sửa không giữ. Không chèn/xóa hàng trong sheet BOQ của BOQ_TRACK: tool đối chiếu cột Item, hàng bị dời thì `build` dừng và không nhận sửa tay.

## 5. Không được
- Ghi/sửa BOQ_TRACK, FILE_INDEX bằng openpyxl/Excel hay viết script build riêng. Mọi dữ liệu đi qua JSON + `td.py build`.
- Sửa `form\COT.json` để qua kiểm. Form chỉ đổi khi user yêu cầu.
- Tải link ngoài (OneDrive/Drive…) — ghi "Chưa có (link)" trong FILE_INDEX.
- Tự lập BOQ cho hạng mục có bản vẽ nhưng không có trong BOQ CĐT (vd Main Lobby SOFITEL) khi user chưa quyết → ghi VIEC_MO.
- Lấy số liệu từ dự án khác hoặc từ thư mục đợt costing cũ (xem `D:\COSTING EXCEL\CLAUDE.md`, `ACTIVE.md`).
- **Sửa file trong `SKILL\` ở máy người dùng** (người dùng chỉ `git pull`; sửa tại chỗ sẽ gây xung đột ở lần pull sau). Xem mục 6.

## 6. Máy người dùng: tool không đọc được mẫu của dự án
Dấu hiệu: `td.py` chạy được nhưng BOQ / RFI / bản vẽ / vật liệu ra 0 dòng (vd mẫu BOQ/TCQ tiếng Việt, tên cột khác SOFITEL). **Dừng, không sửa `td.py`, `JSON.md`, `buoc\*.md`.** Làm như sau:
1. Tạo file riêng **`tool\td_local_<mã dự án>.py`** (vd `td_local_sarene.py`). Chỉ viết phần đọc mẫu khác; phần còn lại `import` từ `td.py` (không chép nguyên `td.py` — bản chép sẽ lỗi thời khi skill cập nhật).
2. Thêm `SKILL/10-tender/tool/td_local_*.py` vào `.git/info/exclude` của máy đó (không sửa `.gitignore` của repo).
3. Ghi vào `README.md` của `10_TENDER` dự án: "Dự án này dùng `td_local_<mã>.py`". Từ đó chạy `py -I "<tool local>" <lệnh> <GOC> …` thay cho `td.py`.
4. Báo người quản lý skill: gửi file `td_local_*.py` và 1 file mẫu BOQ/RFI của dự án (đã được phép dùng) để gộp vào `td.py` thành bản chung. Sau khi `git pull` bản mới có hỗ trợ, bỏ file local.
Người quản lý skill (máy gốc) sửa `td.py` rồi commit + push; người dùng chỉ pull.
