---
name: 10-tender
description: Quét hồ sơ mời thầu cho PHÒNG TENDER (toàn gói, không lọc phạm vi nhà máy) của dự án ghi ở ACTIVE.md — theo dõi thay đổi cuối cùng qua các đợt TCQ/Addendum (vật liệu, bản vẽ, phạm vi), ra <GOC>\10_TENDER\03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx (BOQ CĐT + cột nhà thầu + cột nội bộ + sổ TCQ_LOG/MAT/DWG/KHU/SCOPE) và bản gửi CĐT. Dùng khi user nói "scan tender", "quét HSMT", "tender", "cập nhật TCQ", "có TCQ mới", "làm BOQ tender", "build lại tender".
---

# 10-tender — ĐIỀU PHỐI

Phiên này chỉ **điều phối**: chạy tool, giao việc, đọc kết quả kiểm. **Không tự đọc BOQ / TCQ / bản vẽ / spec ở đây** — mỗi bước giao 1 subagent mới (ngữ cảnh sạch, chỉ đọc file hướng dẫn của bước đó).
Khác `01-scan`: quét **toàn gói**, gắn nhãn chứ không bỏ dòng; KHÔNG áp quy tắc "phạm vi nhà máy 5 xưởng"; các đợt nối tiếp, đợt sau đè đợt trước.

## 0. Chuẩn bị (không hỏi user nếu đã rõ)
- `GOC` = thư mục dự án trong `D:\COSTING EXCEL\ACTIVE.md`. `TOOL` = `D:\COSTING EXCEL\SKILL\10-tender\tool\td.py`, chạy `py -I "<TOOL>" <lệnh> <GOC> …` (không tham số = xem lệnh).
- Chưa có `GOC\10_TENDER\_data\_cauhinh.json` → `init` với thư mục hồ sơ, BOQ CĐT, file TCQ **gộp mới nhất**, bản dịch nội bộ (lấy từ ACTIVE.md; thiếu thì hỏi user 1 câu). Có TCQ gộp mới hơn → chạy lại `init --tcq <file mới>`.
- Thư mục trong hồ sơ không phải hồ sơ thầu (file học, form, dữ liệu dự án khác) → `init --bo-thu-muc "<thư mục>;<thư mục 2>"` (đường dẫn tính từ thư mục hồ sơ, ghi vào khóa `bo_thu_muc` của `_cauhinh.json`, thay cả danh sách cũ). `kiemke` không đi vào các thư mục này và báo XEM nếu tên không khớp thư mục nào. Thư mục đã có trong FILE_INDEX thì lần `kiemke` sau báo file mất.
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
| 8 | build | `build` | (mục 3) | bạn | 7 |
| 9 | goi | `goi <GOC> "<bên>"` | (mục 1a) | bạn, khi tender cần | 8 |

Prompt giao việc (thay `<…>`; không thêm số liệu hay suy đoán của bạn vào prompt):
```
Làm bước <buoc> của skill 10-tender.
Đọc rồi làm đúng theo: D:\COSTING EXCEL\SKILL\10-tender\buoc\<file>.md và D:\COSTING EXCEL\SKILL\10-tender\JSON.md
GOC = <…>   TOOL = D:\COSTING EXCEL\SKILL\10-tender\tool\td.py
[chia nhóm: Nhóm <n>, chỉ làm: <đợt TCQ / mã phòng / sheet>. Ghi file *.<buoc>-<n>.json]
Trả về ≤ 15 dòng: file đã ghi, số dòng, việc chưa làm được và lý do, câu cần hỏi user.
```

**Lõi tìm trang dùng chung với 01-scan** (`SKILL\_core\scan_core.py`: đọc chữ + khung tên, trang trùng, BM25, mã chuẩn hóa). Lệnh: `trang` (cache chữ), `tim --ma <mã>` (trang có mã, mọi cách viết), `tim <regex>` (trang có chữ), `ungvien --phong <mã>` (trang ứng viên mỗi dòng BOQ, chia nhóm loại tờ). Sửa lõi hoặc cách chấm điểm → đo trước/sau: `DOCS\scan-toi-uu\bench.py` (01-scan) và `ungvien <GOC> --do` (độ trùng với tờ LLM đã gán; SOFITEL 29/10: dòng có ≥1 trang trùng @12 61%, @40 80%).

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
- Cột `Link spec` (ở sheet BOQ và BOQ_TONGHOP): `MAT.Nguồn spec` nếu có, không thì trang Định nghĩa mới nhất của mã. Mở bằng add-in SRC_LINKS.
- DASHBOARD đếm mã chưa có trang spec, dòng có mã chưa có link, cặp lệch Area. `MAT.Nguồn spec` trỏ trang không có chữ mã thì ra dòng `XEM spec`. Đo cột Area: `DOCS\scan-toi-uu\tender-test\do_area.py`.

## 2. Sau mỗi bước
1. `py -I "<TOOL>" check <GOC> --den <buoc>`.
2. Có LỖI thuộc bước vừa làm → gửi nguyên danh sách cho subagent đó (SendMessage) sửa JSON của nó. Tối đa 2 vòng; còn lỗi → ghi vào báo cáo, làm tiếp. (LỖI "thiếu cột" của bước SAU là bình thường.)
3. Sau bước 1–3: gom câu hỏi của subagent cần user quyết (vd hạng mục có bản vẽ nhưng không có trong BOQ, file thiếu) → hỏi user **1 câu gộp** trước khi giao bước 6–7. User trả lời → ghi `README.md` mục "Việc chưa quyết" và (nếu là dữ liệu) vào `*.user.json` tương ứng.
4. Bước 6: chia nhóm theo sheet `KHU`/cột `Mã phòng` (vd SK-01+SK-02 / ST-01 / GK+GT / CR / suite L18–20 / hành lang+sảnh thang+cabin). Lấy danh sách bằng `loc <GOC> KHU --cot Khu,Mã phòng`.
5. Bước 7: chạy `doichieu` (nháp từ TCQ_LOG) **sau** bước 3, 5, 6 rồi mới giao subagent.

## 3. Kết thúc
- `py -I "<TOOL>" build <GOC>` → `03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx` (bản cùng ngày tự sao `_backup`), `01_index\FILE_INDEX.xlsx`, `04_out\<DA>_BOQ_<yymmdd>_CDT.xlsx` (BOQ gốc + chỉ 3 cột chính thức).
- BOQ_TRACK có sheet **BOQ_TONGHOP** (sau DASHBOARD): 1 dòng = 1 dòng BOQ CĐT, nhìn ngang mô tả – phạm vi – bản vẽ hiệu lực – Mã VL/spec/Link spec – TCQ; ô nào cũng bấm được (dòng BOQ gốc, tờ PDF, MAT, file spec, TCQ_LOG). Chỉ để xem — dựng lại mỗi lần build, sửa ở SCOPE/MAT/TCQ_LOG.
- Sheet BOQ chi tiết của CĐT có các cột của BOQ_TONGHOP theo cùng thứ tự, không ẩn cột nào, thêm `Loại thay đổi`, `Ngày cập nhật`. Không lặp `Item`, `Mô tả BOQ`, `ĐV`, `KL BOQ`, `Dòng BOQ gốc`, vì đã có trên chính hàng đó. Bản `_CDT` chỉ có 3 cột chính thức.
- **Sheet hiện:** `DASHBOARD`, `BOQ_TONGHOP`, sheet CĐT, `MAT`, `TCQ_LOG`, `VIEC_MO`. **Sheet ẩn** (mỗi lần build ẩn lại, chuột phải tab → Unhide; `DASHBOARD` có mục lục): `SCOPE`, `DWG`, `KHU`, `AREA`, `VL_XUATHIEN`, `_LINKS`, `_CODE`; `CAU_TAO`, `PHAN_VIEC` ẩn khi chưa có mẫu cấu tạo. Sửa ở sheet ẩn vẫn được kéo vào khi build. Link tới sheet ẩn không nhảy, nên link trên sheet hiện không được trỏ vào các sheet này.
- Cho nhẹ file: `SCOPE`, `DWG`, `KHU`, `AREA`, `VL_XUATHIEN` không có hyperlink; `VL_XUATHIEN` bỏ cột `Khóa`, `Dòng BOQ`, `Chữ gốc` ở sheet. JSON trong `_data` vẫn giữ đủ.
- Link file đi qua sheet `_LINKS` (cùng định dạng 01-scan) + add-in `_MASTER\vba\SRC_LINKS.xlam`: bấm ô → mở PDF **đúng trang** bằng Edge; ô nhiều link → menu chọn. Không có add-in thì bấm ô không mở gì (`_LINKS` đã bỏ cột mở dự phòng; đường dẫn file vẫn ở cột E để mở tay). Link 1 đích trong workbook (dòng BOQ gốc, 1 mã VL/TCQ) là hyperlink thường.
- File trong zip được link qua bản giải nén ở `_data\_zip` (build tự giải nén file đang được link). Link ghi tương đối theo thư mục file xlsx — chuyển cả thư mục `10_TENDER` + hồ sơ cùng nhau thì link không gãy.
- Báo user, ngắn: đường dẫn file · số dòng BOQ / TCQ / MAT / DWG · số dòng Code ID highlight · LỖI còn lại · **CẦN XEM** (chép từ sheet DASHBOARD, gộp theo loại) · câu hỏi. Không kể lại nội dung từng dòng.
- Bản gửi CĐT: **không tự gửi**. Nhắc user kiểm cột "Mô tả NT – Đề xuất / KL – Đề xuất" (chỉ user điền) trước khi gửi.

## 4. Chạy lại / đợt mới
- Có TCQ/Addendum mới: đặt thư mục đợt vào thư mục hồ sơ → `init --tcq <TCQ gộp mới>` (nếu đổi) → `kiemke` (ID file cũ giữ nguyên; so chữ ký nội dung với lần trước, kể cả file trong zip → `_thaydoi.json`: file mới / đổi / mất) → bước 1 chỉ gán vai trò file mới → `trang` (chỉ đọc file mới/đổi) → `tcq` → bước 3 **chỉ cho đợt mới** (+ ghi `Bị thay bởi` vào câu cũ) → `dwg`, `mat` → **`thaydoi`**: dòng SCOPE có link tới file đổi/mất hoặc Số BV có bản mới, MAT có nguồn spec đổi, TCQ chưa phân loại, link bản vẽ trỏ bản cũ, đếm theo mã phòng → bước 4, 5 cho phần mới → bước 6 **chỉ cho mã phòng `thaydoi` báo** → `doichieu` → bước 7 cho dòng có `Mã TCQ` mới hoặc `thaydoi` báo → build.
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
