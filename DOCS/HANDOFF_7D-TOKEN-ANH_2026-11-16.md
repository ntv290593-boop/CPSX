# Handoff: 7d — giảm token ảnh/lượt, việc tiếp theo (16/11)

Dự án: GBIA VVIP (ACTIVE.md). Skill: `SKILL\10-tender`, bước 7d đọc sâu theo trang (commit `cf2dc45`, main-1, chưa push).
Nối tiếp: `DOCS\PLAN_DOCSAU-PAGEFIRST_2026-11-16.md` (mục 5.2 = số đo nhóm 1 + 7; file đang sửa, chưa commit).

## 1. Bối cảnh
User dán 2 phản hồi chuyên gia trong phiên 16/11 (không lưu file):
- **PH1 "One-pass Page Extraction"** (PDF → PAGE_DATA → Python mapping → kiểm độ phủ → crop đúng chỗ). Đã kiểm: **phần lớn đã có** trong 7d theo trang. `TRANG_QS` = PAGE_DATA, `ghep` = Python mapping, `PROFILE.json` + `check`/`tincay` = kiểm độ phủ, cột `Crop` = crop đúng chỗ. Cột "Hiện tại" của PH1 tả bản 2 lượt ngày 14/11. **Không đổi kiến trúc.** Lấy 2 ý: ghi món ngoài danh sách, cột `Chữ gốc` (mục 4).
- **PH2 "giảm token: vòng đời ảnh + số lượt"**. Chẩn đoán đúng hướng, nhưng sai 2 dữ kiện (mục 2).

## 2. Đã kiểm chứng (có nguồn)
| Nhận định | Kết quả | Nguồn |
|---|---|---|
| "DWG" trong 10-tender là CAD | **Sai.** Sheet `DWG` = sổ danh mục bản vẽ (Số BV, Tên tờ, Ngày issue hiệu lực, Đợt hiệu lực, File PDF, File DWG) | `form\COT.json` mục `DWG` |
| GBIA có file CAD | **Không:** 0 file .dwg, 21 PDF | `_ACTIVE\GBIA_VVIP\_ZIP_GOC\OneDrive_2026-11-14.zip` (unzip -l) |
| → Hướng D của PH2 (cắt crop theo khung chi tiết CAD) | Không làm được ở GBIA; hoãn tới dự án có CAD | |
| Rev/Số BV đã có sẵn, không cần crop khung tên | **Đúng:** manifest in ngày issue + đợt hiệu lực từ sổ DWG | `tool\docsau.py:228-229` |
| Cái gì làm agent crop khung tên | `7d-docsau.md:21` ("ảnh cả trang… xác nhận tờ, khung tên"), trường `Revision tờ` (`:24`), PAGES yêu cầu `Rev`/`Số BV` (`:30`). `Revision tờ` **không** có trong `PROFILE.json` | `buoc\7d-docsau.md` |
| "JSON không phải thủ phạm" | **Chưa đúng khi nhân giá.** Nhóm 1+7: đọc cache 10,25M×0,1 = 1,03M (48%) · ghi cache 0,46M×1,25 = 0,58M (27%) · output 0,11M×5 = 0,55M (26%) | plan 16/11 mục 5.2 |
| 346 token/dòng do lặp tên cột/Crop/Tờ | **Sai (phiên này nói nhầm).** File đã dùng `_chung` theo trang. Phần nặng là `Món` + `Cụm đọc` lặp lại ở mọi dòng của cùng món, và `Giá trị` viết văn xuôi dài | `GBIA_TEST\…\_data\TRANG_QS.docsau-1.json` (bản sao, scratchpad phiên a33a2059) |
| Đọc tuần tự từng trang | **Quy trình đã viết tuần tự** (xemtrang → ảnh cả trang → phóng → ghi). Agent có mở trước cả 8 trang hay không thì chưa xem log | `7d-docsau.md:19-30` |
| Loại tờ chưa rõ → xem mọi trường | **Đúng là lỗi:** `truong_tren_to` trả về mọi trường khi `loai` rỗng; `xemtrang` in "? (xem mọi trường)" | `docsau.py:197-203`, `:811` |
| Đoán loại tờ | `loai_to` chỉ dùng nhãn link + PAGES (`Loại tờ`, `Tên BV`), **chưa dùng `Tên tờ` của sổ DWG** | `docsau.py:186-194`, `:233` |
| Tham số nhóm | `TRANG_NHOM = 8`, `MON_NHOM = 40`, đổi được bằng `docsau --trang N --mon M` | `docsau.py:140-141`, `:265` |

Lập luận (chưa đo): ảnh nằm lại trong ngữ cảnh tới hết agent, nên lượng ảnh đọc lại tăng theo bình phương số trang của 1 agent. Chia nhóm 8 trang thành 2×4 thì phần này giảm khoảng ½, đổi lại tốn thêm 1 lần ghi cache nền khởi động agent. Phải đo mới biết.

## 2b. ĐÃ LÀM (user duyệt "tiến hành tới khi xong", 16/11) — P0 + pilot, CHƯA COMMIT
**Sửa (P0):** `tool/docsau.py`, `tool/td.py` (usage), `buoc/7d-docsau.md`, flowchart + CHANGELOG, `kiemthu.py` (+`gd7_trang_rev_loai_va_trung_ten`; --mau 92/92, GĐ7 16/16).
- Kiểm lại phát hiện thêm: dòng `rev` trong danh sách trang GBIA **trống ở 196/196 trang** — `lap_manifest` dùng `td.RX_DWG` mặc định thay vì `dwg_regex` của dự án (`FUR-\d{4}`). Nhận định "Rev đã có sẵn" ở mục 2 đúng theo code, sai theo dữ liệu. Đã sửa: 196/196 có Rev, 148/196 có số tờ trong nhãn.
- Sổ DWG GBIA ghi theo **bộ** (18 dòng / 196 trang), nên `Tên tờ` không nói loại từng tờ. `loai_to` chỉ dùng `Tên tờ` khi 1 số BV = 1 trang → GBIA vẫn 41 trang `?`. Bù bằng `xemtrang --loai` (agent tự đặt sau ảnh cả trang; pilot: nhóm 13 dùng cho 2 trang tủ).
- `xemtrang` in `Số BV / Rev (… KHÔNG crop khung tên)`; món cùng profile trùng 40 ký tự đầu trên 1 trang in thêm Mã VL; `ghep` giữ Dẫn chiếu không cụm (`_ghep.json` `dan_chieu_khong_cum`, check XEM tờ chưa đọc, xemcum báo).
- Ghi thêm vào 7d: chạy `td.py` qua Bash (2/4 agent báo PowerShell 5.1 làm mất tham số đường dẫn PDF có 2 dấu cách liền).

**Pilot** (bản sao GBIA_TEST; kết quả 16/11 cất ở `GBIA_TEST\10_TENDER\_do_1611`; `docsau --trang 4` → 49 nhóm; chạy nhóm 1+2 = F0004 p1–8 (đúng 8 trang B02 của nhóm 1 cũ), 12+13 = F0010 p7–14 (nhóm 7 cũ là p8–15, lệch 1 trang); đo bằng script scratchpad `do3.py` từ transcript agent):

| | 16/11 (2 agent × 8 trang) | P0 + `--trang 4` (4 agent × 4 trang) |
|---|---|---|
| Lượt gọi | 98 | 122 (+24%) |
| Ảnh: cả trang / crop / khung tên | 16 / 39 / 7 | 16 / 39 / **0** |
| Image Exposure (Σ lượt còn lại sau mỗi ảnh) | 1927 | **927 (−52%)** |
| cache read / write / output | 10,25M / 0,46M / 113k | 10,15M / 0,51M / 135k |
| Quy giá (cr 0,1 · cw 1,25 · out 5) | 2,17M | 2,33M (**+7%**) |
| Thời gian | 25 + 34 phút (2 agent) | 12–17 phút mỗi agent, song song |
| Dòng TRANG_QS / chỉ có `Cụm ứng viên` | 298 / 66 (22%) | 357 / **15 (4%)** |
| `check` LỖI ở file nhóm | 0 | 0 |

- **Kết luận:** Image Exposure giảm một nửa đúng như lập luận, nhưng **tổng token không giảm**: mỗi lượt đọc lại nền ~80k (prompt, 7d-docsau.md, README, SCAN-NHAN-LOAI mục 4); thêm agent = thêm lượt mang nền. Ảnh không phải phần lớn ngữ cảnh. Đòn bẩy token tiếp theo là **nền mỗi lượt** (rút 7d-docsau.md/SCAN-NHAN-LOAI cho nhóm trang) và **số lượt/trang**, không phải cỡ nhóm.
- Chất lượng tốt lên rõ (`Cụm ứng viên` 22% → 4%; bỏ crop khung tên), nhưng đây là nhờ P0, không tách được với cỡ nhóm. Output +20% vì nhiều dòng hơn (đọc đủ hơn).
- **Khuyến nghị:** giữ mặc định `--trang 8`; dùng `--trang 4` khi cần nhanh (thời gian thực ~½, đắt hơn ~7%). Chưa đổi code mặc định. Mẫu n=1, LLM dao động.
- **Đối chiếu nhóm 6** (lần đầu chạy; `docsau --trang 4 --sheet "B02 1F VVIP Lounge 2"`, 15 cụm, chỉ 8/19 trang B02 đã đọc): 20 lượt, **1 ảnh** (mở lại 1 crop cho !30), 0,57M quy giá, 6 phút. Gán 2 hắt LED ứng viên về !35/!36, 8 issue (V6-06 SPT-03 BV nhôm đục lỗ ≠ MAT; V6-07 WD-02A BV sồi ≠ MAT óc chó), 1 ASK_USER (WP-04 nhà máy làm cả giấy?). Còn KHÔNG THẤY trường giá vì 11 tờ chi tiết chưa đọc — đúng thiết kế (báo về, không tự mở). `check` 0 LỖI ở cụm nhóm 6.
- Lưu ý: `_docsau.json` của bản sao giờ là bản lọc B02 (bản 49 nhóm chép ở scratchpad phiên 7fb336c5 `_docsau_49nhom.*`).

## 2c. ĐÃ LÀM 16/11 (chiều) — agent `doc-trang` + lệnh `luu`, CHƯA COMMIT
Sau phản biện chuyên gia (user: "ok làm đi"). Đo trên transcript pilot: nền khởi động mỗi agent ≈ 30% chi phí 7d, ghi cache 27%, output 29%; cache 5 phút không hết hạn (không gap > 300 s) → bỏ ý TTL 1 giờ.
- **Sửa:** `SKILL\_agents\doc-trang.md` (mới; junction `.claude\agents` → `SKILL\_agents`; tools Read + Bash, `omitClaudeMd`, toàn bộ luật nhóm trang trong system prompt); `docsau.py` lệnh `luu` (ghi 1 trang qua stdin, kiểm như check, tự điền Khóa trang + PAGES, lỗi thì không ghi, gửi lại = thay), `giao` nhóm trang in prompt ngắn + mục Thứ tự ưu tiên/Bẫy của README; `td.py` usage + dispatch; `7d-docsau.md`, `SKILL.md`, flowchart, CHANGELOG; `kiemthu.py` (+`gd7_luu_ghi_trang`, sửa `gd7_xem_gop_giao_va_nhom_chung`) → --mau 93/93.
- **Nền đo 1 lượt:** general-purpose 41.550 → doc-trang 10.827 token (−74%, đã gồm hướng dẫn).
- **Chạy lại nhóm 1 + 7 (8 trang/agent, cùng 16 trang với mốc 2,17M), Opus:**

| | Mốc general-purpose (code trước P0) | doc-trang + luu |
|---|---|---|
| Lượt | 98 | 88 |
| Ảnh (cả trang / crop / khung tên) | 62 (16/39/7) | 54 (16/38/0) |
| Đọc cache / ghi cache / output | 10,25M / 0,46M / 113k | 5,48M / 0,24M / 94k |
| **Quy giá** | **2,17M** | **1,31M (−40%)** |
| Dòng TRANG_QS / chỉ Cụm ứng viên | 298 / 66 (22%) | 333 / 11 (3%) |
| Dòng Kích thước BV / Dẫn chiếu / issue | 32 / 14 / 9 | 42 / 20 / 12 |
| Thời gian | 25 + 34 phút | 26 + 20 phút |

- Mốc chạy code trước P0, nên phần chất lượng (Cụm ứng viên, khung tên) một phần do P0. Phần token: nền mỗi lượt ~41–48k → ~19k (nhóm 1) / 16k (nhóm 7, đọc cache 9k từ nhóm 1 khởi động trước). n=1 mỗi bên.
- **Bẫy đã gặp:** `claude -p` (headless) mặc định chạy **Sonnet** → agent làm ẩu (không crop, bỏ Gemini, 0,35M) — không dùng để so. Phiên đang mở không thấy agent mới tạo cho tới khi Claude Code nạp lại danh sách agent.
- **Còn lại (chưa làm):** `xemcum --gon` cho nhóm Đối chiếu; `effort` của doc-trang (đo kèm chất lượng); thử cỡ nhóm 4/8/16 với doc-trang; check lọc theo nhóm (agent không lọc được LỖI của nhóm mình trong 264 LỖI cả dự án — đã bảo doc-trang không chạy check).
- Kết quả lưu ở bản sao: `10_TENDER\_data\*.docsau-1/7.json` (Opus doc-trang), `_do_sonnet\` (lần Sonnet), `_do_4trang\` (pilot 4 trang), `_do_1611\` (mốc). `_docsau.json` hiện là bản đầy đủ 26 nhóm trang `--trang 8`.

## 2d. ĐÃ LÀM 16/11 (tối) — `xemcum --gon` + `check --nhom`
- `xemcum <GOC> <cụm…> --gon` / `--nhom <n>`: TRANG_QS gom theo (trang, chi tiết, món, khu, crop), ẩn dòng QUAN_SAT do `ghep` chép lại (Model `ghep ←` / `td.py ghep`), mô tả BOQ ≤ 300 ký tự; nhiều cụm 1 lệnh, đường dẫn PDF in 1 lần. Đo bản sao GBIA: cụm B02!21 9,1 → 5,2 KB (−43%); nhóm 27 (25 cụm) 151 → 96 KB (−36%). Output Bash bị cắt ~30k ký tự nên `giao` dặn ≤ 5 cụm/lần. Chưa đo token thật của nhóm Đối chiếu.
- `check --nhom <n>`: chỉ giữ LỖI/XEM nhắc tới trang (`trang_nhom`/`tim_to`), cụm (`so_nhom`/`tim_to`, kèm trang của cụm) hoặc file `.docsau-<n>.json`. LỖI "trang giao chưa đọc" tách theo nhóm (`Đọc trang · nhóm <n>: a/b trang …`). Bản sao GBIA: nhóm 1, 7 → 0 LỖI; nhóm 2 → 1 LỖI (8/8 trang chưa đọc); nhóm 27 → 36 LỖI. `doc-trang` vẫn không chạy check (`luu` đã kiểm).
- kiemthu: sửa `gd7_doc_theo_trang_ghep_va_do_phu` (câu LỖI mới + check --nhom + xemcum --gon), thêm kiểm prompt giao đối chiếu. GĐ7 17/17.
- **Đo A/B nhóm Đối chiếu 6** (bản sao `GBIA_DC` ở scratchpad phiên 21a5de68, dựng lại đúng `_data` của mốc `_do_4trang`, `ghep` lại; prompt chỉ khác 1 dòng `xemcum --gon` + `check --nhom`; Opus general-purpose; đo `do3.py` từ transcript):

| | Mốc (xemcum đủ) | --gon |
|---|---|---|
| Lượt | 20 | 13 |
| Ảnh | 1 | 1 |
| Đọc cache / ghi cache / output | 2,36M / 0,150M / 29k | 1,08M / 0,132M / 22k |
| **Quy giá** | **0,569M** | **0,384M (−33%)** |
| Thời gian | ~6 phút | 4 phút 25 |
| QUAN_SAT / SCOPE / ISSUES / ASK_USER | 43 / 13 / 8 / 1 | 36 / 13 / 6 / 0 |
| check --nhom 6 | — | 0 LỖI ở cụm nhóm (3 LỖI = nhóm trang 3/4/5 chưa đọc) |

  - Review chất lượng: không có kết luận nào mâu thuẫn với mốc; 5 issue trùng nội dung (len chân dưới trụ gỗ, WD-02A sồi/óc chó, SPT-03, BF-01B cabinet/wardrobe, nẹp MT-01 sàn–trần); thêm 1 (phào 15×10 nổi/lõm chưa rõ). **Thiếu so với mốc:** (a) !24/!31 dải cao 150 WD-01 ở E3 có thể là bục BF-02A chứ không phải len chân (nguy cơ tính trùng); (b) !27 WP-04 BV cao 2800 vs BOQ 2,84; (c) !28 trụ WD-01 cao sàn–CH 3000 vs BOQ 2,84 (chỉ ghi ở nẹp !29); (d) ASK_USER WP-04 nhà máy có làm giấy dán không (SCOPE vẫn CHƯA RÕ). 2 issue "tờ chưa đọc" của mốc nay chỉ báo trong phần trả về — đúng thiết kế. n=1 mỗi bên, không tách được do --gon hay dao động LLM.
  - Lô 5 cụm ốp tường ra 35k ký tự → bị lưu file, agent mất 1 lượt Read → đã sửa: `xemcum` tự dừng ở `XEMCUM_MAX` = 24.000 ký tự, in lệnh gọi tiếp.
- **Chạy lặp nhóm 6 (3 lần mỗi bên, 16/11 tối):** A = mốc (7d-docsau.md bản trả lời cũ `scratchpadd_moc.md`, không --gon/--nhom), B = --gon; mỗi lần 1 bản sao `DC_A1/A2/B1/B2` (scratchpad phiên 21a5de68), chạy song song; chấm bằng `scratchpad\cham6.py` (token từ transcript + dò 4 điểm, đã kiểm tay điểm a).

| | A0 | A1 | A2 | **TB A** | B0 | B1 | B2 | **TB B** |
|---|---|---|---|---|---|---|---|---|
| Quy giá (M) | 0,569 | 0,513 | 0,689 | **0,590** | 0,384 | 0,456 | 0,418 | **0,419 (−29%)** |
| Lượt | 20 | 16 | 27 | 21 | 13 | 16 | 13 | 14 |
| Ảnh | 1 | 0 | 4 | | 1 | 1 | 0 | |
| ISSUES | 8 | 7 | 8 | | 6 | 8 | 8 | |
| (a) dải 150 E3 = bục BF-02A? (tính trùng !24/!31) | có | có | có | **3/3** | – | – | có | **1/3** |
| (b) WP-04 BV 2800 vs BOQ 2,84 | có | – | – | 1/3 | – | – | – | 0/3 |
| (c) !28 trụ gỗ sàn–CH 3000 vs 2,84 | có | có | có | 3/3 | – | có | có | 2/3 |
| (d) ASK_USER WP-04 | có | – | – | 1/3 | – | – | – | 0/3 |

  - Token: 2 khoảng không chồng nhau (B cao nhất 0,456 < A thấp nhất 0,513) → −29% là thật.
  - Điểm chung 6/6: SPT-03 BV≠MAT, WD-02A sồi≠óc chó, len chân dưới trụ gỗ, gán khe hắt !35/!36. (b), (d) chỉ 1/6 → dao động, không do --gon.
  - (a) 3/3 vs 1/3: dấu hiệu --gon làm yếu, n nhỏ. Đã kiểm: `--gon` KHÔNG mất dữ kiện (mô tả BOQ BF-02A L3727 H150, TRANG_QS !31 "len 150 dưới WP-04 3727" ở E3, ghi chú BOQ "E3 phần tủ/bục không tính" đều in); khác ở chỗ chế độ đủ in mỗi quan sát 2 lần (TRANG_QS + bản chép QUAN_SAT). Hướng bù (chưa làm): tool tự báo trong xemcum "cụm khác có kích thước BOQ trùng số trên cùng trang" — kiểm máy, không trông vào LLM để ý.
- **ĐÃ LÀM 16–17/11 (đo 17/11 ở mục 2e, đã commit): hướng bù (a)** — `xemcum` in mục `TRÙNG SỐ BOQ CỤM KHÁC` (`docsau.trung_kt_cum_khac`): số (≥ 100, bỏ mã FUR-/WD-/SW3/F0004, thập phân, %) trong TRANG_QS đã gán cụm trùng kích thước BOQ (`Kích thước BoQ`, không có thì số trong Mô tả BOQ) của cụm khác cùng sheet BOQ hoặc cùng trang (Bản vẽ hiệu lực / TRANG_QS); chỉ báo khi phần trùng ngoài kích thước của chính cụm có số không chia hết 50; ≤ 6 cụm, 2 dòng/cụm. Cùng sheet là bắt buộc: BF-02A (!24) không có tờ E3 (F0004|7) trong Bản vẽ hiệu lực nên điều kiện "cùng trang" không bắt được. `giao` + 7d-docsau.md mục 8 dặn xét → ISSUES. Bản sao GBIA_DC: báo 4 cụm — B02!31 → !24 (ca mẫu), B02!27 → !24 (tường WP-04 sau bục 3727), B08!33 → BF-02B 3720 (ca tương tự), B08!32 nhiều cụm tủ cạnh (nhiễu nhất). kiemthu thêm `gd7_xemcum_trung_so_boq_cum_khac` — ca này bắt 1 lỗi: mẫu bỏ mã (`F0004`, `SW3`) xóa luôn `L3727`/`H150` (GBIA vẫn ra đúng nhờ lùi về Mô tả BOQ); đã sửa (giữ tiền tố 1 chữ L/W/H/D/R/C/Ø). **kiemthu --mau 94/94 (17/11, `py -I`).** Đã đo lại nhóm Đối chiếu 6: xem mục 2e.
  - Lưu ý chạy kiemthu: dùng `py -I` (Python 3.13 có pywin32); venv `_MASTER\tools\.venv` thiếu `win32com` → GĐ6 `gd6_build_khong_cau_hinh_van_ra_2_file` hỏng giả.

## 2e. ĐO 17/11 — thử mù 1 sheet ngẫu nhiên + chạy lặp nhóm 6 với TRÙNG SỐ
**Thử mù B04** (user: "chọn nhóm ngẫu nhiên, không biết trước kết quả"): `secrets.choice` trong 7 sheet chưa thử, ≥ 4 cụm (bỏ B02, B08) → **B04 Lift Car EV2-EV3** (13 cụm, FUR-0005 = F0006 p1–5). Bản sao mới của GBIA_TEST (scratchpad phiên 5d934f24 `GBIA_RUN`, kết quả cũ nhóm 1/7 dời ra `_cu_b02b08`), `docsau --trang 8 --sheet "B04 …"` → Trang 1 + Đối chiếu 2; Opus.
| | Trang 1 (doc-trang) | Đối chiếu 2 (--gon) |
|---|---|---|
| Lượt / ảnh | 35 / 21 (5 cả trang, 16 crop, 0 khung tên) | 12 / 1 |
| Quy giá | 0,79M (23,3k / lượt món; mốc general-purpose 41,7k) | 0,386M (29,7k / cụm; mốc xemcum đủ 39,3k) |
| check --nhom | 0 LỖI, 1 XEM | 0 LỖI, 7 XEM |
| Dòng chỉ Cụm ứng viên | 7/158 (4%) | — |
- Prompt nhóm Trang: thay danh sách món bằng "chạy xemtrang" + dòng "chạy td.py qua Bash" (lệch nhẹ so với `giao`); Đối chiếu nguyên văn `giao`.
- Kiểm tay: RCP FUR-3311 đúng "650x448mm CEILING ACCESS PANEL" + "TRAP DOOR HIDDEN IN THE CEILING" như agent báo.
- TRÙNG SỐ báo 3 cụm, đều số 448 (bề rộng tấm da LE-01 = cạnh nắp thăm 650x448): !30 → issue có ích (nắp thăm nằm trong trần da, kiểm trừ diện tích); !22/!23 nhiễu, agent bỏ qua đúng. **Nhiễu còn lại: số lưới tấm không chia hết 50 (448) lọt bộ lọc.**
- Phát hiện B04: !25 "External cladding" BV chỉ có ốp trong cabin; !33 trần kim loại MT-01 không có trên BV (chỉ flat bar); ST-01 BV 20+9 vs MAR 30; khe LED + driver vs "Lighting by MEP"; trap door không có dòng BOQ. **Lỗi tool:** nhãn Phạm vi nhà máy nháp sai — !31 nắp thăm bọc da gắn "Thạch cao", !35 sàn đá gắn "cabinet" (chưa sửa).

**Chạy lặp nhóm Đối chiếu 6 với `--gon` + TRÙNG SỐ** (C1–C3, bản sao GBIA_DC sạch, bỏ `*.docsau-6.json`, `ghep` lại, prompt `giao 6` nguyên văn; chấm `cham6c.py` = cham6.py đổi RUNS, scratchpad phiên 5d934f24):
| | A cũ (xemcum đủ) | B `--gon` | **C `--gon` + TRÙNG SỐ** |
|---|---|---|---|
| Quy giá 3 lần | 0,569 · 0,513 · 0,689 | 0,384 · 0,456 · 0,418 | **0,372 · 0,425 · 0,393** |
| TB | 0,590M | 0,419M (−29%) | **0,397M (−33%)** |
| (a) dải 150 E3 = bục BF-02A | 3/3 | 1/3 | **3/3** |
| (c) !28 trụ gỗ 3000 vs 2,84 | 3/3 | 2/3 | **3/3** |
| (b) WP-04 2800 vs 2,84 · (d) ASK WP-04 | 1/3 · 1/3 | 0/3 · 0/3 | 0/3 · 0/3 |
- C max 0,425 < A min 0,513 → −33% là thật. Điểm chung vẫn 3/3 (SPT-03, WD-02A, len chân dưới trụ, gán khe hắt). check --nhom 6: 0 LỖI ở cụm nhóm (3 LỖI = nhóm trang 3/4/5 chưa đọc). Cả 3 lần kết luận đúng "BOQ E3 đã trừ, KL không trùng".
- **Kết luận:** giữ `--gon` + TRÙNG SỐ làm chuẩn nhóm Đối chiếu. Cả 7d so cách cũ ≈ −40% token. (b), (d) cần kiểm máy riêng nếu muốn bắt chắc. n=3 mỗi bên, chưa có đáp án tuyệt đối.
- Còn lại: lọc nhiễu số lưới tấm trong TRÙNG SỐ; sửa nhãn Phạm vi nhà máy nháp (!31 B04, !35 B04); `effort` của doc-trang; cỡ nhóm 4/8/16 với doc-trang (tốn Opus, cần user duyệt).

## 3. Việc tiếp theo (cũ — P0 + pilot đã làm, xem 2b)
**P0 — sửa nhỏ (nhớ cập nhật flowchart + CHANGELOG):**
1. Bỏ crop khung tên khi sổ DWG đã có Số BV + ngày issue cho trang đó. Chỉ crop khi sổ thiếu hoặc mâu thuẫn. Sửa `7d-docsau.md:21/24/30`: `Rev`/`Số BV` trong PAGES lấy theo dòng `rev` mà `xemtrang` in ra; `Revision tờ` chỉ ghi khi khác sổ.
2. Loại tờ: `loai_to` dùng thêm `Tên tờ` của sổ DWG. Thêm `xemtrang … --loai <MB|MĐ|…>`: agent nhìn ảnh cả trang, tự xác định loại, rồi lấy đúng danh sách trường (không thêm lượt phân loại riêng). Đếm lại số trang chưa rõ loại (trước đây 41/196).
3. Lỗi đã đo 16/11: (a) các cụm trùng tên sau `ten_mon` → in thêm `Mã VL` để phân biệt (26% dòng chỉ có `Cụm ứng viên`); (b) `ghep` bỏ qua `Dẫn chiếu` không có `Cụm đọc` → đưa vào `_ghep.json` + XEM theo số tờ.

**Pilot sau P0:** dùng bản sao GBIA (scratchpad phiên a33a2059 → `GBIA_TEST`; nếu đã mất thì chép lại `_data` trước, **không chạy trên GBIA thật**). Chạy lại đúng nhóm trang 1 (B02) + 7 (B08) với `--trang 4` cộng các sửa ở P0. So với số đo 16/11 theo: token quy giá, số lượt, số ảnh (cả trang / crop / khung tên), **Image Exposure** = Σ số lượt còn lại sau khi mở mỗi ảnh (tính từ transcript agent), số dòng chỉ có `Cụm ứng viên`, `check` 0 LỖI. Sau đó chạy 1 nhóm Đối chiếu (chưa chạy lần nào).
Chưa có đáp án để chấm đúng/sai (không có CAD). Nếu cần thì dò tay 2–3 trang.

**P1:**
- Làm gọn JSON đầu ra: gom theo món (`Món`/`Cụm đọc` ghi 1 lần, các trường bên trong), `Giá trị` ngắn.
- Ghi món ngoài danh sách (sửa `7d-docsau.md:31`: hiện nay "không đoán được cụm thì bỏ qua") thành 1 dòng gọn, phục vụ phát hiện thiếu BOQ/RFI.
- Thêm cột `Chữ gốc` vào TRANG_QS cho mã/nhãn.
- `Crop` gắn hash PDF (bản nhẹ của "ROI Registry"), không âm thầm dùng crop của revision cũ.
- Chia nhóm theo tải dự kiến (ước số crop), làm sau khi có kết quả pilot `--trang 4`.

**Hoãn:** cắt crop theo CAD (hướng D), Gemini so theo trường, so `Rev` PAGES với DWG.
