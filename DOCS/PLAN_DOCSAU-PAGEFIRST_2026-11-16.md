# Plan: 7d đọc sâu THEO TRANG — biết trước cần kiểm gì, quét 1 lần, ghi 1 lần (16/11)

Đầu vào: yêu cầu của user ("trước khi quét bản vẽ A phải biết cần verify gì trên A, rồi quét và ghi 1 lần") + phản hồi chuyên gia "Page-first Evidence Architecture" (user dán trong phiên 16/11, không lưu file). Tiếp nối `PLAN_DOCSAU-READONCE-TOKEN_2026-11-14.md` (A1–A8).

## 1. Kiểm chứng phản hồi (với code và dữ liệu GBIA bản sao)
| Nhận định | Kết quả kiểm |
|---|---|
| `chia_luot1` đã gom cụm chung trang (Union-Find) | **Đúng.** Nhận định cũ "2 cụm chung trang chắc chắn mở 2 lần" sai một phần; trùng chỉ xảy ra khi thành phần bị cắt vì quá to, hoặc khi 1 agent làm theo cụm mở lại trang cho cụm thứ 2. |
| `tincay` không biết trường bị bỏ sót | **Đúng, là lỗ hổng thật.** `kt` chỉ xét trường đã có dòng; cụm thiếu hẳn `Có cánh`/`Ngăn kéo` vẫn có thể CAO. |
| Máy lập checklist từ dữ liệu sẵn có | **Được, nhưng nguồn phải là `Mã cấp 3`**, không phải regex mô tả: GBIA 147 cụm có 144 cụm có Mã cấp 3 (3 KHÔNG RÕ). |
| Callout Index từ lớp text PDF | **Không làm được ở GBIA:** 196/196 trang **không có lớp chữ**. Máy không đọc được callout hay loại tờ từ PDF. |
| Checklist phải xét loại tờ | **Đúng.** Loại tờ đoán được từ nhãn link bước 6 (+ PAGES): GBIA còn 41/196 trang chưa rõ khi chạy mới. |

## 2. Phản biện — chỗ không theo phản hồi
1. **Độ phủ kiểm ở cấp cụm × trường, không ở cấp trang × món × trường.** Một trường bắt buộc (vd `Ngăn kéo`) chỉ cần kết luận ở *một* tờ trong bộ của cụm. Bắt mọi trang kết luận mọi trường áp dụng sẽ đẻ hàng nghìn dòng `KHÔNG THẤY` vô nghĩa (mặt bằng không thể hiện ngăn kéo). Cấp trang chỉ kiểm: **mỗi món được giao trên trang có ≥ 1 dòng** (thấy / không thấy). Loại tờ chỉ để *chỉ chỗ tìm* (`cần:` / `nếu có:`), không phải điều kiện bắt buộc theo trang. 4 mức REQUIRED/CONDITIONAL/NOT_EXPECTED/UNKNOWN rút còn 2 (`BB`, `DK`) + "tờ chưa rõ loại = xem mọi trường".
2. **Không tách "đọc mù" rồi mới mở checklist trong cùng agent.** Tool không ép được thứ tự trong 1 agent (agent gọi lệnh lúc nào cũng được), còn đọc mù không biết món nào cần thì hoặc chép cả trang (tốn token đầu ra) hoặc phải phóng lại vùng ở pha B (tốn ảnh — ngược mục tiêu). Nguồn thiên lệch thật là **giá trị** BOQ, nên: danh sách cần kiểm chỉ có **tên món + tên trường**, tool bỏ kích thước và mã hoàn thiện khỏi tên món; nhóm trang không mở `xemcum`/mô tả BOQ đầy đủ; so với BOQ dời sang nhóm Đối chiếu (agent khác).
3. **Callout Index do LLM ghi, máy đối chiếu**, không do máy dò từ PDF: dòng `Dẫn chiếu` trong TRANG_QS → `ghep` so với bộ tờ của cụm → số tờ ngoài bộ hiện ở `Tờ chưa đọc` + XEM.
4. **Gemini theo trang, chưa theo trường.** Đã gom chế độ số theo trang (`gemini --trang`, 1 lần/crop cho mọi món). So field-level `(món, trường, vị trí, giá trị)` (P5 của phản hồi) cần đổi câu hỏi `--quansat` và định dạng trả lời của `_core/gemini_doc.py` (dùng chung 01-scan đóng băng) → để sau, đo lệch thật trước.
5. **Profile không có "người chịu trách nhiệm" trong code**: file `form\PROFILE.json` có `phien_ban`, `trang_thai` (user duyệt 16/11), `chu_tri`. Không đoán: thiếu Mã cấp 3 → `CHƯA PHÂN LOẠI` (dùng trường chung + XEM), không bỏ trường.

## 3. Kiến trúc đã làm (16/11)
```
docsau (tool)  cụm + hợp đồng kiểm (profile theo Mã cấp 3) + trang → món → trường (theo loại tờ)
   │           mọi trang đúng 1 nhóm chủ (≤ 8 trang, ≤ 40 lượt món); cụm không có trang → nhóm Tìm tờ
   ▼
① Trang · Nhóm n / Tìm tờ · Nhóm n (LLM, song song)  xemtrang = danh sách cần kiểm; ảnh mỗi trang 1 lần
   │           → TRANG_QS (Cụm đọc chắc | Cụm ứng viên; Khu; KL theo BV; Dày tường; Dẫn chiếu; Kiểm chéo Gemini), PAGES
   │           → gemini --trang (Claude ghi trước)
   ▼
② ghep (tool)  TRANG_QS → QUAN_SAT / SCOPE / BOQ_MAP / DAY_TUONG / _gemini_so; độ phủ → _ghep.json
   ▼
③ Đối chiếu · Nhóm n (LLM, không ảnh trừ chỗ thiếu)  so BOQ/MAT, gán món ứng viên, bù trường BB thiếu, KL mâu thuẫn
   ▼
check (độ phủ) → neo → tincay (hạ THAP khi trường giá BB chưa có dòng) → 7e → build
```
File: `tool/docsau.py` (docsau, ghep, gemini --trang, xemtrang, xemcum, giao, kiem, tincay), `tool/td.py` (lệnh `ghep`, usage), `form/PROFILE.json` (mới), `form/COT.json` (TRANG_QS + `Thành phần`, `Cụm ứng viên`, `Khu`; cột ghi được bởi `ghep-tool`; QUAN_SAT `Nguồn` + `TOOL`), `buoc/7d-docsau.md` (viết lại), `SKILL.md`, `JSON.md`, flowchart + CHANGELOG, `kiemthu.py` (GĐ7: 3 ca mới, 2 ca viết lại, 3 ca chỉnh dữ liệu theo hợp đồng kiểm).

Đo trên bản sao GBIA (đã bỏ kết quả 7d cũ): 196 trang → 26 nhóm trang (mỗi trang 1 chủ) + 6 nhóm đối chiếu; profile TU 41, OP_TUONG 39, TRAN 33, KINH 23, SAN 8, CHƯA PHÂN LOẠI 3. Bản 2 lượt cũ: 10 nhóm lượt 0 + 23 nhóm lượt 1 làm theo cụm.

## 4. Tiêu chí nghiệm thu (phản hồi mục 6) → trạng thái
| Tiêu chí | Cơ chế | Trạng thái |
|---|---|---|
| Trang có 2 chủ = 0 | `chia_trang` theo cấu trúc; kiểm thử `gd7_manifest_trang_va_du_to` | làm |
| Trường BB chưa kết luận = 0 khi xong | `thieu_hop_dong` → check LỖI, tincay THAP | làm |
| Quan sát thiếu nguồn = 0 | check LỖI HÌNH/TEXT không tờ (có từ trước) | có |
| Chưa rõ cụm mà tự kết luận = 0 | `Cụm ứng viên` không được ghép; XEM; đối chiếu quyết | làm |
| KL cộng trùng = 0 | cùng khu nhiều trang: bằng lấy 1, khác báo, không cộng | làm |
| m2/md tự suy = 0 | luật 7d; KL m2/md là DK chỉ khi có bảng thống kê | luật, máy chưa chặn |
| QUAN_SAT chép tay = 0 | `ghep` | làm |
| Revision cũ dùng âm thầm = 0 | rev từ DWG in ở danh sách trang; chưa so với `Rev` PAGES | **chưa** |

## 5. Còn lại / chờ user
1. User duyệt `form\PROFILE.json` (trường BB/DK từng loại, map Mã cấp 3 → profile, đơn vị đếm) — **đã duyệt 16/11**.
2. **Chạy thật 7d GBIA theo trang** 1–2 nhóm trang + `ghep` + 1 nhóm đối chiếu, đo: token Claude, token Gemini, số ảnh cả trang, số crop (so với số đo A1–A5 của plan 14/11). Kết quả 7d cũ của GBIA (2 lượt, 19 nhóm) vẫn nguyên trong `_data`; chạy lại `docsau` sẽ đổi `_docsau.json` sang bản 3 — nên chạy trên bản sao hoặc chép `_data` trước.
   **Đo 16/11** (bản sao GBIA ở scratchpad phiên a33a2059, đã xóa `*.docsau-<n>.json` cũ, `docsau` → 26 + 6 nhóm; chạy nhóm trang 1 (B02, 8 trang, 15 cụm) và 7 (B08, 8 trang, 10 cụm) song song, rồi `ghep`; **chưa chạy nhóm đối chiếu**):

   | | Nhóm 1 | Nhóm 7 | Cộng |
   |---|---|---|---|
   | Lượt gọi | 41 | 57 | 98 |
   | Ảnh mở (Read PNG) | 27 | 35 | 62 (~3,9/trang: cả trang + 2–4 crop + crop khung tên) |
   | cache read / write / output | 4,0M / 0,30M / 63k | 6,25M / 0,17M / 50k | 10,25M / 0,46M / 0,11M |
   | Peak ngữ cảnh | 175k | 168k | |
   | Thời gian | 25 phút | 34 phút | |
   | TRANG_QS / chỉ có `Cụm ứng viên` | 182 / 38 | 116 / 40 | 298 / 78 (66 sau ghep) |

   - Quy theo giá (cr 0,1 · cw 1,25 · out 5): ~2,2M tương đương cho 16 trang ≈ 0,136/trang; bản 14/11 ≈ 23,2M / 188 trang ≈ 0,123/trang. **Theo trang chưa rẻ hơn**; ước cả gói 26 nhóm ≈ 28M + 6 nhóm đối chiếu (bản cũ 23,2M toàn bộ). Lợi ích ở chất lượng (độ phủ món/trang, check 0 LỖI ở file nhóm), không ở token.
   - **`Cụm ứng viên` 26%**: do danh sách bỏ mã hoàn thiện nên các cụm cùng mô tả bị cắt giống hệt (!27/!28, !32/!33 = ốp gỗ vs ốp giấy; !35/!36 hắt đèn) → agent không phân biệt được. Hướng sửa (chờ user): khi ≥ 2 cụm trên cùng trang trùng tên sau khi làm sạch, `xemtrang`/`_docsau.txt` in thêm `Mã VL` để phân biệt.
   - Tốn: crop khung tên mỗi trang (Rev/Số BV đã có ở DWG), 15/16 trang không có lớp chữ nên đọc hình; output JSON nhóm 1 lớn (63k).
   - Dẫn chiếu: 14 dòng; `ghep` báo 1 ngoài bộ (!25 → FUR-2114). Các tờ chi tiết tủ FUR-1561/1571/1572/1581 đã nằm trong bộ tờ của cụm (đúng, không báo). **Lỗ hổng:** 7 dòng `Dẫn chiếu` không có `Cụm đọc` (mặt cắt tường FUR-1441/1442/1643, chỉ có hoặc không có `Cụm ứng viên`) bị `ghep` bỏ qua → tờ mặt cắt tường có thể không ai đọc. Hướng sửa: `ghep` đưa dẫn chiếu không cụm vào `_ghep.json` + XEM theo số tờ (kiểm tờ đó có thuộc bộ tờ nào không).
   **Tiếp theo:** xem `DOCS\HANDOFF_7D-TOKEN-ANH_2026-11-16.md` (kiểm phản hồi giảm token ảnh, P0 + pilot `--trang 4`).
   **Đã làm 16/11 (mục 2b handoff):** P0 + pilot `--trang 4` + 1 nhóm đối chiếu. Image Exposure −52% nhưng token quy giá +7%; `Cụm ứng viên` 22% → 4%; khung tên 7 → 0 crop. Đòn bẩy token tiếp theo = nền mỗi lượt, không phải cỡ nhóm.
3. Gemini field-level (mục 2.4), so `Rev` PAGES với DWG, loại tờ cho 41 trang chưa rõ (PAGES `Loại tờ` sau lần đọc đầu sẽ lấp).
