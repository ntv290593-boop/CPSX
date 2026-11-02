# Handoff: CTX tự lập BOQ (skill 10-tender), 02/11/2026

Viết cho phiên Claude tiếp theo. Đọc cả `ACTIVE.md` và `_ACTIVE\OneDrive_2026-11-01\10_TENDER\README.md`.

## 0. Cập nhật 02/11/2026 09:10 (mới nhất, thắng các mục dưới khi khác nhau)
Mục 4 bước 1–7 đã làm xong. Đã build: `03_boq\CTX_BOQ_TRACK_261102.xlsx`, `CTX_DMVT_261102.xlsx`, `04_out\CTX_BOQ_261102_CDT.xlsx` (371 dòng, DMVT 190 dòng, 186 ảnh; `check` 0 LỖI, 1 XEM `Chưa đọc sâu`; `kiem_boq` còn K2 15, K3 1, K8 CG-4001). Đã trả lời user 6 câu. Bản 261101 có comment cột W của user vẫn giữ nguyên.
- Soát lại độc lập: 137/137 tag, `_soat2\kq.json`. Tag tường lệch 3/80. Tag khác tường lệch 29/57, trong đó 20 là GL-03. Kết quả chốt: `_soat_tag.6.json` (30 tag: GL-03 về tường, 7 portal, 3 dải gỗ về tường). Còn 389 tag tường chưa soát lần 2.
- Sửa thêm: bề mặt `portal`, ghi lại tag `khac`; `Kích thước BoQ` lên `BOQ_TONGHOP`; 12 ảnh DMVT chọn lại (`Ảnh chọn`); bản dịch bước 8 `_dmvt_llm.json`; `tap_ma` nhận mã 4 số (build lại từng nhân đôi 122 dòng FF&E); `lap_boq.py` tự trỏ `boq` sang file vừa ghi; `kiem_boq` ghi `VIEC_MO.doichieu-2.json`.
- Bản `_dmvt_snapshot.json` cũ, DMVT 261101 và bản lỗi nhân đôi nằm ở `03_boq\_backup`.
- **User chốt 02/11 (đã làm, bản build 09:32: 376 dòng, DMVT 191 dòng, 187 ảnh, `check` 0 LỖI):** tính sàn sauna WD-08 (công tác FF.WF ngoài ô tick, 2 dòng); thêm dòng GL-11 (WF.MI), con tiện WD-05 (WF.WO), cấu hình `san_tinh`, `chi_tiet`, `ngoai_tick` trong `_lap_boq.json`; portal giữ cách xếp, chờ phòng tender phản hồi; CTX chưa chạy `docsau`, không hỏi lại.
- **User chốt thêm 02/11 (đã làm, bản build mới nhất: 378 dòng, DMVT 192 dòng, 188 ảnh):** tính cả sàn WD-03, WD-07 (FF.WF 4 dòng). Sửa chữ dính DMVT: `dmvt.py` tách dòng, bỏ chữ rác; `dmvt_xlsx.py` chiều cao do Excel đo, dư 2 dòng. Kiểm trên file: `do_cao` 0 hàng thấp hơn chữ.
- Chờ user: FF.WF thuộc gói nào (VIEC_MO). Xóa DXF 12 GB (mục 4 bước 8) khi user duyệt xong.
- Lúc 09:16 có 2 bản chép file đầu ra trong `1. THÔNG TIN`, sau đó đã được dời đi, không phải do tôi; quy tắc tạm trong `FILE_INDEX.kiemke.json` đã bỏ.
- Tool chưa commit: phiên costing-excel-52 commit theo 2 commit user đã chọn.

## 1. Việc đang làm
User yêu cầu lập BOQ cho dự án CTX (bản vẽ ID ghi Sun Feliza Suites) theo `Định hướng BOQ.xlsx`, chi tiết từng hạng mục trong phạm vi BOHO, rồi chạy skill `10-tender`.
Bản đầu đã giao 01/11. User soi ra lỗi, hỏi 6 câu và comment ở cột W sheet `BOQ_TONGHOP`, nói "đó mới chỉ là một số lỗi tôi soi được". Đang sửa theo gốc, **chưa build lại, chưa trả lời user**.

## 2. Quyết định của user (đừng hỏi lại)
- Khối lượng: lập danh mục + số đếm được; m² ốp để đợt 2.
- Phạm vi: theo ô tick của `PHẠM VI CÔNG VIỆC BOHO.xlsx`, phần lệch với Kế hoạch thì ghi lại. Cửa gỗ theo tick DO.WD gói 4.
- `Phần hầm` để trống. Hành lang Tháp 1 vẫn liệt kê, trạng thái Chờ.
- Không viết lại Python cho từng dự án: dùng mẫu `form\MAU_DU_AN.json` (`init --mau`), phần riêng ở `_data\_lap_boq.json`.
- Tên tháp (đã kiểm): Wealth = F1 = Tháp 1; Happiness = F2 = Tháp 2; Peace = F3 = Tháp 3; Health = F5 = Tháp 5; Longevity = F6 = Tháp 6.
- Tầng 1–6 vào sheet `Podium 1-7`, sheet `Tháp n` là tầng điển hình + cabin. Do tôi kiểm rồi chốt (Thuyết minh tr.13–14), user chưa phản hồi kết quả.
- `ACTIVE.md` đã đổi sang CTX. User chưa trả lời trực tiếp việc này; đã báo. SOFITEL chỉ tạm dừng.

## 3. 6 câu của user: nguyên nhân và trạng thái
| # | Câu hỏi | Nguyên nhân thật | Trạng thái |
|---|---|---|---|
| 1 | Link vật liệu có đủ thông tin, ảnh mà DMVT không lấy | `dmvt.py` chỉ đọc mẫu trang SOFITEL | Đã sửa (`nguon_spec: spec_boc` trong `_dmvt.json`): 190 dòng, 186 có ảnh; 4 dòng không ảnh là mã không có trong spec (GL-00, MT-00, ST-00, WD-11). Mới chạy `--chi-boc`, chưa ra xlsx |
| 2 | WD-01 xuất hiện nhiều | Đúng thiết kế: 1 dòng cho mỗi khu + vị trí (tường, trần). Trước đây còn cộng các dòng có WD-01 trong cột vật liệu joinery, đồ rời. Hiện 13 dòng mã WD-01 | Chỉ cần giải thích; có thể hỏi user có muốn thêm dòng tổng theo mã không |
| 3 | SP-05 link sang tờ có SP-06 | Tag ở Model space dùng chung 6 tờ, máy gắn đại vào tờ đầu. Nhiều tag Model còn nằm ngoài mọi viewport (không hiện trên tờ nào) nhưng vẫn bị đếm | Đã sửa: `dwg_tag.py` gán tag về tờ theo viewport. SP-05 giờ chỉ còn ở trần, tờ 1213 |
| 4 | Đồ rời, nhiều dòng không có link nguồn | 3 lỗi: regex phân biệt hoa thường ("Tag tờ") làm mất link 182/378 dòng; `ma_chuan` không nhận mã 4 số (SX-2000); đồ rời chưa có trong MAT | Đã sửa cả 3. Còn 15 dòng không có tag trên bản vẽ (có trong lịch FF&E): kiểm bằng link spec, có cờ |
| 5 | BOQ không có kích thước | Tôi chưa làm | Đã làm: đồ rời từ spec, cửa từ tờ lịch cửa, joinery từ số lớn nhất trên tờ đầu (có ghi chú). Ốp tường chưa đo. Còn thiếu: FUR.OT 7, DO.WD 2, LO.* 5, DE.LI 1 dòng |
| 6 | Đã có bước kiểm kết quả chưa | Chưa. `td.py check` chỉ kiểm cấu trúc; lúc đó tôi chỉ so tay 2 tờ. Phải nói thẳng điều này | Đã thêm `kiem_boq.py` (K1–K8) và soát bề mặt bằng hình |

Comment cột W (đã lưu `_data\_comment_user_261101.json`, file gốc ở `03_boq\_backup\CTX_BOQ_TRACK_261101_co_comment_user.xlsx`):
- SP-05 "bản vẽ là sp-06": đã xử lý.
- WD-04 sảnh Happiness "Cửa": đã xử lý, không còn dòng ốp tường.
- SP-12 Ballroom "trần": đã chuyển sang trần.
- WD-08 Bể bơi "là sàn": tờ 3413 đã loại. **Dòng vẫn còn ở ốp tường** theo tờ 3416, 3428, 3429 (spec ghi "floor and wall"). Phải báo user và xem tận mắt 3 tờ này, đừng nói là đã xong.

**Lỗi gốc:** trước đây bề mặt (tường, trần, sàn) suy từ loại tờ. Nay theo từng tag, xem mục 5.

## 4. Việc tiếp theo, theo thứ tự
1. **Soát lại độc lập chưa xong:** ảnh ở `_data\_soat2\` (137 tag, 70 tờ: toàn bộ 57 tag khác tường + 80 tag tường ngẫu nhiên), chưa có `_soat2\kq.json`. Subagent cũ dừng vì hết hạn mức; giao subagent mới, **không cho đọc `_soat_tag.*.json`**, ghi `_soat2\kq.json` dạng `{"rows":[{"khoa","ma","be_mat","ghi_chu"}]}`. Rồi `py -I SKILL\10-tender\tool\mau_soat.py <GOC> so`. Dòng khác nhau: tự xem hình, ghi kết quả chốt vào `_data\_soat_tag.<n>.json`.
   Lý do cần: ở lần soát đầu, phần lớn các tờ 2120–2233, 2301–2406, 2440–2443, 2500–2501, 2530–2536 không có vòng đỏ (lỗi gốc tọa độ, đã sửa bằng `goc`). Subagent phải tự định vị nên kém chắc hơn.
2. Xem kỹ các tag subagent báo kém chắc: 2524#2 (`khong_ro`), 2524#4, 3418#11, 3425#18, 3425#22, 3428#3 (WD-01 PORTAL FRAME xếp cửa), 2406#3, #7, #12 (con tiện gỗ, `khac`), 2433–2436 (20 tag GL-03 sát sàn xếp `san`, chỉ suy từ mặt đứng), 2520#2, 2522#3, #5 (dải gỗ xếp trần), 3445#11, 3447#13, 3427#17 (tấm chắn tiểu GL-11).
3. Chạy lại (bước `lap_boq` ghi đè file BOQ cùng ngày; `<GOC>` = `_ACTIVE\OneDrive_2026-11-01`):
   `lap_boq.py <GOC>` → `td.py kiemke`, `boq`, `dwg`, `mat` → `lap_boq_buoc.py <GOC>` → `td.py phanloai` → `td.py check` → `kiem_boq.py <GOC>` → `td.py build <GOC>`.
   `kiem_boq.py` đã sửa K7, K8 sau lần chạy cuối, cần chạy lại. Trước `build`, kiểm `~$CTX_BOQ_TRACK_*.xlsx`: có thì user đang mở file, dừng và báo.
4. Xem hình mẫu ảnh DMVT: `Nhiều ảnh, máy chọn` 31 dòng; `Cùng ảnh với dòng khác` 16 mã (CG-2104, CG-3100, CG-4101, CG-5100, LX-1102–1104, SX-3101, SX-5101, WC-01, WC-03, WC-04, WC-07, WC-09, WD-03, WD-05). Mở xlsx DMVT kiểm ảnh đúng ô.
5. Thêm 1 dòng VIEC_MO: mã FF&E trùng giữa 2 file LF: CG-4001 (LF_7240, LF_7252), AX-1000 (LF_7210, LF_7212). MAT và DMVT đang lấy 1 bản; BOQ tra đúng theo khu.
6. Cập nhật tài liệu (quy định CLAUDE.md): `DOCS\Flowchart\10-tender_flowchart.md` phần 9 + `CHANGELOG.md`, `SKILL\10-tender\SKILL.md` mục 1d: thêm `be_mat.py`, `chuan_bi_soat.py`, `soi_tag.py`, `mau_soat.py`, `kiem_boq.py`, bước soát hình, viewport/`goc` trong `dwg_tag.py`, `nguon_spec` của `dmvt.py`. Phần ghi "dmvt chưa đọc spec AID" đã lỗi thời, phải sửa. README `10_TENDER` cũng vậy.
7. Trả lời user đủ 6 câu kèm số thật, nói rõ phần nào soát rồi, phần nào chưa. Đừng nói "đã kiểm hết".
8. Dọn: xóa thư mục DXF 12 GB khi chắc không cần đọc lại:
   `C:\Users\Admin\AppData\Local\Temp\claude\d--COSTING-EXCEL\5bc3db31-5d2f-41f6-9d38-f0c2f33f7d40\scratchpad\dwgtam`. Đổi lại 196 DWG mất khoảng 20–30 phút.

Chưa làm và chưa ai yêu cầu làm ngay: đo m² ốp (đợt 2). Không tự commit: user chưa yêu cầu.

## 5. Luồng hiện tại
```
_core\dwg_tag.py   DWG → DXF (ODA) → _data\_dwg_tag.json (tag theo tờ, gồm MULTILEADER; mũi tên; viewport; goc = khung tên)
spec_boc.py        spec FM, LF, IM, SW → _data\_spec_boc.json
be_mat.py          bề mặt từng tag: nhãn tag → tên khung hình → kết quả soát (_soat_tag.*.json thắng)
chuan_bi_soat.py   ảnh khoanh đỏ các tag can_soat → _data\_soat\
lap_boq.py         phạm vi tick + tag + spec + lịch cửa → 0.BOQ_TU_LAP\…xlsx + _lap_boq_dong.json
lap_boq_buoc.py    ghi MAT, DWG, KHU, banve, doichieu, _phanloai_llm.json
kiem_boq.py        K1 tờ link có mã không · K2 không nguồn · K3 PDF khớp số tờ · K4, K5 sót · K6 đếm lại · K7 kích thước · K8 mã trùng
mau_soat.py        lấy mẫu soát lại độc lập, đo tỉ lệ lệch
```
Kết quả soát hình lần đầu: 526 tag; tường 469, cửa 16, joinery 9, sàn 20, trần 3, khác 8, không rõ 1. Máy tự xử lý thêm 52 tag (trần 31, sàn 21).
BOQ lần lập cuối (01/11 23:00): 368 dòng, 117 dòng có cờ `[XEM]`. Đây là file BOQ tự lập; **BOQ_TRACK và DMVT trong `03_boq` vẫn là bản cũ 01/11** (378 dòng), user đã xem bản đó.

## 6. Bẫy
- **Có phiên khác đang sửa skill:** `td.py`, `SKILL.md`, `COT.json`, `JSON.md` có thêm bước `docsau`, `7c`, `7d`, `7e`, `neo`, `gemini_doc` không phải của tôi. Vì vậy `check` báo XEM `Chưa đọc sâu`. Không sửa, không hoàn tác; sửa `td.py` thì đọc bản hiện tại và thay đúng đoạn. Hỏi user CTX có cần chạy `docsau` không.
- `SKILL.md` của `.claude\skills` và `SKILL` là cùng nội dung; sửa ở `SKILL\`.
- `build` dựng lại `BOQ_TONGHOP`, comment của user ở cột W sẽ mất: đã lưu, và nên nhắc user ghi comment ở đâu để không mất.
- PDF: tag là nét vẽ, không có trong lớp chữ. Số đếm lấy từ DWG cùng đợt 31/07. Memory `pdf-is-primary-input` nói PDF là đường chính; ở đây DWG là ngoại lệ, đã báo user.
- Chạy `py -I` thì script không thấy module cùng thư mục: các tool mới đã tự thêm `sys.path`.
- Trong heredoc Bash, đường dẫn Windows có thể hỏng ký tự `\`: dùng dấu `/`.
- Tầng điển hình và cabin chưa ra được khối lượng: chưa có số tầng áp dụng, chưa có số cabin. Xem sheet VIEC_MO (11 việc).
- Hạn mức: phiên trước hết hạn mức 5 giờ khi chạy subagent; soát hình tốn khoảng 100–220k token mỗi nhóm.
