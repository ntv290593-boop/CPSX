# Cách ghi dữ liệu (mọi bước của 01-scan)

Bạn chỉ ghi **file JSON**. File Excel do `pd.py build` dựng — không tự ghi xlsx, không viết script dựng file.

## Nơi ghi, tên file
`<GOC>\01_scan\_data\<DOT>\<SHEET>.<buoc>.json` — vd `BOQ_TONGHOP.boq.json`, `QUAN_SAT.banve-2.json` (bước 4 chia nhóm).
- Chỉ ghi file của bước mình. Không sửa file của bước khác, không đụng `*.user.json`, `_snapshot.json`.
- Thấy dữ liệu bước khác sai → ghi 1 dòng ISSUES (hoặc báo trong kết quả trả về), không tự sửa.

## Cột được ghi
Chạy `py -I "<TOOL>" cot --buoc <buoc>` → in mọi sheet + cột bước bạn được ghi, kiểu, danh sách giá trị hợp lệ, ghi chú. Chỉ dùng đúng tên cột đó (có dấu, đúng hoa thường). Cột lạ → build báo LỖI.

## Dạng file
```json
{"rows": [
  {"Mã": "V01", "Kích thước BV": "900×550×440", "Bản vẽ đối chứng": [["p13 E3 mặt đứng", "F004", 13], ["MB p10", "F004", 10]]}
]}
```
- Sheet KHÓA (BOQ_TONGHOP, BOQ_MAP, ISSUES, PAGES…): mỗi dòng phải có đủ cột khóa (lệnh `cot` in ra); các bước ghi các cột khác nhau của cùng một dòng, build tự ghép theo khóa.
- Sheet NỐI (QUAN_SAT, SPEC, NOTES): mỗi dòng là một dòng mới.
- Số (kiểu SO) ghi dạng số `12`, không phải `"12"`. Không có giá trị → bỏ cột, đừng ghi `""`, `0` hay `"?"`.

## Link (kiểu LINK) — bắt buộc cho mọi giá trị đọc từ tài liệu
Một link = `["nhãn", "file", trang, "ô đích"]`, bỏ phần cuối nếu không cần:
- `file` = ID trong FILES (`"F004"`) — nên dùng — hoặc đường dẫn tương đối trong `00_input\<DOT>`.
- `trang` = số trang PDF (đếm từ 1, theo file, không phải số in trên tờ).
- `ô đích` = ô trong file Excel nguồn, vd `"'B5.16 SD'!A9"`.
- Nhãn ngắn, người đọc hiểu: `"p13 E3 mặt đứng WC"`, `"02 00 04 p4 §2.02"`, `"BoQ B5.16!9"`.

Ô có nhiều link: ghi list nhiều link (ô hiện `nhãn đầu ▾n`). Muốn tự đặt chữ hiện: `{"text": "…", "links": [...]}`. Không có nguồn: ghi chuỗi `"KHÔNG THẤY – đã tìm ở …"`.

## Tự kiểm trước khi trả kết quả
`py -I "<TOOL>" check <GOC> <DOT> --den <buoc>` → sửa hết LỖI thuộc file của bạn. Dòng "XEM" là ngoại lệ để user xem, không phải lỗi.

## Nháp của tool (`*-tool.json`)
Lệnh `boq-tach`, `doichieu` ghi `<SHEET>.<buoc>-tool.json` (vd `BOQ_CDT.boq-tool.json`). Quy ước ưu tiên khi ghép: **tool < LLM < user** theo từng ô — LLM ghi cùng ô thì thắng, không báo xung đột; ô LLM không ghi thì giữ nháp tool. ID issue của tool dùng tiền tố `BT` (boq) / `DT` (doichieu) để không trùng ID LLM (`B01`, `D01`). Chạy lại lệnh sẽ ghi đè file `-tool` (file LLM không bị đụng). Đo tỉ lệ ô do tool ghi: `py -I "<TOOL>" do <GOC> <DOT>`. Lọc nhanh dòng đã ghép: `py -I "<TOOL>" loc <GOC> <DOT> QUAN_SAT Mã=V01 "Trạng thái=KHÔNG THẤY"` (`cột=gt`, `cột!=gt`, `cột~regex`, `--cot a,b`).

## File phụ trợ tìm trang (không phải sheet, không vào PROJECT_DATA)
Tên bắt đầu bằng `_` nên `build`/`check` không ghép như sheet. Dùng cho `pd.py ungvien`.
| File | Ai ghi | Nội dung |
|---|---|---|
| `_timkiem.json` | tool (`kiemke`, `ungvien`) | `trang_trung` (trang → bản giữ), `file_tong`, `loai` (loại tờ đoán từ tên tờ: danh_muc/thong_ke/mep/mat_bang/hoan_thien/chi_tiet/mat_dung/legend/khac). Không sửa tay. |
| `_timkiem.kiemke.json` | bước kiemke (LLM) | `{"du_phong": ["F004"]}` |
| `_boq_tach.json` | tool (`boq-tach`) | `{"<sheet>!<dòng>": {"ma_boq": "VL1", "dai_rong_cao": ["900","","2375"]}}` — mã cột "Mã SP" và R/S/C của BoQ để bước boq điền `Kích thước BoQ` |
| `_neo.json` | tool (`neo`) | `{"ngay", "hash" (dấu của các link lúc chạy), "tong": {CÓ NEO/LỆCH/KHÔNG KIỂM ĐƯỢC: n}, "khong_kiem": {lý do: n}, "lech": [{sh, dong, cot, ten, nhan, link, ly_do, goi_y}]}` — `check` đọc file này (chỉ XEM). Không sửa tay. |
| `_doichieu_tool.json` | tool (`doichieu`) | `{"muc": {mã: CAO/TB/THAP}, "ly_do": {mã: [lý do]}}` |
| `_timkiem.boq.json` | bước boq (LLM) | `{"khu_file": {Khu: [FID]}, "ma_cha": {mã: [mã cha]}, "tu_khoa": {mã: "từ khóa"}, "ma_bv": {mã: [mã BV hoặc từ]}}` — `check` kiểm mã/file có thật |
