# Cách ghi dữ liệu (mọi bước của 10-tender)

Bạn chỉ ghi **file JSON** (và file `.md` khi bước của bạn nói rõ). File Excel do `td.py build` dựng — không tự ghi xlsx, không viết script dựng file.

## Nơi ghi, tên file
`<GOC>\10_TENDER\_data\<SHEET>.<buoc>.json` — vd `TCQ_LOG.tcq-2.json`, `SCOPE.banve-3.json`, `VIEC_MO.tcq-1.json`.
- Chỉ ghi file của bước mình. Không sửa `*-tool.json` (tool ghi), `*.user.json`, `_snapshot.json`, `_cauhinh.json` (trừ khi bước của bạn cho phép).
- Thấy dữ liệu bước khác sai → ghi 1 dòng VIEC_MO (hoặc báo trong kết quả trả về), không tự sửa file của bước khác.

## Sheet và cột
`py -I "<TOOL>" cot --buoc <buoc>` → in sheet + cột bạn được ghi, kiểu, danh sách hợp lệ, cột BẮT BUỘC. Dùng đúng tên cột (có dấu). Cột lạ / cột không phải của bước → LỖI.
| Sheet | Loại | Khóa | Nội dung |
|---|---|---|---|
| FILE_INDEX | KHÓA | `ID` (F0001… do tool; L01… = link chưa tải) | mỗi file (kể cả file trong zip) |
| SCOPE | KHÓA | `Khóa` = `<sheet BOQ>!<dòng>` | mỗi dòng hạng mục BOQ CĐT |
| TCQ_LOG | KHÓA | `Mã TCQ` = `TCQ04-22` (đợt 2 số – STT; `N1` = ghi chú không số) | mỗi câu TCQ |
| DWG | KHÓA | `Số BV` | mỗi số bản vẽ, bản hiệu lực |
| KHU | KHÓA | `Khu` | thư mục/khu của từng bộ hồ sơ → mã phòng |
| MAT | KHÓA | `Mã VL` | mỗi mã vật liệu, trạng thái cuối |
| VIEC_MO | NỐI | — | việc còn mở, câu cần hỏi |

## Dạng file
```json
{"rows": [
  {"Mã TCQ": "TCQ04-22", "Chủ đề": "Phạm vi", "Bị thay bởi": "TCQ05-24", "Nội dung cuối (VN)": "Đã bị TCQ05-24 thay: 50-501 omit"},
  {"_loc": {"Sheet": "^0\\.Summary$"}, "Trạng thái": "Không áp dụng", "Nhóm thực hiện": "Chung"}
]}
```
- Sheet KHÓA: mỗi dòng có cột khóa, chỉ ghi các cột của bạn; tool ghép theo khóa. Với SCOPE / TCQ_LOG / DWG **không được thêm khóa mới** (khóa phải có trong nháp tool — xem bằng `loc`).
- **Dòng mẫu `_loc`** (tiết kiệm token): `{"_loc": {"cột": "regex", …}, "cột A": "giá trị"}` → áp cho mọi dòng khớp tất cả regex (không phân biệt hoa thường). Dòng có khóa cụ thể luôn ghi đè dòng mẫu trong cùng file. Dùng cho gán hàng loạt (vd mọi file trong 1 thư mục, mọi dòng 1 nhóm BOQ), rồi ghi dòng cụ thể cho ngoại lệ.
- Số (SO) ghi dạng số. Danh sách (DS) đúng 1 giá trị; DSN nhiều giá trị cách `; `. Ngày `yyyy-mm-dd`. Nhiều mã trong 1 ô (Mã VL, Dòng BOQ, Mã TCQ): cách nhau `; `.
- Không có giá trị → **bỏ cột**, đừng ghi `""`, `0`, `"?"`. Muốn **xóa** giá trị nháp sai của tool (vd `Mã TCQ` không liên quan) → ghi `"(xóa)"`.

## Link (kiểu LINK) — nguồn cho mọi giá trị đọc từ tài liệu
`["nhãn", "FID", trang, "ô"]`, bỏ phần cuối nếu không cần; nhiều link = list các link.
- `FID` = ID trong FILE_INDEX (`"F0123"`). File trong zip cũng có FID (giải nén để đọc: `py -I "<TOOL>" zip <GOC> F0123`).
- `trang` = số trang PDF đếm từ 1. `ô` = ô Excel nguồn, vd `"'TC No4_replied'!E29"`.
- Nhãn ngắn, có số bản vẽ nếu là bản vẽ: `"ID-SK-52-301 p1 mặt đứng"`, `"Finishes Schedule p12 WD-01"`.
- Không có nguồn: ghi chuỗi `"KHÔNG THẤY – đã tìm ở …"`. Cột CHU (Nội dung cuối…) thì ghi nguồn ngay trong câu: `"… (TCQ06-1)"`.

## Đọc nhanh dữ liệu đã ghép (đừng mở cả file JSON)
`py -I "<TOOL>" loc <GOC> <SHEET> "cột=gt" "cột!=gt" "cột~regex" "cột=" --cot a,b --max 50` (`cột=` = ô trống).

## Ưu tiên khi ghép
Theo từng ô: **tool (`*-tool.json`) < LLM < user (`*.user.json`)**. Nhiều file LLM cùng ghi 1 ô → file xếp sau theo tên thắng (tránh: mỗi nhóm chỉ ghi phần của mình). Ngoại lệ: các nhóm **cùng bước** (`banve-1`, `banve-2`…) cùng ghi 1 dòng dùng chung nhiều mã phòng → ô LINK (`Bản vẽ hiệu lực`, `Nguồn`) và cột mã (`Mã VL`, `Mã spec/FF&E`) được **gộp**; nhóm cứ ghi link/mã của phòng mình, không cần chép của nhóm khác.

## Tự kiểm trước khi trả kết quả
`py -I "<TOOL>" check <GOC> --den <buoc>` → sửa hết LỖI thuộc file của bạn. "XEM" = ngoại lệ để user xem, không phải lỗi.
