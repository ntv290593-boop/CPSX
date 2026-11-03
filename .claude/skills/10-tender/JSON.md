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
| CAU_TAO | KHÓA | `Khóa mẫu` = `CT-01#1` | thư viện mẫu cấu tạo cho dòng nhiều bên; chỉ tender khai trong xlsx, LLM không ghi |
| AREA | KHÓA | `Tiền tố` = `51` | tiền tố mã → tập phòng; tool dựng từ KHU, user sửa ở sheet AREA, LLM không ghi |
| VL_XUATHIEN | KHÓA | `Khóa` = `<mã>\|<FID>\|<trang>` | mã xuất hiện ở đâu; tool dựng lại mỗi lần build, chỉ để xem, không nhận sửa tay |
| VIEC_MO | NỐI | — | việc còn mở, câu cần hỏi |
| QUAN_SAT | NỐI | — | đọc sâu: 1 dòng = 1 trường của 1 `Cụm đọc`, có `Nguồn`, `Tờ · Chi tiết` |
| BOQ_MAP | KHÓA | `Khóa KL` = `<Khóa>\|<Khu>` | đọc sâu: KL BOQ và KL theo BV từng khu; khóa do tool `docsau`, không thêm mới |
| DAY_TUONG | KHÓA | `Khóa DT` = `<Cụm đọc>\|<Vị trí>` | đọc sâu: dày tường từng vị trí cửa |
| SPEC | NỐI | — | mục spec ảnh hưởng giá, cấu tạo |
| PAGES | KHÓA | `Khóa trang` = `<FID>\|<trang>` | trang đã mở, đọc bằng gì; `Số BV` ở đây thắng khung tên máy đoán khi `neo` kiểm |
| ISSUES | KHÓA | `ID` (`S01`, `V2-01`, `D01`; tool `DT01`) | chỗ lệch cần xử lý, gắn `Cụm đọc` |
| RFI_LOG | KHÓA | `RFI` = `RFI-01` | câu hỏi nháp gửi CĐT, gom từ ISSUES |
| ASK_USER | KHÓA | `ID` | quyết định nội bộ cần user; `Trả lời` user điền ở sheet |

Riêng `Mã cấp 3`, `Nguồn phân loại`, `Cờ phân loại` của SCOPE chỉ tool `phanloai` ghi. LLM trả lời qua `_data\_phanloai_llm.json` (`buoc/2b-phanloai.md`); user sửa `Mã cấp 3` bằng droplist ở sheet BOQ hoặc SCOPE. `Code ID`, `Link vật liệu` do `build` tính. `Bên thực hiện` ở `SCOPE.ben.json` hoặc user sửa ở sheet BOQ/SCOPE; không bước LLM nào ghi cột này.

Đọc sâu theo `Cụm đọc` = `Khóa` dòng đầu của cụm (`_data\_docsau.json`, `_docsau.txt` do tool `docsau`). Các cột `Kích thước BoQ`, `Kích thước BV`, `Bản vẽ đối chứng`, `Spec · ML`, `Thông số riêng`, `Mã hoàn thiện`, `Giả định tính giá` của SCOPE chỉ cần ghi ở dòng đại diện, dòng cùng cụm tự nhận khi ghép, trừ ô đã có giá trị riêng. Tool khác: `_gemini_so\<cụm>.json` (`gemini --cum`), `_neo.json` (`neo`), `_tincay.json` (`tincay`). Không sửa tay.

DMVT không qua `merge`: `_dmvt.json` (cấu hình dự án), `_dmvt_boc.json` và `_dmvt_anh\` (tool, ghi đè mỗi lần), `_dmvt_llm.json` (bước 8, theo `Khóa`), `_dmvt_snapshot.json`, `_dmvt_hang.json` (tool; không xóa snapshot). Xem `buoc/8-dmvt.md`.

## Dạng file
```json
{"rows": [
  {"Mã TCQ": "TCQ04-22", "Chủ đề": "Phạm vi", "Bị thay bởi": "TCQ05-24", "Nội dung cuối (VN)": "Đã bị TCQ05-24 thay: 50-501 omit"},
  {"_loc": {"Sheet": "^0\\.Summary$"}, "Trạng thái": "Không áp dụng"}
]}
```
- Sheet KHÓA: mỗi dòng có cột khóa, chỉ ghi các cột của bạn; tool ghép theo khóa. Với SCOPE / TCQ_LOG / DWG **không được thêm khóa mới** (khóa phải có trong nháp tool — xem bằng `loc`).
- **Giao lại một phần bước** (vd bước 6 cho mã phòng `thaydoi` báo): file nhóm mới đánh số lớn hơn mọi nhóm cũ của bước, thêm `"giao_lai": true` cạnh `"rows"`. Mỗi dòng file này chạm tới (khóa cụ thể hoặc `_loc`) bỏ hết kết quả các nhóm trước cùng bước (ô quay về nháp tool), rồi mới ghi. Không có cờ thì nhóm cùng bước được **gộp** (link, mã cộng dồn). Dòng dùng chung nhiều mã phòng: nhóm giao lại phải ghi đủ cho cả dòng; `check` báo XEM các dòng này. Chỉ áp cho sheet có khóa (không áp VIEC_MO).
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
