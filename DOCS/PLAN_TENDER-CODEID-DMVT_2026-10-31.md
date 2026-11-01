# PLAN — Skill 10-tender: Code ID, DMVT, gộp 01-scan (2026-10-31)

Trạng thái: đã chốt qua grill với user 31/10, đã review và đo lại cùng ngày (quyết định 2, 9 đổi, 10 giữ nguyên, thêm 14–16). **GĐ0 và GĐ1 xong 31/10** (`bo_thu_muc`, `ma_chuan`/`regex_ma` trong `_core`, `tim --ma`, vai trò `Scope`; kiểm bằng `kiemthu.py`; chưa commit, `kiemke` thật chưa chạy lại; PVCV mới thử bằng file giả, SOFITEL chưa có file thật). **GĐ2 xong phần code và đo 31/10 (mục 2a), `nho.json` chờ user duyệt `NHO_DUYET.xlsx`.** Chưa chạy `phanloai` trên `_data` thật. **GĐ3 và GĐ3b xong 31/10 (mục 2b, 2c), `build` thật đã chạy, danh sách bên đang TẠM chờ user gửi. GĐ4 xong 01/11 (mục 2d), `build` thật đã chạy. Từ GĐ5 trở đi chưa làm.**
Dự án đo: SOFITEL (`ACTIVE.md`). File Sarene (PVCV) và hidden sheet Sheraton trong `YYMMDD_DUAN_List of material.xlsx` chỉ là form, không lấy số liệu.

## 1. Quyết định đã chốt

| # | Chủ đề | Quyết định |
|---|---|---|
| 1 | Đầu ra | Skill 10-tender xuất 2 file riêng: BOQ_TRACK và DMVT |
| 2 | Phân loại | **Sửa 31/10:** 2 chiều độc lập, không suy từ chiều này sang chiều kia. Hạng mục thuộc nhóm nào: `Mã cấp 3` (D) và `Code ID` (G) từ `Datawork.xlsx`, mỗi dòng 1 mã. Ai làm: `Bên thực hiện`, thay `Nhóm thực hiện` + `NTP chi tiết` (mục 6a). Giữ `Trạng thái` (Base/VE/Loại trừ/Omit/Chờ) |
| 3 | Code ID | Đúng cột G, không đánh số đuôi. `Mã cấp 3` là droplist, sửa tay thì `Code ID` nhảy theo |
| 4 | Dòng cần gán | Mọi dòng có đơn vị ở 4 sheet Finishing/Furniture, kể cả dòng loại trừ. GR&Prelims và Provisional sum để trống |
| 5 | Bộ nhớ | Dùng chung ở `_MASTER`, chỉ gồm mẫu chữ → mã (không khối lượng, giá, mã phòng). Nạp từ SSF sau khi user duyệt dòng nghi ngờ; sau đó chỉ ô user sửa droplist mới vào bộ nhớ (ghi `nho_local.json` của máy, người quản lý gộp vào `nho.json`, xem mục 3); LLM chỉ nhớ trong dự án. Dự án sau: **highlight chỗ máy không chắc** |
| 6 | File scope | Kiểm kê nhận file có `PVCV` trong tên là file scope (CĐT hoặc tender lập), trừ thư mục bị loại bởi cấu hình `bo_thu_muc`. Thêm giá trị `Scope` vào danh sách `Vai trò`, không thêm cột |
| 7 | DMVT phạm vi | Toàn bộ vật tư: Finishes, Sanitaryware, Ironmongery, FF&E |
| 8 | DMVT form | `20261031-DMVT SOFITEL BASE (1).xlsx`: A:I = No, Code, Area, Material, Location, Description, NCC, Image, Remarks; 1 Alternative Option; không có cột ✓ phòng |
| 9 | Cột Area | **Đổi 31/10 sau khi đo (mục 4):** lấy theo tiền tố mã → tập phòng, bảng lập riêng từng dự án, luôn gắn cờ. Ca khó mới đọc tag trên mặt bằng. Không suy từ chữ trên trang PDF |
| 10 | Link gửi CĐT | 1 file; tender tự xóa cột link trước khi gửi. User 31/10: không làm lệnh xuất bản sạch |
| 11 | Build lại DMVT | Khớp dòng theo **tập mã** chuẩn hóa của dòng; dòng không mã đi theo dòng có mã phía trên. Giữ cột từ J trở đi. Ô A:I và ảnh sửa tay không bị ghi đè (nhận ra bằng snapshot, mục 5); spec mới khác thì ghi cột `Spec mới khác` |
| 12 | Chọn ảnh | Ưu tiên ảnh chụp, không có mới dùng bản vẽ. Ca khó: máy lọc ứng viên, LLM chọn, luôn gắn cờ |
| 13 | 01-scan | Gộp tool vào `SKILL\_core` trước; đóng băng 01-scan; chỉ xóa sau khi 1 dự án costing thật chạy được từ BOQ_TRACK |
| 14 | Dòng nhiều bên | Không chia nhóm cột, không tách dòng BOQ. Sheet `CAU_TAO` là thư viện mẫu cấu tạo, nhiều dòng BOQ trỏ chung 1 mẫu. Ban đầu để trống; chỉ tạo mẫu khi tender cần chia gói cho dòng đó. Không có khối lượng, giá, Code ID |
| 15 | Gửi nhà thầu phụ | Lệnh `td.py goi <GOC> "<bên>"` xuất mỗi bên 1 file riêng, máy tự ghép BOQ và `CAU_TAO`. Tender không ghép tay |
| 16 | Xem nội bộ | Sheet `PHAN_VIEC` do máy dựng lại mỗi lần build, chỉ để xem và lọc theo bên, giống `BOQ_TONGHOP` |
| 17 | Kiểm thử | Thêm 31/10: mỗi GĐ thêm phần kiểm vào 1 script chạy lại được, tập trung vào chỗ sai là mất dữ liệu hoặc lộ thông tin (mục 8a). Mốc đo không nới sau khi thấy kết quả; đổi mốc do user quyết định |

## 2. Số đo làm căn cứ

**Code ID** — đáp án: `SSF-BOQ-CODE ID.xlsx` cột AT (591 dòng, bỏ đuôi `.xxx`), bảng mã: `Datawork.xlsx` D, G, I.

| Cách | Đúng | Sai | Còn lại |
|---|---|---|---|
| Chỉ từ khóa cột I (mô tả dòng) | 24% | | |
| Từ khóa cột I + tên mục lớn | 31% | | |
| Xếp 3 lớp, học L4-17 → đoán L18-20 (351) | 287 (82%) | 22 (6%) | 42 (12%) |
| Xếp 3 lớp, học L18-20 → đoán L4-17 (240) | 209 (87%) | 13 (5%) | 18 (8%) |

Sai theo lớp (2 chiều cộng lại):

| Lớp | Đúng | Sai | Tỷ lệ sai |
|---|---|---|---|
| Luật | 205 | 0 | 0%, nhưng luật viết sau khi xem file nên lạc quan |
| Nhớ đúng mô tả | 213 | 10 | 4% |
| Nhớ đầu mục | 78 | 25 | 24% → **luôn highlight** |

Giới hạn: cùng một BOQ, cùng người lập. **Chưa đo trên BOQ dự án khác.**

Đo lại khi review 31/10 (script tạm, nối 586/591 nhãn vào SCOPE; đáp án vẫn là SSF cột AT):
- Chỉ chạy các luật liệt kê ở mục 3: L4-17 → L18-20 đúng 72%, sai 3%; chiều ngược đúng 83%, sai 4%. Luật đúng 165 dòng, so với 205 ở bảng trên, và luật `5x-3xx` không khớp dòng nào. Danh sách luật ở mục 3 có thể chưa đủ, hoặc cách dựng khác; không có script gốc nên chưa phân định được.
- Độ đúng cao phần lớn nhờ mô tả trùng: L18-20 trùng L4-17 ở 210/350 dòng, chiều ngược 179/236.
- Học Finishing → đoán Furniture: 36/216 dòng có kết quả (17%), bộ nhớ không trúng dòng nào. Chiều ngược 129/370 (35%). Đây là phép đo gần đúng cho BOQ khác: **có thể 65–83% dòng rơi xuống LLM và bị highlight.**
- Quyết định 4 có 659 dòng (`SCOPE`), đáp án chỉ phủ 591. Trong 71 dòng nối được nhưng không có nhãn, 64 dòng là Loại trừ. Nhóm này không chấm được.

**Ảnh** — so hash 57 ảnh của `DMVT BASE (1)` với ảnh máy cắt từ PDF:
- 50 trúng; 2 lệch hash (`52-500C`, vải voan), có thể do ảnh trơn, chưa xem bằng mắt.
- 5 không có khối để cắt: `51-200`, `50-200`, `53-200`, lót thảm, ray rèm.
- 18/19 dòng IR có từ 2 ảnh trở lên. Sanitary: 21/24 trang có từ 2 ảnh trở lên.

**Sheet MAT** (`SSP_BOQ_TRACK_261028.xlsx`): 90/92 mã tìm thấy trang bằng `DRAWING REFERENCE` (F0042, F0043, F0044). Chưa thấy `MR-05`, `SW-05`; có ít nhất 1 trang Sanitary mà regex chưa đọc được mã.

## 2a. Kết quả GĐ2 (31/10)

Đo bằng `DOCS\scan-toi-uu\tender-test\do_codeid.py`, chạy đúng code `SKILL\_core\code_id.py` + `_MASTER\code_id\luat.json`, dựng bộ nhớ từ phần học, không đọc `nho.json`. Đáp án SSF cột AT nối được 588 nhãn (2 nhãn theo độ giống), 3 nhãn không có ở BOQ CĐT.

| Học chéo tầng | n | Đúng | Sai | Xuống LLM |
|---|---|---|---|---|
| L4-17 → L18-20 | 350 | 301 (86,0%) | 10 (2,9%) | 39 (11,1%) |
| L18-20 → L4-17 | 238 | 219 (92,0%) | 6 (2,5%) | 13 (5,5%) |
| **Gộp, mốc đúng ≥ 80%, sai ≤ 6%: ĐẠT** | 588 | 520 (88,4%) | 16 (2,7%) | 52 (8,8%) |

| Lớp | Đúng | Sai | Tỷ lệ sai |
|---|---|---|---|
| Luật | 210 | 0 | 0% |
| Nhớ mô tả | 167 | 2 | 1,2% |
| Nhớ đầu mục | 143 | 14 | 8,9%, luôn highlight |

**Phân định luật 165/205:** cùng bộ luật nhóm, khác chỗ đọc tiền tố FF&E. Mã rug và rèm nằm ở `Mục phụ`, không ở dòng con, nên đọc cột `Mã spec/FF&E` sót 11 dòng rug và 34 dòng rèm. Tái lập đúng cả 2 số: 165 = 129 luật nhóm/IR/SW (nhận `IR-` cả ở đầu mô tả) + 36 FF&E đọc từ cột `Mã spec/FF&E`; 205 = 124 (`IR-` chỉ xét `Mã VL`) + 81 FF&E đọc thêm `Mục phụ`. Bản chạy thật: 129 + 81 = 210.

**Đổi so với cách viết ở mục 3, do user chưa xét:**
- Khóa nhớ đầu mục có thêm họ mã của dòng (`ST`, `WD`, `F4`, rỗng). Không thêm họ mã: chỉ áp dòng kích thước thì đúng 73%, dưới mốc; áp mọi dòng thì sai 6,1%, trên mốc.
- Dòng chỉ có kích thước không qua nhớ mô tả, vì `overall size #x#` không nói được là món gì.
- Không thêm luật `5x-4xx`, dù đo được đúng 90%: đạt mốc rồi thì không thêm luật viết sau khi xem đáp án.

**Giới hạn:** chọn cách trên sau khi xem đáp án, trên cùng BOQ, cùng người lập, nên 88,4% vẫn lạc quan. Finishing ↔ Furniture: 37,2% và 34,9% có kết quả, 0 sai, còn lại 63–65% xuống LLM. **Dự án mới có thể phần lớn dòng bị highlight**, rủi ro mục 9 chưa giảm.

**71 dòng không nhãn** (42 HOE/OSE, 23 Loose, 3 Doors, 3 Finishes): không chấm được. Thử với bộ nhớ đầy đủ thì 65 dòng xuống LLM.

**Trạng thái SOFITEL khi chưa nạp bộ nhớ:** 659 dòng, 213 theo Luật, 446 chờ. Build trên bản sao: 0 LỖI, 0 dòng đổi `Ngày cập nhật`, sheet BOQ chi tiết chưa có cột mới.

**Bảng mã:** 81 mã. Cảnh báo giữ nguyên: dòng 49 `Cửa` (HW.DO), dòng 57 `LS. HA`, dòng 69 `Rug`; 17 dòng nhóm không có mã cột G, trong đó có `OFW`.

**Nạp `nho.json`:** `nap_codeid.py deduyet` ra `10_TENDER\02_extract\NHO_DUYET.xlsx`, 78 dòng: 61 cùng đầu mục và họ mã nhưng khác mã, 12 cùng mô tả nhưng khác mã, 3 không có ở CĐT, 2 khác chữ với CĐT. Điền `Mã chốt` thì dòng vào bộ nhớ theo mã đó; để trống thì không vào. Rồi chạy `nap_codeid.py ghi`.

**Việc GĐ3 phải xử lý:**
- BOQ_TRACK thật (lưu 31/10 11:39) có 2 hàng trống chèn thêm ở `2.1 Finishing-L18-20` (hàng 148, 201), làm 157 dòng lệch `_snapshot`, nên `build` thật đang dừng. Cần user xóa 2 hàng hoặc xác nhận bỏ file đó trước khi build thật.
- Sau `pull_user` phải chạy lại `phanloai`, nếu không `Nguồn phân loại` còn ghi nguồn cũ khi user vừa sửa droplist.
- `Code ID` (cột G) chưa có trong SCOPE; `phanloai` chỉ ghi `Mã cấp 3`.

## 2b. Kết quả GĐ3 (31/10)

**Chuyển dữ liệu** (`DOCS\scan-toi-uu\tender-test\chuyen_ben.py`, sao lưu `_data\_backup\truoc-gd3-20261031-203316`): 355 dòng JSON có cột cũ (351 SCOPE gồm cả `_loc`, 4 MAT; plan ước 350), 724 khóa, 659 có `Bên thực hiện`, 65 khóa `Chung` bỏ theo quy tắc. Đối chiếu độc lập từng khóa: 0 sai. File BOQ_TRACK đã lưu không có ô sửa tay nào khác `_snapshot` ở các sheet sổ, nên bỏ 2 cột cũ không mất sửa tay.

**Tự quyết, khác với plan, do user chưa xét:**
- `Bên thực hiện` ở file riêng `SCOPE.ben.json` (bước `ben`), không ở `SCOPE.doichieu-*`: `SKILL.md` mục 4 cho xóa JSON của bước khi chạy lại, để chung thì mất bên.
- Giữ cả 2 khi chính dòng ghi rõ cả nhóm và chi tiết: `3. Furniture-L4-17!66` = `Nhà máy; Gương MR-01` (vanity kèm gương). `2.1 Finishing-L18-20!20`, `!26` chỉ thừa hưởng `Phá dỡ` từ dòng mẫu rồi bị ghi đè thành Owner, MEP, nên bỏ phần chi tiết.
- `Bên thực hiện` đã sửa được ở sheet BOQ ngay GĐ3, chưa có droplist; droplist theo danh sách bên làm ở GĐ3b.
- `_CODE` để hiện, không ẩn, để tender tra nghĩa mã.

**Kiểm:** `kiemthu.py` 34 mẫu + 9 SOFITEL đạt. Bản sao SOFITEL: `check` 0 LỖI; lần build đầu 0/724 dòng đổi `Ngày cập nhật` so với snapshot cũ; droplist 659 ô; highlight 659 = DASHBOARD (chưa có `nho.json` và chưa `--duyet-luat` nên mọi dòng đều có cờ); Excel mở được cả BOQ_TRACK và `_CDT` không báo sửa lỗi. Đổi droplist làm Code ID nhảy mới thử bằng Excel thật trên file mẫu.

**Lỗi tìm ra khi kiểm:** thẻ `<dataValidations count="0"/>` ở sheet không có ô droplist làm Excel từ chối mở cả file (openpyxl vẫn đọc được); `pull_user` không đóng workbook nên build lại cùng ngày bị Windows từ chối ghi đè (lỗi có sẵn). Cả 2 đã sửa, có phép kiểm.

**Việc GĐ3b phải xử lý:**
- Đã xử lý ở mục 2c: chữ tự do 35 giá trị, tách `; ` trong `goi`.
- Còn lại: highlight là tô tĩnh, đổi droplist xong đến lần build sau mới hết màu.

## 2c. Kết quả GĐ3b + build thật (31/10)

**Đầu vào user 31/10:** đã xóa 2 hàng ở BOQ_TRACK (hết lệch ở cả 6 sheet). `NHO_DUYET.xlsx` trả lời "OK", file không đổi, 78 ô `Mã chốt` trống, nên hiểu là chấp nhận nguyên trạng: 78 dòng không vào bộ nhớ. `nho.json`: 219 mẫu mô tả, 104 mẫu đầu mục. User điền `Mã chốt` sau thì chạy lại `nap_codeid.py ghi`. Danh sách bên: dùng bản đề xuất, user gửi bản chuẩn sau.

**File thật** `03_boq\SSP_BOQ_TRACK_261031.xlsx`, `04_out\SSP_BOQ_261031_CDT.xlsx`: 0 LỖI, 0 XEM; 724 dòng giữ `Ngày cập nhật` 2026-10-28; Excel mở không báo sửa lỗi và tính ra Code ID đủ 659 dòng. Theo nguồn: Luật 213, Nhớ-mô-tả 249, Nhớ-đầu-mục 38, LLM 159. Highlight 410/659: Luật chưa duyệt 213, LLM 159, Nhớ-đầu-mục 38. Bước 2b do phiên chính trả lời 159 dòng, chưa ai duyệt, nên đều có cờ `LLM`.

**Danh sách bên tạm** (`_data\_ben.json`): 18 bên, gộp từ 35 chữ cũ; chữ gốc còn nguyên ở `SCOPE.ben.json`. Máy đề xuất, không có nguồn từ hồ sơ. Không có dòng còn làm: MEP, Gói khác. Cờ `Nhiều bên?`: 40 dòng (plan ước 56 dòng Base theo dấu hiệu rộng hơn; bảng `loai` không tính `Sơn`, `Vải/Da`, vì 1 mã sơn có thể thuộc nhà máy hoặc NTP). Mẫu `CAU_TAO` thật: chưa có, tạo khi tender cần chia gói. 5 dòng đá kèm khung và kính kèm khung đang gộp về bên chính; dòng còn làm có cờ, chờ tender.

**Tự quyết, khác với plan, do user chưa xét:**
- Thêm cột `Dòng BOQ` vào file `goi`: `Item` trùng giữa các sheet, thiếu cột này thì báo giá về không nối lại được dòng. Khóa chỉ gồm tên sheet và số hàng.
- `goi` bỏ mã TCQ khỏi chữ mô tả, vì `Mô tả NT – Base` có ghi `(TCQ05-5)` mà plan cấm lộ mã TCQ.
- Mô tả nhắc tên bên khác (`by MEP`) chỉ báo XEM, không xóa chữ BOQ.
- `Ngày cập nhật` không tính `Bên thực hiện`, `Mã cấu tạo` (đảo quyết định ở GĐ3): đổi danh sách bên sẽ làm hàng trăm dòng hiện như vừa đổi.
- Xóa dòng ở `CAU_TAO` trong xlsx được nhận (sổ khác không); droplist bên chỉ cảnh báo để gõ được `Nhà máy; NTP đá`.

**Kiểm:** thử mẫu WC-01 3 thành phần trên 2 dòng theo đúng tiêu chí GĐ3b; 3 file `goi` kiểm cả XML trong zip; đối chứng máy soát bắt được Code ID lọt vào. Bản sao SOFITEL: 16 file `goi` đều qua máy soát. Chưa tạo file `goi` thật, vì danh sách bên còn tạm.

**Chưa làm:** sổ theo dõi đã gửi gói nào (mục 6a).

## 2d. Kết quả GĐ4 (01/11)

**File thật** `03_boq\SSP_BOQ_TRACK_261101.xlsx`, `04_out\SSP_BOQ_261101_CDT.xlsx`: 0 LỖI, 0 XEM, 724 dòng giữ `Ngày cập nhật` 2026-10-28, Excel mở được cả 2 file. `kiemthu.py` 58/58 (thêm 7 phép GĐ4).

**`VL_XUATHIEN`:** 4.368 dòng, 409 mã, chạy khoảng 1,4 giây nên `build` tự chạy lại. Theo nguồn: Spec 1.680, Bản vẽ 1.726, BOQ 553, Khác 353 (đều ở F0003, tài liệu mời thầu), TCQ 56. Bỏ 11 lượt viết liền (`SW10 0XE`, `ST09`) và 1 lượt mã phòng (`ST-01` ở Room schedule F0003 p99). Mọi dòng có nguồn; `check` báo LỖI nếu thiếu.

**Đo mới, khác giả định của plan:** lớp chữ của cả 342 tờ bản vẽ không có mã VL nào, tag là nét vẽ. Vì vậy 1.726 dòng `Bản vẽ` đều đi qua `Bản vẽ hiệu lực` của dòng BOQ (do LLM gán ở bước 6) và có cờ `Qua dòng BOQ`. Máy chưa xác nhận mã có mặt trên tờ đó.

**`Link spec`:** 551 link ở sheet BOQ chi tiết. Kiểm bằng script trên file thật: 551/551 file đích có thật, trang đích có chữ mã. Chưa ai bấm thử qua add-in trong Excel. 3/457 dòng có mã không có link, do 3 mã không có trang spec trong hồ sơ: `50-501` (TCQ06-2 dẫn link ngoài), `90-601`, `MR-05` (MAT ghi Không thấy spec). `MAT.Nguồn spec`: 0 trang sai.

**Area** (`do_area.py`, bảng dựng từ `KHU`, không học từ đáp án): **28/38** dòng khớp đúng tập phòng. Plan ghi 37 dòng, đếm lại C8:C51 có 38 ô có tên. Quy ước so: `STANDARD KING 02 / 02A` tính là SK-02; hành lang và sảnh thang mỗi loại gộp 1. 10 dòng lệch đều là máy thừa phòng, không thiếu: hàng 9, 11 thừa sảnh thang (tiền tố 50 hợp FF&E và Thảm, form chỉ ghi hành lang); hàng 13, 19, 22, 26, 27, 28, 50, 51 thừa do tiền tố đúng ở mức khu, không phân biệt từng món (đúng rủi ro mục 9). Bảng thật có 6 tiền tố: 01, 50–54. `90-` không có bảng nên để trống.

**Ca lệch:** 17 cặp mã–dòng `Lệch Area` (chủ yếu SK-02A dùng mã 51-, CR-01A dùng mã 52-). GĐ4 chỉ gắn cờ; đọc tag ca lệch là việc LLM của GĐ5.

**Tự quyết, khác với plan, do user chưa xét:**
- Bảng tiền tố là sheet `AREA` của BOQ_TRACK, sửa trực tiếp trong xlsx và đi qua `AREA.user.json`; không thêm file cấu hình riêng.
- `VL_XUATHIEN` chỉ xem, không nhận sửa tay, vì khóa đổi theo hồ sơ.
- Thêm giá trị nguồn `Khác` ngoài BOQ/bản vẽ/spec/TCQ, cho trang không phải spec và không phải tờ bản vẽ (F0003).
- Bỏ PDF TCQ khỏi phần quét trang: TCQ lấy từ TCQ_LOG, có `Mã VL` LLM đã đọc và ô nguồn.
- `regex_ma` không còn tính mã con `52-500.1` vào `52-500`, nên `tim --ma 52-500` ra ít trang hơn trước; `ma_chuan` nhận mã con. Phép kiểm GĐ1 `khong_mat_trang` đổi theo cho khớp, số vẫn phải không sót trang.
- Dạng viết cách (`SP 02`, `GL 02`) được nhận và có cờ; dạng viết liền bỏ, trừ mã của chính trang spec.
- Trang `DRAWING REFERENCE:` có mã nằm chỗ khác trên trang (F0043 p7, `SW-05`): nhận khi cả trang có đúng 1 mã VL. Plan mục 2 đo thiếu `SW-05` vì lý do này.
- `ungvien` và `RX_FFE` của lệnh `mat` vẫn chưa đổi (mục 4); đổi `ungvien` cần đo trước/sau, chưa cần cho GĐ4.

**Sửa sau khi user xem file (01/11):**
- Cột thêm vào sheet BOQ không wrap: trước đó hàng của CĐT giãn từ 13,2 lên tới 409,6 pt ở 3 sheet. Bây giờ 0 hàng lệch chiều cao so với file gốc, ở cả BOQ_TRACK và `_CDT`. Sheet của tool cao cố định 15, chữ dài xem bằng cách click.
- User chọn ẩn: `SCOPE`, `DWG`, `KHU`, `AREA`, `VL_XUATHIEN`, `_LINKS`, `_CODE`; `CAU_TAO`, `PHAN_VIEC` khi chưa có mẫu. **Không ẩn** `Cover`, `0.Summary`, `Room schedule`. Còn 14 sheet hiện, 9 sheet ẩn.
- User cho phép bỏ bớt thông tin ở sheet hệ thống: bỏ hyperlink ở 5 sheet dữ liệu ẩn, bỏ 2 cột mở dự phòng ở `_LINKS` (không có add-in thì bấm ô không mở gì, đường dẫn còn ở cột E), bỏ `Khóa`, `Dòng BOQ`, `Chữ gốc` khỏi sheet `VL_XUATHIEN`. JSON vẫn đủ. File từ 2.485 KB xuống 1.532 KB, `_LINKS` từ 18.193 xuống 9.077 dòng.
- Link tới sheet ẩn không nhảy và không báo gì (đã thử bằng Excel), nên link `DASHBOARD` đổi sang `BOQ_TONGHOP`. Sheet CĐT không tô màu tab, giữ format gốc.
- User yêu cầu sheet BOQ chi tiết có cột như `BOQ_TONGHOP`: bỏ nhóm cột ẩn, thêm `Vùng`, `Nhóm`, `Mã phòng`, `Spec · VL`, xếp cùng thứ tự `BOQ_TONGHOP`; `Mã VL`, `Mã TCQ` bấm được. Đổi cách viết ở mục 3 (cột nội bộ khác ẩn) và danh sách cột ở mục 6. Tự quyết: không lặp `Dòng BOQ gốc`, `Item`, `Mô tả BOQ`, `ĐV`, `KL BOQ` vì đã có trên chính hàng đó; giữ `Loại thay đổi`, `Ngày cập nhật` ở cuối. Tìm thêm lỗi: file gốc khai ẩn M:O ở `4. Provisional sum`, lấn vào cột tool thêm, ở cả bản `_CDT`; đã cắt khai báo lấn. Bản build thử trong thư mục tạm: 22 cột/sheet ở 6 sheet, 0 cột ẩn, 0 hàng lệch chiều cao, 13.032 ô so với `BOQ_TONGHOP` khác 0; file từ 1.532 KB lên 1.614 KB, `_LINKS` từ 9.077 lên 9.827 dòng.

## 3. Phân loại Code ID (máy trước)

Lệnh mới `td.py phanloai`, chạy sau `boq`. Thứ tự, dừng ở lớp đầu tiên có kết quả:

1. **User đã sửa** (`*.user.json`) — thắng tất cả.
2. **Luật** từ tên mục lớn + tiền tố mã (bảng dữ liệu, không viết cứng theo dự án): `IR-` → HW.DO; `SW-`/SANITARYWARE → EQP.SA; DEMOLITIONS → STR.REN; MASONRY → WET.BR; THERMAL AND MOISTURE → WET.WP; FF&E `5x-2xx` → LO.RUG (wall to wall → FF.CP); `5x-3xx` → DE.CU; `5x-7xx` → DE.LI.
3. **Nhớ mô tả**: mô tả + mục lớn, số được thay bằng `#`.
4. **Nhớ đầu mục**: mục lớn + dòng đầu mục không có đơn vị phía trên + họ mã của dòng (sửa khi làm GĐ2, xem mục 2a). Luôn highlight.
5. **LLM**: phần còn lại, chọn trong danh sách đóng của cột D hoặc `KHÔNG RÕ`.

Lưu `Nguồn phân loại` (User / Luật / Nhớ-mô-tả / Nhớ-đầu-mục / LLM) theo từng dòng.
**Highlight**: nhớ đầu mục, LLM, `KHÔNG RÕ`, nhớ mô tả lấy từ dự án khác, và **luật ở dự án mới cho tới khi user duyệt bằng `phanloai --duyet-luat`** (`5x-yyy` là cách đánh số của SOFITEL). Lưu ở cột `Cờ phân loại`. Duyệt theo ngoại lệ: DASHBOARD đếm số dòng highlight.

Trả lời mục 6: có học thêm, nhưng phần đóng góp chính là ngữ cảnh (mục lớn, đầu mục, tiền tố), không phải từ đồng nghĩa. Từ đồng nghĩa chỉ bổ sung cho LLM. Từ khóa trùng cần gỡ trong bảng luật: "side table" (LO.TB, LO.CA), "curtain" (CF.GY, DE.CU), "accessories" (HW.OT, DE.ART).

**Dữ liệu dùng chung:** `_MASTER\code_id\` gồm bảng mã (chép từ Datawork thành JSON, vì Datawork nằm trong `_ACTIVE` không lên git), `luat.json`, `nho.json`. Khi triển khai **phải thêm ngoại lệ này vào `CLAUDE.md`** (cho phép bộ nhớ chữ → mã dùng chung).

`_MASTER` được git theo dõi, còn máy người dùng chỉ `git pull` (`SKILL.md` mục 5). Vì vậy `nho.json` chỉ người quản lý sửa. Chỗ user sửa droplist ghi vào `nho_local.json` của máy (thêm vào `.gitignore`); `phanloai` đọc local trước rồi đến chung; người quản lý gộp rồi commit.

**Nạp bộ nhớ ban đầu:** từ SSF, không lấy dòng nghi ngờ: khác chữ với BOQ CĐT (`3. Furniture` dòng 68, 70: "to vanity" / "to minibar"), mục `Additional items` không có ở CĐT, nhãn mâu thuẫn trong dữ liệu (`54-404 Lamp table` = DE.LI). Xuất danh sách cho user duyệt trước khi nạp. Nối theo nội dung dòng, không theo số dòng.

**Trong xlsx:**
- `Mã cấp 3`: droplist cột D. `Code ID`: `=IFERROR(INDEX(_CODE!G:G,MATCH(<ô>,_CODE!D:D,0)),"")`; sheet `_CODE` đi kèm file.
- Đặt vào vị trí cũ của `Nhóm thực hiện`, `NTP chi tiết`, theo thứ tự `Mã cấp 3` · `Code ID` · `Bên thực hiện` · `Mã cấu tạo`.
- `Mã cấp 3` là cột nội bộ đầu tiên phải được `pull_user` đọc ngược từ sheet BOQ (các cột nội bộ khác đang là công thức, sửa không giữ). Giá trị sửa đi vào `*.user.json` và `nho_local.json`.
- 4 cột trên không nằm trong nhóm cột ẩn (`add_boq_cols` đang ẩn mọi cột nội bộ), nếu không user sẽ không thấy droplist và highlight.
- D và G hiện không trùng. Lưu ý Datawork: G57 = `LS. HA` có dấu cách, giữ đúng theo quyết định 3 và cảnh báo; dòng 99 `OFW` chưa có mã G; D49 chỉ ghi "Cửa", droplist dễ nhầm với công tác cửa.

## 4. Sổ xuất hiện vật liệu (cho link) và cột Area

Sheet/JSON mới `VL_XUATHIEN`: 1 dòng = mã VL + nguồn (BOQ/bản vẽ/spec/TCQ) + khu/mã phòng + file + trang. Máy lập từ cache `trang`, sheet `KHU`/`DWG`, `SCOPE`.

Dùng cho:
- DMVT cột link bản vẽ.
- BOQ_TRACK cột `Link vật tư`, `Link spec`.

Chuẩn hóa mã trước khi tìm: `51 - 500B` = `51-500B`, `ST 01` = `ST-01`, tên file `AREA53` = `AREA 53`. Cần hàm mới vì `ncode` trong `scan_core` trả `None` với `51-500B`; `RX_FFE` trong `td.py` cũng không nhận khoảng trắng. `tim` nhận regex nên thêm tùy chọn `--ma` thay vì đổi regex của người gọi.
**Sửa lỗi hiện có:** `td.py tim "51-500B"` chỉ ra 2/7 trang, sót F0046 p6, p9, p66, p67 và F1167 p3 (bản mới hơn). Đã sửa ở GĐ1 bằng `tim --ma`.

**Đo khi làm GĐ1 (31/10), việc GĐ4 phải xử lý:**
- `regex_ma` cố ý ưu tiên không sót vì `tim` có người đọc cột chữ khớp. VL viết liền có bắt nhầm: `SW10` là mã bưu điện London (`SW10 0XE`), `ST09` là mã màu vải; khoảng 12 trong 26 lượt trang tăng thêm của nhóm VL. `VL_XUATHIEN` ghi tự động nên không được tin dạng viết liền nếu chưa xem ngữ cảnh.
- `52 - 500.1` khớp `52-500`: mã con tính vào mã cha.
- `ncode` trả `None` với mã FF&E nên `ungvien` không dùng được mã FF&E để chấm điểm; `RX_FFE` của lệnh `mat` không nhận `51 - 500B`. GĐ1 chưa đổi 2 chỗ này vì ngoài tiêu chí; đổi `ungvien` thì đo trước/sau.

**Cột Area (đo 31/10 trên 37 dòng có tên phòng, `DMVT BASE (1)` C8:C51):**

| Nguồn | Khớp đúng tập phòng |
|---|---|
| Mã phòng có trên cùng trang PDF với mã VL (chưa thử tìm theo tên phòng) | 0/37 |
| `Mã phòng` trong SCOPE | 6/37 |
| Tiền tố mã → tập phòng | 29/37, lạc quan vì bảng lập từ chính đáp án |

Cách làm:
1. Mỗi dự án có bảng tiền tố → tập phòng, máy dựng nháp từ sheet `KHU` (`Bộ hồ sơ` = FF&E/Thảm) để user sửa. Ví dụ SOFITEL theo đáp án: 51 → SK-01, SK-02, ST-01, CR-01A; 52 → CR-01, GK-01, GT-01; 53 → JS-01; 54 → PS-01, PS-02, PS-02A. Hàng này chỉ là ví dụ, không viết cứng.
2. Mọi ô Area lấy theo cách này đều gắn cờ. Mã không có tiền tố khu (`IR-`, `WD-`…): lấy `Mã phòng` từ dòng BOQ và gắn cờ; không có thì để trống (form để trống cột này ở 19 dòng IR).
3. Ca khó = tiền tố và `Mã phòng` BOQ lệch nhau: đọc tag bằng ảnh theo QUY_TRINH bước 3, chỉ cho các mã đó. Tag là nét vẽ nên không có cách đọc bằng chữ.

Bẫy khi tra mã: `ST-01` vừa là mã đá vừa là mã phòng, nên phải xem ngữ cảnh trước khi ghi vào `VL_XUATHIEN`. Số Area của bộ IDD và bộ FF&E khác nhau (IDD Area 51 = Lift Lobby, FF&E AREA51 = Standard King), nên chỉ nối cùng `Bộ hồ sơ`.

## 5. DMVT

**Bóc A:I bằng máy** theo loại trang:
- `MATERIAL REFERENCE`: mỗi khối `ITEM 5x - yyy` → Design, Colour, Content, Width, Rubs, Supplier.
- `DRAWING REFERENCE: XX-nn` (Finishes, Sanitary, Ironmongery): Item Description, Reference, Finish, Supplier.
- `SPECIFICATION`: thêm cả SPECIAL INSTRUCTIONS (vật liệu có thể chỉ xuất hiện ở đó).

**Phiên bản:** lấy bản mới nhất **theo từng mã**, không theo cả file (260529 → 260605 → 261019 → 261023 Area 50 trong TCQ06).
**Gộp dòng:** FF&E theo Design + Colour (theo QUY_TRINH); các mã khác 1 dòng/mã.
**Áp kết luận TCQ:** spec mới nhất chưa phải kết luận cuối. DMVT đọc `Trạng thái cuối` của sheet MAT (hiện 7 mã Đổi theo TCQ, 3 Omit, 1 VE): Omit giữ dòng và tô xám theo QUY_TRINH; Đổi theo TCQ/VE giữ nội dung spec, ghi kết luận vào Remarks và gắn cờ. Mã FF&E chưa có trong MAT thì gắn cờ, không suy.
**Nhóm mục:** form hiện có mục I–VI và X (38 dòng FF&E + 19 dòng IR, số thứ tự nhảy từ 38 lên 61). Finishes và Sanitary (quyết định 7) chưa có mẫu nhóm mục; khi làm GĐ5 phải lấy mẫu từ user.
**LLM chỉ làm:** dịch cột Material và Location sang tiếng Việt, ghi chú, ca ảnh khó, đọc tag ca khó của cột Area.

**Crop ảnh (mục 7):**
1. Xác định loại trang rồi xác định khối: ảnh thuộc khối khi tâm ảnh nằm giữa 2 tiêu đề `ITEM`; trang `DRAWING REFERENCE` lấy ảnh trong vùng `Sample Swatch`. Không đoán vùng cắt.
2. Lấy ảnh gốc theo xref, không chụp cả trang. Không có ảnh raster thì render vùng 150 dpi.
3. Nhiều ảnh: ảnh chụp trước (phân biệt bằng độ đa dạng màu với bản vẽ nét), lớn nhất, đủ rõ.
4. Kiểm máy, sai thì gắn cờ:
   - Khối không có ảnh (đã gặp `51-501A`, `52-501A`).
   - Ảnh gần như trắng, độ phân giải thấp, ảnh quá nhỏ.
   - Hai mã khác Design/Colour cùng hash ảnh; cùng Design/Colour nhưng khác hash.
   - Sau khi chèn: ô trên-trái = ô dưới-phải và ở đúng cột H; số ảnh = số dòng; bỏ qua shape comment.
5. Sheet `_SOI_ANH`: chỉ các dòng cờ (mã, ảnh thu nhỏ, link trang nguồn) để user duyệt ngoại lệ.
6. Người thay ảnh: giữ ảnh, ghi nhận, build lại không đè (quyết định 11).

**Ghi file:** theo QUY_TRINH dùng Excel COM (`DispatchEx`), không lưu bằng openpyxl; chèn ảnh sau cùng, `Placement=1`, dán định dạng trước rồi mới merge; kiểm khóa `~$` trước khi chạy.

**Form và file ra:** form sạch (không dữ liệu dự án) đặt ở `SKILL\10-tender\form\`, thêm ngoại lệ vào `.gitignore` vì `*.xlsx` đang bị chặn. File ra: `03_boq\<DA>_DMVT_<yymmdd>.xlsx`. Form đã dùng tới cột AF, nên `Spec mới khác` và cột link đặt từ AG.

**Build lại:** file snapshot riêng lưu giá trị A:I máy đã ghi và hash ảnh theo từng dòng. Ô hoặc ảnh khác snapshot = người đã sửa. Giữ dòng không mã (đang có ở dòng 60–63, dưới IR-07). Khi spec đổi làm tách/gộp tập mã của 1 dòng: không tự chuyển sửa tay, gắn cờ cho user.

**Link trước khi gửi CĐT:** ngoài cột link, tender còn phải xóa `_LINKS` và `_SOI_ANH`, vì chứa đường dẫn nội bộ. Skill nhắc đủ 3 mục này khi báo kết quả; dòng cuối sheet form không ghi hướng dẫn nội bộ.

## 6. BOQ_TRACK

Sheet BOQ chi tiết của CĐT có các cột nội bộ:
`Mã cấp 3` · `Code ID` · `Bên thực hiện` · `Mã cấu tạo` · `Bản vẽ hiệu lực` · `Link spec` · `Link vật tư` · `Mã VL` · `Trạng thái` · `Mã TCQ` · `Nội dung cuối (VN)` · `Loại thay đổi` · `Việc còn mở` · `Rev/ngày/đợt`.

- Giữ nguyên 3 cột chính thức gửi CĐT và 2 cột nội bộ đang có: `Mã spec/FF&E`, `Ngày cập nhật`.
- `Link vật tư` trỏ sang dòng DMVT. Hai file build cùng lần, cùng `yymmdd`, nên link được dựng lại mỗi lần build. Thứ tự build: DMVT trước để có số dòng, BOQ_TRACK sau. `link_item` chưa hỗ trợ file ngoài `FILE_INDEX`, phải thêm loại đích này.
- Bỏ `Nhóm thực hiện`/`NTP chi tiết` ở cả `SCOPE`, `MAT`, `BOQ_TONGHOP`, `DASHBOARD`, `form\COT.json`, `JSON.md`, `buoc\5-mat.md`, `buoc\7-doichieu.md`. Sheet `SCOPE` vẫn là kho dữ liệu.
- **Chuyển dữ liệu cũ:** 350 dòng trong `SCOPE.doichieu-*.json` (gồm cả dòng mẫu `_loc`) và `MAT.mat.json` đang ghi 2 cột này, bỏ khỏi COT thì `check` báo "cột lạ". Sao lưu, rồi dùng script **chuyển** giá trị SCOPE sang `Bên thực hiện` trước khi đổi COT, không xóa mất: NTP → giá trị `NTP chi tiết` (325 dòng SCOPE đang có); Nhà máy, Owner, MEP, Gói khác giữ nguyên; Chung bỏ. Riêng MAT xóa cột (chỉ 4/92 mã có giá trị). `derive` so sánh bỏ qua cột mới, để lần build đầu không đổi `Ngày cập nhật` của mọi dòng.
- 3 dòng IR-01 (Owner + Base) giữ `Bên thực hiện` = Owner, cùng `Nội dung cuối (VN)` đang có.
- Bước `doichieu` không còn gán Nhóm nên bớt LLM.
- BOQ đầu vào vẫn là BOQ CĐT. `SSF-BOQ-CODE ID` chỉ làm dữ liệu học.

## 6a. Bên thực hiện và gửi nhà thầu phụ

**Chỗ nhập chỉ có 2 nơi:** sheet BOQ và `CAU_TAO`. Khối lượng và giá chỉ nằm ở dòng BOQ, không cộng theo thành phần.

| Nơi | Cột | Cách có giá trị |
|---|---|---|
| Sheet BOQ | `Bên thực hiện` | Dòng 1 bên: droplist theo danh sách bên của dự án. Dòng có `Mã cấu tạo`: máy gộp từ `CAU_TAO`, ví dụ `NCC giấy; Nhà máy; NTP sắt`; sửa ở BOQ không có tác dụng và bị báo XEM |
| Sheet BOQ | `Mã cấu tạo` | Tender chọn cho dòng nhiều bên, ví dụ `CT-01` |
| `CAU_TAO` | `Mã cấu tạo` · `Thành phần` · `Mã VL` · `Bên thực hiện` · `Nguồn` | Tender khai, mỗi mẫu 1 lần. Khóa `CT-01#1`. Thành phần chỉ có trên bản vẽ phải ghi nguồn tờ/trang |

**`Bên thực hiện` không suy từ Code ID.** Cùng `WF.WP` nhưng `2. Finishing-L4-17!203` đang gán NTP, còn `!457` gán Nhà máy.

**Nháp:** SOFITEL lấy từ dữ liệu cũ (mục 6). Dự án mới: `Mã VL` → `Loại` trong MAT → bảng `Loại → Bên` riêng từng dự án, luôn gắn cờ; không có `Mã VL` thì để trống.

**Cờ dòng nghi nhiều bên:** máy chỉ gắn cờ, không tự tạo mẫu. Dấu hiệu: từ 2 mã VL khác bên, hoặc mô tả có `plywood/MDF`, `framework`, `metal frame`, `consisting of`… SOFITEL đo 31/10 (4 sheet, 659 dòng): 56 dòng Base có dấu hiệu, gộp được 28–36 mẫu, khoảng 70–100 dòng `CAU_TAO` nếu khai hết. Từ 2 mã VL chưa chắc 2 bên, và cấu tạo chỉ có trên bản vẽ chưa được đếm, nên đây chỉ là ước lượng.

**DASHBOARD:** số dòng theo từng bên; số dòng Base ở 4 sheet **chưa có bên**; số dòng còn cờ nhiều bên.

**Sheet `PHAN_VIEC`** (BOQ_TRACK, chỉ xem): 1 dòng = dòng BOQ + phần việc + bên; dòng 1 bên có phần việc là cả dòng. Có lặp BOQ nhưng do máy dựng lại, không có trong `_snapshot` và `pull_user`.

**File gửi** `td.py goi <GOC> "<bên>"` → `04_out\<DA>_GOI_<bên>_<yymmdd>.xlsx`, workbook mới, không chép từ BOQ_TRACK:

| Cột | Lấy từ |
|---|---|
| Item BOQ | SCOPE |
| Mô tả | `Mô tả NT – Base` nếu có (đã theo TCQ), không thì `Mô tả BOQ` |
| Phần việc của bên | `CAU_TAO`, dòng 1 bên ghi cả dòng |
| Mã VL | `CAU_TAO`, dòng 1 bên lấy từ SCOPE |
| Bản vẽ | Số BV dạng chữ từ `Bản vẽ hiệu lực` |
| ĐV, KL BOQ (tham khảo) | SCOPE |
| Ghi chú | `VE` hoặc `Chờ xác nhận` khi dòng ở trạng thái đó |
| ĐV NTP, KL NTP, Đơn giá | Để trống cho bên nhận điền |

- Lấy dòng Base, VE, Chờ; bỏ Loại trừ, Omit, Không áp dụng.
- Không có Code ID, Mã TCQ, `Nội dung cuối`, `Việc còn mở`, link nội bộ, giá, hay dòng và tên của bên khác.
- KL NTP không cộng ngược vào BOQ, vì đơn vị có thể khác (khung sắt tính kg, BOQ tính m2).
- Tender kiểm file trước khi gửi; skill không tự gửi.

**Chưa làm:** sổ theo dõi đã gửi gói nào, báo giá về chưa, TCQ đổi dòng nào sau khi gửi. Làm sau khi lệnh `goi` chạy ổn.

## 7. Máy / LLM (mục 10)

| Việc | Máy | LLM |
|---|---|---|
| Kiểm kê, nhận PVCV, bản hiệu lực | Toàn bộ | Vai trò file lạ |
| Code ID | Luật, nhớ | Phần còn lại |
| Bên thực hiện | Nháp, gộp từ `CAU_TAO`, gắn cờ | Không; tender quyết định |
| File gửi NTP, `PHAN_VIEC` | Toàn bộ | Không |
| A:I DMVT | Bóc khối có cấu trúc | Dịch, ghi chú |
| Ảnh | Cắt khối, chọn, kiểm | Ca khó, luôn gắn cờ |
| Link | `VL_XUATHIEN` | Không |
| Area | Bảng tiền tố, `Mã phòng` BOQ, gắn cờ | Đọc tag ca lệch |
| TCQ | Tách câu, nối mã | Kết luận nội dung cuối, như hiện tại |

## 8. Thứ tự triển khai và tiêu chí xong

| GĐ | Việc | Xong khi |
|---|---|---|
| 0 | **Làm trước mọi GĐ:** thêm `bo_thu_muc` vào cấu hình, loại `new\New folder` (`kiemke` duyệt cả thư mục này, sẽ nhận file Sarene PVCV, SSF, Datawork, DMVT vào FILE_INDEX SOFITEL) | `kiemke` không thêm file nào từ `New folder` |
| 1 | `_core`: chuẩn hóa mã, sửa `tim`; nhận `PVCV` trong `kiemke` | `tim --ma "51-500B"` ra đủ 7 trang, gồm F0046 p9 và F1167 p3; chuẩn hóa đúng `51 - 500B`, `ST 01`, `AREA53`; file PVCV giả dựng để thử nhận vai trò `Scope` (SOFITEL chưa có PVCV thật, nên chạy lại với file thật khi có) |
| 2 | `_MASTER\code_id`, lệnh `phanloai`, ngoại lệ `CLAUDE.md`, script đo vào `DOCS\scan-toi-uu\` | **Trước khi code `phanloai`:** đưa script đo vào repo, phân định vì sao luật lệch 165 và 205 dòng (mục 2). **Đo trước khi nạp SSF vào `nho.json`** (nạp rồi mới đo là lộ đáp án), học chéo L4-17 ↔ L18-20: đúng ≥ 80%, sai ≤ 6% tính gộp 2 chiều (mốc plan 35/591 = 5,9%). Thứ tự lớp đúng: User thắng tất cả, `nho_local` đọc trước `nho`, mỗi dòng có `Nguồn phân loại`. Báo riêng độ phủ Finishing ↔ Furniture và 71 dòng không nhãn |
| 3 | Chuyển dữ liệu cũ, đổi cột BOQ_TRACK, droplist, `pull_user`, highlight | `check` 0 LỖI sau chuyển; đối chiếu từng giá trị cũ của 350 dòng với `Bên thực hiện` theo quy tắc mục 6, không dòng nào mất; lần build đầu không đổi `Ngày cập nhật`; đổi droplist → Code ID nhảy và giá trị vào `*.user.json` + `nho_local.json`; build lại vẫn giữ; highlight đúng các lớp ở mục 3 và DASHBOARD đếm khớp số dòng highlight; bản `_CDT` không có cột nội bộ |
| 3b | `Bên thực hiện`, `Mã cấu tạo`, sheet `CAU_TAO`, `PHAN_VIEC`, lệnh `goi` (mục 6a). Làm trước DMVT vì `Mã VL` của `CAU_TAO` nối sang DMVT | Thử mẫu WC-01 3 thành phần trên 2 dòng BOQ: sửa mẫu 1 lần thì cả 2 dòng đổi; build lại vẫn giữ; 3 file `goi` chỉ có đúng phần việc của từng bên, máy soát không có giá, Code ID, cột nội bộ, link nội bộ, dòng hay tên của bên khác; DASHBOARD đếm đúng dòng Base chưa có bên |
| 4 | `VL_XUATHIEN`, bảng tiền tố Area dựng từ `KHU`; thêm cột `Link spec` vào sheet BOQ chi tiết | Đo lại Area với bảng dựng từ `KHU`, không lập từ đáp án, báo số khớp trên 37 dòng; không có dòng không nguồn; `ST-01` mã đá không ghi thành mã phòng và ngược lại; không nối Area khác `Bộ hồ sơ` (IDD Area 51 khác FF&E AREA51); bấm `Link spec` mở đúng trang |
| 5 | DMVT: bóc, áp trạng thái MAT, crop, build COM, giữ sửa tay; thêm cột `Link vật tư` vào sheet BOQ chi tiết | So với `DMVT BASE (1)`: ảnh tính bằng perceptual hash 16×16, lệch ≤ 20 bit, ≥ 47/57 (mốc đo 31/10; md5 chỉ trúng 19/57 nên không dùng), mọi ảnh không trúng có cờ trong `_SOI_ANH`; 57 dòng A:I đối chiếu, riêng cột C báo số khớp và không tính là điều kiện qua; Finishes/Sanitary kiểm mẫu vì chưa có đáp án; áp đúng `Trạng thái cuối` (Omit giữ dòng và tô xám; Đổi theo TCQ/VE có Remarks và cờ); sửa tay giữ qua build lại; tách/gộp tập mã thì gắn cờ, không tự chuyển sửa tay; mở lại bằng Excel không lỗi; bấm `Link vật tư` tới đúng dòng DMVT |
| 6 | Gộp pd.py vào `_core`, đóng băng 01-scan | 10-tender không còn gọi `01-scan\tool\pd.py`; các lệnh `mail`, `chu`, `xem`, `anh` chạy qua `_core` cho kết quả giống bản cũ trên cùng file đầu vào |

Mỗi giai đoạn đổi skill: cập nhật `DOCS\Flowchart\` + `CHANGELOG.md` trong cùng phiên, theo `CLAUDE.md`. 10-tender chưa có flowchart, lần đổi đầu phải tạo mới.
Không commit lẫn phần 01-scan đang dở (xem `DOCS\HANDOFF_TOOL-HOA-01SCAN_2026-10-31.md`).

## 8a. Kiểm thử

**Script:** `DOCS\scan-toi-uu\tender-test\kiemthu.py`, chạy cả bộ bằng 1 lệnh. Mỗi GĐ thêm phần kiểm của mình vào đây. GĐ chỉ tính xong khi chạy lại vẫn đạt cả phần của các GĐ trước.

**Dữ liệu:**
- `--mau`: logic thuần, dùng dữ liệu mẫu nhỏ tự tạo, không lấy số liệu dự án. Lên git được; lưu JSON hoặc để script tự dựng, vì `*.xlsx` đang bị chặn.
- `--sofitel`: phần cần dữ liệu thật, chỉ chạy trên máy có `_ACTIVE`. Luôn chạy trên bản sao trong thư mục tạm, không ghi vào `10_TENDER` thật.

**Tự động, không đạt thì GĐ chưa xong:**

| GĐ | Kiểm gì | Nếu sai thì |
|---|---|---|
| 0 | `kiemke` không thêm file nào từ `New folder` | Số liệu dự án khác lẫn vào FILE_INDEX |
| 1 | Chuẩn hóa mã, `tim --ma` đủ 7 trang | Sót trang, kéo theo sai ở GĐ4 và GĐ5 |
| 2 | Thứ tự lớp phân loại, `nho_local` trước `nho` | Mã sai nhưng trông như đúng |
| 3 | Chuyển 350 dòng từng giá trị, `Ngày cập nhật` không đổi | Mất dữ liệu trong khi `check` vẫn báo 0 LỖI |
| 3, 5 | Droplist, ô A:I, ảnh sửa tay còn nguyên sau build lại | Công người sửa bị ghi đè |
| 3b | File `goi` không có thông tin nội bộ, giá, hay bên khác | Lộ thông tin ra nhà thầu phụ |
| 4 | Bẫy `ST-01`, không nối Area khác `Bộ hồ sơ` | Link và Area trỏ sai khu |
| 5 | Áp `Trạng thái cuối` của MAT, gắn cờ khi tách/gộp tập mã | Mã đã Omit hoặc đã đổi vẫn hiện như spec gốc |
| 6 | Cùng đầu vào, `_core` ra giống bản cũ | Gộp tool làm hỏng các bước đọc file |

**Đo và báo số:** độ đúng Code ID (có mốc ở GĐ2), ảnh so với `DMVT BASE (1)` (có mốc ở GĐ5), cột Area (chỉ báo số). Script đo lưu trong repo để chạy lại được.

**Kiểm bằng mắt 1 lần mỗi GĐ liên quan:** mở file bằng Excel, bấm link, chất lượng phần LLM dịch và chọn ảnh.

**Chưa kiểm được:** độ đúng trên BOQ dự án khác, vì SOFITEL không trả lời được câu này. Đo khi có dự án thứ 2.

## 9. Rủi ro còn mở

- GĐ2 đạt mốc nhờ khóa nhớ đầu mục chọn sau khi xem đáp án (mục 2a); chưa có số đo độc lập.
- Độ đúng Code ID trên BOQ khác chưa biết; kế hoạch dựa vào highlight, không dựa vào số 82–87%. Số dòng highlight có thể lớn (65–83% theo phép đo gần đúng ở mục 2), khi đó không còn là duyệt theo ngoại lệ.
- Cột Area theo tiền tố đúng ở mức khu, không phân biệt được từng món; 29/37 là mức lạc quan (mục 4), nên ô nào cũng gắn cờ.
- Lọc "5 xưởng" của costing chuyển sang lọc `Bên thực hiện` = Nhà máy, không dùng bảng Code ID → xưởng. Trước khi xóa 01-scan phải thử trên 1 dự án costing thật (quyết định 13).
- File gửi NTP chỉ đúng khi `Bên thực hiện` đúng. Dòng Base chưa có bên sẽ không vào file nào, nên tender phải xem số này trên DASHBOARD trước khi gửi. Cấu tạo chỉ có trên bản vẽ phải có người khai, máy không tự thấy.
- Một file DMVT + xóa tay trước khi gửi có thể làm lộ đường dẫn nếu tender quên (quyết định 10, user đã chọn). Cùng file đó còn có `Ghi chú giá` (L), `UNIT RATE` (AB:AC) và `Spec mới khác`; user 31/10 không cần lệnh xuất bản sạch, tender tự kiểm trước khi gửi.
- Main Lobby (BOQ thứ 2, F1175) nằm ngoài plan này, làm riêng theo README `10_TENDER`.
- `Trạng thái` đang dựa vào scope hiện tại; khi có PVCV, bước `doichieu` đọc nó như nguồn scope thay vì gắn cột mới.
