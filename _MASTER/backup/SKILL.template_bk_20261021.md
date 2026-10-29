---
name: "cpsx-master"
description: "Skill DUY NHẤT cho workbook costing/BoQ nội thất (có sheet Chiet tinh, Material, Door, Template, MATRIX…). Chạy nhanh bằng thư viện CPSX_LIB (sheet ẩn _LIB): tính giá hạng mục, chép block, link BoQ, mô tả, bỏ tô vàng, ghi quy tắc. Thay thế các skill cpsx-project-rules, costing-door-*, costing-chiet-tinh-hao-hut, cpsx-ve, boq-melamine-laminate-toggle."
---

# CPSX MASTER v2 — chạy nhanh bằng recipe

**Mục tiêu tốc độ:** mỗi việc lặp lại = **1 lệnh `execute_office_js` gọi CPSX_LIB** (+1 lệnh nếu cần sửa thêm). Không dò sheet, không viết script mới, không đọc lại từng ô, không chụp ảnh khi recipe đã trả về kết quả.

Không đọc thêm skill CPSX nào khác: nội dung của chúng đã gộp vào đây.

## 0. Khởi động: 1 lệnh duy nhất mỗi phiên

Chạy đoạn này **ngay lệnh đầu tiên** (không gọi read_surface_context, get_range_as_csv hay đọc `_RULES` riêng):

```js
const s=context.workbook.worksheets.getItemOrNullObject("_LIB");await context.sync();
if(s.isNullObject) return "NO_LIB";
const r=s.getRange("A1:A20");r.load("values");await context.sync();
const L=new Function("context","Excel",r.values.map(x=>x[0]).join(""))(context,Excel);
return await L.map({codes:""});   // codes: regex lọc mã BoQ, vd "^D0" — để "" nếu cần cả bảng
```

`L.map` trả về:
- `rules`: các quy tắc Active/Pending trong `_RULES`.
- `masterRules`: các quy tắc chung đang Active.
- `boq`: tên sheet BoQ và các cột đơn giá, thành tiền, khối lượng, mô tả.
- `items`: mỗi dòng gồm `[dòng, mã, mô tả, KL, công thức đơn giá, có mô tả?]`.
- `chietTinh.blocks`: mỗi block gồm `[hdr, END, mã, dòng BoQ, giá]`.
- `library`: các block mẫu trong Door/Template/Kitchen, mỗi block gồm `[tên, hàng đầu, END]`.

**Giữ kết quả này trong hội thoại**, không đọc lại các vùng đã có.

Các lệnh sau dùng cùng 3 dòng nạp `L` rồi đổi dòng cuối, ví dụ `return await L.newItem({...})`.

| Kết quả lệnh đầu | Xử lý |
|---|---|
| `"NO_LIB"` | Báo 1 dòng: *"File chưa có _LIB: chạy Alt+F8 → PERSONAL.XLSB!InstallCpsxLib rồi Ctrl+S (1 lần/file)"*. Trong lúc chờ: dùng **chế độ dự phòng** (§6). |
| Lỗi có chữ `eval`, `Function`, `unsafe`, `CSP` | Môi trường chặn `new Function`. Dùng chế độ dự phòng cho cả phiên, ghi `_RULES` (MOITRUONG) bằng `L.addRules`. |
| `lib` khác `{{VER}}` | Báo người dùng chạy lại `InstallCpsxLib` (thư viện cũ). Vẫn dùng được. |
| Không có `rules` (chưa có `tblRules`) | Tạo `_RULES` theo §5.3, rồi tiếp tục. |

## 1. Bảng recipe (tham số đầy đủ)

| Việc người dùng nói | Gọi |
|---|---|
| "tính D03 tương tự D01", "lấy mẫu … sang chiết tính, link BoQ" | `L.newItem({src:{code:"D01"}, boqRow:157, edits:[…]})` hoặc `src:{sheet:"Door",row:266}` / `{sheet:"Template",row:1321}` (row = dòng đầu block có số ở B, lấy từ `library`) |
| "tính giá hạng mục 30–35" (nhiều dòng) | `L.newItems({items:[{src,boqRow:30,edits},{src,boqRow:31,edits},…]})`: tính lại 1 lần cho cả lô |
| `L.map` nhận sai cột BoQ | `L.setBoq({sheet:"BOQ",hdrRow:9,code:"A",label:"B",qty:"I",price:"J",amt:"K",desc:"L",w:"E",d:"F",h:"G"})`: lưu vào `_LIB!B3`, các lệnh sau dùng luôn |
| "sửa …", "đổi số cánh / kích thước / mã" trong block đã có | `L.edit({code:"D03", edits:[…]})` hoặc `{hdr:3935, …}` |
| "mô tả D04", "mô tả lại BoQ" | `L.desc({get:["D01"]})` để lấy mẫu → viết → `L.desc({items:[{code:"D04",text:"…"}]})` |
| "bỏ note vàng", "bỏ highlight" | `L.unmark({})` (mọi ô Claude đã tô) hoặc `{sheet:"Chiet tinh",from:4304,to:4578}` |
| Bỏ vàng ở block làm **trước khi có _EDITLOG** | `L.unmark({legacy:{sheet:"Chiet tinh",range:"A4972:BL5131",srcSheet:"Template",srcStartRow:183}})` |
| "cập nhật rule", "ghi bài học" | `L.addRules({rows:[{topic:"DOOR",rule:"…",scope:"RIENG"}]})` |
| "nạp AI_INPUT", "tính các hạng mục đã bóc" | Xem §1b: đọc `tblAiInputView` → `L.newItems` |
| Tô vàng ô tự sửa ngoài recipe | `L.mark({cells:[["Chiet tinh","E3937"],["MATRIX FITFUR-LOOSE","AH157"]]})` |

**`edits`**: danh sách `[hàng, cột, giá trị]`.
- **Hàng** là chữ đầu cột A của hàng trong block (không phân biệt hoa thường, so khớp phần đầu), ví dụ `"Cánh cửa"`. `"Đố đứng#2"` là lần xuất hiện thứ 2. `"hdr"` là dòng đầu block, `"+12"` là dòng đầu + 12.
- **Giá trị** là số, chữ, hoặc công thức `"=..."`. Trong công thức dùng được `{r}` (hàng này), `{h}` (dòng đầu), `{boq}` (dòng BoQ), `{B}` (tên sheet BoQ đã đặt nháy).

**`L.newItem` tự làm toàn bộ** (đừng làm lại bằng tay):
1. Tắt tính toán tự động.
2. Chép block (giá trị, công thức, định dạng, chiều cao hàng) xuống **dưới END cuối của Chiet tinh**, không chèn hàng.
3. Dịch công thức cho đúng block mới:
   - Tham chiếu tương đối trong sheet: dời theo block mới.
   - Tham chiếu `$` trỏ vào trong block mẫu (`$B$40`, `$J$9`): dời theo.
   - Tham chiếu sang sheet khác (Material, OverView…): giữ nguyên.
   - Tham chiếu tới dòng BoQ của mẫu: đổi sang `boqRow`.
4. Bảng gộp tấm (`SUMIFS(AN$…)` ở cột G): tách dải để loại dòng đầu block mới, tránh vòng lặp. Mẫu lấy từ Door/Template thì thay dải toàn cột bằng các dải của Chiet tinh.
5. Mẫu chưa link BoQ (block thư viện) thì đặt dòng đầu:
   - A = mã BoQ, hoặc tên hạng mục nếu cột mã chỉ là số thứ tự.
   - B/C/D = các cột W/D/H của BoQ (nếu có).
   - E = SL.
   - BoQ không có cột W/D/H (mô tả dạng "overall size") thì truyền B/D qua `edits`. K/L luôn truyền qua `edits`.
   - Hàng chi tiết trống của Template (vd `"+3"`) điền qua `edits`: A, C, D, E, K, M, N, S.
6. Đặt đơn giá BoQ = `='Chiet tinh'!H<hdr>`.
7. Áp `edits`.
8. Group các hàng con và thu gọn (`collapse:false` để không thu gọn).
9. Tô vàng các ô đã đổi, ghi nhật ký màu cũ vào `_EDITLOG` (`mark:false` để không tô).
10. Bật lại Automatic và tính lại toàn bộ.
11. Trả về `price`, `boq{qty,price,amount}`, `errors` (lỗi có sẵn được ghi `(có sẵn)`), `rows` (các hàng có E>0: `row,A,B,C,D,E,F,G,I`) và `notes`.

## 1b. Nạp kết quả bóc tách từ Cowork (sheet `AI_INPUT`)

Cowork đọc spec/bản vẽ và ghi `<dự án>\02_takeoff\AI_INPUT.xlsx`. Power Query `qAiInput` nạp vào sheet `AI_INPUT`, bảng `tblAiInputView` (chỉ đọc, không sửa tay).

| Cột | Dùng cho |
|---|---|
| `Ma`, `DongBoQ` | `boqRow` = `DongBoQ`; trống thì tìm `Ma` trong `boq.items` của `L.map` |
| `Loai`, `ThongSo` | Chọn mẫu theo §3.3 (số cánh, EI30/EI60, kính, khung bao…) |
| `W`, `D`, `H`, `SL` | So với BoQ; BoQ không có cột W/D/H thì truyền qua `edits` |
| `VatLieu`, `HoanThien` | Mã vật tư `<dày><LOẠI>` (cột K), mã hoàn thiện (L ở dòng đầu) |
| `Nguon`, `DoTinCay`, `Trang` | `DoTinCay = THAP` hoặc `Nguon = SUY_LUAN` → đưa vào câu hỏi gộp (§2.1), kèm `Trang` để người dùng tra |

Quy trình (2 lệnh):
1. Cùng lệnh với `L.map`: đọc thêm `tblAiInputView` (`getItemOrNullObject("AI_INPUT")`). Bảng trống hoặc không có sheet → báo: *"Chưa có dữ liệu: chạy Alt+F8 → PERSONAL.XLSB!RefreshAiInput (sau khi Cowork ghi xong)"*.
2. Lập bảng đối chiếu ngắn (mã, mẫu chọn, kích thước, điểm nghi vấn) → người dùng "ok" → `L.newItems` cả lô.
- Lệch giữa `AI_INPUT` và BoQ (kích thước, SL): **BoQ thắng**, ghi lệch vào báo cáo cuối.
- Không ghi ngược vào `AI_INPUT` (sẽ bị ghi đè khi làm mới).

## 2. Cách làm việc nhanh (bắt buộc)

1. **Hỏi gộp 1 lần**, chỉ khi thiếu thông số làm đổi giá: cửa hay cánh, số cánh, có khung bao không, loại hoàn thiện, giá kính, phạm vi CHUNG/RIENG. Đưa kèm giả định mặc định để người dùng chỉ cần trả lời "ok". Những gì đã có trong `rules` hoặc trong mô tả BoQ thì không hỏi lại.
2. **Việc có recipe thì làm luôn**, không lập plan chờ duyệt. Recipe chỉ ghi vùng trống dưới END cuối, 1 ô đơn giá BoQ và các ô trong `edits`, và mọi ô đều được tô vàng, ghi nhật ký. Lập plan chờ duyệt chỉ khi: xoá hoặc chèn hàng giữa dữ liệu cũ, sửa block thư viện (Door/Template/Kitchen), hoặc sửa hàng loạt ngoài recipe.
3. **Không kiểm lại bằng tay** những gì recipe đã trả về. Không `getImage`, không đọc lại ô, không quét `#` riêng.
4. **Không có recipe** thì gom mọi thứ cần đọc vào 1 lệnh (nhiều `load` trước 1 `sync`) và mọi thứ cần ghi vào 1 lệnh. Bọc việc ghi bằng `calculationMode = manual … automatic`. Tô vàng bằng `L.mark`.
5. **Báo cáo ngắn** dạng bảng: giá/bộ, SL, thành tiền, các ô đã đổi, lỗi mới (bỏ lỗi `(có sẵn)`), và những gì chưa xác định. Không kể lại từng bước.
6. Hội thoại dài thì gọi `context_snip` cho các kết quả map/report cũ.
7. Nếu kết quả có `error` chứa "chế độ sửa ô" thì nhờ người dùng bấm Enter/Esc rồi chạy lại đúng lệnh đó.

## 3. Quy tắc nghiệp vụ

### 3.1 Chung
| # | Quy tắc |
|---|---|
| 1 | Kích thước và mọi giá trị tính được: nhập công thức, không gõ cứng. |
| 2 | Ô Claude sửa: tô `#FFFF00` qua `L.mark` hoặc recipe. "Bỏ note vàng" = `L.unmark`, không xoá vàng sẵn của mẫu (cột F hao hụt, bảng tổng hợp tấm). |
| 3 | Không đổi cấu trúc công thức có sẵn; chỉ đổi ô nhập liệu hoặc tham chiếu cần thiết. |
| 4 | Không truy cập ổ đĩa; cần file khác thì nhờ người dùng mở trong Excel. |
| 5 | Không sửa sheet backup: `*_Snapshot`, `*_BK_*`, `Auto_Snapshot_*`, `Door_BK_truocsua`. |
| 6 | Định dạng số tiền: theo ô liền kề cùng cột. Nếu không có ô liền kề thì dùng `#,##0;(#,##0);-`. |
| 7 | Không quét toàn bộ sheet Door bằng `showOutlineLevels` hoặc đọc định dạng 900 hàng trong 1 script (quá 90 giây sẽ timeout). |
| 8 | Xoá hàng trống trong block: **luôn chừa 1 hàng trống đệm ở cuối mỗi vùng SUM/SUBTOTAL** (vật tư mộc, vật tư phụ, gia công ngoài…). Chỉ xoá hàng ở giữa vùng. Việc xoá hàng phải lập plan chờ duyệt, ghi rõ hàng sẽ xoá. |

### 3.2 BoQ ↔ Chiet tinh
- BoQ khác nhau theo dự án; `L.map` tự nhận theo hàng tiêu đề, kết quả nằm trong `boq`. Hai dạng đã gặp:

| Dạng | Mã | Tên/mô tả CĐT | Kích thước | SL | Đơn giá | Thành tiền | Mô tả costing |
|---|---|---|---|---|---|---|---|
| `MATRIX…` (dự án nội địa) | A (D01…) | C ("overall size WxH mm") | tách từ C | AG | AH | AI | AJ |
| `BOQ` (xuất khẩu, tiêu đề hàng 8–9) | A (STT) | B Item, D Description | E/F/G = W/D/H | I | J (Unit Price) | K | L (Remark); M = ghi chú yêu cầu của người dùng |

- Đơn giá BoQ = `='Chiet tinh'!H<hdr>`. Thành tiền có sẵn công thức; không sửa tiêu đề hay đơn vị tiền nếu người dùng không yêu cầu.
- Sheet `Matrix` (viết thường) và `Spec` **không** phải BoQ của chủ đầu tư.
- `Chiet tinh` gồm nhiều block nối tiếp, mỗi block từ dòng đầu đến `END`. Các cột ở dòng đầu:
  - A: `='BoQ'!A<dòng>`
  - B, D: rộng, cao, đọc từ C của BoQ bằng `LET(t,…,SEARCH("overall size ",t)…)`
  - C: sâu tường (mặc định 110, theo ghi chú nếu có)
  - E: `='BoQ'!AG<dòng>`
  - H: đơn giá; I: H×E
  - K: loại cửa (`EI30`/`EI60`/`Cửa thường`); L: mã hoàn thiện (WD-01, LA-01…); M: dày khung bao (30/35)
- Cột chi tiết:
  - A: tên; B: dày; C: rộng; D: dài; E: SL/cánh; F: hao hụt; G: KL; H: đơn giá; I: thành tiền
  - J: ĐVT; K: mã vật tư (tra `Material` A→F); L/N/O: hoàn thiện; Q: tên vật liệu
  - AL: DR/DRF; AN…AU: cột phụ hao hụt
- **Mô tả costing** (cột `boq.desc`: AJ hoặc L; khách đọc; viết tiếng Việt):
  - **Không ghi đè mô tả của chủ đầu tư** (cột C hoặc D).
  - Dòng đầu luôn là quy cách tổng: cửa ghi "Cửa đi N cánh <loại> kích thước : W×D×H mm."; hạng mục khác ghi "W…xD…xH… mm :". Sau đó mỗi cấu phần một dòng.
  - Phụ kiện ghi **mã hãng thật**, lấy từ sheet `PK` theo mã dự toán (ví dụ "Hafele 283.33.930"). Không ghi mã nội bộ như `SOP940`.
  - Vật tư gia công ngoài (sắt, inox) ghi là "gia công ngoài".
  - Hạng mục không bao gồm ghi ở mục "Chưa bao gồm".
  - Không ghi quy cách chi tiết nội bộ, ví dụ không viết "veneer tạp mặt đáy".
  - Định dạng lấy theo ô mô tả có sẵn; mặc định không tô vàng.
- Mô tả bản vẽ mơ hồ: bỏ kích thước và độ dày, chỉ mô tả vật liệu.

### 3.3 Cửa (block Door và Chiet tinh)
- Block: dòng đầu là tên (A), K là loại, M là dày khung bao, E là SL (E=0 thì mọi số tấm = 0). Kết thúc bằng `END`.
- Mã vật tư có dạng `<dày><LOẠI>` (`28Honeycomb`, `50ASH`, `6MDF`, `3MDF`, `TRAMGHEPD1`). Mã không có trong Material sẽ báo #N/A.
- Cột N của mọi hàng lớp da/hoàn thiện = `=+L<dòng đầu>`.
- Chọn mẫu:

| Trường hợp | Mẫu |
|---|---|
| Giống một cửa đã tính | `src:{code:"Dxx"}` (ưu tiên, giữ trang trí và hiệu chỉnh) |
| Cửa có kính | Door "Cửa phẳng 1 cánh - Kính". **Không** dùng "Cửa gỗ tự nhiên 1 cánh (Có kính)" (gỗ Ash, không có bảng hao hụt tấm) |
| Chống cháy ("fire rated") | Door "Cửa phẳng 1 cánh - new"; K = EI30/EI60; trang trí (veneer, mây, panel) thêm từ hàng trang trí của mẫu |
| Panel đố gỗ | Door "Cửa panel đố gỗ": đố đứng bản lề 145, đố đứng khoá 140, đố ngang ×3 rộng 140, ván nền 2 mặt SL 4, lõi SL 2 |
| Cửa 2 cánh | Theo bảng ở §3.3a bên dưới |

- §3.3a Cửa N cánh bằng nhau: lấy block 1 cánh cùng loại và đặt `edits:[["Cánh cửa","E",N]]`.
  - Rộng cánh = (cửa − khung bao − khe 6)/N.
  - Rebate ngang và nhân công "nhiều cánh" tự tính theo SUMIF hàng "Cánh cửa".
  - Giá các cửa dùng chung loại tấm sẽ nhích nhẹ, vì bảng gộp tấm dùng chung số tấm.
- Kích thước cánh (mẫu Kính): rộng = B(cửa) − (B khung đứng + B da + B FR)×2 − 6; cao = D(cửa) − (…) − 10. Nếu đề bài ghi 850×2100: hỏi gộp đó là cửa hay cánh.
- Honeycomb: dày = dày cánh − 2×dày tấm mặt. Mã có sẵn: 10, 15, 20, 22, 25, 27, 28, 30, 32, 34, 35, 40, 50.
- Đố tràm + hoàn thiện: mỗi loại đố có 2 hàng.
  - Lõi `TRAMGHEPD1`: dày = dày cánh − 2×dày hoàn thiện.
  - `3MDF` phủ 2 mặt: E = 2×SL đố; rộng và dài = lõi + 20.
- Ô kính:
  - Rộng = rộng cánh − Σ đố đứng.
  - Cao = (cao cánh − đố trên − đố giữa×(n−1) − đố đáy)/n, với n nhập ở E hàng "Ô kính".
  - Mã kính tra theo `KINH` + Material.
- Phần kim loại (nẹp, nan, bánh xe, ray): SL = 0 khi người dùng bảo bỏ qua. Khung sắt tính ở mục **gia công ngoài** (sắt = MTGC, inox = SSGC).
- TK+VIT chỉ = 0 khi không có khung bao.

### 3.4 Hao hụt tấm (bảng tổng hợp ngay dưới mỗi block)
- Các cột:
  - A: `=+Q<hàng>` (khoá tên vật liệu); F: "Tấm"
  - G: `LET(BaseSum,SUM(SUMIFS(AN…) theo các dải),ROUNDUP(BaseSum*1.05,0))`; MDF ×1,05, tràm ×1,03
  - H: `SUM(SUMIFS(AU…))`; I: G/H
- F của chi tiết: `=+VLOOKUP(Q,A<tổng hợp>:I<…>,9,0)`.
- Không có dòng tổng hợp thì F rơi về công thức mặc định: PN & N∈{SL,UV,0} → 1,2; có K → `Material!I`.
- Dải SUMIFS **không được chứa dòng đầu block nào**, vì AN dòng đầu là tiền sẽ gây vòng lặp. `L.newItem` tự tách dải.
- Mỗi block mới thêm 1 dải vào mọi công thức gộp. Nếu `notes` báo "quá dài" (gần 8.000 ký tự): báo người dùng, không tự gộp.
- Khoá A trỏ vào hàng E=0 (Q trống) gây vòng lặp hoặc #N/A. Phải trỏ vào hàng có E>0.
- Thêm loại tấm mới:
  1. Chèn dòng ngay dưới dòng tổng hợp liền kề và chép dòng đó.
  2. Sửa khoá A.
  3. Đổi F các chi tiết sang VLOOKUP và mở rộng vùng tra.
- Cột phụ:
  - AN = `AQ/MAX(INT(1220/(AO+5))*INT(2440/(AP+5)),INT(2440/(AO+5))*INT(1220/(AP+5)))`
  - AO = D; AP = C; AQ = E*AG; AT = Q; AU = AO*AP*AQ/(1220*2440)

### 3.5 Keo, chỉ cạnh
- `KEO.MDF`, cửa thường: (Σ m² lớp da/hoàn thiện cánh + da khung bao + 6MDF rebate)/2,98. Mẫu **cộng cứng số hàng**, nên block mới có vị trí hàng khác phải sửa lại danh sách hàng.
- `KEO.MGO`: chỉ tính cho cửa chống cháy.
- `KEOEPCC`: mẫu nhân `*0`, không tính.
- `KEOCANH` = SUM(BJ). BJ theo mã chỉ cạnh R ở hàng "Cánh cửa"; R lấy L, O từ hàng lớp da đầu tiên. Nếu hàng đó là lõi tràm thì trỏ về hàng hoàn thiện đầu tiên.
- Chỉ cạnh theo dày cánh: dày ≥42 dùng `55<LA..>`.

### 3.6 Melamine ↔ Laminate (file BoQ tổng)
- `OverView!F27` (`Melamine`/`Laminate`) điều khiển `Chiet tinh!K = IF(F27="Melamine",BN,BM)`:
  - BM: mã ván Laminate gốc.
  - BN: mã Melamine tự sinh `<dày><ván><số mặt>M-ME-<màu>`.
- Block Melamine mới trong Material: chép block mẫu (vd HHF-56), chèn trước "VÁN KHÁC", đổi mã và mô tả.
- Tốc độ:
  - Không để IF hoặc LET dài trong K.
  - Giới hạn vùng tham chiếu, ví dụ `$K$1:$K$30000`.
  - Thay hàng loạt: hướng dẫn người dùng Ctrl+H với Calculation Manual.
- Đổi F27 → `calculate(full)` → đọc D10/D24 → trả F27 về giá trị cũ.

### 3.6b Lỗi mẫu đã biết (sửa trong block khi gặp, báo người dùng sửa tận gốc ở Template)
- Template, cột B hàng chi tiết: `--IF(LEFT(K,1)<"3",LEFT(K,2),LEFT(K,1))` tách sai độ dày với mã từ 30mm trở lên (`35MDF` ra 3). Đúng là `=IF(K{r}="",0,IFERROR(--IF(ISNUMBER(--MID(K{r},2,1)),LEFT(K{r},2),LEFT(K{r},1)),0))`.

### 3.7 Lỗi có sẵn — bỏ qua
- Chiet tinh hoặc Door:
  - Cột B `=kd(…)` báo `#NAME?` (hàm VBA không chạy trong Office.js).
  - AB hàng "Cánh cửa" báo #N/A.
- OverView U12, U14 báo #DIV/0!.
- Material K376, K895 báo #VALUE!; K755, K756 báo #DIV/0!.

## 4. Nhật ký tô vàng `_EDITLOG`
- Sheet ẩn hẳn, do thư viện tạo.
- Mỗi ô Claude tô được ghi một dòng `Sheet | Ô | Màu cũ | Thời gian | Ghi chú`.
- `L.unmark` trả lại đúng màu cũ rồi xoá dòng nhật ký, nên vàng sẵn của mẫu không bị mất.
- Không sửa tay sheet này.

## 5. Quy tắc 3 tầng và `_RULES`
1. Thứ tự ưu tiên: `_RULES` (`tblRules`, Active/Pending) > `_RULES_MASTER` (`tblRulesMasterView`, Active) > skill này. Có mâu thuẫn thì báo người dùng.
2. Quy tắc mới:
   - Ghi bằng `L.addRules` (tự bỏ qua dòng trùng, trạng thái Pending).
   - Hỏi `Phạm vi` (`CHUNG` = mọi dự án, `RIENG` = file này) trong cùng câu hỏi gộp.
   - Viết 1 câu, nêu cụ thể ô, cột, công thức.
   - Mã chủ đề: `CHUNG, BOQ, DOOR, HAOHUT, KEO, MELAMINE, MOITRUONG, FILE`.
   - Không điền `ID` và `Dự án gốc`: macro tự điền khi Save.
   - Dòng đã có ID thì sửa ở `RULES_MASTER.xlsx`.
   - Cowork (Claude ngoài Excel) ghi rule CHUNG thẳng vào `RULES_MASTER.xlsx`, các file nhận qua `_RULES_MASTER` khi mở/Refresh All. Rule `RIENG` chỉ ghi trong add-in.
3. Chưa có `_RULES`: tạo sheet `_RULES` (cuối workbook, tab xám).
   - A1 = "SỔ QUY TẮC – Claude tự đọc/ghi."
   - `tblRules` tại G3 với cột `Ngày | Chủ đề | Quy tắc | Nguồn | Trạng thái`.
   - Báo người dùng chạy `PERSONAL.XLSB!SetupProjectRules` để thêm `ID · Phạm vi · Dự án gốc` và `_RULES_MASTER`.

## 6. Chế độ dự phòng (khi chưa có `_LIB` hoặc bị chặn `new Function`)
Dán **nguyên văn** mã dưới đây vào `execute_office_js` theo mẫu:

```
const L=(()=>{ <MÃ CPSX_LIB> })();
return await L.newItem({...});
```

Muốn cài `_LIB` từ trong Excel (khi người dùng không chạy được macro):
- Tạo sheet `_LIB`, đặt `numberFormat` cột A = `"@"`.
- Ghi toàn bộ mã vào `A1` (một ô chứa được 32.767 ký tự), ghi `B1 = "{{VER}}"`.
- Đặt `visibility = "VeryHidden"`.

```js
{{LIB}}
```
