# Tổng hợp để làm bước tiếp: nối cơ chế module ↔ AI_INPUT ↔ takeoff (2026-10-24)

## 1. Đã có
- **Tri thức module** ở `_MASTER\modules\` (đọc `_index.md`, `NGUYEN-TAC.md`): ghép hộp, modun cánh mở/góc, ngăn kéo, len, phụ kiện, vách phẳng / ghép hộp + hộp âm tivi, kim loại gia công, bọc vải, cửa (cánh, khung bao, phụ trợ, biến thể), bàn, sofa, ghế solid, giường, gối, Kitchen (Spec, Matrix, bếp trên/dưới/phụ).
- **Rule** `RULES_MASTER.xlsx`: MAU (R0075–R0133: danh mục 22 + 11 + 24 mẫu), TRACUU (R0135–R0149), FILE-00 (R0151), CHUNG/BOQ/DOOR.
- **Mẫu `260903_DA_CPSX_RE00.xlsm` đã sửa (Kitchen, 2026-10-24):** block SINK lấy thùng/hậu theo Spec cột G; ngăn kéo (K12, 13, 14, 19) thùng NK theo cột F, mặt ngoài + chặn NK theo cột H; N hậu K1–K5, K16 và đợt ngang K18, K19 đúng cột. Hậu góc chữ L giữ 17MDF. Bản sao trước từng bước ở `_MASTER\backup\*_2026-10-24.xlsm`.

## 2. Nguyên tắc không đổi
Sản phẩm cuối = Excel CPSX có công thức sống, link tối đa, ít value. Chỉ nhập các **ô nhập** của mẫu (R0075); không sửa block mẫu; tên cột A là khóa SUMIF; nhân công giữ như mẫu (R0150); nghiệp vụ không rõ thì hỏi; engine/script chỉ để **đối chiếu**.

## 3. Thiết kế (user chốt 2026-10-24)
**Nguyên tắc vai:** 1 bên quyết định (model scan) · 1 bên duyệt tổng quan (user, KHÔNG duyệt từng dòng) · 1 máy ghi (script). Bước ghi vào Excel không có model đọc hiểu lại → tránh "2–3 người cùng làm, lệch thông tin".

**3.1 AI_INPUT — khóa là BLOCK, không phải hạng mục BoQ**
- `tblAiInput` (1 dòng/block, để đọc tổng quan): 15 cột cũ + `BlockId` (vd `D01-1`, `D01-2`) + `Mau` (vd `Door#3`; không khớp → `KHONG_CO_MAU`, dừng hỏi) + `RunId`. `Ma`/`DongBoQ` là khóa ngoại về BoQ. 1 mã BoQ được nhiều block (cửa + ô fix Door#10/#11; tách block theo loại tường R0166; dãy bếp chia block).
- `tblAiModule` (dạng dọc, 1 dòng = 1 ô nhập, khớp 1:1 `edits` của `L.newItem`): `BlockId, Module (tên cột A đúng như mẫu, hoặc "hdr"), Lan (#k), Cot (B/C/D/E/K/L/N…), GiaTri, Nguon, Trang, DoTinCay`. Công tắc header (K loại cửa, R khung bao gỗ đặc, K loại module…) = dòng `Module="hdr"`. Mã vật tư ghi ở đây đã áp R0074.
- Tham số toàn dự án (OverView, vd C52 `KETCAU_CUA`) không đưa vào AI_INPUT.
- Power Query `qAiModule` song song `qAiInput`; `SyncAiInput`/`RefreshAiInput` làm mới cả hai.

**3.2 Chọn mẫu theo HÌNH DẠNG, loại EI là ô nhập K header**
- Cửa phẳng #1/#3/#4/#6/#7 dùng chung công tắc K (`Cửa thường` | EI30…EI120) → K tự đổi lớp FR, lõi, dày cánh, mã kính (CUA-CANH §2). Chọn mẫu: 1 cánh → #1; 1 cánh có kính → #3; 2 cánh → #4; 2 cánh có kính → #5 (R0134); ô chớp → #6; panel đố → #7; gỗ tự nhiên nguyên khối → #8 (1 cánh) / #9 (2 cánh); ô fix → #10 (tấm gỗ) / #11 (kính khung ASH). Tổ hợp không có mẫu → `KHONG_CO_MAU`.
- Cần sửa câu chữ R0099 ("Door#1 dùng khi chống cháy") và R0103 ("Door#5 … chống cháy"): EI không phải tiêu chí chọn mẫu.
- Tủ → Template#1 trên, #2 dưới, #3 cabinet, #4 tủ áo, #5 locker, #6 tủ cao, #7 vanity, #8 tủ gương hộp, #9 gương. Vách/tấm → #10, #11 (ghép hộp, tivi), #12, #13. Bàn #14, sofa/bench #15, ghế solid #16, giường #17, gối #18. Chưa có mẫu → #19 TRỐNG (R0028). **Kitchen#1–#24 chỉ khi bóc ma trận ở Matrix** (R0076).

**3.3 Lệch BV ↔ BoQ: tùy dự án, hỏi CĐT** — script không tự chọn bên thắng; ghi Issue/RFI, tạm tính theo giả định ghi trong RFI (R0163). Skill §1b "BoQ thắng" cần sửa theo.

**3.4 Kiểm tự động (máy kiểm từng dòng thay user)** — Python `py -I`, pass/fail theo BlockId:
- Trước khi ghi: mẫu hợp lệ, không thuộc nhóm cấm (Door#2, Door#8 cho cửa kính thường, Template#22) · Module có trong danh mục module của mẫu (quét tự động Template/Door/Kitchen từ tên block → END) · nhãn không trùng phần đầu trong block (vd "Cánh mở" ↔ "Cánh mở - Ngăn kéo") · số lần `#k` không vượt số dòng sẵn có (vượt = phải chèn hàng R0039 → plan chờ duyệt) · mã vật tư có trong Material và có giá (R0074) · "Giá theo" của mẫu khớp ĐVT BoQ (cái ≈ bộ) · Σ SL các block của 1 mã = SL BoQ (kiểm mới) · Template#1/#2: Σ B module = B header · cửa có dày tường theo vị trí (R0166) · mọi dòng có Nguon + Trang + DoTinCay.
- Sau khi ghi: **kiểm khứ hồi** — đọc lại các ô nhập trong Chiet tinh, so từng ô với `tblAiModule`, khác = lỗi · R0048 (BoQ ↔ Material, ngưỡng 1.000.000).
- `RunId` ghi vào mỗi block; làm mới AI_INPUT có `RunId` mới → liệt kê block cần dựng lại.

**3.5 User duyệt tổng quan — chỉ xem ngoại lệ đã lọc** (báo cáo 1 trang, không xem từng dòng):
1. Bảng tổng theo nhóm (cửa/tủ/vách/…): số block, Σ thành tiền, **đơn giá/ĐVT so khoảng tham chiếu** (từ NORTON + dự án đã duyệt, theo Mau). Ngoài khoảng → nêu block.
2. Danh sách phải xem: DoTinCay = THAP hoặc Nguon = SUY_LUAN; Issue/RFI; `KHONG_CO_MAU`; check fail.
3. Hạng mục giá trị lớn (nhóm chiếm ~80% tổng tiền): **đọc chéo độc lập** — model thứ hai chỉ đọc lại kích thước chính (W/D/H, SL, dày tường) từ bản vẽ, KHÔNG quyết định; lệch với bản scan → đưa vào danh sách 2. (Người kiểm, không phải người làm thứ hai.)
4. Nhất quán: cùng mã/cùng type thì cùng mẫu, cùng cấu tạo; khác → nêu.
5. Mỗi Mau 1 block mẫu để user mở xem nhanh (spot check).
Block không thuộc danh sách nào → đi thẳng, không cần user xem.

**3.6 Luồng:** scan (`01_scan\PROJECT_DATA`) → phân loại → `02_takeoff\AI_INPUT.xlsx` → kiểm trước khi ghi (3.4) + đọc chéo hạng mục lớn → báo cáo tổng quan (3.5) → user duyệt ngoại lệ / trả lời RFI → `L.fromAiInput` (code: chuyển `tblAiModule` → `edits` → `L.newItems`; lỗi thì dừng, không đoán; Nguon/Trang vào comment ô tô vàng R0051) → kiểm khứ hồi + R0048 → chiết tính.

## 4. Checklist "takeoff cần khai" theo nhóm (lấy từ từng file module)
Cửa: rộng, cao, **sâu khung bao theo vị trí**, SL, loại EI, ván nền, mã hoàn thiện, số cánh, kiểu cánh, kính/lam. Tủ: rộng từng module, sâu, cao, số cánh/ngăn kéo, hông lộ (0/1/2), giáp tường/trần, ván thùng, mã hoàn thiện, chân nhựa/gỗ, phụ kiện (hãng). Vách: m², loại (hoàn thiện gỗ/sơn hay bọc vải), khung sắt/gỗ, ô trống tivi (rộng × cao), hoàn thiện. Bàn/ghế/giường/sofa: kích thước, ghép hộp hay không, loại gỗ, kiểu chân, bọc nệm (foam/vải), công tắc Khó/Cong/loại sofa.

## 5. Việc tồn đọng (từ mục "Chưa xác nhận" của từng file — cải thiện dần khi bóc thực tế)
- Kitchen: Matrix link giá module (nhiều bề rộng dùng chung block hay copy), Kitchen#4/#6/#7 chưa đọc từng hàng, #18/#19 chưa đối chiếu Template#6/#4.
- Cửa: Door#8/#9 phần gia công + phụ kiện; Door#5 chi tiết ô kính.
- Chung: cột G dòng module (tạm bỏ), `AG1083` (chiều dài vải sơ đồ), dòng bọc nệm cho giường, khe đèn vách tivi, "Độ khó" kim loại gia công.
- Lỗi mẫu đã biết ngoài Kitchen: R0058 (cột B với mã ≥30mm), R0037 (`#REF!` ở Chiet tinh AH324…), các `#REF!` tiềm ẩn kế thừa từ Template.

## 6. Thứ tự đề xuất
1) ✅ Schema 3.1 (2026-10-24). 2) ✅ Script (`_MASTER\takeoff\`: module_catalog, cpsx_blocks, extract_ai_input, check_ai_input) + `L.fromAiInput` (lib v2.3) + VBA `qAiModule`; chạy thử trên **SOFITEL REV01** (mục 7). Chưa: cài lib v2.3 vào `_LIB` file dự án, import modAiInput v2 vào PERSONAL, thử `L.fromAiInput` trong Excel. 3) Dựng khoảng đơn giá tham chiếu theo Mau từ NORTON. 4) Thử 3 ca (cửa EI, dãy tủ bếp thường Template#1/#2, vách tivi) — ưu tiên ca đã có costing tay làm đáp án. 5) Nhân ra nhóm còn lại; mỗi ca thật bổ sung vào file module. Song song: sửa câu chữ R0099/R0103, skill §1b (bỏ "BoQ thắng"), `schemas/README.md`.

## 7. Thử trên SOFITEL `03_costing\261023_SOFITEL_CPSX_REV01.xlsm` (2026-10-24)
Trích ngược 109 block Chiet tinh → AI_INPUT v2 → chạy bộ kiểm (8 giây).
- **Thư viện mẫu của file dự án ≠ RE00**: không có Door#10/#11 (FX), Kitchen#8; có thêm "Cửa lùa gỗ tràm 1 cánh - kính sọc" (dùng cho D02, D07–D18). → chọn mẫu theo `MauTen` (tên block), danh mục quét từ chính file dự án.
- **Kitchen!A407 thiếu END** (RE00 có): chép Kitchen#7/#8 từ file này sẽ ra block gộp 2 mẫu.
- **Nhãn trùng phần đầu** 21 chỗ (vd Template#12 "Len chân" khớp header "Len chân, thanh lam gỗ") → ghi theo `Off`, không theo nhãn.
- **Thực tế bóc REV01 vượt xa "chỉ ô nhập"**: 1.829 ô nhập (ghi được bằng AI_INPUT) so với 899 ô đè số lên công thức, 1.068 công thức bị sửa, 843 dòng thêm, 451 dòng bớt. (ĐÃ SỬA 2026-10-24) Phát hiện cũ "Kitchen 'Tủ Cao' dùng cho tủ áo/minibar, 370 ô K và 244 ô N bị gõ thẳng" là SAI: script đoán mẫu theo chuỗi nhãn và để Kitchen thắng Template!Cabinet với biên rất nhỏ. Trích lại không có Kitchen (Kitchen chỉ dùng khi user cho phép Matrix, R0076): số ô đè số giảm 511 → 221, các block khớp Template!Cabinet. Đừng dựa vào số liệu ô đè số cũ.
- **Đơn vị**: `L.newItem` link đơn giá BoQ vào H, E = KL BoQ → ô M header chỉ là nhãn. Chỉ F009, F014 (md) link N với E quy đổi.
- **Quan hệ block ↔ dòng BoQ**: có block TÁCH (chia SL) và block THÀNH PHẦN (đơn giá BoQ = Σ H nhiều block, vd F039 = H1082 + H1373; block 1373 dùng chung F038 + F039 qua E = E1082 + E1179) → thêm cột `QuanHe`.
- Lỗi/nghi vấn bộ kiểm bắt được: F011.1 E = AG18 + AG34 (chép từ F011; nên là AG19, và F026 181,1 m bị tính thêm); D01, D03 C header = 110 mặc định của mẫu (chưa kiểm dày tường R0166); mã giá 0: GC.02 (F041, F074), TAYNAM (F047, F048); mã vải không có trong Material: FB.51-500A/E (L006–L009, L033), FB.02 (L036); 31 block sửa phần nhân công (trái R0150).

## 8. Làm tiếp file (2026-10-24, user: "làm tiếp file Excel, không tính mới từ đầu")
- Block cũ giữ nguyên (công của user, không biểu diễn lại). Câu hỏi (a)/(b) ở mục 7 chỉ còn cho dự án mới chép thẳng từ thư viện.
- `ThaoTac`: `MOI` chép từ nguồn (thư viện hoặc `Chiet tinh!<mã>` / `Chiet tinh!#<hàng>`, mang theo mọi chỗ sửa của block nguồn) · `SUA` sửa block có sẵn (Off theo block đó).
- Chép từ Chiet tinh **bắt buộc khai E header** (chặn lỗi kiểu F011.1: E nguồn cộng nhiều dòng BoQ).
- Thử ca D15 (tách khỏi D02, WD-04): bộ kiểm SL đạt (D15 = 2; D02 = 255 + 6); mock Office.js trên dữ liệu REV01: block mới 8005–8095 link dòng 169, E = AG169, L = WD-04, BoQ AH169 → H8005, không còn trỏ dòng 156; ca thiếu E bị từ chối. Excel thật (bản sao `_TEST\SOFITEL_D15`): `_LIB` v2.3, qAiInput 2 dòng, qAiModule 3 dòng. Còn: chạy `L.fromAiInput` bằng Claude trong Excel trên bản sao.
- Đã cài: modAiInput v2 vào PERSONAL.XLSB (backup `_MASTER\backup\PERSONAL_2026-10-24.XLSB`); `cpsx-master.zip` đóng lại (cần upload lên claude.ai).
- **Chạy thật (user, Claude trong Excel, 2026-10-24 14:01)** trên `_TEST\SOFITEL_D15`: kiểm khứ hồi `roundtrip_check.py` đạt 3/3 ô, 0 lỗi; D15 block 8005–8095 giá 5.784.793/bộ (WD-04), BoQ AH169 → H8005; D02 E = AG156 + AG170 (261), giá 6.643.760 → 6.640.660 (bể gộp tấm, R0179). Comment RunId không ghi được trên Excel này → lib 2.4 ghi vào sheet ẩn `_AIRUN`.
