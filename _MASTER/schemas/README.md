# schemas — Chuẩn dữ liệu trao đổi Cowork → Excel (v2, 2026-10-24)

Mẫu: `AI_INPUT_TEMPLATE.xlsx` (v1 ở `_MASTER\backup\AI_INPUT_TEMPLATE_v1_2026-10-24.xlsx`). Thiết kế: `_MASTER\modules\BUOC-TIEP-AI-INPUT-TAKEOFF.md` §3.

Vai: model bóc là bên DUY NHẤT quyết định (mẫu, module, giá trị) · script kiểm từng dòng · user duyệt ngoại lệ · `L.fromAiInput` ghi máy móc, không diễn giải lại.

## AI_INPUT.xlsx — sheet `AI_INPUT`, bảng `tblAiInput` (1 dòng = 1 BLOCK)

| Cột | Ghi chú |
|---|---|
| BlockId | Khóa, vd `D01-1`, `D01-2` (1 mã BoQ có thể nhiều block: cửa + ô fix, tách theo loại tường R0166, dãy bếp chia block) |
| ThaoTac | `MOI` (mặc định: chép nguồn xuống cuối Chiet tinh) · `SUA` (sửa block Chiet tinh có sẵn = MauTen; Off tính theo block đó) |
| Ma | Mã hạng mục BoQ |
| DongBoQ | Dòng BoQ, nhiều dòng cách bằng `,` (dòng đầu = dòng link A và đơn giá) |
| QuanHe | `TACH` (mặc định: nhiều block chia SL của dòng BoQ) · `THANH_PHAN` (đơn giá BoQ = Σ H nhiều block, mỗi block mang đủ KL) |
| Ten, Loai | Từ BoQ (đọc tổng quan) |
| Mau | Mã danh mục RULES (vd `Door#3`) — chỉ để đọc; số thứ tự có thể khác giữa các file |
| MauTen | **Nguồn**: `<sheet>!<tên block>` đúng như file CPSX của dự án (vd `Door!Cửa phẳng 1 cánh - Kính`), hoặc block có sẵn khi làm tiếp file: `Chiet tinh!D02` / `Chiet tinh!#101` (mã trùng → dùng hàng). Chép từ Chiet tinh phải khai lại E header (không kế thừa E nguồn) |
| W / D / H | mm (rộng / sâu / cao, = B / C / D header) = hình học của **1 đơn vị BoQ** (m2 → 1000×1000, m → dài 1000, Nr → kích thước cái). Thành phần (số thanh, tấm) khai ở dòng module với số lượng thật — không nén vào header (user sửa F028 SOFITEL 2026-10-25) |
| SL | Theo đơn vị BoQ khi GiaTheo trống |
| GiaTheo | Trống = đơn giá BoQ link H, E = KL BoQ (cách `L.newItem` làm). `md` / `m2` = link N, E quy đổi (chưa hỗ trợ ghi tự động) |
| VatLieu, HoanThien, ThongSo | Tóm tắt để đọc; giá trị thật nằm ở `tblAiModule` |
| Nguon | CAD / SCAN / SPEC / SUY_LUAN / CPSX |
| DoTinCay | CAO / TB / THAP (THAP → danh sách user xem) |
| Trang | Trang bản vẽ/spec để truy ngược |
| GhiChu | |
| RunId | Mã lần ghi; ghi vào ghi chú ô A block, đổi RunId → liệt kê block cần dựng lại |

## sheet `AI_MODULE`, bảng `tblAiModule` (1 dòng = 1 Ô NHẬP)

| Cột | Ghi chú |
|---|---|
| BlockId | Khóa ngoại về tblAiInput |
| Module | Nhãn cột A **của mẫu** tại dòng đó (`hdr` = dòng header) — để đọc và để bộ kiểm so |
| Lan | Lần xuất hiện thứ k của nhãn khớp CHÍNH XÁC trong mẫu |
| Off | **Khóa ghi**: dòng tính từ header của nguồn (0 = header). Không ghi theo nhãn: nhãn trùng phần đầu ("Len chân" ↔ header "Len chân, thanh lam gỗ") làm ghi nhầm hàng |
| Cot | B / C / D / E / K / L / M / N … |
| GiaTri | Số, chữ, hoặc công thức: `{h}` = header, `{h+n}` = dòng n của block, `{boq}` = dòng BoQ, `{B}` = tên sheet BoQ |
| Loai | NHAP (đổi hằng số) / LINK (hằng số → công thức) |
| Nguon, Trang, DoTinCay, GhiChu | Như trên, theo từng ô |

Chỉ ghi vào **ô nhập của mẫu** (ô hằng số, ngoài phần nhân công R0150). Đè công thức, sửa công thức, thêm/bớt dòng không biểu diễn được ở đây → bộ kiểm báo NGOAI (cần duyệt, làm tay).

## Gộp block ≠ gộp mô tả (user 2026-10-25) — chi tiết ở `_MASTER\mo_ta\`

Nhiều dòng BoQ cùng cấu tạo → 1 block (`DongBoQ` liệt kê các dòng, E = Σ AG; gộp vào block có sẵn bằng `SUA`). Đó chỉ là gộp **tính giá**.
- **Chỉ gộp khi chi phí như nhau**: khác biệt nằm ở phần không tính (giấy, gương, kim loại ngoài phạm vi) hoặc user chấp nhận coi như nhau. Hoàn thiện khác giá (veneer WD-01 ↔ WD-04, gỗ ↔ sơn) → tách block.
- **Mô tả costing (cột AJ "GHI CHÚ COSTING") viết riêng từng dòng** theo BoQ của dòng: mã hoàn thiện của dòng, "Không bao gồm" theo vị trí (lấy từ QUAN_SAT: Phần ngoài phạm vi), ghi rõ khi tạm tính (RFI). Không chép mô tả của block dùng chung.
- Hiện ghi bằng `L.desc({items:[{row,text}], mark:true})` sau khi ghi block (vd `_MASTER\takeoff\mo_ta_SOFITEL_2a.py`). Việc cần làm: đưa vào AI_INPUT thành bảng `tblAiDesc` (DongBoQ, Ma, MoTa, Nguon) để `L.fromAiInput` ghi luôn và `roundtrip_check` kiểm.

## Công cụ (`_MASTER\takeoff\`, chạy `py -I`)

| Script | Việc |
|---|---|
| `module_catalog.py <CPSX> <json> [RULES]` | Danh mục block mẫu + ô nhập + nhãn trùng đầu + lỗi thư viện (thiếu END) của đúng file dự án |
| `check_ai_input.py <AI_INPUT> <CPSX> <KIEM.xlsx> [RULES]` | Kiểm MAU, OFF, O_NHAP, MA_VT, DVT, SL, NGUON, CUA, NGOAI → TONG_QUAN + NGOAI_LE |
| `extract_ai_input.py <CPSX> <AI_INPUT.xlsx> [RULES] [RunId]` | Trích ngược Chiet tinh → AI_INPUT v2 (+ NGOAI_EDITS): dữ liệu thử và kiểm khứ hồi |
| `cpsx_blocks.py` | Lõi dùng chung: đọc block, ghép block với mẫu, phân loại khác biệt |
