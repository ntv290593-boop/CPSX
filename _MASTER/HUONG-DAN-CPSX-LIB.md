# CPSX_LIB — Claude trong Excel chạy nhanh

Claude không còn tự viết script mới cho mỗi việc. Workbook chứa sẵn thư viện (sheet ẩn `_LIB`), nên mỗi việc lặp lại chỉ cần **1 lệnh**.

## Cài đặt (1 lần/máy)

| # | Việc | Thao tác |
|---|---|---|
| 1 | Cập nhật skill | claude.ai → Settings → Capabilities → Skills: xoá `cpsx-master` cũ, upload `D:\COSTING EXCEL\_MASTER\cpsx-master.zip` |
| 2 | Tắt skill trùng | Trong Skills, **tắt hoặc xoá**: `cpsx-project-rules`, `costing-door-rules`, `costing-door-casestudies`, `costing-door-new-leaf-block`, `costing-chiet-tinh-hao-hut`, `cpsx-ve`, `boq-melamine-laminate-toggle`. Nội dung của chúng đã gộp vào `cpsx-master` |
| 3 | Import macro | Alt+F11 → `VBAProject (PERSONAL.XLSB)` → File → Import: `modCpsxLib.bas` |
| 4 | Cập nhật sự kiện | Trong PERSONAL.XLSB: xoá module `clsAppEvents` cũ, rồi Import lại `clsAppEvents.cls` (bản mới có thêm sự kiện mở file) → Save → tắt Excel rồi mở lại |
| 5 | Instructions (nếu có) | Trong Claude Excel, xoá đoạn "Quy trình thêm hạng mục vào Chiet tinh…" ở custom instructions (đã nằm trong skill) |

Sau bước 4: mỗi khi mở một file costing (có sheet `Chiet tinh`), macro tự cài hoặc cập nhật `_LIB`. Bạn chỉ cần Ctrl+S.

Muốn cài tay cho file đang mở: Alt+F8 → `PERSONAL.XLSB!InstallCpsxLib`.

## Kiểm tra đã cài đúng chưa (làm ngay sau khi cài)

| Kiểm | Cách | Đúng khi |
|---|---|---|
| Skill mới | claude.ai → Skills → mở `cpsx-master` | Phần mô tả bắt đầu bằng "Skill DUY NHẤT…", và trong nội dung có "CPSX MASTER v2" |
| Macro | Alt+F11 → PERSONAL.XLSB | Có module `modCpsxLib`; trong `clsAppEvents` có `App_WorkbookOpen` |
| `_LIB` trong file | Mở file costing → Alt+F8 → `PERSONAL.XLSB!InstallCpsxLib` | Hiện hộp thoại "Da cai CPSX_LIB v2" → Ctrl+S |
| Claude dùng thư viện | Mở **phiên Claude mới** trong Excel, gõ một yêu cầu | Lệnh đầu tiên của Claude là `execute_office_js` có dòng `new Function(...)` và `L.map`, không còn `get_range_as_csv _RULES` |

Phiên Claude đang mở sẽ giữ skill cũ. Sau khi upload skill mới, phải bắt đầu phiên mới.

## Dùng hằng ngày

Nói với Claude như cũ. Claude gọi recipe tương ứng:

| Bạn nói | Claude gọi | Số lệnh |
|---|---|---|
| (đầu phiên) | `L.map` — đọc quy tắc, BoQ, các block Chiet tinh, thư viện mẫu | 1 |
| "tính D05 tương tự D01, 2 cánh" | `L.newItem` — chép block, dịch công thức, tách bảng gộp tấm, link BoQ, group, tô vàng, tính lại, báo giá | 1 |
| "tính giá hạng mục 30–35" | `L.newItems` — cả lô, tính lại 1 lần | 1 |
| "sửa khung đứng 250" | `L.edit` | 1 |
| "mô tả D04" | `L.desc` | 1–2 |
| "bỏ note vàng" | `L.unmark` — trả đúng màu cũ, giữ vàng sẵn của mẫu | 1 |
| "cập nhật rule" | `L.addRules` (mặc định CHUNG) | 1 |
| (tự động) bạn nhắc lỗi trong chat | Sửa + `L.addRules` ngay trong lượt đó, báo 1 dòng "Đã ghi rule…" | 1 |
| (tự động đầu phiên) bạn đã sửa ô Claude ghi | `L.map` trả `userEdits` → Claude ghi rule, hỏi 1 câu nếu không rõ lý do → `L.learned` | 1–2 |

## Thay đổi hành vi
- Việc có recipe thì Claude **làm luôn**, không lập plan chờ duyệt. Claude chỉ hỏi gộp 1 lần khi thiếu thông số làm đổi giá.
- Block mới luôn được thêm **dưới END cuối** của Chiet tinh, không chèn hàng giữa dữ liệu cũ.
- Mọi ô Claude tô vàng được ghi vào sheet ẩn `_EDITLOG` (kèm màu cũ), nên lệnh bỏ vàng chỉ cần 1 lệnh.

## Sửa thư viện
- Mã nguồn: `cpsx-master\lib\cpsx_lib.js`. Khi sửa, tăng `VER`.
- Build lại skill: `node skill-src\build.js` (chạy trong `_MASTER`), sau đó zip lại và upload.
- Các file mở sau đó sẽ tự nhận `_LIB` mới.
- Bản cũ lưu ở `backup\`.
