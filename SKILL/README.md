# SKILL — mỗi lệnh 1 skill, mỗi skill 1 việc

Claude Code đọc thư mục này qua liên kết `D:\COSTING EXCEL\.claude\skills` → `SKILL` (junction). Sửa file ở đây là có hiệu lực ngay ở phiên mới.

| Gõ lệnh | Skill | Ra file | Trạng thái |
|---|---|---|---|
| `scan` / `scan thư mục <đợt>` | `01-scan` | `01_scan\PROJECT_DATA_<đợt>.xlsx` | ✅ 27/10/2026 |
| `takeoff` / `bóc` | `02-takeoff` | `02_takeoff\…` (AI_INPUT) | chưa làm |
| (trong Excel) | `cpsx-master` (Claude trong Excel) | `03_costing\*.xlsm` | giữ nguyên, tách sau |
| `rfi` | `04-rfi` | `04_rfi\…RFI_REV00.xlsx` | chưa làm |
| `mail` | `05-mail` | `05_mail\…` | chưa làm |
| `hỏi giá ncc` | `06-ncc` | `06_ncc\…` | chưa làm |
| `scan tender` / `quét HSMT` / `cập nhật TCQ` | `10-tender` | `10_TENDER\03_boq\<DA>_BOQ_TRACK_<yymmdd>.xlsx` + bản gửi CĐT | ✅ 28/10/2026 (tool đã chạy thử SOFITEL; các bước LLM chưa chạy) |

Số đầu tên skill = thư mục của dự án mà skill ghi vào.

## Nguyên tắc chung (vì sao chia như vậy)
1. **Một lệnh của user → skill tự chạy hết các bước.** Bước nặng giao cho subagent mới (ngữ cảnh sạch), mỗi subagent chỉ đọc file hướng dẫn của bước đó.
2. **Form cố định.** Đầu ra theo form xlsx trong `<skill>\form\`, không để model tự nghĩ cấu trúc mỗi lần.
3. **Model quyết định, máy ghi.** Model ghi JSON; script cố định dựng xlsx, tạo link, kiểm lỗi. Mỗi cột chỉ một bước được ghi.
4. **User duyệt ngoại lệ.** File kết quả có DASHBOARD → "CẦN XEM", user không phải đọc từng dòng.
5. **Sửa tay của user luôn được giữ** khi chạy lại (`*.user.json`).

10-tender dùng `form\COT.json` (thứ tự/tên cột, bước được ghi, danh sách) thay cho form xlsx.

## Cấu trúc một skill
```
<skill>\
├── SKILL.md      điều phối: bước nào, ai làm, kiểm gì, báo gì (ngắn)
├── JSON.md       cách ghi dữ liệu (chung cho mọi bước)
├── buoc\         1 file / bước — chỉ subagent của bước đó đọc
├── form\         form xlsx (đổi cột/thứ tự cột: sửa ở đây, sheet ẩn _COT)
└── tool\         script cố định
```

## Đổi form PROJECT_DATA
Mở `01-scan\form\PROJECT_DATA_FORM.xlsx`:
- Đổi **thứ tự cột**: kéo cột ở hàng 1 của sheet; giữ đúng tên.
- Thêm/đổi **cột**: thêm ở hàng 1 **và** thêm dòng trong sheet ẩn `_COT` (bước được ghi, bắt buộc, kiểu, danh sách).
- Cột xám = nhóm thu gọn (bấm dấu + phía trên để mở).
