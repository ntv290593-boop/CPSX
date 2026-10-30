# Flowchart dự án COSTING EXCEL

Thư mục này giữ sơ đồ luồng của từng skill. **Mỗi lần đổi tính năng của skill (sửa SKILL.md, `buoc/*.md`, `tool/*.py`, thêm/bớt lệnh, đổi thứ tự bước, đổi luật merge/check), phải cập nhật flowchart tương ứng trong cùng phiên làm việc.**

## Danh mục

| File | Skill | Nguồn code |
|---|---|---|
| [01-scan_flowchart.md](01-scan_flowchart.md) | 01-scan (quét hồ sơ thầu → PROJECT_DATA) | `SKILL/01-scan/` (SKILL.md, `buoc/`, `tool/pd.py`) |
| [CHANGELOG.md](CHANGELOG.md) | Nhật ký thay đổi flowchart | — |

Skill chưa có flowchart: `10-tender` (thêm `10-tender_flowchart.md` khi cần, dùng mẫu bên dưới).

## Quy tắc cập nhật

1. Sửa tính năng xong → mở flowchart của skill đó, sửa đúng phần bị ảnh hưởng (số phần ở đầu file: 0 luồng tổng, 1–5 các bước, 4 ứng viên, 6 đối chiếu, 7 neo/check/build).
2. Đổi dòng "Cập nhật lần cuối" và "Trạng thái" ở đầu file flowchart.
3. Thêm 1 dòng vào [CHANGELOG.md](CHANGELOG.md): ngày · skill · phần đã sửa · lý do.
4. Chỉ ghi điều có trong code/hướng dẫn hiện hành. Việc chưa làm ghi rõ "chưa làm", không vẽ như đã có.
5. Sơ đồ vẽ bằng khối ASCII trong code fence để xem được ở mọi nơi (VS Code, GitHub, Notepad).

## Mẫu cho skill mới

```
# Flowchart skill <tên>
> Cập nhật lần cuối: <ngày> · Theo code: <đường dẫn>
> Trạng thái: <việc đã / chưa làm>

# 0. LUỒNG TỔNG
# 1..n. TỪNG BƯỚC / THUẬT TOÁN LÕI
# n+1. ĐIỂM THIẾT KẾ ĐÁNG CHÚ Ý
```
