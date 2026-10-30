# COSTING EXCEL

Đầu mỗi phiên, đọc `ACTIVE.md` để biết dự án đang làm.

## Chống nhầm dự án
- Chỉ lấy số liệu, mã, đơn giá, kích thước từ file của dự án ghi trong `ACTIVE.md`.
- Mọi số liệu đưa ra phải nêu nguồn (file / sheet / ô). Không có nguồn thì ghi "chưa có", không suy từ dự án khác.
- Không mở dự án khác (kể cả trong `_TEST`, `01. DOOR`) để lấy số liệu. Muốn tham chiếu thì hỏi người dùng trước.
- Memory nhắc dự án khác với `ACTIVE.md` thì coi là cũ, kiểm tra lại trước khi dùng.
- `_ARCHIVE` là lưu trữ, bị chặn đọc.

## Flowchart
- Flowchart các skill nằm ở `DOCS\Flowchart\` (xem `README.md` ở đó).
- Mỗi khi đổi tính năng của skill (SKILL.md, `buoc/*.md`, `tool/*.py`), cập nhật flowchart tương ứng + thêm 1 dòng vào `DOCS\Flowchart\CHANGELOG.md` trong cùng phiên.
