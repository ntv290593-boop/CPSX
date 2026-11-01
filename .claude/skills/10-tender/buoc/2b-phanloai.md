# Bước 2b — phanloai: gán Mã cấp 3 cho dòng máy chưa phân loại được

**Ghi:** `_data\_phanloai_llm.json`, `VIEC_MO.phanloai.json` nếu có câu cần hỏi. Không ghi `SCOPE.*.json`.
Tool `phanloai` gán trước theo thứ tự: User > Luật > Nhớ mô tả > Nhớ đầu mục. Bạn chỉ làm các dòng còn lại. Đây chỉ là phân loại **hạng mục thuộc nhóm nào**, không suy ra ai làm.

## Làm
1. Đọc README. Lấy danh sách: `py -I "<TOOL>" loc <GOC> SCOPE "Cờ phân loại=Chờ LLM" --cot Khóa,Nhóm,Mục phụ,Mô tả BOQ,Mã VL --max 800`.
2. Lấy danh sách đóng: `py -I "<TOOL>" cot --buoc phanloai-tool`, cột `Mã cấp 3`. Mô tả từng mã ở `D:\COSTING EXCEL\_MASTER\code_id\bang_ma.json` (`mo_ta`).
3. Mỗi dòng chọn đúng 1 tên trong danh sách, xét theo mục lớn, đầu mục, vật liệu của dòng. Không chắc thì ghi `KHÔNG RÕ`, không chọn đại.
4. Ghi:
```json
{"rows": [
  {"Khóa": "2. Finishing-L4-17!123", "Mã cấp 3": "Hoàn thiện sàn gỗ"},
  {"Khóa": "3. Furniture-L4-17!159", "Mã cấp 3": "KHÔNG RÕ"}
]}
```
5. Chạy lại `py -I "<TOOL>" phanloai <GOC>`. Tool báo LỖI nếu tên ngoài danh sách hoặc khóa không có. Dòng cùng mô tả và đầu mục dùng chung câu trả lời, nên có thể chỉ cần trả lời 1 dòng đại diện, rồi chạy lại xem còn dòng nào.

## Không được
- Không ghi `Nguồn phân loại`, `Cờ phân loại`: tool tự ghi.
- Không đưa câu trả lời vào `_MASTER\code_id`. LLM chỉ nhớ trong dự án.
- Không lấy mã từ dự án khác.
- Không sửa ô ở sheet xlsx. Ô user đã chọn droplist (`Nguồn phân loại` = User) luôn thắng, không trả lời lại.

## Trước khi trả kết quả
- Trả về: số dòng đã trả lời, số `KHÔNG RÕ`, số dòng còn chờ sau khi chạy lại.
