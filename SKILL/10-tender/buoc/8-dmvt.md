# Bước 8 — dmvt: dịch, ghi chú, chọn ảnh ca khó

**Ghi:** `_data\_dmvt_llm.json`. Không sửa `_dmvt_boc.json`, `_dmvt_snapshot.json`, `_dmvt.json`, không tự ghi xlsx.
Tool đã bóc A:I, cắt ảnh, áp `Trạng thái cuối` của MAT. Bước này chỉ làm phần máy không làm được.

## Làm
1. Đọc README. Xem dòng: `_data\_dmvt_boc.json`, mỗi dòng có `Khóa`, `Mã`, `Material`, `Location`, `Cờ ảnh`.
2. Với mỗi dòng ghi 1 mục theo `Khóa`, đúng nguyên văn:
```json
{"51-500A": {"Material VN": "Sofa: vải bọc", "Location VN": "- Sofa: vải bọc – 51-500A", "Ghi chú": "..."},
 "50-200":  {"Material VN": "Thảm Axminster đặt riêng trải kín", "Location VN": "- ...", "Ảnh chọn": 0}}
```
   - `Material VN`: dịch `Material`. Máy ghép dưới chữ tiếng Anh.
   - `Location VN`: dịch `Location`, **thay** chữ máy ở cột E, nên phải giữ đủ các mã.
   - `Ghi chú`: chỉ điều đọc được từ chữ spec (vd lỗi chính tả, repeat lớn phải tính hao hụt). Không đoán. Máy nối sau kết luận TCQ ở cột I.
   - `Ảnh chọn`: chỉ số trong danh sách ứng viên, chỉ cho dòng có cờ `Ảnh ca khó`. Xem ứng viên bằng ảnh; ưu tiên ảnh mẫu vật liệu, không lấy mặt bằng. Dòng này luôn có cờ `LLM chọn ảnh`.
3. Không dịch tên riêng: tên mẫu vải, mã màu, tên hãng, model.
4. Không có căn cứ thì bỏ khóa đó. Thiếu `Material VN` thì dòng có cờ `Chưa dịch`, không để trống.

## Cột Area ca lệch (khi user yêu cầu)
Cờ `Area theo tiền tố` đúng ở mức khu, chưa phân biệt từng món. Ca `Lệch Area` ở `VL_XUATHIEN` phải đọc tag trên mặt bằng bằng ảnh, tag là nét vẽ nên không tìm được bằng chữ. Đọc xong ghi vào cột C của file, build lại vẫn giữ. Không suy tờ chưa xem giống tờ đã xem.

## Trước khi trả kết quả
- `py -I "<TOOL>" dmvt <GOC>` chạy qua, cờ `Chưa dịch` về 0 hoặc nêu lý do.
- Trả về: số dòng đã dịch, dòng chọn ảnh, điểm cần xác nhận với thiết kế.
