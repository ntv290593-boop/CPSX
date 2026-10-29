# Nguyên tắc (user chốt 2026-10-24)

1. **Sản phẩm cuối là file Excel (CPSX) có công thức**, như người dùng thao tác tay: chép block mẫu, nhập ô đầu vào, mọi thứ còn lại là hàm. Không dán số chết.
2. **Hạn chế value tối đa.** Cái gì liên quan nhau thì LINK: kích thước/SL đầu block = công thức đọc BoQ (R0023), mã vật liệu tra Material qua K (R0073), tham số dự án lấy từ OverView. Chỉ gõ tay các ô đầu vào mà mẫu quy định là ô nhập (R0075).
3. **Không viết lại công thức của mẫu.** Dùng module có sẵn (R0039); chỉ thêm dòng khi chưa có, chèn trước hàng đệm (R0022). Công thức tham chiếu tương đối, không khóa `$` (R0032).
4. **File trong thư mục này chỉ là tri thức + bộ kiểm.** Script/engine tính lại số lượng chỉ để ĐỐI CHIẾU với Excel, không thay công thức trong Excel. Lệch → báo user, không tự sửa mẫu (FILE-00 mục 5).
5. Ô Claude sửa tô vàng #FFFF00 (R0051); nghiệp vụ không rõ thì hỏi (R0075).
