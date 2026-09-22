# Đánh giá dữ liệu và thiết kế nhập liệu

## Dữ liệu mẫu đi kèm dự án

Kết quả rà soát file `node_outputs/PM khoa CTCH T5.xlsx` sau khi sửa bộ đọc dữ liệu:

| Sheet | Dòng dữ liệu | Dòng đầy đủ | Dòng thiếu | Nhận xét |
|---|---:|---:|---:|---|
| `phauthuat` | 118 | 110 | 8 | 8 dòng thiếu `Loại phẫu thuật` |
| `thu thuat` | 1.197 | 1.197 | 0 | Có 3 dòng trống trước dòng Tổng Cộng, không đưa lên giao diện |
| `tieuphau` | 79 | 79 | 0 | Có 4 dòng trống và 1 dòng chỉ có công thức mẫu, không đưa lên giao diện |

## Lỗi dữ liệu tuổi đã phát hiện và sửa

Trong sheet `phauthuat`, tiêu đề `Tuổi` nằm ở cột D nhưng một số dòng lưu tuổi ở cột E. Bộ đọc cũ chỉ kiểm tra cột D nên báo thiếu tuổi sai ở nhiều dòng.

Phiên bản này đọc và ghi tuổi theo cả D/E:

- Dòng có tuổi ở D: giữ nguyên D.
- Dòng có tuổi ở E: đọc và cập nhật đúng E.
- Không còn tính các dòng này là thiếu tuổi.

Sau khi sửa, số dòng thiếu của sheet phẫu thuật giảm từ 60 xuống còn 8 dòng thiếu thực sự.

## Trạng thái file tháng mới sau khi xóa mẫu phụ mổ

Nếu file tháng mới được xóa cả dữ liệu mẫu `PTV chính`, bảng có thể hiển thị toàn bộ dòng là thiếu. Đây không phải lỗi: cột PTV chính là cột bắt buộc và đang chờ nhập dữ liệu thật.

Khung `Kiểm tra dữ liệu` trên giao diện sẽ giải thích rõ cột đang thiếu và số dòng trống/công thức mẫu bị loại khỏi bảng.

## Nguyên tắc UX đã áp dụng

- Bảng nhập liệu là nội dung chính và chiếm phần lớn chiều cao màn hình.
- Chọn file, upload, sheet và chế độ hiển thị nằm trên một thanh gọn.
- Thêm dòng, nhân bản, xóa và lưu nằm ngay trên bảng.
- Không còn khung `Nhập liệu`, `Dòng đang chọn` và `Nhập nhanh bí danh`.
- Tác vụ ít dùng được gom vào menu để giảm nhiễu.
- Gợi ý nhân sự lọc theo đúng vai trò và loại sheet.
