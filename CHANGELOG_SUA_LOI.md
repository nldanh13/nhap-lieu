# Các lỗi đã sửa

## Web app và tự động lưu

- Khóa thao tác lưu theo đúng file và sheet hiện tại.
- Lưu thành công trước khi cho phép đổi file hoặc đổi sheet.
- Không xóa trạng thái thay đổi khi ghi Excel thất bại.
- Thay lưu từng dòng bằng API lưu nhiều dòng trong một lần mở/lưu workbook.
- Tải toàn bộ dòng phù hợp thay vì giới hạn 900 dòng.
- Hiển thị các cột bắt buộc được dùng để xác định thiếu dữ liệu.
- Lưu đường dẫn autosave tương đối; bỏ qua đường dẫn cũ không còn tồn tại.
- Trả cảnh báo khi không cập nhật được công thức tổng hoặc gặp header không tồn tại.

## Thống kê

- `tong_hop_thang_excel.py` dùng đúng thư mục theo `NAM_DU_LIEU` trong thông báo lỗi.
- `xuat_excel.py` xuất thêm sheet `Chưa phân loại`.
- Xuất sheet `Trùng PT-BHYT` khi có báo cáo trùng.

## PTV/phụ mổ

- Không xóa trắng PTV/phụ mổ trước khi đối chiếu.
- Giữ nguyên dữ liệu nhập tay khi không khớp hoặc dữ liệu nguồn trống.

## Đóng gói

- Thêm `requirements.txt`.
- Thêm hướng dẫn cài đặt và chạy trong `README.md`.

## Thiết kế lại giao diện và làm sạch file tháng mới – 22/07/2026

- Bỏ tiêu đề cố định `T5`; tiêu đề tự xác định tháng/năm và loại sheet.
- Chia quy trình thành các bước: chọn file, upload, chọn sheet và chọn chế độ hiển thị.
- Chỉ hiện công cụ phẫu thuật hoặc thủ thuật phù hợp với sheet đang mở.
- Bổ sung bảng chỉ số: dòng hiển thị, dòng thiếu dữ liệu và dòng đang sửa.
- Bổ sung cảnh báo khi file nguồn có dữ liệu sẵn trong các cột phụ mổ.
- Bổ sung nút **Xóa dữ liệu mẫu phụ mổ** với màn hình xem trước và xác nhận an toàn.
- Mặc định chỉ xóa `Phụ mổ 1`, `Phụ mổ 2`, `Phụ mổ 3`; `PTV chính` là lựa chọn riêng.
- Chỉ xử lý các dòng dữ liệu trước `Tổng Cộng`; giữ nguyên công thức.
- Lưu kết quả vào file `_NHAP_LIEU`, không sửa file nguồn upload.
- Bổ sung hỗ trợ liệt kê và mở file `.xlsm`, giữ VBA khi lưu.
- Ẩn file báo cáo phụ khỏi danh sách file Excel làm việc.

## Tối ưu UX nhập liệu và quản lý nhân sự – 22/07/2026

- Thiết kế lại màn hình theo hướng ưu tiên bảng nhập liệu; giảm chiều cao phần chọn file và đưa bảng lên gần đầu trang.
- Bỏ các khung `Nhập liệu`, `Dòng đang chọn` và `Nhập nhanh bí danh`.
- Gom thêm dòng, nhân bản, xóa và lưu thành một thanh công cụ ngay trên bảng.
- Các tác vụ tự động ít dùng được đưa vào menu `Thêm tác vụ`.
- Bổ sung bảng đánh giá dữ liệu: số dòng đầy đủ, cột đang thiếu và số dòng trống/công thức mẫu bị bỏ qua.
- Sửa cách đọc cột `Tuổi` của mẫu phẫu thuật khi tuổi nằm ở cột D hoặc E.
- Bổ sung tab `Quản lý nhân sự`: thêm, sửa, xóa, tạm ngưng, tìm kiếm và lọc theo vai trò.
- Danh sách nhân sự được lưu trong `nhan_su_web.json` và được dùng chung cho web app lẫn các script tự động.
- Khi nhập sheet phẫu thuật, gợi ý ở các cột PTV/phụ mổ chỉ gồm `Bác sĩ` và `BSNT`; không còn gợi ý Điều dưỡng hoặc KTV.
- Khi nhập sheet tiểu phẫu, cột Bác sĩ chỉ gợi ý Bác sĩ/BSNT; cột Điều dưỡng và KTV dùng đúng danh sách theo vai trò.
