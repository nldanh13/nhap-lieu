# Nhập liệu phẫu thuật – thủ thuật

Web app nội bộ hỗ trợ nhập liệu sổ phẫu thuật, thủ thuật, tiểu phẫu và quản lý danh mục nhân sự Khoa Ngoại CTCH–TK.

## Yêu cầu

- Node.js 18+
- Python 3.10+
- Google Chrome hoặc Microsoft Edge nếu sử dụng chức năng EMR

## Cài đặt

```bash
npm install
python -m pip install -r requirements.txt
```

Sao chép `emr_config.example.json` thành `emr_config.json`, sau đó nhập cấu hình trên máy cục bộ. Không commit `emr_config.json` vì có thông tin đăng nhập.

## Chạy ứng dụng

```bash
npm start
```

Mở: http://127.0.0.1:3002

## Cấu trúc chính

- `server.js`: API Node/Express.
- `public/`: giao diện web.
- `node_excel_api.py`: đọc, cập nhật và hoàn tất file Excel.
- `emr_tieuphau_fill.py`: bổ sung bác sĩ từ EMR.
- `emr_integration/`: mô-đun tìm kiếm EMR.
- `nhan_su_web.json`: danh mục nhân sự và bí danh.
- `ten_cls_tieuphau.json`: danh mục CLS tiểu phẫu.

## An toàn dữ liệu

Không đưa lên GitHub mật khẩu EMR, file Excel người bệnh, ảnh sổ phẫu thuật, thư mục upload/output/autosave hoặc dữ liệu nhận dạng người bệnh.
