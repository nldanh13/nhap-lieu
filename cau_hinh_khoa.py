# =========================
# CẤU HÌNH KHOA CHỈ ĐỊNH THEO TỪNG FILE
# =========================
# Cách hoạt động:
# - Sổ Thủ Thuật sẽ lọc theo danh sách khoa riêng của Sổ Thủ Thuật.
# - Sổ Phẫu Thuật sẽ lọc theo danh sách khoa riêng của Sổ Phẫu Thuật.
# - Chương trình so khớp dạng "chứa cụm từ", không phân biệt hoa thường và dấu tiếng Việt.
#   Ví dụ "Khoa Ngoại Chấn Thương" sẽ khớp với
#   "Khoa Ngoại Chấn Thương Chỉnh Hình và Thần Kinh".

KHOA_CHI_DINH_THEO_NGUON = {
    "Sổ Thủ Thuật": [
        "Khoa Khám Bệnh",
        "Khoa Ngoại Chấn Thương",
    ],
    "Sổ Phẫu Thuật": [
        "Khoa Cấp Cứu",
        "Khoa Ngoại Chấn Thương",
    ],
}

# Danh sách dự phòng nếu sau này có thêm nguồn khác không nằm trong KHOA_CHI_DINH_THEO_NGUON.
# Mặc định để trống nghĩa là nguồn đó sẽ lấy tất cả khoa.
KHOA_CHI_DINH_CAN_LAY_MAC_DINH = []

# Có xuất sheet thống kê theo khoa chỉ định hay không.
XUAT_SHEET_THEO_KHOA_CHI_DINH = True


def get_khoa_chi_dinh_can_lay(source_name):
    """
    Trả về danh sách khoa cần lấy theo từng nguồn dữ liệu.
    """
    return KHOA_CHI_DINH_THEO_NGUON.get(source_name, KHOA_CHI_DINH_CAN_LAY_MAC_DINH)
