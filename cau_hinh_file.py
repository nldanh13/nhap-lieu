from pathlib import Path
import re
import unicodedata

# Thư mục chứa file code.
BASE_DIR = Path(__file__).resolve().parent

# =========================
# CẤU HÌNH CHẠY FILE THEO NĂM / THÁNG
# =========================

# Bạn đặt dữ liệu trong thư mục có tên đúng bằng NAM_DU_LIEU, cùng cấp với file code.
# Ví dụ:
# E:/Thong_ke/2026/TH1.2026.xlsx
# E:/Thong_ke/2026/TH2.2026.xlsx
# ...
NAM_DU_LIEU = 2026
THU_MUC_TONG_HOP_THANG = BASE_DIR / str(NAM_DU_LIEU)

# Chọn tháng cần chạy RIÊNG cho từng nguồn dữ liệu.
#
# 1) THANG_CAN_CHAY_TONG_HOP:
#    Dùng khi chạy file tổng hợp tháng bằng:
#       python tong_hop_thang_excel.py
#
# 2) THANG_CAN_CHAY_THEO_SO:
#    Dùng khi chạy thư mục TheoSo bằng:
#       python chay_theo_so.py
#
# Cách khai báo tháng:
#   [] hoặc "tat_ca"  => chạy cả năm
#   "1-5"             => chạy tháng 1 đến tháng 5
#   [6]               => chỉ chạy tháng 6
#   "1,3,5"           => chạy tháng 1, 3, 5
#
# Tình huống hiện tại của bạn:
# - Tháng 1 đến tháng 5 đã có file tổng hợp tháng.
# - Tháng 6 chưa có file tổng hợp tháng nên chạy theo sổ.
THANG_CAN_CHAY_TONG_HOP = [1, 2, 3, 4, 5]
THANG_CAN_CHAY_THEO_SO = [6]

# Giữ biến cũ để các file code cũ import không lỗi.
# Mặc định biến cũ được hiểu là cấu hình cho file tổng hợp tháng.
THANG_CAN_CHAY = THANG_CAN_CHAY_TONG_HOP

# Nếu muốn chỉ định file thủ công thì thêm vào đây.
# Nếu danh sách này có file, chương trình sẽ ưu tiên dùng danh sách này,
# không quét thư mục theo NAM_DU_LIEU nữa.
DANH_SACH_FILE_TONG_HOP_THANG = []

# Các file bắt đầu bằng các tiền tố này sẽ bị bỏ qua khi quét thư mục.
BO_QUA_FILE_BAT_DAU_BANG = ["~$", "thong_ke_"]

# Chỉ xử lý các sheet bạn cần quan tâm.
SHEET_TONG_HOP_THANG_CAN_XU_LY = [
    "phauthuat",
    "phauthuat_BHYT",
    "thu thuat",
    "tieuphau",
]

# =========================
# FILE ĐẦU RA
# =========================

def _parse_month_selection(value):
    """
    Trả về None nếu chạy tất cả tháng.
    Trả về list[int] nếu chỉ chạy một số tháng.
    """
    if value is None:
        return None

    if value == [] or value == "":
        return None

    if isinstance(value, str):
        s = value.strip().lower()
        if s in {"tat_ca", "tất cả", "all", "ca_nam", "cả năm", ""}:
            return None

        months = []
        for part in s.split(","):
            part = part.strip()
            if not part:
                continue

            if "-" in part:
                start, end = part.split("-", 1)
                start = int(start.strip())
                end = int(end.strip())
                months.extend(range(start, end + 1))
            else:
                months.append(int(part))
    elif isinstance(value, int):
        months = [value]
    else:
        months = [int(x) for x in value]

    months = sorted(set(months))
    invalid = [m for m in months if m < 1 or m > 12]
    if invalid:
        raise ValueError(f"Tháng không hợp lệ: {invalid}. Chỉ nhận từ 1 đến 12.")

    return months


def _output_suffix_from_months(months):
    if months is None:
        return f"{NAM_DU_LIEU}_ca_nam"

    if len(months) == 1:
        return f"{NAM_DU_LIEU}_T{months[0]}"

    consecutive = months == list(range(months[0], months[-1] + 1))
    if consecutive:
        return f"{NAM_DU_LIEU}_T{months[0]}_den_T{months[-1]}"

    return f"{NAM_DU_LIEU}_" + "_".join(f"T{m}" for m in months)


THANG_DA_CHON_TONG_HOP = _parse_month_selection(THANG_CAN_CHAY_TONG_HOP)
THANG_DA_CHON_THEO_SO = _parse_month_selection(THANG_CAN_CHAY_THEO_SO)

# Giữ biến cũ để các file code cũ import không lỗi.
THANG_DA_CHON = THANG_DA_CHON_TONG_HOP

OUTPUT_FILE = BASE_DIR / f"thong_ke_phan_loai_{_output_suffix_from_months(THANG_DA_CHON_TONG_HOP)}.xlsx"
OUTPUT_FILE_THEO_SO = BASE_DIR / f"thong_ke_phan_loai_TheoSo_{_output_suffix_from_months(THANG_DA_CHON_THEO_SO)}.xlsx"

# =========================
# TÙY CHỌN XUẤT SHEET
# =========================

XUAT_SHEET_THEO_NHAN_SU = False
XUAT_SHEET_THEO_NHAN_SU_NHOM = False
XUAT_SHEET_THEO_VAI_TRO = False
XUAT_SHEET_DU_LIEU_LOC = True
XUAT_SHEET_DU_LIEU_VAI_TRO = False

# =========================
# CHẾ ĐỘ XỬ LÝ
# =========================

# File TH theo tháng chỉ phân loại kỹ thuật, không lọc nhân sự.
XU_LY_FILE_TONG_HOP_THANG = True

# Tắt chế độ HIS cũ.
XU_LY_SO_THU_THUAT = False
XU_LY_SO_PHAU_THUAT = False

# Giữ lại để các file code cũ import không lỗi, nhưng chế độ này đang tắt.
FILE_SO_THU_THUAT = BASE_DIR / "SoThuThuat_20260703164043.xlsx"
FILE_SO_PHAU_THUAT = BASE_DIR / "PTTT_SoPhauThuat_TT50_20260703164153(1).xlsx"

# =========================
# KIỂM TRA / XỬ LÝ TRÙNG PHẪU THUẬT - BHYT
# =========================

# Mỗi file tháng có thể có cùng một ca nằm ở cả sheet phauthuat và phauthuat_BHYT.
# Khóa trùng: Ngày + Họ tên người bệnh + Phương pháp phẫu thuật.
# True  = khi thống kê tổng sẽ giữ dòng ở phauthuat và loại dòng trùng ở phauthuat_BHYT.
# False = vẫn giữ đủ cả hai sheet nhưng có sheet báo cáo trùng để kiểm tra.
LOAI_TRUNG_PHAUTHUAT_BHYT = True

# Xuất sheet báo cáo các dòng trùng giữa phauthuat và phauthuat_BHYT.
XUAT_SHEET_TRUNG_LAP_PHAUTHUAT_BHYT = True

# =========================
# HÀM TÌM FILE THEO THÁNG
# =========================

def _remove_accents(text):
    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    return text


def _extract_month_from_filename(file_path):
    """
    Cố gắng lấy tháng từ tên file.
    Hỗ trợ các dạng tên phổ biến:
    - TH1.2025.xlsx
    - TH01.2025.xlsx
    - T1.xlsx
    - Thang 1.xlsx / Tháng 1.xlsx
    """
    name = _remove_accents(Path(file_path).stem).lower()

    patterns = [
        rf"(?:^|[^a-z0-9])th\s*0?([1-9]|1[0-2])(?:[^0-9]|$)",
        rf"(?:^|[^a-z0-9])t\s*0?([1-9]|1[0-2])(?:[^0-9]|$)",
        rf"thang\s*0?([1-9]|1[0-2])(?:[^0-9]|$)",
    ]

    for pattern in patterns:
        match = re.search(pattern, name)
        if match:
            return int(match.group(1))

    return None


def _file_is_allowed_by_month(file_path):
    months = THANG_DA_CHON_TONG_HOP
    if months is None:
        return True

    file_month = _extract_month_from_filename(file_path)
    if file_month is None:
        return False

    return file_month in months


def get_file_tong_hop_thang_list():
    """
    Trả về danh sách file tổng hợp tháng cần xử lý.
    Bản này KHÔNG dùng FILE_TONG_HOP_THANG đơn lẻ nữa.
    Chỉ chạy theo:
    1. DANH_SACH_FILE_TONG_HOP_THANG nếu có khai báo thủ công
    2. Thư mục năm theo THANG_CAN_CHAY_TONG_HOP
    """
    if DANH_SACH_FILE_TONG_HOP_THANG:
        result = []
        for p in DANH_SACH_FILE_TONG_HOP_THANG:
            file_path = Path(p)
            if _file_is_allowed_by_month(file_path):
                result.append(file_path)
        return result

    if not THU_MUC_TONG_HOP_THANG.exists():
        return []

    files = []
    for pattern in ("*.xlsx", "*.xls"):
        files.extend(THU_MUC_TONG_HOP_THANG.glob(pattern))

    result = []
    for file_path in sorted(files):
        name = file_path.name
        if any(name.startswith(prefix) for prefix in BO_QUA_FILE_BAT_DAU_BANG):
            continue
        if not _file_is_allowed_by_month(file_path):
            continue
        result.append(file_path)

    return result
