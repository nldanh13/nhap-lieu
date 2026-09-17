# =========================
# DANH SÁCH BÁC SĨ
# =========================

BAC_SI = [
    "Nguyễn Thành Tấn",
    "Hồ Điền",
    "Nguyễn Chí Nguyện",
    "Nguyễn Giang Tử",
    "Nguyễn Lê Hoan",
    "Nguyễn Tư Thái Bảo",
    "Phan Văn Tuấn",
    "Phạm Việt Tân",
    "Trần Nguyễn Anh Duy",
    "Trần Quang Sơn",
    "Trần Quốc Toản",
]


# =========================
# DANH SÁCH ĐIỀU DƯỠNG
# =========================

DIEU_DUONG = [
    "Phan Văn Hiếu",
    "Võ Thị Yến Nhi",
    "Võ Phương Duy",
    "Nguyễn Kim Ngân",
    "Lê Ngọc Diệu",
    "Trần Quỳnh Minh Thư",
    "Lê Thị Tuyết Đoan",
    "Thạch Thị Thúy Đa",
    "Lê Kim Hoàn",
    "Nguyễn Lê Duy Anh",
    "Hồ Nguyễn Tố Như",
]

KTV = [
    "Nguyễn Duy Linh",
    "Thạch Trần Thị Cẩm Tú",
    "Nguyễn Thị Nhân Ái",
]


# =========================
# DANH SÁCH BÁC SĨ NỘI TRÚ / BSNT
# =========================
# Dùng để nhập nhanh phụ mổ và chuyển bí danh nhân sự trong web app.

BSNT = [
    {"ho_ten": "Lê Nhỉ Khang", "bi_danh": "KHANG", "nhom": "BSNT 3"},
    {"ho_ten": "Lê Quốc Khánh", "bi_danh": "KHÁNH", "nhom": "BSNT 3"},
    {"ho_ten": "Nguyễn Hoài Linh", "bi_danh": "LINH", "nhom": "BSNT 3"},
    {"ho_ten": "Trần Trung Nghĩa", "bi_danh": "NGHĨA", "nhom": "BSNT 3"},
    {"ho_ten": "Nguyễn Văn Nghiêm", "bi_danh": "NGHIÊM", "nhom": "BSNT 3"},
    {"ho_ten": "Nguyễn Trung Nhị", "bi_danh": "NHỊ", "nhom": "BSNT 3"},
    {"ho_ten": "Võ Ngọc Thiện", "bi_danh": "THIỆN", "nhom": "BSNT 3"},
    {"ho_ten": "Phạm Văn Hưởng", "bi_danh": "V.HƯỞNG", "nhom": "BSNT 3"},
    {"ho_ten": "Lê Cường Thạnh", "bi_danh": "THẠNH", "nhom": "BSNT 2"},
    {"ho_ten": "Phạm Hữu Trọng", "bi_danh": "TRỌNG", "nhom": "BSNT 2"},
    {"ho_ten": "Công Chỉnh", "bi_danh": "CHỈNH", "nhom": "BSNT 1"},
    {"ho_ten": "Dương Minh Hậu", "bi_danh": "HẬU", "nhom": "BSNT 1"},
    {"ho_ten": "Đặng Duy Lộc", "bi_danh": "LỘC", "nhom": "BSNT 1"},
    {"ho_ten": "Trần Huỳnh Minh Thiện", "bi_danh": "M.THIỆN", "nhom": "BSNT 1"},
    {"ho_ten": "Trịnh Thị Ý Như", "bi_danh": "NHƯ", "nhom": "BSNT 1"},
    {"ho_ten": "Nguyễn Thị Thanh Thuý", "bi_danh": "THUÝ", "nhom": "BSNT 1"},
    {"ho_ten": "Nguyễn Trần Việt Tiến", "bi_danh": "TIẾN", "nhom": "BSNT 1"},
]

# =========================
# TÊN VIẾT TẮT / BÍ DANH TRONG FILE TỔNG HỢP THÁNG
# =========================
# File "Phau thuat-Thu thuat-Tieu thu thuat TH1.2025" dùng tên viết tắt
# như HOAN, NGUYỆN, V.TÂN... thay vì họ tên đầy đủ.
# Nếu sau này gặp tên viết tắt khác, thêm vào đây.
# Lưu ý: chương trình so khớp bí danh theo kiểu còn dấu để tránh nhầm
# TẤN với V.TÂN.

BI_DANH_NHAN_SU = {
    "Nguyễn Thành Tấn": ["TẤN"],
    "Hồ Điền": ["ĐIỀN"],
    "Nguyễn Chí Nguyện": ["NGUYỆN"],
    "Nguyễn Giang Tử": ["G.TỬ"],
    "Nguyễn Lê Hoan": ["HOAN"],
    "Nguyễn Tư Thái Bảo": ["BẢO"],
    "Phan Văn Tuấn": ["TUẤN"],
    "Phạm Việt Tân": ["V.TÂN"],
    "Trần Nguyễn Anh Duy": ["DUY"],
    "Trần Quang Sơn": ["SƠN"],
    "Trần Quốc Toản": ["TOẢN"],
}

# Tự động đưa BSNT vào bảng bí danh để web app và các script có thể tra cứu.
for _item in BSNT:
    BI_DANH_NHAN_SU.setdefault(_item["ho_ten"], [])
    if _item["bi_danh"] not in BI_DANH_NHAN_SU[_item["ho_ten"]]:
        BI_DANH_NHAN_SU[_item["ho_ten"]].insert(0, _item["bi_danh"])


def get_bi_danh(person_name):
    return BI_DANH_NHAN_SU.get(person_name, [])

# =========================
# CẤU HÌNH NHÂN SỰ TỪ WEB APP
# =========================
# Khi file nhan_su_web.json tồn tại, danh sách được quản lý trong tab
# "Quản lý nhân sự" sẽ thay thế cấu hình mặc định phía trên. Các script tự động
# tiếp tục import BAC_SI / DIEU_DUONG / KTV / BSNT như cũ nên không cần sửa luồng.
def _nap_cau_hinh_nhan_su_web():
    import json
    from pathlib import Path

    config_path = Path(__file__).resolve().with_name("nhan_su_web.json")
    if not config_path.exists():
        return
    try:
        payload = json.loads(config_path.read_text(encoding="utf-8"))
        rows = payload.get("staff", []) if isinstance(payload, dict) else []
        rows = [row for row in rows if isinstance(row, dict) and row.get("active", True)]
    except Exception:
        return

    BAC_SI[:] = [str(row.get("hoTen", "")).strip() for row in rows if row.get("vaiTro") == "bac_si" and str(row.get("hoTen", "")).strip()]
    DIEU_DUONG[:] = [str(row.get("hoTen", "")).strip() for row in rows if row.get("vaiTro") == "dieu_duong" and str(row.get("hoTen", "")).strip()]
    KTV[:] = [str(row.get("hoTen", "")).strip() for row in rows if row.get("vaiTro") == "ktv" and str(row.get("hoTen", "")).strip()]
    BSNT[:] = [
        {
            "ho_ten": str(row.get("hoTen", "")).strip(),
            "bi_danh": str(row.get("biDanh", "")).strip(),
            "nhom": str(row.get("nhom", "") or "BSNT").strip(),
        }
        for row in rows
        if row.get("vaiTro") == "bsnt" and str(row.get("hoTen", "")).strip()
    ]

    BI_DANH_NHAN_SU.clear()
    for row in rows:
        name = str(row.get("hoTen", "")).strip()
        alias = str(row.get("biDanh", "")).strip()
        if not name:
            continue
        BI_DANH_NHAN_SU.setdefault(name, [])
        if alias and alias not in BI_DANH_NHAN_SU[name]:
            BI_DANH_NHAN_SU[name].append(alias)


_nap_cau_hinh_nhan_su_web()
