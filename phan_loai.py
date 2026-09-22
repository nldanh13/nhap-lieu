import pandas as pd

from cau_hinh_phuong_phap import (
    NHOM_PHAN_LOAI_THEO_THU_TU,
    PHAN_LOAI_PHUONG_PHAP_UU_TIEN,
)
from utils_text import normalize_text


CHUA_PHAN_LOAI = "Chưa phân loại"


def kiem_tra_nhom_hop_le(nhom):
    if nhom == CHUA_PHAN_LOAI:
        return

    if nhom not in NHOM_PHAN_LOAI_THEO_THU_TU:
        raise ValueError(
            f"Nhóm phân loại không hợp lệ: {nhom}. "
            f"Hãy kiểm tra lại trong cau_hinh_phuong_phap.py"
        )


def _normalize_source_key(source_name):
    return normalize_text(source_name or "")


# Tạo index để phân loại nhanh, tránh dò 677 kỹ thuật cho từng dòng dữ liệu.
_EXACT_ALL = {}
_EXACT_BY_SOURCE = {}
_CONTAINS_ALL = []
_CONTAINS_BY_SOURCE = {}

for item in PHAN_LOAI_PHUONG_PHAP_UU_TIEN:
    phuong_phap = item.get("phuong_phap", "")
    nhom = item.get("nhom", "")
    kieu_khop = item.get("kieu_khop", "chinh_xac")
    nguon = item.get("nguon", "") or item.get("nguon_ap_dung", "")

    if not phuong_phap or not nhom:
        continue

    kiem_tra_nhom_hop_le(nhom)
    pp_norm = normalize_text(phuong_phap)

    data = (nhom, phuong_phap)

    if isinstance(nguon, str):
        nguon_list = [nguon] if nguon else []
    elif isinstance(nguon, (list, tuple, set)):
        nguon_list = list(nguon)
    else:
        nguon_list = []

    if kieu_khop == "chinh_xac":
        if nguon_list:
            for n in nguon_list:
                _EXACT_BY_SOURCE.setdefault(_normalize_source_key(n), {})[pp_norm] = data
        else:
            _EXACT_ALL[pp_norm] = data
    elif kieu_khop == "chua_cum_tu":
        if nguon_list:
            for n in nguon_list:
                _CONTAINS_BY_SOURCE.setdefault(_normalize_source_key(n), []).append((pp_norm, nhom, phuong_phap))
        else:
            _CONTAINS_ALL.append((pp_norm, nhom, phuong_phap))


def classify_method(method_text, source_name=""):
    """
    Phân loại theo TÊN KỸ THUẬT CỤ THỂ, không dùng keyword.

    Thứ tự:
    1. Dò tên chính xác trong PHAN_LOAI_PHUONG_PHAP_UU_TIEN.
    2. Dò cụm từ chỉ khi bạn chủ động đặt kieu_khop="chua_cum_tu".
    3. Nếu không khớp -> Chưa phân loại.

    Như vậy khi có kỹ thuật mới hoặc kỹ thuật bị xếp sai, bạn chỉ cần sửa
    danh sách trong cau_hinh_phuong_phap.py, không lo keyword bắt nhầm.
    """
    text_norm = normalize_text(method_text)
    source_key = _normalize_source_key(source_name)

    # Ưu tiên cấu hình có giới hạn nguồn trước.
    source_exact = _EXACT_BY_SOURCE.get(source_key, {})
    if text_norm in source_exact:
        nhom, phuong_phap = source_exact[text_norm]
        return nhom, phuong_phap, "Theo tên kỹ thuật - chính xác"

    if text_norm in _EXACT_ALL:
        nhom, phuong_phap = _EXACT_ALL[text_norm]
        return nhom, phuong_phap, "Theo tên kỹ thuật - chính xác"

    for pp_norm, nhom, phuong_phap in _CONTAINS_BY_SOURCE.get(source_key, []):
        if pp_norm and pp_norm in text_norm:
            return nhom, phuong_phap, "Theo tên kỹ thuật - chứa cụm từ"

    for pp_norm, nhom, phuong_phap in _CONTAINS_ALL:
        if pp_norm and pp_norm in text_norm:
            return nhom, phuong_phap, "Theo tên kỹ thuật - chứa cụm từ"

    return CHUA_PHAN_LOAI, "", ""


def make_priority_method_table():
    if not PHAN_LOAI_PHUONG_PHAP_UU_TIEN:
        return pd.DataFrame(columns=["Phương pháp", "Nhóm phân loại", "Kiểu khớp", "Nguồn áp dụng"])

    rows = []
    for item in PHAN_LOAI_PHUONG_PHAP_UU_TIEN:
        rows.append({
            "Phương pháp": item.get("phuong_phap", ""),
            "Nhóm phân loại": item.get("nhom", ""),
            "Kiểu khớp": item.get("kieu_khop", "chinh_xac"),
            "Nguồn áp dụng": item.get("nguon", "Tất cả"),
        })
    return pd.DataFrame(rows)


def make_keyword_table():
    """
    Phiên bản này không dùng keyword nữa.
    Hàm này giữ lại để tránh lỗi nếu module khác còn gọi.
    """
    return pd.DataFrame(columns=["Nhóm phân loại", "Keyword"])
