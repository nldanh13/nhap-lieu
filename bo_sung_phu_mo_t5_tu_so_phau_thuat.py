# -*- coding: utf-8 -*-
"""
Bổ sung PTV chính / phụ mổ cho sheet phauthuat của file PM khoa CTCH T5
từ Sổ Phẫu Thuật.

Chạy file này độc lập, không ảnh hưởng các file thống kê hiện có:
    python bo_sung_phu_mo_t5_tu_so_phau_thuat.py

Luồng xử lý:
1. Đọc sheet phauthuat của file T5.
2. Đọc Sổ Phẫu Thuật.
3. Đối chiếu theo ngày phẫu thuật + họ tên người bệnh + phương pháp phẫu thuật.
4. Chỉ cập nhật PTV chính / phụ mổ khi tìm được ca khớp và nguồn có bí danh.
5. Giữ nguyên dữ liệu đã nhập thủ công nếu ca không khớp hoặc nguồn đang trống.
6. Xuất ra file Excel mới, giữ nguyên file gốc.
"""

from __future__ import annotations

import re
import unicodedata
from collections import defaultdict
from copy import copy
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

try:
    from cau_hinh_file import BASE_DIR
except Exception:
    BASE_DIR = Path(__file__).resolve().parent

try:
    from cau_hinh_khoa import get_khoa_chi_dinh_can_lay
except Exception:
    get_khoa_chi_dinh_can_lay = None

try:
    from danh_sach_nhan_su import BI_DANH_NHAN_SU
except Exception:
    BI_DANH_NHAN_SU = {}

# =========================
# CẤU HÌNH FILE
# =========================

FILE_PM_T5 = BASE_DIR / "2026" / "PM khoa CTCH T5.xlsx"
THU_MUC_THEO_SO = BASE_DIR / "TheoSo"
FILE_SO_PHAU_THUAT = None  # Nếu muốn chỉ định thủ công, điền Path(...) tại đây.

SHEET_PHAUTHUAT_T5 = "phauthuat"
SHEET_BAO_CAO = "Bao_cao_bo_sung_phu_mo"

FILE_DAU_RA = BASE_DIR / "2026" / "PM khoa CTCH T5_da_bo_sung_phu_mo.xlsx"
FILE_BAO_CAO = BASE_DIR / "bao_cao_bo_sung_phu_mo_T5.xlsx"

# Nếu True thì ghi đè trực tiếp lên FILE_PM_T5.
# Khuyến nghị để False để giữ file gốc.
GHI_DE_FILE_GOC = False

# Sổ Phẫu Thuật nên lọc cùng logic đang dùng cho thống kê khoa Ngoại Chấn Thương.
# Mặc định lấy từ cau_hinh_khoa.py: Khoa Cấp Cứu + Khoa Ngoại Chấn Thương.
# Nếu muốn lấy toàn bộ sổ, đổi thành [].
KHOA_CHI_DINH_CAN_LAY = None

# Các cột được phép cập nhật từ Sổ Phẫu Thuật.
# Chương trình không xóa trắng dữ liệu cũ khi không tìm được giá trị nguồn.
COT_T5_CAN_LAM_SACH = ["PTV chính", "Phụ mổ 1", "Phụ mổ 2", "Phụ mổ 3"]

# Ánh xạ cột đích trong file T5 với cột nguồn trong Sổ Phẫu Thuật.
ANH_XA_COT_NHAN_SU = {
    "PTV chính": "Phẫu thuật viên",
    "Phụ mổ 1": "Phụ PTTT1",
    "Phụ mổ 2": "Phụ PTTT2",
    "Phụ mổ 3": "Phụ PTTT3",
}

# Cột dùng để đối chiếu trong file T5.
COT_T5_NGAY = "NGÀY"
COT_T5_HO_TEN = "Họ và tên bệnh nhân"
COT_T5_PHUONG_PHAP = "Chẩn đoán và phương pháp phẫu thuật"

# Cột dùng để đối chiếu trong Sổ Phẫu Thuật.
COT_SO_THOI_GIAN = "Thời gian thực hiện"
COT_SO_HO_TEN = "Họ và tên"
COT_SO_PHUONG_PHAP = "Phương pháp phẫu thuật"
COT_SO_KHOA_CHI_DINH = "Khoa phòng chỉ định"

# Điểm tối thiểu để nhận là khớp khi cùng ngày + cùng họ tên.
# Không đặt quá cao vì file T5 thường viết tắt: KHX, PTNS, PT...
DIEM_KHOP_PHUONG_PHAP_TOI_THIEU = 0.20

# =========================
# TIỆN ÍCH CHUẨN HÓA
# =========================


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and str(value) == "nan":
        return ""

    text = str(value).strip().lower()
    if not text:
        return ""

    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    text = re.sub(r"\s+", " ", text).strip()
    return text


def normalize_header(value: Any) -> str:
    return normalize_text(value).replace(" ", "")


def clean_staff_name(value: Any) -> str:
    """Làm sạch tên nhân sự lấy từ Sổ Phẫu Thuật."""
    if value is None:
        return ""

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none", "-"}:
        return ""

    text = re.sub(r"^\s*[-–—]+\s*", "", text)
    text = re.sub(r"\s+", " ", text).strip()

    if not text or text in {"-", "–", "—"}:
        return ""
    return text


def remove_staff_titles(value: str) -> str:
    """Bỏ học hàm/học vị/chức danh trước tên nhân sự để so khớp bí danh.

    Lý do cần làm kỹ:
    - Sổ Phẫu Thuật thường ghi: "TS BS. Nguyễn Tư Thái Bảo",
      "BS CKI. Nguyễn Giang Tử", "BS CKII. Nguyễn Lê Hoan",
      "Ths. BS. Phạm Việt Tân"...
    - Danh sách bí danh lại lưu theo họ tên thật, ví dụ: "Nguyễn Giang Tử".
    - Nếu chỉ bỏ được 1 tiền tố đầu tiên thì tên còn lại vẫn là
      "BS. Nguyễn..." hoặc "CKI. Nguyễn..." và sẽ không khớp bí danh.
    """
    text = clean_staff_name(value)
    if not text:
        return ""

    # Chuẩn hóa dấu chấm trong cụm chức danh để các dạng
    # "TS.BS.", "TS BS.", "Ths. BS CKI." đều xử lý được.
    text = re.sub(r"\s+", " ", text).strip()

    title_patterns = [
        r"^pgs\s*\.?\s*",
        r"^gs\s*\.?\s*",
        r"^ts\s*\.?\s*",
        r"^ths\s*\.?\s*",
        r"^thạc\s*sĩ\s*\.?\s*",
        r"^bác\s*sĩ\s*\.?\s*",
        r"^bs\s*ck\s*ii\s*\.?\s*",
        r"^bs\s*ck\s*i\s*\.?\s*",
        r"^bsckii\s*\.?\s*",
        r"^bscki\s*\.?\s*",
        r"^bsck2\s*\.?\s*",
        r"^bsck1\s*\.?\s*",
        r"^bs\s*\.?\s*",
        r"^dr\s*\.?\s*",
        r"^cn\s*\.?\s*",
        r"^dd\s*\.?\s*",
        r"^đd\s*\.?\s*",
        r"^ktv\s*\.?\s*",
        r"^phcn\s*\.?\s*",
        r"^phẫu\s*thuật\s*viên\s*\.?\s*",
    ]

    # Bỏ lặp nhiều lần vì có thể có nhiều chức danh liên tiếp, ví dụ "TS BS.".
    changed = True
    while changed:
        changed = False
        before = text
        for pattern in title_patterns:
            text = re.sub(pattern, "", text, flags=re.IGNORECASE).strip()
        if text != before:
            changed = True

    text = re.sub(r"\s+", " ", text).strip(" .,-–—")
    return text


def build_staff_alias_lookup() -> dict[str, str]:
    """Tạo bảng tra cứu: họ tên đầy đủ hoặc bí danh đều trả về bí danh chính."""
    lookup = {}
    for full_name, aliases in BI_DANH_NHAN_SU.items():
        aliases = list(aliases or [])
        if not aliases:
            continue

        primary_alias = str(aliases[0]).strip()
        if not primary_alias:
            continue

        keys = [full_name, *aliases]
        for key in keys:
            key_norm = normalize_text(remove_staff_titles(str(key)))
            if key_norm:
                lookup[key_norm] = primary_alias

    return lookup


STAFF_ALIAS_LOOKUP = build_staff_alias_lookup()


def staff_to_alias(value: Any) -> str:
    """Chuyển tên nhân sự từ sổ về bí danh chính trong BI_DANH_NHAN_SU.

    Nếu không tìm được bí danh, giữ trống để người dùng tự kiểm tra/thêm cấu hình,
    tránh ghi họ tên đầy đủ vào file tháng.
    """
    text = remove_staff_titles(clean_staff_name(value))
    if not text:
        return ""

    key = normalize_text(text)
    if key in STAFF_ALIAS_LOOKUP:
        return STAFF_ALIAS_LOOKUP[key]

    # Một số file có thể nhập dạng "Nguyễn Chí Nguyện - ..." hoặc "BS. Nguyễn Chí Nguyện".
    # So khớp bao chứa để vẫn trả về bí danh, nhưng chỉ khi đủ rõ ràng.
    matched_aliases = []
    for known_key, alias in STAFF_ALIAS_LOOKUP.items():
        if not known_key:
            continue
        if known_key in key or key in known_key:
            matched_aliases.append(alias)

    matched_aliases = list(dict.fromkeys(matched_aliases))
    if len(matched_aliases) == 1:
        return matched_aliases[0]

    return ""


def simplify_method(value: Any) -> str:
    """Chuẩn hóa tên phương pháp để so khớp tốt hơn giữa file T5 và Sổ Phẫu Thuật."""
    text = normalize_text(value)
    if not text:
        return ""

    # Bỏ các phần tỷ lệ 50%, 30%... trong file tổng hợp tháng.
    text = re.sub(r"\b\d+\s*%\b", " ", text)

    replacements = {
        "ptns": "phau thuat noi soi",
        "pttt": "phau thuat thu thuat",
        "pt": "phau thuat",
        "khx": "ket hop xuong",
        "dc khx": "dung cu ket hop xuong",
        "vpm": "vet phan mem",
    }

    # Thay cụm dài trước.
    for old, new in sorted(replacements.items(), key=lambda x: -len(x[0])):
        text = re.sub(rf"\b{re.escape(old)}\b", new, text)

    text = re.sub(r"\s+", " ", text).strip()
    return text


def token_overlap_score(a: str, b: str) -> float:
    tokens_a = {x for x in a.split() if len(x) >= 2}
    tokens_b = {x for x in b.split() if len(x) >= 2}
    if not tokens_a or not tokens_b:
        return 0.0
    return len(tokens_a & tokens_b) / max(len(tokens_a), len(tokens_b))


def method_score(target_method: Any, source_method: Any) -> float:
    a = simplify_method(target_method)
    b = simplify_method(source_method)
    if not a or not b:
        return 0.0

    if a == b:
        return 1.0
    if a in b or b in a:
        return 0.90

    ratio = SequenceMatcher(None, a, b).ratio()
    overlap = token_overlap_score(a, b)
    return max(ratio, overlap)


def excel_serial_to_date(value: int | float) -> date | None:
    try:
        # Excel/Windows date base used by openpyxl.
        return (datetime(1899, 12, 30) + timedelta(days=int(value))).date()
    except Exception:
        return None


def parse_date_t5(value: Any) -> date | None:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value
    if isinstance(value, (int, float)):
        return excel_serial_to_date(value)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return None

    # Sửa lỗi nhập nhầm năm kiểu 07/5/5026 -> 07/5/2026.
    text = re.sub(r"/(5\d{3})$", "/2026", text)

    for fmt in ("%d/%m/%Y", "%d/%m/%y"):
        try:
            return datetime.strptime(text, fmt).date()
        except Exception:
            pass

    # Bắt lại ngày/tháng/năm trong chuỗi nếu có thêm giờ.
    m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", text)
    if m:
        d, mth, y = map(int, m.groups())
        if y >= 3000:
            y = 2026
        try:
            return date(y, mth, d)
        except Exception:
            return None

    return None


def parse_datetime_so(value: Any) -> datetime | None:
    if value is None:
        return None

    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, time.min)

    text = str(value).strip()
    if not text or text.lower() in {"nan", "none"}:
        return None

    # Dạng thường gặp: 11:40 07/05/2026
    m = re.search(r"(\d{1,2}:\d{2})\s+(\d{1,2}/\d{1,2}/\d{4})", text)
    if m:
        raw = f"{m.group(1)} {m.group(2)}"
        try:
            return datetime.strptime(raw, "%H:%M %d/%m/%Y")
        except Exception:
            pass

    # Nếu chỉ có ngày.
    d = parse_date_t5(text)
    if d:
        return datetime.combine(d, time.min)

    return None


def cell_is_blank(value: Any) -> bool:
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text.lower() in {"nan", "none"}


# =========================
# ĐỌC WORKBOOK / HEADER
# =========================


def find_sheet_case_insensitive(wb, wanted_name: str) -> str:
    wanted_key = normalize_header(wanted_name)
    for name in wb.sheetnames:
        if normalize_header(name) == wanted_key:
            return name
    raise ValueError(f"Không tìm thấy sheet '{wanted_name}'. Các sheet hiện có: {wb.sheetnames}")


def find_header_row_and_columns(ws) -> tuple[int, dict[str, int]]:
    for row in ws.iter_rows(min_row=1, max_row=min(ws.max_row, 30)):
        values = [cell.value for cell in row]
        if any(normalize_header(v) == "stt" for v in values):
            header_row = row[0].row
            mapping = {}
            for cell in row:
                if not cell.value:
                    continue
                key = normalize_header(cell.value)
                if key and key not in mapping:
                    mapping[key] = cell.column
            return header_row, mapping
    raise ValueError("Không tìm thấy dòng tiêu đề có cột STT trong sheet phauthuat.")


def require_col(header_map: dict[str, int], col_name: str) -> int:
    key = normalize_header(col_name)
    if key not in header_map:
        raise ValueError(f"Không tìm thấy cột '{col_name}' trong sheet phauthuat.")
    return header_map[key]


def get_cell_by_col_name(ws, row_idx: int, header_map: dict[str, int], col_name: str):
    return ws.cell(row=row_idx, column=require_col(header_map, col_name))


def clone_cell_style(src_cell, dst_cell) -> None:
    if src_cell.has_style:
        dst_cell.font = copy(src_cell.font)
        dst_cell.fill = copy(src_cell.fill)
        dst_cell.border = copy(src_cell.border)
        dst_cell.alignment = copy(src_cell.alignment)
        dst_cell.number_format = src_cell.number_format
        dst_cell.protection = copy(src_cell.protection)


# =========================
# ĐỌC SỔ PHẪU THUẬT
# =========================


def tim_file_so_phau_thuat() -> Path:
    if FILE_SO_PHAU_THUAT:
        path = Path(FILE_SO_PHAU_THUAT)
        if not path.exists():
            raise FileNotFoundError(f"Không tìm thấy FILE_SO_PHAU_THUAT: {path}")
        return path

    candidates = []
    if THU_MUC_THEO_SO.exists():
        candidates.extend(THU_MUC_THEO_SO.glob("*SoPhauThuat*.xlsx"))
        candidates.extend(THU_MUC_THEO_SO.glob("*SoPhauThuat*.xls"))
        candidates.extend(THU_MUC_THEO_SO.glob("*PhauThuat*.xlsx"))
        candidates.extend(THU_MUC_THEO_SO.glob("*PhauThuat*.xls"))

    candidates = [p for p in candidates if not p.name.startswith("~$")]
    candidates = sorted(set(candidates), key=lambda p: p.stat().st_mtime, reverse=True)

    if not candidates:
        raise FileNotFoundError(
            f"Không tìm thấy file Sổ Phẫu Thuật trong thư mục {THU_MUC_THEO_SO}. "
            "Hãy điền FILE_SO_PHAU_THUAT trong file code."
        )

    return candidates[0]


def read_so_phau_thuat_rows(file_path: Path) -> list[dict[str, Any]]:
    """Đọc Sổ Phẫu Thuật bằng openpyxl theo dạng stream để chạy nhanh."""
    wb = load_workbook(file_path, data_only=True, read_only=True)
    ws = wb.active
    max_col = min(ws.max_column or 60, 80)

    # Sổ Phẫu Thuật hiện tại có header 2 dòng ở Excel dòng 6-7.
    # Ta tự tìm dòng có cột TT và Họ và tên để tránh lệ thuộc tuyệt đối vào số dòng.
    header_row_1 = None
    for row_idx, row_values in enumerate(
        ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), max_col=max_col, values_only=True),
        start=1,
    ):
        normalized_values = {normalize_header(v) for v in row_values if v is not None}
        if "tt" in normalized_values and "hovaten" in normalized_values:
            header_row_1 = row_idx
            break

    if header_row_1 is None:
        # Fallback giống his_excel.py: header dòng 6-7.
        header_row_1 = 6

    header_row_2 = header_row_1 + 1
    header_rows = list(
        ws.iter_rows(min_row=header_row_1, max_row=header_row_2, max_col=max_col, values_only=True)
    )
    header_1_values = header_rows[0]
    header_2_values = header_rows[1]

    headers = []
    used = defaultdict(int)

    for col_idx in range(max_col):
        h1 = header_1_values[col_idx] if col_idx < len(header_1_values) else None
        h2 = header_2_values[col_idx] if col_idx < len(header_2_values) else None
        parts = [str(x).strip() for x in (h1, h2) if x is not None and str(x).strip()]

        if not parts:
            name = f"Cot_{col_idx + 1}"
        elif parts[0].lower().startswith("nhân viên") or parts[0].lower().startswith("nhan vien"):
            name = parts[-1]
        elif len(parts) == 1:
            name = parts[0]
        else:
            name = " ".join(parts)

        used[name] += 1
        if used[name] > 1:
            name = f"{name}_{used[name]}"
        headers.append(name)

    rows = []
    for values in ws.iter_rows(min_row=header_row_2 + 1, max_row=ws.max_row, max_col=max_col, values_only=True):
        record = dict(zip(headers, values))

        # Chỉ giữ dòng dữ liệu thật: cột TT là số.
        try:
            float(str(record.get("TT", "")).strip())
        except Exception:
            continue

        rows.append(record)

    wb.close()
    return rows


def khoa_duoc_lay(khoa_value: Any, danh_sach_khoa: list[str]) -> bool:
    if not danh_sach_khoa:
        return True
    khoa_norm = normalize_text(khoa_value)
    return any(normalize_text(k) in khoa_norm for k in danh_sach_khoa)


@dataclass
class SourceRecord:
    source_index: int
    ngay: date
    ho_ten_key: str
    ho_ten: str
    phuong_phap: str
    thoi_gian: datetime | None
    khoa: str
    staff: dict[str, str]


def get_khoa_filter() -> list[str]:
    if KHOA_CHI_DINH_CAN_LAY is not None:
        return list(KHOA_CHI_DINH_CAN_LAY)

    if get_khoa_chi_dinh_can_lay is not None:
        try:
            return list(get_khoa_chi_dinh_can_lay("Sổ Phẫu Thuật") or [])
        except Exception:
            pass

    return ["Khoa Cấp Cứu", "Khoa Ngoại Chấn Thương"]


def build_source_records(rows: list[dict[str, Any]]) -> list[SourceRecord]:
    khoa_filter = get_khoa_filter()
    records = []

    for i, row in enumerate(rows, start=1):
        dt = parse_datetime_so(row.get(COT_SO_THOI_GIAN))
        if not dt:
            continue

        ho_ten = row.get(COT_SO_HO_TEN)
        ho_ten_key = normalize_text(ho_ten)
        if not ho_ten_key:
            continue

        khoa = row.get(COT_SO_KHOA_CHI_DINH)
        if not khoa_duoc_lay(khoa, khoa_filter):
            continue

        staff = {}
        for target_col, source_col in ANH_XA_COT_NHAN_SU.items():
            staff[target_col] = staff_to_alias(row.get(source_col))

        records.append(
            SourceRecord(
                source_index=i,
                ngay=dt.date(),
                ho_ten_key=ho_ten_key,
                ho_ten=str(ho_ten).strip() if ho_ten is not None else "",
                phuong_phap=str(row.get(COT_SO_PHUONG_PHAP) or "").strip(),
                thoi_gian=dt,
                khoa=str(khoa or "").strip(),
                staff=staff,
            )
        )

    return records


# =========================
# ĐỐI CHIẾU VÀ CẬP NHẬT
# =========================


def choose_best_match(
    target_date: date | None,
    target_name: str,
    target_method: Any,
    index_by_date_name: dict[tuple[date, str], list[SourceRecord]],
) -> tuple[SourceRecord | None, float, str]:
    if not target_date:
        return None, 0.0, "Không đọc được ngày"

    name_key = normalize_text(target_name)
    if not name_key:
        return None, 0.0, "Không có họ tên"

    candidates = index_by_date_name.get((target_date, name_key), [])
    if not candidates:
        return None, 0.0, "Không tìm thấy theo ngày + họ tên"

    if len(candidates) == 1:
        candidate = candidates[0]
        score = method_score(target_method, candidate.phuong_phap)
        if score < DIEM_KHOP_PHUONG_PHAP_TOI_THIEU:
            return candidate, score, "Khớp ngày + họ tên, phương pháp khác nhiều"
        return candidate, score, "Khớp"

    scored = []
    for candidate in candidates:
        scored.append((method_score(target_method, candidate.phuong_phap), candidate))
    scored.sort(key=lambda x: x[0], reverse=True)

    best_score, best_candidate = scored[0]
    if best_score < DIEM_KHOP_PHUONG_PHAP_TOI_THIEU:
        return best_candidate, best_score, "Có nhiều dòng cùng ngày + họ tên, phương pháp khác nhiều"

    # Nếu 2 dòng có điểm bằng nhau, vẫn lấy dòng đầu tiên nhưng ghi chú để người dùng kiểm tra.
    if len(scored) >= 2 and abs(scored[0][0] - scored[1][0]) < 0.001:
        return best_candidate, best_score, "Có nhiều dòng gần giống nhau, cần kiểm tra lại"

    return best_candidate, best_score, "Khớp"


def make_report_workbook(report_rows: list[dict[str, Any]], output_file: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = SHEET_BAO_CAO

    headers = [
        "STT file T5",
        "Dòng Excel T5",
        "Ngày T5",
        "Họ tên T5",
        "Phương pháp T5",
        "Trạng thái",
        "Điểm khớp",
        "Thời gian trong sổ",
        "Phương pháp trong sổ",
        "Khoa chỉ định trong sổ",
        "PTV chính",
        "Phụ mổ 1",
        "Phụ mổ 2",
        "Phụ mổ 3",
        "Ghi chú",
    ]
    ws.append(headers)

    for item in report_rows:
        ws.append([item.get(h, "") for h in headers])

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    header_font = Font(bold=True)
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    widths = {
        "A": 12,
        "B": 14,
        "C": 14,
        "D": 28,
        "E": 42,
        "F": 22,
        "G": 10,
        "H": 20,
        "I": 42,
        "J": 34,
        "K": 28,
        "L": 28,
        "M": 28,
        "N": 28,
        "O": 36,
    }
    for col, width in widths.items():
        ws.column_dimensions[col].width = width
    for row in ws.iter_rows():
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)


def bo_sung_phu_mo() -> None:
    if not FILE_PM_T5.exists():
        raise FileNotFoundError(f"Không tìm thấy file T5: {FILE_PM_T5}")

    file_so = tim_file_so_phau_thuat()
    print(f"File T5: {FILE_PM_T5}")
    print(f"Sổ Phẫu Thuật: {file_so}")

    so_rows = read_so_phau_thuat_rows(file_so)
    source_records = build_source_records(so_rows)
    print(f"Số dòng Sổ Phẫu Thuật sau khi lọc: {len(source_records):,}")

    index_by_date_name = defaultdict(list)
    for record in source_records:
        index_by_date_name[(record.ngay, record.ho_ten_key)].append(record)

    wb = load_workbook(FILE_PM_T5)
    sheet_name = find_sheet_case_insensitive(wb, SHEET_PHAUTHUAT_T5)
    ws = wb[sheet_name]

    header_row, header_map = find_header_row_and_columns(ws)

    stt_col = require_col(header_map, "STT")
    ngay_col = require_col(header_map, COT_T5_NGAY)
    ho_ten_col = require_col(header_map, COT_T5_HO_TEN)
    phuong_phap_col = require_col(header_map, COT_T5_PHUONG_PHAP)
    target_cols = {col: require_col(header_map, col) for col in ANH_XA_COT_NHAN_SU}

    for col_name in COT_T5_CAN_LAM_SACH:
        require_col(header_map, col_name)

    report_rows = []
    count_data = 0
    count_matched = 0
    count_filled_any = 0
    count_unmatched = 0
    count_need_check = 0
    count_preserved_manual = 0

    for row_idx in range(header_row + 1, ws.max_row + 1):
        stt_value = ws.cell(row=row_idx, column=stt_col).value
        if cell_is_blank(stt_value):
            continue
        try:
            float(str(stt_value).strip())
        except Exception:
            continue

        count_data += 1
        target_date = parse_date_t5(ws.cell(row=row_idx, column=ngay_col).value)
        target_name = ws.cell(row=row_idx, column=ho_ten_col).value
        target_method = ws.cell(row=row_idx, column=phuong_phap_col).value

        old_staff_values = {
            col_name: ws.cell(row=row_idx, column=target_cols[col_name]).value
            for col_name in COT_T5_CAN_LAM_SACH
        }

        match, score, note = choose_best_match(
            target_date=target_date,
            target_name=str(target_name or ""),
            target_method=target_method,
            index_by_date_name=index_by_date_name,
        )

        status = "Không khớp"
        filled_any = False
        preserved_manual = False

        if match is not None:
            count_matched += 1
            status = "Đã cập nhật"

            for target_col_name, col_idx in target_cols.items():
                staff_value = match.staff.get(target_col_name, "")
                if staff_value:
                    ws.cell(row=row_idx, column=col_idx).value = staff_value
                    filled_any = True
                elif not cell_is_blank(old_staff_values.get(target_col_name)):
                    preserved_manual = True

            if filled_any:
                count_filled_any += 1
            else:
                status = "Khớp nhưng nguồn trống - giữ dữ liệu cũ"

            if "kiểm tra" in note.lower() or "khác nhiều" in note.lower():
                count_need_check += 1
        else:
            count_unmatched += 1
            if any(not cell_is_blank(value) for value in old_staff_values.values()):
                preserved_manual = True
                status = "Không khớp - giữ dữ liệu cũ"

        if preserved_manual:
            count_preserved_manual += 1

        report_rows.append(
            {
                "STT file T5": stt_value,
                "Dòng Excel T5": row_idx,
                "Ngày T5": target_date.strftime("%d/%m/%Y") if target_date else "",
                "Họ tên T5": target_name or "",
                "Phương pháp T5": target_method or "",
                "Trạng thái": status,
                "Điểm khớp": round(score, 3),
                "Thời gian trong sổ": match.thoi_gian.strftime("%H:%M %d/%m/%Y") if match and match.thoi_gian else "",
                "Phương pháp trong sổ": match.phuong_phap if match else "",
                "Khoa chỉ định trong sổ": match.khoa if match else "",
                "PTV chính": match.staff.get("PTV chính", "") if match else "",
                "Phụ mổ 1": match.staff.get("Phụ mổ 1", "") if match else "",
                "Phụ mổ 2": match.staff.get("Phụ mổ 2", "") if match else "",
                "Phụ mổ 3": match.staff.get("Phụ mổ 3", "") if match else "",
                "Ghi chú": note + (
                    " | Không xóa dữ liệu nhân sự cũ khi nguồn trống/không khớp."
                    if preserved_manual
                    else ""
                ),
            }
        )

    output_file = FILE_PM_T5 if GHI_DE_FILE_GOC else FILE_DAU_RA
    output_file.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_file)
    make_report_workbook(report_rows, FILE_BAO_CAO)

    print("\nHOÀN TẤT")
    print(f"- Số dòng phauthuat đã xử lý: {count_data:,}")
    print(f"- Số dòng khớp theo sổ: {count_matched:,}")
    print(f"- Số dòng đã điền được ít nhất 1 bí danh nhân sự: {count_filled_any:,}")
    print(f"- Số dòng không khớp: {count_unmatched:,}")
    print(f"- Số dòng có ít nhất 1 giá trị nhân sự cũ được giữ lại: {count_preserved_manual:,}")
    print(f"- Số dòng cần kiểm tra lại: {count_need_check:,}")
    print(f"- File T5 sau bổ sung: {output_file}")
    print(f"- File báo cáo đối chiếu: {FILE_BAO_CAO}")


if __name__ == "__main__":
    bo_sung_phu_mo()
