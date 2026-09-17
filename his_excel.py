from pathlib import Path

import pandas as pd

from cau_hinh_cot import COT_KHOA_CHI_DINH_CO_THE, COT_PHUONG_PHAP_CO_THE
from utils_text import clean_header_value, normalize_text


def flatten_his_columns(columns):
    """
    Gộp header 2 dòng của file HIS thành tên cột dễ xử lý.

    Ví dụ:
    ('Nhân viên thủ thuật', 'TT chính') -> 'TT chính'
    ('Phương pháp thủ thuật', '') -> 'Phương pháp thủ thuật'
    """
    result = []

    for col in columns:
        if isinstance(col, tuple):
            parts = [clean_header_value(x) for x in col]
            parts = [p for p in parts if p]

            if not parts:
                name = ""
            elif parts[0].lower().startswith("nhân viên"):
                name = parts[-1]
            elif len(parts) == 1:
                name = parts[0]
            else:
                name = " ".join(parts)
        else:
            name = clean_header_value(col)

        result.append(name)

    final_cols = []
    counts = {}

    for i, name in enumerate(result, start=1):
        if not name:
            name = f"Cot_{i}"

        if name in counts:
            counts[name] += 1
            name = f"{name}_{counts[name]}"
        else:
            counts[name] = 1

        final_cols.append(name)

    return final_cols


def read_his_excel(file_path, source_name):
    """
    Đọc đúng header của 2 loại file HIS.

    Sổ Thủ Thuật:
      - Header dòng Excel 5-6
      - Bỏ dòng đánh số 7 và dòng trống 8

    Sổ Phẫu Thuật:
      - Header dòng Excel 6-7
      - Bỏ dòng đánh số 8 và dòng trống 9
    """
    file_path = Path(file_path)

    if not file_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    if source_name == "Sổ Thủ Thuật":
        df = pd.read_excel(file_path, header=[4, 5], skiprows=[6, 7])
    elif source_name == "Sổ Phẫu Thuật":
        df = pd.read_excel(file_path, header=[5, 6], skiprows=[7, 8])
    else:
        df = pd.read_excel(file_path)

    df.columns = flatten_his_columns(df.columns)

    # Chỉ giữ dòng dữ liệu thật có STT dạng số.
    if "TT" in df.columns:
        df = df[pd.to_numeric(df["TT"], errors="coerce").notna()].copy()

    return df


def find_column(df, possible_names):
    col_map = {normalize_text(c): c for c in df.columns}

    for name in possible_names:
        key = normalize_text(name)
        if key in col_map:
            return col_map[key]

    return None


def find_method_column(df):
    col = find_column(df, COT_PHUONG_PHAP_CO_THE)
    if col:
        return col

    print("Các cột hiện có:")
    print(df.columns.tolist())

    raise ValueError(
        "Không tìm thấy cột phương pháp. "
        "Hãy kiểm tra tên cột trong file Excel và thêm vào COT_PHUONG_PHAP_CO_THE."
    )


def find_department_column(df):
    return find_column(df, COT_KHOA_CHI_DINH_CO_THE)


def resolve_existing_columns(df, wanted_columns):
    col_map = {normalize_text(c): c for c in df.columns}
    found = []

    for name in wanted_columns:
        key = normalize_text(name)
        if key in col_map and col_map[key] not in found:
            found.append(col_map[key])

    return found
