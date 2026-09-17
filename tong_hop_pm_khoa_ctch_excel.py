from pathlib import Path

import pandas as pd

from utils_text import normalize_text
from xu_ly_du_lieu import process_dataframe


# File cấu trúc mới: ví dụ "PM khoa CTCH T5.xlsx"
# Đặc điểm:
# - Có sheet: phauthuat, thu thuat, tieuphau
# - Không có sheet: phauthuat_BHYT
# - Có thể có thêm sheet GB, YCBS nhưng không lấy
#
# File này chỉ PHÂN LOẠI THEO KỸ THUẬT, không lọc nhân sự.

SHEET_PM_KHOA_CTCH_CAN_XU_LY = [
    "phauthuat",
    "thu thuat",
    "tieuphau",
]

CAU_HINH_SHEET_PM_KHOA_CTCH = {
    "phauthuat": {
        "method_col": [
            "Chẩn đoán và phương pháp phẫu thuật",
            "Chẩn đoán và phương pháp PT",
            "Phương pháp phẫu thuật",
            "Phương pháp PT",
            "Tên CLS",
            "Tên dịch vụ",
        ],
        "staff_columns": [],
        "allow_keep_without_staff": True,
    },
    "thu thuat": {
        "method_col": [
            "Tên CLS",
            "Tên dịch vụ",
            "PHƯƠNG PHÁP THỦ THUẬT",
            "Phương pháp thủ thuật",
        ],
        "staff_columns": [],
        "allow_keep_without_staff": True,
    },
    "tieuphau": {
        "method_col": [
            "Tên CLS",
            "Tên dịch vụ",
            "Phương pháp PT",
            "Phương pháp phẫu thuật",
        ],
        "staff_columns": [],
        "allow_keep_without_staff": True,
    },
}


def _make_unique_columns(columns):
    final_cols = []
    counts = {}
    for i, col in enumerate(columns, start=1):
        name = str(col).strip() if col is not None else ""
        if not name or name.lower().startswith("nan"):
            name = f"Cot_{i}"
        if name in counts:
            counts[name] += 1
            name = f"{name}_{counts[name]}"
        else:
            counts[name] = 1
        final_cols.append(name)
    return final_cols


def _normalize_sheet_name(value):
    return normalize_text(value).replace(" ", "")


def _normalize_col_name(value):
    return normalize_text(value).replace(" ", "")


def file_label_from_path(file_path):
    return Path(file_path).stem


def get_existing_sheet_names(file_path):
    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names


def la_file_pm_khoa_ctch(file_path):
    """
    Nhận diện file cấu trúc mới.
    Dùng để tách xử lý riêng, không trộn logic vào file TH cũ.
    """
    try:
        sheet_names = get_existing_sheet_names(file_path)
    except Exception:
        return False

    normalized = {_normalize_sheet_name(name) for name in sheet_names}

    co_3_sheet_chinh = {"phauthuat", "thuthuat", "tieuphau"}.issubset(normalized)
    khong_co_bhyt = "phauthuat_bhyt" not in normalized and "phauthuatbhyt" not in normalized
    co_sheet_phu_dac_trung = bool({"gb", "ycbs"} & normalized)

    return co_3_sheet_chinh and (khong_co_bhyt or co_sheet_phu_dac_trung)


def resolve_method_column(df, configured_method_col, sheet_name=""):
    if isinstance(configured_method_col, str):
        candidates = [configured_method_col]
    else:
        candidates = list(configured_method_col or [])

    for col in candidates:
        if col in df.columns:
            return col

    normalized_map = {_normalize_col_name(col): col for col in df.columns}
    for col in candidates:
        key = _normalize_col_name(col)
        if key in normalized_map:
            return normalized_map[key]

    fallback_keywords = [
        "chandoanvaphuongphapphauthuat",
        "chandoanvaphuongphappt",
        "phuongphapphauthuat",
        "phuongphapthuthuat",
        "phuongphappt",
        "tencls",
        "tendichvu",
    ]
    for raw_col in df.columns:
        key = _normalize_col_name(raw_col)
        if any(keyword in key for keyword in fallback_keywords):
            return raw_col

    raise ValueError(
        f"Không tìm thấy cột phương pháp trong sheet {sheet_name}. "
        f"Các cột hiện có: {df.columns.tolist()}. "
        f"Hãy thêm tên cột vào CAU_HINH_SHEET_PM_KHOA_CTCH trong tong_hop_pm_khoa_ctch_excel.py"
    )


def read_pm_sheet(file_path, sheet_name):
    """
    Đọc sheet cấu trúc mới.
    Tự tìm dòng có STT, sau đó lấy vùng dữ liệu từ cột STT đến cột cuối có tiêu đề.
    """
    raw = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

    header_row_idx = None
    stt_col_idx = None

    for idx, row in raw.iterrows():
        for col_idx, value in enumerate(row.tolist()):
            if str(value).strip().upper() == "STT":
                header_row_idx = idx
                stt_col_idx = col_idx
                break
        if header_row_idx is not None:
            break

    if header_row_idx is None:
        raise ValueError(f"Không tìm thấy dòng tiêu đề STT trong sheet {sheet_name}")

    header_values = raw.iloc[header_row_idx].tolist()

    last_col_idx = stt_col_idx
    for i in range(stt_col_idx, len(header_values)):
        value = header_values[i]
        if pd.notna(value) and str(value).strip():
            last_col_idx = i

    headers = _make_unique_columns(header_values[stt_col_idx:last_col_idx + 1])

    data = raw.iloc[header_row_idx + 1:, stt_col_idx:last_col_idx + 1].copy()
    data.columns = headers
    data["_sheet_goc"] = sheet_name
    data["_dong_excel_goc"] = data.index + 1

    # Chỉ giữ dòng dữ liệu thật có STT dạng số.
    data = data[pd.to_numeric(data[headers[0]], errors="coerce").notna()].copy()

    return data.reset_index(drop=True)


def process_pm_khoa_ctch_workbook(file_path, file_label=None):
    file_path = Path(file_path)
    file_label = file_label or file_label_from_path(file_path)

    existing_sheets = get_existing_sheet_names(file_path)
    existing_sheet_map = {_normalize_sheet_name(name): name for name in existing_sheets}

    all_filtered = []
    all_roles = []

    for sheet_name in SHEET_PM_KHOA_CTCH_CAN_XU_LY:
        cfg = CAU_HINH_SHEET_PM_KHOA_CTCH[sheet_name]
        real_sheet_name = existing_sheet_map.get(_normalize_sheet_name(sheet_name))

        if real_sheet_name is None:
            print(f"File {file_path.name} không có sheet {sheet_name} -> bỏ qua")
            continue

        print(f"Đang xử lý file cấu trúc mới: {file_path.name} | sheet: {real_sheet_name}")
        df = read_pm_sheet(file_path, real_sheet_name)
        method_col = resolve_method_column(df, cfg["method_col"], real_sheet_name)
        source_name = f"{file_label} - {sheet_name}"

        df_filtered, df_roles = process_dataframe(
            df=df,
            source_name=source_name,
            staff_columns=cfg["staff_columns"],
            method_col=method_col,
            department_col=None,
            allow_keep_without_staff=cfg.get("allow_keep_without_staff", False),
        )

        for table in (df_filtered, df_roles):
            if table is not None and not table.empty:
                table.insert(0, "File nguồn", file_path.name)
                table.insert(1, "Tháng/File", file_label)
                table.insert(2, "Sheet nguồn", sheet_name)
                table.insert(3, "Kiểu file", "PM khoa CTCH")

        all_filtered.append(df_filtered)
        all_roles.append(df_roles)

    if not all_filtered:
        empty = pd.DataFrame()
        empty.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()
        return empty, pd.DataFrame()

    df_filtered = pd.concat(all_filtered, ignore_index=True)
    df_roles = pd.concat(all_roles, ignore_index=True)
    df_filtered.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()

    return df_filtered, df_roles
