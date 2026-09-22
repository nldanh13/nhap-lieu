from pathlib import Path

import pandas as pd

from cau_hinh_file import LOAI_TRUNG_PHAUTHUAT_BHYT
from utils_text import normalize_text
from xu_ly_du_lieu import process_dataframe


# Cấu hình riêng cho các file "Phau thuat-Thu thuat-Tieu thu thuat" theo tháng.
# File này chỉ PHÂN LOẠI THEO KỸ THUẬT, không lọc/không thống kê theo nhân sự.
#
# Bản này hỗ trợ cả 2 kiểu file:
# 1) File cũ có sheet phauthuat_BHYT.
# 2) File mới KHÔNG có sheet phauthuat_BHYT.
# Nếu sheet không tồn tại, chương trình tự bỏ qua, không báo lỗi.

CAU_HINH_SHEET_TONG_HOP_THANG = {
    "phauthuat": {
        "method_col": [
            "Phương pháp phẫu thuật",
            "Chẩn đoán và phương pháp phẫu thuật",
            "Chẩn đoán và phương pháp PT",
            "Phương pháp PT",
            "Tên CLS",
        ],
        "staff_columns": [],
        "allow_keep_without_staff": True,
    },
    "phauthuat_BHYT": {
        "method_col": [
            "Phương pháp phẫu thuật",
            "Chẩn đoán và phương pháp phẫu thuật",
            "Chẩn đoán và phương pháp PT",
            "Phương pháp PT",
            "Tên CLS",
        ],
        "staff_columns": [],
        "allow_keep_without_staff": True,
    },
    "thu thuat": {
        "method_col": [
            "PHƯƠNG PHÁP THỦ THUẬT",
            "Phương pháp thủ thuật",
            "Tên CLS",
            "Tên dịch vụ",
        ],
        "staff_columns": [],
        "allow_keep_without_staff": True,
    },
    "tieuphau": {
        "method_col": [
            "Phương pháp PT",
            "Tên CLS",
            "Tên dịch vụ",
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


def file_label_from_path(file_path):
    """Lấy nhãn tháng/file để đưa vào báo cáo."""
    return Path(file_path).stem


def _normalize_col_name(value):
    return normalize_text(value).replace(" ", "")


def resolve_method_column(df, configured_method_col, sheet_name=""):
    """
    Tìm cột phương pháp theo cấu hình.
    configured_method_col có thể là string hoặc list string.
    So khớp 2 lớp:
    1. Đúng tên cột.
    2. Chuẩn hóa bỏ dấu/hoa thường/khoảng trắng.
    """
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

    # Dự phòng: nếu tên cột có chứa các cụm quen thuộc.
    fallback_keywords = [
        "phuongphapphauthuat",
        "phuongphapthuthuat",
        "phuongphappt",
        "chandoanvaphuongphapphauthuat",
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
        f"Hãy thêm tên cột vào CAU_HINH_SHEET_TONG_HOP_THANG trong tong_hop_thang_excel.py"
    )


def get_existing_sheet_names(file_path):
    """Lấy danh sách sheet thực tế trong file Excel."""
    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names


def read_summary_sheet(file_path, sheet_name):
    """
    Đọc một sheet tổng hợp tháng có header nằm rải rác ở dòng khác nhau.
    Hàm tự tìm dòng có STT rồi lấy từ cột STT đến cột cuối có header.
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

    # Lấy từ cột STT đến cột cuối có tên tiêu đề.
    last_col_idx = stt_col_idx
    for i in range(stt_col_idx, len(header_values)):
        value = header_values[i]
        if pd.notna(value) and str(value).strip():
            last_col_idx = i

    headers = header_values[stt_col_idx:last_col_idx + 1]
    headers = _make_unique_columns(headers)

    data = raw.iloc[header_row_idx + 1:, stt_col_idx:last_col_idx + 1].copy()
    data.columns = headers
    data["_sheet_goc"] = sheet_name
    data["_dong_excel_goc"] = data.index + 1

    # Chỉ giữ dòng dữ liệu thật có STT dạng số.
    data = data[pd.to_numeric(data[headers[0]], errors="coerce").notna()].copy()

    return data.reset_index(drop=True)


def _first_not_empty(row, column_names):
    """Lấy giá trị đầu tiên không trống trong danh sách cột."""
    for col in column_names:
        if col in row.index:
            value = row[col]
            if pd.notna(value) and str(value).strip():
                return value
    return ""


def _make_duplicate_key(row):
    """Khóa kiểm tra trùng giữa phauthuat và phauthuat_BHYT trong cùng một file tháng."""
    file_nguon = _first_not_empty(row, ["File nguồn"])
    ngay = _first_not_empty(row, ["Ngày", "NGÀY", "Ngay"])
    ho_ten = _first_not_empty(row, ["Họ Tên", "Họ Tên ", "Họ và tên", "HỌ TÊN", "Ho ten", "Họ và tên bệnh nhân"])
    phuong_phap = _first_not_empty(row, [
        "Phương pháp dùng phân loại",
        "Phương pháp phẫu thuật",
        "Chẩn đoán và phương pháp phẫu thuật",
        "Tên CLS",
    ])
    return "||".join([
        normalize_text(file_nguon),
        normalize_text(ngay),
        normalize_text(ho_ten),
        normalize_text(phuong_phap),
    ])


def _row_display_info(row):
    return {
        "File nguồn": _first_not_empty(row, ["File nguồn"]),
        "Tháng/File": _first_not_empty(row, ["Tháng/File"]),
        "Sheet nguồn": _first_not_empty(row, ["Sheet nguồn"]),
        "Ngày": _first_not_empty(row, ["Ngày", "NGÀY", "Ngay"]),
        "Họ tên": _first_not_empty(row, ["Họ Tên", "Họ Tên ", "Họ và tên", "HỌ TÊN", "Ho ten", "Họ và tên bệnh nhân"]),
        "Phương pháp": _first_not_empty(row, [
            "Phương pháp dùng phân loại",
            "Phương pháp phẫu thuật",
            "Chẩn đoán và phương pháp phẫu thuật",
            "Tên CLS",
        ]),
        "STT": _first_not_empty(row, ["STT"]),
        "Dòng Excel": _first_not_empty(row, ["_dong_excel_goc"]),
        "PTV chính": _first_not_empty(row, ["PTV chính"]),
        "Phụ 1": _first_not_empty(row, ["Phụ mổ 1"]),
        "Phụ 2": _first_not_empty(row, ["Phụ mổ 2"]),
        "Phụ 3": _first_not_empty(row, ["Phụ mổ 3"]),
    }


def _sheet_mask(df, sheet_name):
    if "Sheet nguồn" in df.columns:
        return df["Sheet nguồn"].eq(sheet_name)
    if "Nguồn" in df.columns:
        return df["Nguồn"].astype(str).str.endswith(f"- {sheet_name}")
    return pd.Series(False, index=df.index)


def xu_ly_trung_phauthuat_bhyt(df_filtered):
    """
    Kiểm tra trùng giữa 2 sheet phauthuat và phauthuat_BHYT.
    Nếu file không có phauthuat_BHYT thì tự bỏ qua phần kiểm tra trùng.
    """
    if df_filtered.empty:
        df_filtered.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()
        return df_filtered

    df = df_filtered.copy()
    mask_pt = _sheet_mask(df, "phauthuat")
    mask_bhyt = _sheet_mask(df, "phauthuat_BHYT")

    if not mask_pt.any() or not mask_bhyt.any():
        df.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()
        return df

    df["Khóa trùng PT-BHYT"] = df.apply(_make_duplicate_key, axis=1)
    df["Tình trạng trùng PT-BHYT"] = ""

    pt_map = {}
    for idx, row in df.loc[mask_pt].iterrows():
        key = row["Khóa trùng PT-BHYT"]
        if key:
            pt_map.setdefault(key, []).append((idx, row))

    duplicate_report_rows = []
    bhyt_duplicate_indexes = []

    for idx_bhyt, row_bhyt in df.loc[mask_bhyt].iterrows():
        key = row_bhyt["Khóa trùng PT-BHYT"]
        if key not in pt_map:
            continue

        for idx_pt, row_pt in pt_map[key]:
            info_pt = _row_display_info(row_pt)
            info_bhyt = _row_display_info(row_bhyt)
            duplicate_report_rows.append({
                "File nguồn": info_pt["File nguồn"],
                "Tháng/File": info_pt["Tháng/File"],
                "Ngày": info_pt["Ngày"],
                "Họ tên": info_pt["Họ tên"],
                "Phương pháp": info_pt["Phương pháp"],
                "Dòng phauthuat": info_pt["Dòng Excel"],
                "STT phauthuat": info_pt["STT"],
                "PTV chính phauthuat": info_pt["PTV chính"],
                "Phụ 1 phauthuat": info_pt["Phụ 1"],
                "Phụ 2 phauthuat": info_pt["Phụ 2"],
                "Phụ 3 phauthuat": info_pt["Phụ 3"],
                "Dòng phauthuat_BHYT": info_bhyt["Dòng Excel"],
                "STT BHYT": info_bhyt["STT"],
                "PTV chính BHYT": info_bhyt["PTV chính"],
                "Phụ 1 BHYT": info_bhyt["Phụ 1"],
                "Phụ 2 BHYT": info_bhyt["Phụ 2"],
                "Đề xuất xử lý": "Giữ phauthuat; loại dòng trùng ở phauthuat_BHYT khi thống kê tổng",
            })
            df.at[idx_pt, "Tình trạng trùng PT-BHYT"] = "Có dòng trùng ở phauthuat_BHYT"
            df.at[idx_bhyt, "Tình trạng trùng PT-BHYT"] = "Trùng với phauthuat"
            bhyt_duplicate_indexes.append(idx_bhyt)

    duplicate_report = pd.DataFrame(duplicate_report_rows)

    if LOAI_TRUNG_PHAUTHUAT_BHYT and bhyt_duplicate_indexes:
        df = df.drop(index=sorted(set(bhyt_duplicate_indexes))).copy()

    df.attrs["trung_lap_phauthuat_bhyt"] = duplicate_report
    return df


def process_summary_workbook(file_path, sheet_names, file_label=None):
    file_path = Path(file_path)
    file_label = file_label or file_label_from_path(file_path)

    existing_sheets = get_existing_sheet_names(file_path)
    existing_sheet_map = {normalize_text(name): name for name in existing_sheets}

    all_filtered = []
    all_roles = []

    for sheet_name in sheet_names:
        if sheet_name not in CAU_HINH_SHEET_TONG_HOP_THANG:
            print(f"Bỏ qua sheet chưa cấu hình: {sheet_name}")
            continue

        real_sheet_name = existing_sheet_map.get(normalize_text(sheet_name))
        if real_sheet_name is None:
            print(f"File {file_path.name} không có sheet {sheet_name} -> bỏ qua")
            continue

        cfg = CAU_HINH_SHEET_TONG_HOP_THANG[sheet_name]
        print(f"Đang xử lý file: {file_path.name} | sheet: {real_sheet_name}")
        df = read_summary_sheet(file_path, real_sheet_name)

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

        all_filtered.append(df_filtered)
        all_roles.append(df_roles)

    if not all_filtered:
        empty = pd.DataFrame()
        empty.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()
        return empty, pd.DataFrame()

    df_filtered = pd.concat(all_filtered, ignore_index=True)
    df_roles = pd.concat(all_roles, ignore_index=True)

    df_filtered = xu_ly_trung_phauthuat_bhyt(df_filtered)

    return df_filtered, df_roles


def process_one_summary_workbook(file_path, sheet_names):
    """
    Chọn đúng bộ đọc theo cấu trúc file.
    - File TH cũ: đọc bằng process_summary_workbook.
    - File PM khoa CTCH cấu trúc mới: đọc bằng tong_hop_pm_khoa_ctch_excel.py.
    """
    try:
        from tong_hop_pm_khoa_ctch_excel import la_file_pm_khoa_ctch, process_pm_khoa_ctch_workbook
    except ImportError:
        la_file_pm_khoa_ctch = None
        process_pm_khoa_ctch_workbook = None

    if la_file_pm_khoa_ctch and la_file_pm_khoa_ctch(file_path):
        return process_pm_khoa_ctch_workbook(file_path)

    return process_summary_workbook(file_path, sheet_names)


def process_multiple_summary_workbooks(file_paths, sheet_names):
    all_filtered = []
    all_roles = []
    duplicate_reports = []

    for file_path in file_paths:
        df_filtered, df_roles = process_one_summary_workbook(file_path, sheet_names)

        report = df_filtered.attrs.get("trung_lap_phauthuat_bhyt")

        # Xóa attrs trước khi concat để tránh pandas so sánh DataFrame trong attrs.
        df_filtered_clean = df_filtered.copy()
        df_filtered_clean.attrs = {}
        df_roles_clean = df_roles.copy()
        df_roles_clean.attrs = {}

        all_filtered.append(df_filtered_clean)
        all_roles.append(df_roles_clean)

        if report is not None and not report.empty:
            duplicate_reports.append(report)

    if not all_filtered:
        empty = pd.DataFrame()
        empty.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()
        return empty, pd.DataFrame()

    df_all = pd.concat(all_filtered, ignore_index=True)
    roles_all = pd.concat(all_roles, ignore_index=True)

    if duplicate_reports:
        df_all.attrs["trung_lap_phauthuat_bhyt"] = pd.concat(duplicate_reports, ignore_index=True)
    else:
        df_all.attrs["trung_lap_phauthuat_bhyt"] = pd.DataFrame()

    return df_all, roles_all


def main():
    """Cho phép chạy trực tiếp: python tong_hop_thang_excel.py"""
    from cau_hinh_file import (
        OUTPUT_FILE,
        SHEET_TONG_HOP_THANG_CAN_XU_LY,
        THU_MUC_TONG_HOP_THANG,
        get_file_tong_hop_thang_list,
    )
    from xu_ly_du_lieu import make_summary
    from xuat_excel import export_excel

    file_paths = get_file_tong_hop_thang_list()

    if not file_paths:
        raise ValueError(
            "Không tìm thấy file tháng để xử lý. "
            f"Hãy bỏ file Excel vào thư mục '{THU_MUC_TONG_HOP_THANG}' "
            "hoặc khai báo DANH_SACH_FILE_TONG_HOP_THANG trong cau_hinh_file.py"
        )

    print("Đang chạy chế độ nhiều file tháng - chỉ phân loại kỹ thuật")
    print("Số file đầu vào:", len(file_paths))
    for item in file_paths:
        print("-", item)
    print("File đầu ra:", OUTPUT_FILE)

    df_filtered, df_roles = process_multiple_summary_workbooks(
        file_paths,
        SHEET_TONG_HOP_THANG_CAN_XU_LY,
    )

    if df_filtered.empty:
        raise ValueError("Không có dòng dữ liệu nào được xử lý. Hãy kiểm tra tên file/sheet trong cau_hinh_file.py")

    summaries = make_summary(df_filtered, df_roles)
    export_excel(OUTPUT_FILE, summaries, df_filtered, df_roles)

    duplicate_report = df_filtered.attrs.get("trung_lap_phauthuat_bhyt")
    so_dong_trung = 0 if duplicate_report is None or duplicate_report.empty else len(duplicate_report)

    print("Đã xuất file:", OUTPUT_FILE)
    print("Tổng dòng phân loại:", len(df_filtered))
    print("Số dòng chưa phân loại:", (df_filtered["Nhóm phân loại"] == "Chưa phân loại").sum())
    print("Số dòng trùng PT-BHYT đã phát hiện:", so_dong_trung)


if __name__ == "__main__":
    main()
