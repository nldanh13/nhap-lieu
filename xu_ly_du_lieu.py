import pandas as pd

from cau_hinh_khoa import (
    KHOA_CHI_DINH_CAN_LAY_MAC_DINH,
    KHOA_CHI_DINH_THEO_NGUON,
    get_khoa_chi_dinh_can_lay,
)
from danh_sach_nhan_su import BAC_SI, DIEU_DUONG, get_bi_danh

try:
    from danh_sach_nhan_su import KTV
except ImportError:
    KTV = []
from his_excel import find_department_column, find_method_column, read_his_excel, resolve_existing_columns
from phan_loai import classify_method, make_priority_method_table
from utils_text import contains_alias_text, contains_text, normalize_text


def contains_name(cell_value, name):
    return contains_text(cell_value, name)


def get_matching_people(row, people_list, columns):
    matched = []

    for person in people_list:
        for col in columns:
            if col in row.index and contains_name(row[col], person):
                matched.append(person)
                break

    return matched


def get_matching_roles(row, people_list, columns):
    records = []

    for person in people_list:
        for col in columns:
            if col in row.index and contains_name(row[col], person):
                records.append({
                    "Nhân sự": person,
                    "Vai trò/Cột": col,
                    "Giá trị trong ô": row[col],
                })

    return records


def department_allowed(value, allowed_departments):
    if not allowed_departments:
        return True

    value_norm = normalize_text(value)
    allowed_norm = [normalize_text(x) for x in allowed_departments]

    return any(x and x in value_norm for x in allowed_norm)


def _scan_staff_by_full_name_and_alias(df, dept_mask, existing_staff_columns, use_alias=True):
    """Quét nhân sự theo họ tên đầy đủ và tùy chọn bí danh viết tắt."""
    normalized_staff_cols = {
        col: df[col].fillna("").map(normalize_text)
        for col in existing_staff_columns
    }

    matched_bs_by_idx = {idx: set() for idx in df.index}
    matched_dd_by_idx = {idx: set() for idx in df.index}
    matched_ktv_by_idx = {idx: set() for idx in df.index}
    role_rows = []

    def scan_people(people_list, group_name, matched_map):
        for person in people_list:
            person_norm = normalize_text(person)
            aliases = get_bi_danh(person) if use_alias else []

            for col in existing_staff_columns:
                # 1) So khớp họ tên đầy đủ, không phân biệt dấu/hoa thường.
                if person_norm:
                    mask_full_name = normalized_staff_cols[col].str.contains(person_norm, regex=False, na=False) & dept_mask
                else:
                    mask_full_name = False

                hit_indexes = set(df.index[mask_full_name]) if person_norm else set()

                # 2) So khớp bí danh viết tắt, GIỮ DẤU để tránh nhầm TẤN với V.TÂN.
                for alias in aliases:
                    mask_alias = df[col].fillna("").map(lambda value: contains_alias_text(value, alias)) & dept_mask
                    hit_indexes.update(df.index[mask_alias])

                for idx in hit_indexes:
                    matched_map[idx].add(person)
                    role_rows.append({
                        "Nguồn": "",
                        "Nhóm nhân sự": group_name,
                        "Nhân sự": person,
                        "Vai trò/Cột": col,
                        "Giá trị trong ô": df.at[idx, col],
                        "Khoa chỉ định": "",
                        "_index_goc": idx,
                    })

    scan_people(BAC_SI, "Bác sĩ", matched_bs_by_idx)
    scan_people(DIEU_DUONG, "Điều dưỡng", matched_dd_by_idx)
    scan_people(KTV, "KTV", matched_ktv_by_idx)

    return matched_bs_by_idx, matched_dd_by_idx, matched_ktv_by_idx, role_rows


def process_dataframe(
    df,
    source_name,
    staff_columns,
    method_col=None,
    department_col=None,
    allow_keep_without_staff=False,
    use_alias=True,
):
    """Xử lý một DataFrame đã đọc sẵn từ Excel."""
    if method_col is None:
        method_col = find_method_column(df)

    if department_col is None:
        department_col = find_department_column(df)

    existing_staff_columns = resolve_existing_columns(df, staff_columns)

    if not existing_staff_columns and not allow_keep_without_staff:
        print("Các cột hiện có trong file:")
        print(df.columns.tolist())
        raise ValueError(
            f"Không tìm thấy cột nhân sự cho {source_name}. "
            f"Hãy kiểm tra lại cấu hình cột."
        )

    # Lọc khoa chỉ định theo từng nguồn nếu người dùng có cấu hình.
    allowed_departments = get_khoa_chi_dinh_can_lay(source_name)

    if allowed_departments:
        if not department_col:
            raise ValueError(
                f"Đã cấu hình lọc khoa cho {source_name}, nhưng không tìm thấy cột khoa chỉ định. "
                f"Hãy kiểm tra cau_hinh_cot.py -> COT_KHOA_CHI_DINH_CO_THE."
            )
        dept_mask = df[department_col].map(lambda value: department_allowed(value, allowed_departments))
    else:
        dept_mask = pd.Series(True, index=df.index)

    if existing_staff_columns:
        matched_bs_by_idx, matched_dd_by_idx, matched_ktv_by_idx, role_rows = _scan_staff_by_full_name_and_alias(
            df,
            dept_mask,
            existing_staff_columns,
            use_alias=use_alias,
        )
    else:
        matched_bs_by_idx = {idx: set() for idx in df.index}
        matched_dd_by_idx = {idx: set() for idx in df.index}
        matched_ktv_by_idx = {idx: set() for idx in df.index}
        role_rows = []

    for item in role_rows:
        item["Nguồn"] = source_name
        item["Khoa chỉ định"] = df.at[item["_index_goc"], department_col] if department_col else ""

    if existing_staff_columns:
        keep_indexes = [
            idx for idx in df.index
            if dept_mask.loc[idx] and (matched_bs_by_idx[idx] or matched_dd_by_idx[idx] or matched_ktv_by_idx[idx])
        ]
    else:
        # Một số sheet tổng hợp tháng, ví dụ "thu thuat", không có cột nhân sự.
        # Khi allow_keep_without_staff=True, vẫn lấy toàn bộ dòng dữ liệu để thống kê theo nhóm kỹ thuật.
        keep_indexes = [idx for idx in df.index if dept_mask.loc[idx]]

    filtered_rows = []
    category_by_idx = {}
    method_by_idx = {}

    for idx in keep_indexes:
        row = df.loc[idx]
        method_text = row[method_col]
        category, matched_keyword, classify_type = classify_method(method_text, source_name)
        category_by_idx[idx] = category
        method_by_idx[idx] = method_text

        matched_bs = sorted(matched_bs_by_idx[idx])
        matched_dd = sorted(matched_dd_by_idx[idx])
        matched_ktv = sorted(matched_ktv_by_idx[idx])

        new_row = row.to_dict()
        new_row["Nguồn"] = source_name
        new_row["Cột phương pháp"] = method_col
        new_row["Phương pháp dùng phân loại"] = method_text
        new_row["Nhóm phân loại"] = category
        new_row["Keyword/Phương pháp khớp"] = matched_keyword
        new_row["Kiểu phân loại"] = classify_type
        new_row["Bác sĩ khớp"] = ", ".join(matched_bs)
        new_row["Điều dưỡng khớp"] = ", ".join(matched_dd)
        new_row["KTV khớp"] = ", ".join(matched_ktv)
        new_row["Có Bác sĩ"] = "Có" if matched_bs else "Không"
        new_row["Có Điều dưỡng"] = "Có" if matched_dd else "Không"
        new_row["Có KTV"] = "Có" if matched_ktv else "Không"
        new_row["Kiểu dò nhân sự"] = "Họ tên đầy đủ + bí danh" if use_alias else "Chỉ họ tên đầy đủ"

        if department_col:
            new_row["Cột khoa chỉ định"] = department_col
            new_row["Khoa chỉ định dùng lọc"] = row[department_col]
        else:
            new_row["Cột khoa chỉ định"] = ""
            new_row["Khoa chỉ định dùng lọc"] = ""

        if not existing_staff_columns:
            new_row["Ghi chú nhân sự"] = "Sheet không có cột nhân sự; lấy toàn bộ dòng dữ liệu"

        filtered_rows.append(new_row)

    # Gắn thêm thông tin phương pháp/nhóm cho bảng role.
    role_rows_final = []
    keep_set = set(keep_indexes)
    for item in role_rows:
        idx = item.pop("_index_goc")
        if idx not in keep_set:
            continue
        item["Nhóm phân loại"] = category_by_idx.get(idx, "")
        item["Phương pháp"] = method_by_idx.get(idx, "")
        role_rows_final.append(item)

    return pd.DataFrame(filtered_rows), pd.DataFrame(role_rows_final)


def process_file(file_path, source_name, staff_columns, use_alias=True):
    print(f"Đang xử lý: {file_path}")
    df = read_his_excel(file_path, source_name)
    return process_dataframe(df, source_name, staff_columns, use_alias=use_alias)

def safe_groupby_count(df, cols, value_name):
    if df.empty:
        return pd.DataFrame(columns=cols + [value_name])

    return (
        df.groupby(cols, dropna=False)
        .size()
        .reset_index(name=value_name)
        .sort_values(value_name, ascending=False)
    )


def make_summary(df_filtered, df_roles):
    duplicate_report = df_filtered.attrs.get("trung_lap_phauthuat_bhyt") if hasattr(df_filtered, "attrs") else None
    so_dong_trung_pt_bhyt = 0 if duplicate_report is None or duplicate_report.empty else len(duplicate_report)

    if df_filtered.empty:
        summary_overview = pd.DataFrame([
            ["Tổng dòng lọc được", 0],
            ["Số dòng chưa phân loại", 0],
            ["Dòng trùng phauthuat_BHYT đã phát hiện", so_dong_trung_pt_bhyt],
        ], columns=["Nội dung", "Số lượng"])
    else:
        summary_overview = pd.DataFrame([
            ["Tổng dòng lọc được", len(df_filtered)],
            ["Số dòng chưa phân loại", (df_filtered["Nhóm phân loại"] == "Chưa phân loại").sum()],
            ["Dòng trùng phauthuat_BHYT đã phát hiện", so_dong_trung_pt_bhyt],
        ], columns=["Nội dung", "Số lượng"])

    summary_by_source = safe_groupby_count(df_filtered, ["Nguồn"], "Số dòng")

    if "File nguồn" in df_filtered.columns:
        summary_by_file = safe_groupby_count(df_filtered, ["File nguồn"], "Số dòng")
        summary_by_file_category = safe_groupby_count(df_filtered, ["File nguồn", "Nhóm phân loại"], "Số dòng")
    else:
        summary_by_file = pd.DataFrame(columns=["File nguồn", "Số dòng"])
        summary_by_file_category = pd.DataFrame(columns=["File nguồn", "Nhóm phân loại", "Số dòng"])

    summary_by_category = safe_groupby_count(df_filtered, ["Nguồn", "Nhóm phân loại"], "Số dòng")

    # Các bảng dưới đây vẫn được tạo để có thể bật xuất sheet trong cau_hinh_file.py.
    summary_by_staff = safe_groupby_count(df_roles, ["Nguồn", "Nhóm nhân sự", "Nhân sự"], "Số lượt")
    summary_by_staff_category = safe_groupby_count(
        df_roles,
        ["Nguồn", "Nhóm nhân sự", "Nhân sự", "Nhóm phân loại"],
        "Số lượt",
    )
    summary_by_role = safe_groupby_count(df_roles, ["Nguồn", "Nhóm nhân sự", "Vai trò/Cột"], "Số lượt")

    if "Khoa chỉ định dùng lọc" in df_filtered.columns:
        summary_by_department = safe_groupby_count(
            df_filtered,
            ["Nguồn", "Khoa chỉ định dùng lọc", "Nhóm phân loại"],
            "Số dòng",
        )
    else:
        summary_by_department = pd.DataFrame(columns=["Nguồn", "Khoa chỉ định dùng lọc", "Nhóm phân loại", "Số dòng"])

    if not df_filtered.empty:
        unclassified = df_filtered[df_filtered["Nhóm phân loại"] == "Chưa phân loại"].copy()
        if not unclassified.empty:
            sample_unclassified = (
                unclassified["Phương pháp dùng phân loại"]
                .dropna()
                .astype(str)
                .value_counts()
                .reset_index()
            )
            sample_unclassified.columns = ["Phương pháp chưa phân loại", "Số lần"]
            sample_unclassified["Gợi ý thêm vào file"] = (
                "cau_hinh_phuong_phap.py -> PHAN_LOAI_PHUONG_PHAP_UU_TIEN"
            )
        else:
            sample_unclassified = pd.DataFrame(columns=["Phương pháp chưa phân loại", "Số lần", "Gợi ý thêm vào file"])
    else:
        sample_unclassified = pd.DataFrame(columns=["Phương pháp chưa phân loại", "Số lần", "Gợi ý thêm vào file"])

    khoa_rows = []
    for nguon, ds_khoa in KHOA_CHI_DINH_THEO_NGUON.items():
        if ds_khoa:
            for khoa in ds_khoa:
                khoa_rows.append({
                    "Nguồn": nguon,
                    "Khoa chỉ định cần lấy": khoa,
                    "Ghi chú": "So khớp chứa cụm từ, không phân biệt dấu/hoa thường",
                })
        else:
            khoa_rows.append({
                "Nguồn": nguon,
                "Khoa chỉ định cần lấy": "Lấy tất cả khoa",
                "Ghi chú": "Danh sách đang để trống",
            })

    if KHOA_CHI_DINH_CAN_LAY_MAC_DINH:
        for khoa in KHOA_CHI_DINH_CAN_LAY_MAC_DINH:
            khoa_rows.append({
                "Nguồn": "Nguồn khác",
                "Khoa chỉ định cần lấy": khoa,
                "Ghi chú": "Cấu hình mặc định",
            })

    if not khoa_rows:
        khoa_rows.append({
            "Nguồn": "Tất cả",
            "Khoa chỉ định cần lấy": "Đang để trống: lấy tất cả khoa",
            "Ghi chú": "",
        })

    khoa_table = pd.DataFrame(khoa_rows)

    return {
        "Tổng quan": summary_overview,
        "Theo file": summary_by_file,
        "Theo file - nhóm": summary_by_file_category,
        "Theo nguồn": summary_by_source,
        "Theo nhóm kỹ thuật": summary_by_category,
        "Theo khoa chỉ định": summary_by_department,
        "Chưa phân loại": sample_unclassified,
        "Phân loại ưu tiên": make_priority_method_table(),
        "Cấu hình khoa": khoa_table,
    }
