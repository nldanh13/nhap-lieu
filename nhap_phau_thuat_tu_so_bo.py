# -*- coding: utf-8 -*-
"""Nhập các dòng từ file "Danh sách sơ bộ" (trích xuất bằng OCR/AI từ ảnh sổ
phẫu thuật viết tay, đã chuẩn hóa 4 cột BS bằng xu_ly_so_phau_thuat.py)
thành dòng mới trong sheet phauthuat của file PM khoa CTCH đang dùng.

Chạy:
    python nhap_phau_thuat_tu_so_bo.py \
        --file "PM khoa CTCH T8_NHAP_LIEU.xlsx" \
        --so-bo "Danh_sach_phau_thuat_da_chuan_hoa.xlsx" \
        --output "PM khoa CTCH T8_NHAP_LIEU.xlsx"

CHƯA có trong dữ liệu OCR nên luôn để trống, cần tự điền/kiểm tra sau:
- Loại phẫu thuật (Loại 1/2/3/Đặc biệt) — dùng để tính tiền, không có cách
  suy luận chắc chắn từ tên phương pháp.
- Số tiền thực tính thù lao cho khoa — lấy từ bảng giá/HIS, sổ viết tay
  không có. Công thức tính hoa hồng (cột Số tiền/Thực lãnh) vẫn được sao
  chép nguyên vẹn từ dòng mẫu nên sẽ tự tính đúng ngay khi điền số tiền.

Tên bác sĩ (PTV chính/Phụ mổ 1-3) được đối chiếu qua danh sách nhân sự
(nhan_su_web.json, qua staff_to_alias) để lấy đúng bí danh; nếu không khớp
ai sẽ giữ nguyên chữ đọc được và đánh dấu trong file báo cáo.
"""
from __future__ import annotations

import argparse
from pathlib import Path
from typing import Any

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

from chuyen_thu_thuat_sang_tieuphau_t5 import (
    copy_row_style,
    find_header_row_and_columns,
    find_sheet_case_insensitive,
    find_total_row,
    insert_rows_preserve_merges,
    normalize_header,
    normalize_text,
    translate_formula,
    value_is_blank,
)
from bo_sung_phu_mo_t5_tu_so_phau_thuat import staff_to_alias


SHEET_PHAU_THUAT = "phauthuat"
SHEET_SO_BO_MAC_DINH = "Danh sách sơ bộ"

# Tên cột thật trong sheet phauthuat của file PM khoa CTCH.
COT_PHAU_THUAT_CAN = [
    "NGÀY",
    "Họ và tên bệnh nhân",
    "Tuổi",
    "Chẩn đoán và phương pháp phẫu thuật",
    "PTV chính",
    "Phụ mổ 1",
    "Phụ mổ 2",
    "Phụ mổ 3",
]

# Các cột công thức tính hoa hồng theo thứ tự cột trong sheet phauthuat.
# J=Số tiền PTV chính, L=Số tiền phụ mổ 1, N=Số tiền phụ mổ 2,
# P=Số tiền phụ mổ 3, Q=ĐD+HL, R=Thực lãnh.
COT_CONG_THUC_TIEN = [10, 12, 14, 16, 17, 18]

DOCTOR_HEADERS_SO_BO = [
    "BS phẫu thuật 1",
    "BS phẫu thuật 2",
    "BS phẫu thuật 3",
    "BS phẫu thuật 4",
]


def header_map_so_bo(ws) -> dict[str, int]:
    header_row = 5
    return {
        str(cell.value).strip(): cell.column
        for cell in ws[header_row]
        if not value_is_blank(cell.value)
    }


def find_col_so_bo(headers: dict[str, int], *names: str) -> int | None:
    for name in names:
        if name in headers:
            return headers[name]
    return None


def doi_bac_si_sang_bi_danh(raw: Any) -> tuple[str, bool]:
    """Trả về (bí danh hoặc chữ gốc, đã khớp được hay chưa)."""
    text = str(raw or "").strip()
    if not text or text == "-":
        return "-", True
    alias = staff_to_alias(text)
    if alias:
        return alias, True
    return text, False


def ghep_chan_doan_phuong_phap(chan_doan: Any, phuong_phap: Any) -> str:
    chan_doan_text = str(chan_doan or "").strip()
    phuong_phap_text = str(phuong_phap or "").strip()
    if chan_doan_text and phuong_phap_text and chan_doan_text != phuong_phap_text:
        return f"{chan_doan_text}. {phuong_phap_text}"
    return phuong_phap_text or chan_doan_text


def doc_dong_so_bo(ws_so_bo) -> list[dict[str, Any]]:
    headers = header_map_so_bo(ws_so_bo)
    col_ngay = find_col_so_bo(headers, "Ngày ghi sổ", "Ngày")
    col_ho_ten = find_col_so_bo(headers, "Họ tên người bệnh", "Họ và tên")
    col_tuoi = find_col_so_bo(headers, "Tuổi")
    col_chan_doan = find_col_so_bo(headers, "Chẩn đoán đọc từ sổ", "Chẩn đoán")
    col_phuong_phap = find_col_so_bo(headers, "Phương pháp phẫu thuật", "Phương pháp")
    col_bac_si = [find_col_so_bo(headers, name) for name in DOCTOR_HEADERS_SO_BO]
    col_bac_si_chuoi = find_col_so_bo(headers, "Kíp phẫu thuật (đọc theo chuỗi)", "Bác sĩ phẫu thuật")

    if col_ngay is None or col_ho_ten is None:
        raise ValueError(
            "Không tìm thấy cột Ngày/Họ tên người bệnh trong sheet danh sách sơ bộ. "
            "Kiểm tra lại tên cột hoặc dùng --cot-bac-si nếu file có tên cột khác."
        )
    if not any(col_bac_si) and col_bac_si_chuoi is None:
        raise ValueError(
            "Không tìm thấy cột bác sĩ (BS phẫu thuật 1-4 hoặc cột chuỗi Bác sĩ phẫu thuật). "
            "Chạy xu_ly_so_phau_thuat.py để tách 4 cột BS trước."
        )

    rows: list[dict[str, Any]] = []
    for row_idx in range(6, ws_so_bo.max_row + 1):
        ho_ten = ws_so_bo.cell(row_idx, col_ho_ten).value
        ngay = ws_so_bo.cell(row_idx, col_ngay).value
        if value_is_blank(ho_ten) and value_is_blank(ngay):
            continue

        if any(col_bac_si):
            bac_si_raw = [ws_so_bo.cell(row_idx, c).value if c else "-" for c in col_bac_si]
        else:
            # Chưa qua xu_ly_so_phau_thuat.py: tách tạm chuỗi bằng dấu gạch ngang.
            chuoi = str(ws_so_bo.cell(row_idx, col_bac_si_chuoi).value or "")
            parts = [p.strip() for p in chuoi.replace("\r", "\n").split(" - ") if p.strip()]
            bac_si_raw = (parts + ["-", "-", "-", "-"])[:4]

        rows.append(
            {
                "source_row": row_idx,
                "ngay": ngay,
                "ho_ten": ho_ten,
                "tuoi": ws_so_bo.cell(row_idx, col_tuoi).value if col_tuoi else None,
                "chan_doan_phuong_phap": ghep_chan_doan_phuong_phap(
                    ws_so_bo.cell(row_idx, col_chan_doan).value if col_chan_doan else None,
                    ws_so_bo.cell(row_idx, col_phuong_phap).value if col_phuong_phap else None,
                ),
                "bac_si_raw": bac_si_raw,
            }
        )
    return rows


def tim_dong_du_lieu_cuoi(ws, header_row: int, total_row: int, cols: list[int]) -> int:
    for row_idx in range(total_row - 1, header_row, -1):
        if any(not value_is_blank(ws.cell(row_idx, c).value) for c in cols):
            return row_idx
    return header_row


def tim_dong_mau_cong_thuc(ws, header_row: int, total_row: int) -> int:
    for row_idx in range(header_row + 1, total_row):
        if any(
            isinstance(ws.cell(row_idx, c).value, str) and ws.cell(row_idx, c).value.startswith("=")
            for c in COT_CONG_THUC_TIEN
        ):
            return row_idx
    return header_row + 1


def cap_nhat_tong_cong(ws, header_row: int, total_row: int) -> None:
    start = header_row + 1
    end = total_row - 1
    ws.cell(total_row, 8).value = f"=SUM(H{start}:H{end})"
    # Các cột J/L/N/P/Q/R dùng SUBTOTAL trong mẫu gốc; giữ nguyên kiểu SUM để
    # chắc chắn cộng đủ dòng mới thêm (SUBTOTAL trong file gốc có thể đang bị
    # giới hạn phạm vi cũ, xem README/báo cáo).
    for col_idx in COT_CONG_THUC_TIEN:
        col_letter = ws.cell(total_row, col_idx).column_letter
        ws.cell(total_row, col_idx).value = f"=SUM({col_letter}{start}:{col_letter}{end})"


def nhap_phau_thuat_tu_so_bo(
    file_path: Path,
    so_bo_path: Path,
    output_path: Path,
    sheet_so_bo: str = SHEET_SO_BO_MAC_DINH,
) -> dict[str, Any]:
    wb = load_workbook(file_path)
    sheet_name = find_sheet_case_insensitive(wb, SHEET_PHAU_THUAT)
    ws = wb[sheet_name]

    wb_so_bo = load_workbook(so_bo_path, data_only=True)
    if sheet_so_bo not in wb_so_bo.sheetnames:
        sheet_so_bo = wb_so_bo.sheetnames[0]
    ws_so_bo = wb_so_bo[sheet_so_bo]

    dong_so_bo = doc_dong_so_bo(ws_so_bo)

    header_row, header_map = find_header_row_and_columns(ws, ["STT", *COT_PHAU_THUAT_CAN])
    col_stt = header_map[normalize_header("STT")]
    col_ngay = header_map[normalize_header("NGÀY")]
    col_ho_ten = header_map[normalize_header("Họ và tên bệnh nhân")]
    col_tuoi = header_map[normalize_header("Tuổi")]
    col_chan_doan = header_map[normalize_header("Chẩn đoán và phương pháp phẫu thuật")]
    col_ptv_chinh = header_map[normalize_header("PTV chính")]
    col_phu_mo = [header_map[normalize_header(f"Phụ mổ {i}")] for i in (1, 2, 3)]

    total_row_before = find_total_row(ws)
    data_cols = [col_ngay, col_ho_ten, col_ptv_chinh, *col_phu_mo]
    last_data_row = tim_dong_du_lieu_cuoi(ws, header_row, total_row_before, data_cols)
    insert_at = last_data_row + 1
    amount = len(dong_so_bo)

    current_stt = [
        ws.cell(r, col_stt).value
        for r in range(header_row + 1, total_row_before)
        if isinstance(ws.cell(r, col_stt).value, (int, float))
    ]
    next_stt = (max(current_stt) if current_stt else 0) + 1

    style_template_row = header_row + 1
    formula_template_row = tim_dong_mau_cong_thuc(ws, header_row, total_row_before)

    if amount:
        insert_rows_preserve_merges(ws, insert_at, amount)

    bao_cao_rows: list[dict[str, Any]] = []
    for offset, item in enumerate(dong_so_bo):
        row_idx = insert_at + offset
        copy_row_style(ws, style_template_row, row_idx, ws.max_column)

        ws.cell(row_idx, col_stt).value = next_stt + offset
        ws.cell(row_idx, col_ngay).value = item["ngay"]
        ws.cell(row_idx, col_ho_ten).value = item["ho_ten"]
        ws.cell(row_idx, col_tuoi).value = item["tuoi"]
        ws.cell(row_idx, col_chan_doan).value = item["chan_doan_phuong_phap"]
        # Loại phẫu thuật (cột kế Chẩn đoán) và Số tiền thực tính thù lao cố
        # ý để trống — không có cách suy luận đúng từ dữ liệu OCR, cần người
        # dùng tự điền/kiểm tra bằng bảng giá hoặc HIS.

        alias_ptv, khop_ptv = doi_bac_si_sang_bi_danh(item["bac_si_raw"][0])
        ws.cell(row_idx, col_ptv_chinh).value = alias_ptv

        khop_phu_mo = []
        for phu_mo_idx, col in enumerate(col_phu_mo):
            alias, khop = doi_bac_si_sang_bi_danh(item["bac_si_raw"][phu_mo_idx + 1])
            ws.cell(row_idx, col).value = alias
            khop_phu_mo.append(khop)

        for col_idx in COT_CONG_THUC_TIEN:
            src = ws.cell(formula_template_row, col_idx)
            dst = ws.cell(row_idx, col_idx)
            if isinstance(src.value, str) and src.value.startswith("="):
                dst.value = translate_formula(src.value, src, dst)

        can_kiem_tra = []
        if not khop_ptv:
            can_kiem_tra.append(f"PTV chính '{item['bac_si_raw'][0]}' không khớp ai trong danh sách nhân sự")
        for idx, khop in enumerate(khop_phu_mo, start=1):
            if not khop and item["bac_si_raw"][idx] not in ("-", ""):
                can_kiem_tra.append(f"Phụ mổ {idx} '{item['bac_si_raw'][idx]}' không khớp ai trong danh sách nhân sự")

        bao_cao_rows.append(
            {
                "Dòng Excel (sheet mới)": row_idx,
                "Dòng nguồn (danh sách sơ bộ)": item["source_row"],
                "Họ và tên": item["ho_ten"],
                "PTV chính": alias_ptv,
                "Phụ mổ 1": ws.cell(row_idx, col_phu_mo[0]).value,
                "Phụ mổ 2": ws.cell(row_idx, col_phu_mo[1]).value,
                "Phụ mổ 3": ws.cell(row_idx, col_phu_mo[2]).value,
                "Cần kiểm tra": "; ".join(can_kiem_tra) if can_kiem_tra else "",
            }
        )

    total_row_after = find_total_row(ws)
    cap_nhat_tong_cong(ws, header_row, total_row_after)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_path)

    return {
        "so_dong_them": amount,
        "dong_chen_tu": insert_at if amount else None,
        "dong_chen_den": insert_at + amount - 1 if amount else None,
        "tong_cong_truoc": total_row_before,
        "tong_cong_sau": total_row_after,
        "so_dong_can_kiem_tra": sum(1 for r in bao_cao_rows if r["Cần kiểm tra"]),
        "bao_cao_rows": bao_cao_rows,
    }


def tao_bao_cao(bao_cao_rows: list[dict[str, Any]], output_report_path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Bao_cao_nhap_phau_thuat"
    if bao_cao_rows:
        headers = list(bao_cao_rows[0].keys())
    else:
        headers = ["Dòng Excel (sheet mới)", "Dòng nguồn (danh sách sơ bộ)", "Họ và tên", "PTV chính", "Phụ mổ 1", "Phụ mổ 2", "Phụ mổ 3", "Cần kiểm tra"]
    ws.append(headers)
    for row in bao_cao_rows:
        ws.append([row.get(h, "") for h in headers])

    header_fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    widths = [16, 20, 26, 16, 14, 14, 14, 50]
    for idx, width in enumerate(widths, start=1):
        ws.column_dimensions[ws.cell(1, idx).column_letter].width = width

    output_report_path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(output_report_path)


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Nhập dòng từ danh sách sơ bộ OCR vào sheet phauthuat.")
    parser.add_argument("--file", required=True, type=Path, help="File PM khoa CTCH đang dùng.")
    parser.add_argument("--so-bo", required=True, type=Path, help="File danh sách sơ bộ đã chuẩn hóa 4 cột BS.")
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--sheet-so-bo", default=SHEET_SO_BO_MAC_DINH)
    parser.add_argument("--report", type=Path, default=None, help="File báo cáo, mặc định cạnh --output.")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    result = nhap_phau_thuat_tu_so_bo(args.file, args.so_bo, args.output, args.sheet_so_bo)

    report_path = args.report or args.output.with_name(args.output.stem + "_bao_cao_nhap_phau_thuat.xlsx")
    tao_bao_cao(result["bao_cao_rows"], report_path)

    print("ĐÃ NHẬP DANH SÁCH SƠ BỘ VÀO SHEET PHAUTHUAT")
    print(f"- Số dòng đã thêm: {result['so_dong_them']}")
    if result["so_dong_them"]:
        print(f"- Chèn từ dòng {result['dong_chen_tu']} đến {result['dong_chen_den']}")
    print(f"- Dòng Tổng Cộng: {result['tong_cong_truoc']} -> {result['tong_cong_sau']}")
    print(f"- Số dòng cần kiểm tra tên bác sĩ: {result['so_dong_can_kiem_tra']}")
    print("- CẦN TỰ ĐIỀN: cột 'Loại phẫu thuật' và 'Số tiền thực tính thù lao cho khoa' cho các dòng mới.")
    print(f"- File đầu ra: {args.output}")
    print(f"- File báo cáo: {report_path}")


if __name__ == "__main__":
    main()
