# -*- coding: utf-8 -*-
"""Đối chiếu tên bác sĩ trong file "danh sách sơ bộ" (AI trích từ ảnh sổ viết
tay) với danh sách nhân sự thật (nhan_su_web.json), để phát hiện sớm tên bị
đọc sai/không khớp ai trước khi đưa vào dùng chính thức.

Vấn đề thường gặp: chữ viết tay dễ bị đọc nhầm dấu thanh (vd "Toàn" thay vì
"Toản"), AI lại không tự đánh dấu "chưa chắc" cho cột Bác sĩ như đã làm với
cột Chẩn đoán.

Chạy:
    python doi_chieu_nhan_su_so_bo.py --file "Danh_sach_phau_thuat_trich_tu_anh_so_bo.xlsx"

Có thể chỉ định thêm:
    --sheet "Danh sách sơ bộ"   (mặc định: sheet đầu tiên có cột "Bác sĩ phẫu thuật")
    --cot "Bác sĩ phẫu thuật"   (mặc định: tự dò các tên cột thường gặp)
    --nguong 0.72               (ngưỡng độ khớp tối thiểu để coi là gợi ý đáng tin, 0-1)

Kết quả: file mới "<tên gốc>_da_doi_chieu.xlsx", giữ nguyên sheet gốc và
thêm 2 cột "Khớp nhân sự" / "Gợi ý nếu chưa khớp", cùng 1 sheet mới
"Cần kiểm tra tên BS" liệt kê riêng các dòng cần xem lại.
"""
from __future__ import annotations

import argparse
import json
import re
import unicodedata
from dataclasses import dataclass
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

BASE_DIR = Path(__file__).resolve().parent
STAFF_FILE = BASE_DIR / "nhan_su_web.json"

# Tên cột chứa bác sĩ hay gặp trong các file trích xuất từ ảnh sổ.
COT_BAC_SI_CO_THE = [
    "Bác sĩ phẫu thuật",
    "Bác sĩ",
    "PTV chính",
    "Bác sĩ mổ chính",
]

NGUONG_KHOP_MAC_DINH = 0.72


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    if not text or text in {"nan", "none"}:
        return ""
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d")
    text = re.sub(r"[^a-z0-9]+", " ", text)
    return re.sub(r"\s+", " ", text).strip()


TITLE_PATTERN = re.compile(
    r"^(pgs|gs|ts|ths|bsckii|bscki|bsck2|bsck1|bs\s*ck\s*ii|bs\s*ck\s*i|bs|dr)\.?\s*",
    re.IGNORECASE,
)


def strip_title(value: Any) -> str:
    text = str(value or "").strip()
    changed = True
    while changed:
        before = text
        text = TITLE_PATTERN.sub("", text).strip()
        changed = text != before
    return text


@dataclass
class NhanSu:
    ho_ten: str
    bi_danh: str
    nhom: str
    vai_tro: str


def load_staff(only_active: bool = True, vai_tro_can_lay: set[str] | None = None) -> list[NhanSu]:
    if not STAFF_FILE.exists():
        raise FileNotFoundError(f"Không tìm thấy {STAFF_FILE}")
    data = json.loads(STAFF_FILE.read_text(encoding="utf-8"))
    result = []
    for item in data.get("staff", []):
        if only_active and not item.get("active", True):
            continue
        vai_tro = str(item.get("vaiTro") or "").strip()
        if vai_tro_can_lay and vai_tro not in vai_tro_can_lay:
            continue
        result.append(
            NhanSu(
                ho_ten=str(item.get("hoTen") or "").strip(),
                bi_danh=str(item.get("biDanh") or "").strip(),
                nhom=str(item.get("nhom") or "").strip(),
                vai_tro=vai_tro,
            )
        )
    return result


# Cột "Bác sĩ phẫu thuật"/"PTV chính" chỉ nên so khớp với Bác sĩ/BSNT — không
# tính Điều dưỡng/KTV, để tránh nhầm khi tên trùng một từ (vd "Duy" là bí danh
# của bác sĩ Trần Nguyễn Anh Duy nhưng cũng là tên lót của điều dưỡng Võ
# Phương Duy).
VAI_TRO_BAC_SI = {"bac_si", "bsnt"}


@dataclass
class KetQuaDoiChieu:
    trang_thai: str  # "Khớp chính xác" | "Có khả năng khớp" | "Không khớp ai" | "Bỏ trống"
    ten_goc: str
    goi_y: str  # "Họ tên (BÍ DANH)" gần nhất, rỗng nếu không có
    diem_khop: float


def doi_chieu_mot_ten(ten_goc: str, danh_sach: list[NhanSu]) -> KetQuaDoiChieu:
    sach = strip_title(ten_goc)
    key = normalize_text(sach)
    if not key:
        return KetQuaDoiChieu("Bỏ trống", ten_goc, "", 0.0)

    best: tuple[float, NhanSu | None] = (0.0, None)
    exact_matches: list[NhanSu] = []

    for ns in danh_sach:
        # Chỉ bí danh, họ tên đầy đủ, và "tên không kèm họ" (cách gọi phổ biến,
        # vd "Giang Tử" cho "Nguyễn Giang Tử") mới tính "khớp chính xác". Từng
        # từ tách lẻ (vd chỉ "Duy") rất dễ trùng giữa nhiều người khác nhau nên
        # chỉ dùng để tính điểm khớp mờ, không dùng để kết luận khớp tuyệt đối.
        ten_khong_ho = " ".join(ns.ho_ten.split()[1:])
        exact_candidates = [ns.bi_danh, ns.ho_ten, ten_khong_ho]
        is_exact = any(normalize_text(c) == key for c in exact_candidates if c)

        fuzzy_candidates = [ns.bi_danh, ns.ho_ten, *ns.ho_ten.split()]
        best_for_this_person = 0.0
        for cand in fuzzy_candidates:
            cand_key = normalize_text(cand)
            if not cand_key:
                continue
            ratio = 1.0 if cand_key == key else SequenceMatcher(None, key, cand_key).ratio()
            best_for_this_person = max(best_for_this_person, ratio)

        if is_exact:
            exact_matches.append(ns)
        if best_for_this_person > best[0]:
            best = (best_for_this_person, ns)

    # Nhiều người có cùng bí danh/họ tên trùng khớp tuyệt đối là tình huống
    # cấu hình nhân sự bị trùng, không phải lỗi của file này — vẫn báo rõ.
    if len(exact_matches) == 1:
        ns = exact_matches[0]
        goi_y = f"{ns.ho_ten} ({ns.bi_danh})"
        # "Toàn" và "Toản" chuẩn hóa giống nhau (bỏ dấu thanh) nên khớp đúng
        # người, nhưng nếu giữ nguyên chữ đọc từ ảnh khi gõ vào file chính
        # thức thì công thức Excel so khớp chính xác từng ký tự (vd
        # H7="TOẢN") sẽ không nhận ra. Cảnh báo riêng khi chữ viết lệch dấu.
        dung_dau = sach.strip().casefold() in {
            c.casefold() for c in (ns.bi_danh, ns.ho_ten, " ".join(ns.ho_ten.split()[1:])) if c
        }
        if not dung_dau:
            return KetQuaDoiChieu(
                "Khớp đúng người nhưng lệch dấu",
                ten_goc,
                f"Đúng dấu phải ghi: {goi_y}",
                0.99,
            )
        return KetQuaDoiChieu("Khớp chính xác", ten_goc, goi_y, 1.0)
    if len(exact_matches) > 1:
        goi_y = "; ".join(f"{ns.ho_ten} ({ns.bi_danh})" for ns in exact_matches)
        return KetQuaDoiChieu("Khớp chính xác nhiều người", ten_goc, goi_y, 1.0)

    score, ns = best
    if ns is not None and score >= NGUONG_KHOP_MAC_DINH:
        return KetQuaDoiChieu("Có khả năng khớp", ten_goc, f"{ns.ho_ten} ({ns.bi_danh})", score)

    goi_y = f"{ns.ho_ten} ({ns.bi_danh})" if ns is not None else ""
    return KetQuaDoiChieu("Không khớp ai", ten_goc, goi_y, score)


def doi_chieu_o_bac_si(gia_tri: Any, danh_sach: list[NhanSu]) -> list[KetQuaDoiChieu]:
    """Một ô có thể ghi nhiều khả năng cách nhau bằng '/' khi AI không chắc,
    ví dụ 'BS Sơn/BS Toản'. Đối chiếu riêng từng khả năng."""
    text = str(gia_tri or "").strip()
    if not text:
        return [KetQuaDoiChieu("Bỏ trống", "", "", 0.0)]

    phan = [p.strip() for p in re.split(r"[/;]", text) if p.strip()]
    return [doi_chieu_mot_ten(p, danh_sach) for p in phan]


def tim_cot_bac_si(headers: list[str]) -> str | None:
    header_keys = {normalize_text(h).replace(" ", ""): h for h in headers}
    for ten in COT_BAC_SI_CO_THE:
        key = normalize_text(ten).replace(" ", "")
        if key in header_keys:
            return header_keys[key]
    return None


def tim_cot_ghi_chu(ws, header_row: int) -> int | None:
    for col_idx in range(1, ws.max_column + 1):
        header = str(ws.cell(header_row, col_idx).value or "").strip()
        if normalize_text(header).replace(" ", "") == "ghichu":
            return col_idx
    return None


def tim_sheet_va_cot(wb, sheet_arg: str | None, cot_arg: str | None) -> tuple[str, int, str, int]:
    """Trả về (tên sheet, dòng tiêu đề, tên cột Bác sĩ, số cột Bác sĩ)."""
    sheet_names = [sheet_arg] if sheet_arg else wb.sheetnames

    for sheet_name in sheet_names:
        if sheet_name not in wb.sheetnames:
            continue
        ws = wb[sheet_name]
        for row_idx in range(1, min(ws.max_row, 20) + 1):
            headers = [str(ws.cell(row_idx, c).value or "").strip() for c in range(1, ws.max_column + 1)]
            if sum(1 for h in headers if h) < 3:
                continue
            cot_bac_si = cot_arg if cot_arg in headers else tim_cot_bac_si(headers)
            if cot_bac_si and cot_bac_si in headers:
                col_idx = headers.index(cot_bac_si) + 1
                return sheet_name, row_idx, cot_bac_si, col_idx

    raise ValueError(
        "Không tìm thấy cột chứa tên bác sĩ (thử các tên: "
        + ", ".join(COT_BAC_SI_CO_THE)
        + "). Dùng --cot để chỉ định tên cột chính xác."
    )


def _format_summary_sheet(ws) -> None:
    header_fill = PatternFill("solid", fgColor="FCE4D6")
    header_font = Font(bold=True)
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for cell in ws[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
    widths = {"A": 10, "B": 30, "C": 34, "D": 14, "E": 40}
    for col, width in widths.items():
        ws.column_dimensions[col].width = width


def run(file_path: Path, sheet_arg: str | None, cot_arg: str | None, nguong: float) -> dict[str, Any]:
    global NGUONG_KHOP_MAC_DINH
    NGUONG_KHOP_MAC_DINH = nguong

    danh_sach = load_staff(vai_tro_can_lay=VAI_TRO_BAC_SI)
    if not danh_sach:
        raise ValueError("Danh sách nhân sự đang trống (nhan_su_web.json).")

    wb = load_workbook(file_path)
    sheet_name, header_row, cot_bac_si, col_bac_si = tim_sheet_va_cot(wb, sheet_arg, cot_arg)
    ws = wb[sheet_name]
    col_ghi_chu = tim_cot_ghi_chu(ws, header_row)

    col_khop = ws.max_column + 1
    col_goi_y = ws.max_column + 2
    ws.cell(header_row, col_khop).value = "Khớp nhân sự"
    ws.cell(header_row, col_goi_y).value = "Gợi ý nếu chưa khớp"
    for col_idx in (col_khop, col_goi_y):
        cell = ws.cell(header_row, col_idx)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="D9EAF7")
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.column_dimensions[ws.cell(header_row, col_khop).column_letter].width = 18
    ws.column_dimensions[ws.cell(header_row, col_goi_y).column_letter].width = 40

    ok_fill = PatternFill("solid", fgColor="E2EFDA")
    warn_fill = PatternFill("solid", fgColor="FFF2CC")
    bad_fill = PatternFill("solid", fgColor="FCE4D6")

    thong_ke = {
        "khop_chinh_xac": 0,
        "lech_dau": 0,
        "co_kha_nang": 0,
        "khong_khop": 0,
        "bo_trong": 0,
        "nhieu_nguoi": 0,
    }
    can_kiem_tra: list[dict[str, Any]] = []

    for row_idx in range(header_row + 1, ws.max_row + 1):
        gia_tri = ws.cell(row_idx, col_bac_si).value
        if all(ws.cell(row_idx, c).value in (None, "") for c in range(1, ws.max_column + 1)):
            continue

        ket_qua_list = doi_chieu_o_bac_si(gia_tri, danh_sach)
        trang_thai_list = [kq.trang_thai for kq in ket_qua_list]

        if all(t == "Bỏ trống" for t in trang_thai_list):
            trang_thai_tong = "Bỏ trống"
            thong_ke["bo_trong"] += 1
        elif all(t == "Khớp chính xác" for t in trang_thai_list):
            trang_thai_tong = "Khớp chính xác"
            thong_ke["khop_chinh_xac"] += 1
        elif any(t == "Không khớp ai" for t in trang_thai_list):
            trang_thai_tong = "Không khớp ai"
            thong_ke["khong_khop"] += 1
        elif any(t == "Khớp chính xác nhiều người" for t in trang_thai_list):
            trang_thai_tong = "Trùng cấu hình nhân sự"
            thong_ke["nhieu_nguoi"] += 1
        elif all(t in ("Khớp chính xác", "Khớp đúng người nhưng lệch dấu") for t in trang_thai_list):
            trang_thai_tong = "Đúng người nhưng lệch dấu — nên sửa lại chính tả"
            thong_ke["lech_dau"] += 1
        else:
            trang_thai_tong = "Có khả năng khớp"
            thong_ke["co_kha_nang"] += 1

        goi_y_text = " | ".join(
            f"{kq.ten_goc or '(trống)'} -> {kq.goi_y or 'không tìm được ai gần giống'}"
            for kq in ket_qua_list
            if kq.trang_thai != "Khớp chính xác"
        )

        ws.cell(row_idx, col_khop).value = trang_thai_tong
        ws.cell(row_idx, col_goi_y).value = goi_y_text

        cell_khop = ws.cell(row_idx, col_khop)
        if trang_thai_tong == "Khớp chính xác":
            cell_khop.fill = ok_fill
        elif trang_thai_tong in (
            "Có khả năng khớp",
            "Bỏ trống",
            "Đúng người nhưng lệch dấu — nên sửa lại chính tả",
        ):
            cell_khop.fill = warn_fill
        else:
            cell_khop.fill = bad_fill

        # Ghi thẳng vào cột "Ghi chú" sẵn có (chỗ AI đang dùng để cảnh báo ảnh
        # mờ/chưa đọc rõ) để thấy ngay khi lướt bảng chính, không phải mở
        # thêm sheet riêng. Giữ lại ghi chú cũ nếu đã có.
        if col_ghi_chu and trang_thai_tong not in ("Khớp chính xác", "Bỏ trống"):
            note_moi = f"[Kiểm tra tên BS] {trang_thai_tong}: {goi_y_text}"
            note_cu = str(ws.cell(row_idx, col_ghi_chu).value or "").strip()
            ws.cell(row_idx, col_ghi_chu).value = f"{note_cu} | {note_moi}" if note_cu else note_moi

        if trang_thai_tong not in ("Khớp chính xác", "Bỏ trống"):
            can_kiem_tra.append(
                {
                    "Dòng Excel": row_idx,
                    "Tên bác sĩ đọc từ ảnh": str(gia_tri or "").strip(),
                    "Trạng thái": trang_thai_tong,
                    "Độ khớp cao nhất": round(max((kq.diem_khop for kq in ket_qua_list), default=0.0), 2),
                    "Gợi ý": goi_y_text,
                }
            )

    summary_sheet_name = "Cần kiểm tra tên BS"
    if summary_sheet_name in wb.sheetnames:
        del wb[summary_sheet_name]
    ws_summary = wb.create_sheet(summary_sheet_name)
    summary_headers = ["Dòng Excel", "Tên bác sĩ đọc từ ảnh", "Trạng thái", "Độ khớp cao nhất", "Gợi ý"]
    ws_summary.append(summary_headers)
    for item in can_kiem_tra:
        ws_summary.append([item[h] for h in summary_headers])
    _format_summary_sheet(ws_summary)

    output_file = file_path.with_name(file_path.stem + "_da_doi_chieu" + file_path.suffix)
    wb.save(output_file)

    return {
        "file_dau_ra": str(output_file),
        "sheet": sheet_name,
        "cot_bac_si": cot_bac_si,
        "tong_dong": sum(thong_ke.values()),
        "thong_ke": thong_ke,
        "so_dong_can_kiem_tra": len(can_kiem_tra),
        "da_ghi_vao_cot_ghi_chu": col_ghi_chu is not None,
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Đối chiếu tên bác sĩ trong file trích xuất từ ảnh với danh sách nhân sự.")
    parser.add_argument("--file", required=True, help="File Excel danh sách sơ bộ trích từ ảnh.")
    parser.add_argument("--sheet", default=None, help="Tên sheet cần đối chiếu (mặc định: tự dò).")
    parser.add_argument("--cot", default=None, help="Tên cột chứa bác sĩ (mặc định: tự dò).")
    parser.add_argument("--nguong", type=float, default=NGUONG_KHOP_MAC_DINH, help="Ngưỡng độ khớp gợi ý (0-1).")
    return parser


def main() -> None:
    args = build_parser().parse_args()
    file_path = Path(args.file)
    if not file_path.exists():
        raise FileNotFoundError(f"Không tìm thấy file: {file_path}")

    result = run(file_path, args.sheet, args.cot, args.nguong)

    print("ĐÃ ĐỐI CHIẾU TÊN BÁC SĨ")
    print(f"- Sheet: {result['sheet']} · Cột: {result['cot_bac_si']}")
    print(f"- Tổng dòng có dữ liệu: {result['tong_dong']}")
    tk = result["thong_ke"]
    print(f"- Khớp chính xác: {tk['khop_chinh_xac']}")
    print(f"- Đúng người nhưng lệch dấu, nên sửa lại chính tả: {tk['lech_dau']}")
    print(f"- Có khả năng khớp (đề nghị xem lại): {tk['co_kha_nang']}")
    print(f"- Không khớp ai trong danh sách nhân sự: {tk['khong_khop']}")
    print(f"- Trùng cấu hình nhân sự (2 người cùng bí danh/tên): {tk['nhieu_nguoi']}")
    print(f"- Bỏ trống: {tk['bo_trong']}")
    print(f"- Số dòng cần kiểm tra (xem sheet 'Cần kiểm tra tên BS'): {result['so_dong_can_kiem_tra']}")
    if result["da_ghi_vao_cot_ghi_chu"]:
        print("- Đã ghi chú các dòng cần kiểm tra thẳng vào cột 'Ghi chú' của sheet gốc.")
    else:
        print("- Không tìm thấy cột 'Ghi chú' trong sheet gốc nên chỉ ghi vào 2 cột mới và sheet tổng hợp.")
    print(f"- File đầu ra: {result['file_dau_ra']}")


if __name__ == "__main__":
    main()
