# -*- coding: utf-8 -*-
"""Tiến trình Python chạy nền, giữ sẵn module đã nạp để xử lý liên tiếp nhiều
lệnh Excel mà không phải mở tiến trình Python mới + import lại openpyxl mỗi
lần (đo thực tế: import + mở + lưu 1 workbook 3 sheet/200 dòng tốn khoảng
0.6-0.7 giây CHỈ RIÊNG phần mở tiến trình/thư viện — đây là nguyên nhân
chính gây lag khi gõ dữ liệu, đổi sheet, tự động lưu).

Giao thức: mỗi dòng đọc từ stdin là 1 JSON {"id": <int>, "argv": [...]},
"argv" giống hệt danh sách tham số dòng lệnh cũ (vd ["read-sheet", "--file",
...]). Mỗi dòng ghi ra stdout là JSON {"id": <int>, ...} tương ứng — dùng
lại nguyên vẹn build_parser()/commands/respond() của node_excel_api.py nên
hành vi từng lệnh giống hệt bản chạy rời trước đây, chỉ khác chỗ không phải
mở tiến trình mới.

Cố tình KHÔNG xử lý lệnh "chay-ocr-anh-tho" ở đây (server.js vẫn mở tiến
trình riêng cho lệnh đó) vì nó gọi mạng ra Google Document AI, có thể mất
vài phút — nếu chạy qua tiến trình nền dùng chung này sẽ làm nghẽn toàn bộ
thao tác gõ/lưu khác của người dùng trong lúc chờ.
"""
from __future__ import annotations

import contextlib
import io
import json
import sys

import node_excel_api as api

COMMANDS = {
    "staff-list": api.command_staff_list,
    "save-staff-list": api.command_save_staff_list,
    "cls-list": api.command_cls_list,
    "save-cls-list": api.command_save_cls_list,
    "finalize-workbook": api.command_finalize_workbook,
    "list-files": api.command_list_files,
    "workbook-info": api.command_workbook_info,
    "normalize-dates": api.command_normalize_dates,
    "read-sheet": api.command_read_sheet,
    "update-row": api.command_update_row,
    "update-rows": api.command_update_rows,
    "clear-columns": api.command_clear_columns,
    "insert-row": api.command_insert_row,
    "delete-row": api.command_delete_row,
    "task-tat-ca": api.command_task_tat_ca,
    "task-bo-sung-phu-mo": api.command_task_bo_sung_phu_mo,
    "task-chuyen-tieu-phau": api.command_task_chuyen_tieu_phau,
    "task-cap-nhat-cls-tieuphau": api.command_task_cap_nhat_cls_tieuphau,
    "task-dien-bs": api.command_task_dien_bs,
    "task-nhap-phau-thuat-so-bo": api.command_task_nhap_phau_thuat_so_bo,
    "extract-zip-images": api.command_extract_zip_images,
    "doi-chieu-phau-thuat-anh": api.command_doi_chieu_phau_thuat_anh,
    "tach-cot-anh-so-phau-thuat": api.command_tach_cot_anh_so_phau_thuat,
}


def handle_one(argv: list) -> str:
    """Chạy 1 lệnh, trả về text JSON cuối cùng mà respond() đã in ra."""
    parser = api.build_parser()
    args = parser.parse_args(argv)
    if args.cmd not in COMMANDS:
        return json.dumps({"ok": False, "error": f"Lệnh không hỗ trợ qua worker: {args.cmd}"})

    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        try:
            COMMANDS[args.cmd](args)
        except SystemExit:
            pass  # fail() đã respond() (ghi vào buf) trước khi raise SystemExit
        except Exception as exc:  # noqa: BLE001 - phải bắt mọi lỗi để worker không chết
            try:
                api.fail(str(exc))
            except SystemExit:
                pass

    text = buf.getvalue().strip()
    if not text:
        return json.dumps({"ok": False, "error": "Lệnh không trả về phản hồi nào."})
    return text.splitlines()[-1]


def main() -> None:
    for raw_line in sys.stdin:
        line = raw_line.strip()
        if not line:
            continue
        try:
            request = json.loads(line)
            req_id = request.get("id")
            argv = request.get("argv") or []
        except Exception as exc:  # noqa: BLE001
            sys.stdout.write(json.dumps({"ok": False, "error": f"Yêu cầu không hợp lệ: {exc}"}) + "\n")
            sys.stdout.flush()
            continue

        try:
            result_text = handle_one(argv)
            result = json.loads(result_text)
        except SystemExit as exc:
            result = {"ok": False, "error": f"Lệnh thoát bất thường (mã {exc.code})."}
        except Exception as exc:  # noqa: BLE001
            result = {"ok": False, "error": str(exc)}

        result["id"] = req_id
        sys.stdout.write(json.dumps(result, ensure_ascii=False) + "\n")
        sys.stdout.flush()


if __name__ == "__main__":
    main()
