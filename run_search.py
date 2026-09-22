# -*- coding: utf-8 -*-
"""CLI/JSON bridge cho emr_list_search.py.

Ví dụ dòng lệnh:
  python run_search.py --config config.json --type procedure --patient 26069905 --keep-open

Ví dụ nhận JSON từ stdin:
  echo {"config":"config.json","list_type":"surgery","patient":"26069905"} | python run_search.py --stdin-json
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Dict

from emr_list_search import EmrListSearcher, load_config


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="Tìm người bệnh trên D/s Thủ thuật hoặc D/s Phẫu thuật")
    parser.add_argument("--config", default="config.json", help="Đường dẫn config JSON")
    parser.add_argument("--type", dest="list_type", choices=("procedure", "surgery"))
    parser.add_argument("--patient", help="Mã hoặc tên người bệnh")
    parser.add_argument("--months", type=int, default=3)
    parser.add_argument("--headless", action="store_true")
    parser.add_argument("--keep-open", action="store_true", help="Giữ Chrome tại trang kết quả đến khi nhấn Enter")
    parser.add_argument("--output", help="Lưu JSON kết quả ra file")
    parser.add_argument("--stdin-json", action="store_true", help="Đọc yêu cầu JSON từ stdin")
    return parser


def _read_request(args: argparse.Namespace) -> Dict[str, Any]:
    request: Dict[str, Any] = {}
    if args.stdin_json:
        raw = sys.stdin.read().strip()
        if not raw:
            raise ValueError("stdin không có JSON")
        loaded = json.loads(raw)
        if not isinstance(loaded, dict):
            raise ValueError("JSON đầu vào phải là object")
        request.update(loaded)
    request.setdefault("config", args.config)
    request.setdefault("list_type", args.list_type)
    request.setdefault("patient", args.patient)
    request.setdefault("months", args.months)
    request.setdefault("headless", args.headless)
    request.setdefault("keep_open", args.keep_open)
    request.setdefault("output", args.output)
    return request


def main() -> int:
    args = _parser().parse_args()
    client = None
    try:
        req = _read_request(args)
        if req.get("list_type") not in {"procedure", "surgery"}:
            raise ValueError("Thiếu list_type: procedure hoặc surgery")
        if not str(req.get("patient") or "").strip():
            raise ValueError("Thiếu patient: mã hoặc tên người bệnh")

        config = load_config(str(req.get("config") or "config.json"))
        client = EmrListSearcher(config, headless=bool(req.get("headless")))
        client.start()
        result = client.search(
            str(req["list_type"]),
            str(req["patient"]),
            months=int(req.get("months") or 3),
        ).to_dict()

        output_text = json.dumps(result, ensure_ascii=False, indent=2)
        output_path = str(req.get("output") or "").strip()
        if output_path:
            Path(output_path).write_text(output_text, encoding="utf-8")
        print(output_text, flush=True)

        if bool(req.get("keep_open")) and not bool(req.get("headless")):
            input("Chrome đang dừng tại danh sách kết quả. Nhấn Enter để đóng...")
        return 0
    except Exception as exc:
        print(json.dumps({"ok": False, "error": str(exc)}, ensure_ascii=False), flush=True)
        return 1
    finally:
        if client is not None:
            client.close()


if __name__ == "__main__":
    raise SystemExit(main())
