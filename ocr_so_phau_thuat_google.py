#!/usr/bin/env python3
"""
Đọc chữ viết tay trong ZIP ảnh bằng Google Cloud Document AI.

Cài đặt:
    pip install google-cloud-documentai

Chạy thử:
    python ocr_so_phau_thuat_google.py \
      --zip img.zip \
      --project-id TEN_PROJECT \
      --location us \
      --processor-id ID_PROCESSOR \
      --credentials service-account.json \
      --output ocr_raw.json

Script tạo JSON thô gồm toàn văn, từng dòng, vị trí tương đối và confidence.
Bước tách bệnh nhân/bác sĩ thành các cột Excel nên dùng dữ liệu JSON này
qua một bước kiểm tra, vì chữ viết tay và các dòng ghi nối trong sổ có thể
không được OCR nhận dạng hoàn hảo.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import zipfile
from pathlib import Path

from google.api_core.client_options import ClientOptions
from google.cloud import documentai_v1 as documentai


def extract_text(document, anchor) -> str:
    parts = []
    for segment in anchor.text_segments:
        start = int(segment.start_index or 0)
        end = int(segment.end_index or 0)
        parts.append(document.text[start:end])
    return "".join(parts).strip()


def vertices_to_box(layout):
    polygon = layout.bounding_poly
    points = []
    for vertex in polygon.normalized_vertices:
        points.append({"x": float(vertex.x), "y": float(vertex.y)})
    return points


def process_image(client, processor_name, filename, content):
    request = documentai.ProcessRequest(
        name=processor_name,
        raw_document=documentai.RawDocument(content=content, mime_type="image/jpeg"),
        process_options=documentai.ProcessOptions(
            # Trả thêm điểm chất lượng ảnh để biết ảnh nào cần chụp lại.
            enable_image_quality_scores=True,
        ),
    )
    result = client.process_document(request=request)
    doc = result.document

    pages = []
    for page_index, page in enumerate(doc.pages, start=1):
        lines = []
        for line_index, line in enumerate(page.lines, start=1):
            lines.append(
                {
                    "line": line_index,
                    "text": extract_text(doc, line.layout.text_anchor),
                    "confidence": float(line.layout.confidence),
                    "box": vertices_to_box(line.layout),
                }
            )
        pages.append(
            {
                "page": page_index,
                "width": float(page.dimension.width),
                "height": float(page.dimension.height),
                "lines": lines,
            }
        )

    quality = None
    if getattr(doc, "image_quality_scores", None):
        quality = {
            "quality_score": float(doc.image_quality_scores.quality_score),
            "defects": [
                {"type": str(defect.type_), "confidence": float(defect.confidence)}
                for defect in doc.image_quality_scores.detected_defects
            ],
        }

    return {
        "file": filename,
        "text": doc.text,
        "pages": pages,
        "image_quality": quality,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--zip", required=True, type=Path)
    parser.add_argument("--project-id", required=True)
    parser.add_argument("--location", default="us")
    parser.add_argument("--processor-id", required=True)
    parser.add_argument("--credentials", type=Path)
    parser.add_argument("--output", default="ocr_raw.json", type=Path)
    parser.add_argument("--limit", type=int, default=0, help="Chỉ xử lý N ảnh để kiểm tra thử")
    args = parser.parse_args()

    if args.credentials:
        os.environ["GOOGLE_APPLICATION_CREDENTIALS"] = str(args.credentials.resolve())

    endpoint = f"{args.location}-documentai.googleapis.com"
    client_options = ClientOptions(api_endpoint=endpoint)
    client = documentai.DocumentProcessorServiceClient(client_options=client_options)
    processor_name = client.processor_path(args.project_id, args.location, args.processor_id)

    with zipfile.ZipFile(args.zip) as archive:
        image_names = sorted(
            [
                name
                for name in archive.namelist()
                if name.lower().endswith((".jpg", ".jpeg", ".png"))
            ],
            key=lambda value: Path(value).name.casefold(),
        )
        if args.limit > 0:
            image_names = image_names[: args.limit]

        results = []
        for index, name in enumerate(image_names, start=1):
            print(f"[{index}/{len(image_names)}] {Path(name).name}")
            result = process_image(client, processor_name, Path(name).name, archive.read(name))
            results.append(result)

    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(
        json.dumps(
            {
                "processor": processor_name,
                "image_count": len(results),
                "results": results,
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )
    print(f"Đã lưu OCR: {args.output}")


if __name__ == "__main__":
    main()
