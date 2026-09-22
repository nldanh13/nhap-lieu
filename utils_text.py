import re
import unicodedata

import pandas as pd


def remove_accents(text):
    text = str(text)
    text = unicodedata.normalize("NFD", text)
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.replace("đ", "d").replace("Đ", "D")
    return text


def normalize_text(text):
    if pd.isna(text):
        return ""

    text = remove_accents(text).lower()
    text = re.sub(r"\s+", " ", text)
    text = text.strip()
    return text


def clean_header_value(x):
    if pd.isna(x):
        return ""

    s = str(x).replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)

    if s.startswith("Unnamed:"):
        return ""

    return s


def contains_text(cell_value, keyword):
    cell_norm = normalize_text(cell_value)
    keyword_norm = normalize_text(keyword)
    return bool(keyword_norm) and keyword_norm in cell_norm


def normalize_alias_text(text):
    """Chuẩn hóa bí danh nhưng GIỮ DẤU để tránh nhầm TẤN và TÂN."""
    if pd.isna(text):
        return ""
    s = str(text).upper().replace("\n", " ").strip()
    s = re.sub(r"\s+", " ", s)
    return s


def contains_alias_text(cell_value, alias):
    cell = normalize_alias_text(cell_value)
    alias_norm = normalize_alias_text(alias)
    return bool(alias_norm) and alias_norm in cell
