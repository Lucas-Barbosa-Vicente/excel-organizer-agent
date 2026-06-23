import io
import os
import re
import unicodedata
import zipfile
from typing import Optional

import pandas as pd

MAX_PDF_BYTES = 20 * 1024 * 1024  # 20 MB per PDF


def _sanitize_filename_part(value: str) -> str:
    """Remove path separators and parent-dir sequences from a filename component."""
    value = value.replace(os.sep, "_").replace("/", "_").replace("\\", "_")
    value = value.replace("..", "_")
    return value.strip()


def normalize(text: str) -> str:
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9 ]", "", text.lower())


def load_employees(excel_bytes: bytes, name_col: str, id_col: str) -> list[dict]:
    df = pd.read_excel(io.BytesIO(excel_bytes), dtype={id_col: str})
    employees = []
    for _, row in df.iterrows():
        emp_id = str(row[id_col]).strip()
        emp_name = str(row[name_col]).strip()
        employees.append({
            "id": emp_id,
            "name": emp_name,
            "norm_id": normalize(emp_id),
            "norm_name": normalize(emp_name),
        })
    return employees


def find_match(filename_stem: str, employees: list[dict]) -> Optional[dict]:
    norm_stem = normalize(filename_stem).replace(" ", "")
    for emp in employees:
        norm_id = emp["norm_id"].replace(" ", "")
        if norm_id and norm_id in norm_stem:
            return emp
    for emp in employees:
        norm_name = emp["norm_name"].replace(" ", "")
        if norm_name and norm_name in norm_stem:
            return emp
    return None


def build_renamed_zip(
    pdf_files: list[tuple[str, bytes]],
    employees: list[dict],
    output_pattern: str = "{id} - {name}",
) -> tuple[bytes, int, int, list[str]]:
    buf = io.BytesIO()
    matched = 0
    unmatched = 0
    unmatched_files: list[str] = []
    used_names: dict[str, int] = {}

    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        for orig_name, content in pdf_files:
            stem = orig_name.rsplit(".", 1)[0]
            emp = find_match(stem, employees)
            if emp is None:
                unmatched += 1
                unmatched_files.append(orig_name)
                continue
            safe_id = _sanitize_filename_part(emp["id"])
            safe_name = _sanitize_filename_part(emp["name"])
            new_stem = output_pattern.format(id=safe_id, name=safe_name)
            if new_stem in used_names:
                used_names[new_stem] += 1
                new_stem = f"{new_stem} ({used_names[new_stem]})"
            else:
                used_names[new_stem] = 0
            zf.writestr(f"{new_stem}.pdf", content)
            matched += 1

    return buf.getvalue(), matched, unmatched, unmatched_files
