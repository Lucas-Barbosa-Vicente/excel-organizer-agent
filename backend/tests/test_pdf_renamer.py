import io
import zipfile
import pandas as pd
import pytest

from app.services.pdf_renamer import (
    normalize,
    load_employees,
    find_match,
    build_renamed_zip,
)


def make_excel(rows: list[dict]) -> bytes:
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


# ── normalize ────────────────────────────────────────────────────────────────

def test_normalize_strips_accents():
    assert normalize("João") == "joao"

def test_normalize_lowercases():
    assert normalize("SILVA") == "silva"

def test_normalize_removes_special_chars():
    assert normalize("doc_00123.pdf") == "doc00123pdf"

def test_normalize_handles_numbers():
    assert normalize("00123") == "00123"


# ── load_employees ────────────────────────────────────────────────────────────

def test_load_employees_returns_list():
    data = make_excel([
        {"Nome": "Ana Silva", "Matricula": "001"},
        {"Nome": "Carlos Souza", "Matricula": "002"},
    ])
    employees = load_employees(data, name_col="Nome", id_col="Matricula")
    assert len(employees) == 2
    assert employees[0]["id"] == "001"
    assert employees[0]["name"] == "Ana Silva"
    assert employees[0]["norm_id"] == "001"
    assert employees[0]["norm_name"] == "ana silva"

def test_load_employees_raises_on_missing_column():
    data = make_excel([{"Nome": "Ana", "Matricula": "001"}])
    with pytest.raises(KeyError):
        load_employees(data, name_col="NomeErrado", id_col="Matricula")


# ── find_match ────────────────────────────────────────────────────────────────

EMPLOYEES = [
    {"id": "001", "name": "Ana Silva", "norm_id": "001", "norm_name": "ana silva"},
    {"id": "002", "name": "Carlos Souza", "norm_id": "002", "norm_name": "carlos souza"},
]

def test_find_match_by_id():
    result = find_match("holerite_001_jan", EMPLOYEES)
    assert result is not None
    assert result["id"] == "001"

def test_find_match_by_name():
    result = find_match("holerite_carlos_souza", EMPLOYEES)
    assert result is not None
    assert result["id"] == "002"

def test_find_match_by_name_with_accent():
    employees = [{"id": "003", "name": "João Müller", "norm_id": "003", "norm_name": "joao muller"}]
    result = find_match("holerite_joao_muller", employees)
    assert result is not None

def test_find_match_returns_none_when_no_match():
    result = find_match("documento_xyz_9999", EMPLOYEES)
    assert result is None

def test_find_match_id_priority_over_name():
    employees = [
        {"id": "001", "name": "Outro Nome", "norm_id": "001", "norm_name": "outro nome"},
        {"id": "002", "name": "Doc001", "norm_id": "002", "norm_name": "doc001"},
    ]
    result = find_match("doc_001", employees)
    assert result["id"] == "001"


# ── build_renamed_zip ─────────────────────────────────────────────────────────

def test_build_renamed_zip_renames_matched():
    pdf_files = [("holerite_001.pdf", b"%PDF matched")]
    zip_bytes, matched, unmatched, unmatched_files = build_renamed_zip(
        pdf_files, EMPLOYEES, "{id} - {name}"
    )
    assert matched == 1
    assert unmatched == 0
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = zf.namelist()
    assert "001 - Ana Silva.pdf" in names

def test_build_renamed_zip_tracks_unmatched():
    pdf_files = [("sem_correspondencia.pdf", b"%PDF nope")]
    zip_bytes, matched, unmatched, unmatched_files = build_renamed_zip(
        pdf_files, EMPLOYEES, "{id} - {name}"
    )
    assert matched == 0
    assert unmatched == 1
    assert "sem_correspondencia.pdf" in unmatched_files

def test_build_renamed_zip_deduplicates_names():
    employees = [
        {"id": "001", "name": "Ana Silva", "norm_id": "001", "norm_name": "ana silva"},
        {"id": "001", "name": "Ana Silva", "norm_id": "001", "norm_name": "ana silva"},
    ]
    pdf_files = [
        ("doc_001_jan.pdf", b"%PDF a"),
        ("doc_001_fev.pdf", b"%PDF b"),
    ]
    zip_bytes, matched, _, _ = build_renamed_zip(pdf_files, employees, "{id} - {name}")
    assert matched == 2
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = zf.namelist()
    assert len(names) == 2
    assert len(set(names)) == 2  # no duplicate ZIP entries

def test_build_renamed_zip_custom_pattern():
    pdf_files = [("arq_001.pdf", b"%PDF")]
    zip_bytes, matched, _, _ = build_renamed_zip(pdf_files, EMPLOYEES, "{name}")
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        names = zf.namelist()
    assert "Ana Silva.pdf" in names


def test_build_renamed_zip_sanitizes_path_traversal():
    employees = [
        {"id": "../../evil", "name": "Legit Name", "norm_id": "evil", "norm_name": "legit name"},
    ]
    pdf_files = [("holerite_evil.pdf", b"%PDF")]
    zip_bytes, matched, _, _ = build_renamed_zip(pdf_files, employees, "{id} - {name}")
    with zipfile.ZipFile(io.BytesIO(zip_bytes)) as zf:
        for name in zf.namelist():
            assert ".." not in name
            assert name == name.lstrip("/\\")
