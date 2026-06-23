import io
import zipfile
import pandas as pd
import pytest
from fastapi.testclient import TestClient

from main import app

client = TestClient(app)


def make_excel_bytes(rows: list[dict]) -> bytes:
    df = pd.DataFrame(rows)
    buf = io.BytesIO()
    df.to_excel(buf, index=False)
    return buf.getvalue()


def make_pdf_bytes(content: str = "%PDF-1.4 fake") -> bytes:
    return content.encode()


# ── POST /api/rename-pdfs ─────────────────────────────────────────────────────

def test_rename_pdfs_success():
    excel = make_excel_bytes([
        {"Nome": "Ana Silva", "Matricula": "001"},
        {"Nome": "Carlos Souza", "Matricula": "002"},
    ])
    response = client.post(
        "/api/rename-pdfs",
        data={"name_column": "Nome", "id_column": "Matricula"},
        files=[
            ("excel_file", ("funcionarios.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("pdf_files", ("holerite_001.pdf", make_pdf_bytes(), "application/pdf")),
            ("pdf_files", ("holerite_002.pdf", make_pdf_bytes(), "application/pdf")),
        ],
    )
    assert response.status_code == 200
    body = response.json()
    assert body["status"] == "success"
    assert body["matched"] == 2
    assert body["unmatched"] == 0
    assert body["total"] == 2
    assert "download_token" in body


def test_rename_pdfs_partial_match():
    excel = make_excel_bytes([{"Nome": "Ana Silva", "Matricula": "001"}])
    response = client.post(
        "/api/rename-pdfs",
        data={"name_column": "Nome", "id_column": "Matricula"},
        files=[
            ("excel_file", ("funcionarios.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("pdf_files", ("holerite_001.pdf", make_pdf_bytes(), "application/pdf")),
            ("pdf_files", ("desconhecido.pdf", make_pdf_bytes(), "application/pdf")),
        ],
    )
    assert response.status_code == 200
    body = response.json()
    assert body["matched"] == 1
    assert body["unmatched"] == 1
    assert "desconhecido.pdf" in body["unmatched_files"]


def test_rename_pdfs_all_unmatched_returns_400():
    excel = make_excel_bytes([{"Nome": "Ana Silva", "Matricula": "001"}])
    response = client.post(
        "/api/rename-pdfs",
        data={"name_column": "Nome", "id_column": "Matricula"},
        files=[
            ("excel_file", ("funcionarios.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("pdf_files", ("xyz_nenhum.pdf", make_pdf_bytes(), "application/pdf")),
        ],
    )
    assert response.status_code == 400


def test_rename_pdfs_invalid_column_returns_422():
    excel = make_excel_bytes([{"Nome": "Ana", "Matricula": "001"}])
    response = client.post(
        "/api/rename-pdfs",
        data={"name_column": "ColunaErrada", "id_column": "Matricula"},
        files=[
            ("excel_file", ("func.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("pdf_files", ("doc_001.pdf", make_pdf_bytes(), "application/pdf")),
        ],
    )
    assert response.status_code == 422


# ── GET /api/download-zip/{token} ─────────────────────────────────────────────

def test_download_zip_returns_valid_zip():
    excel = make_excel_bytes([{"Nome": "Ana Silva", "Matricula": "001"}])
    post = client.post(
        "/api/rename-pdfs",
        data={"name_column": "Nome", "id_column": "Matricula"},
        files=[
            ("excel_file", ("f.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("pdf_files", ("holerite_001.pdf", make_pdf_bytes(), "application/pdf")),
        ],
    )
    token = post.json()["download_token"]
    dl = client.get(f"/api/download-zip/{token}")
    assert dl.status_code == 200
    assert dl.headers["content-type"] == "application/zip"
    with zipfile.ZipFile(io.BytesIO(dl.content)) as zf:
        names = zf.namelist()
    assert "001 - Ana Silva.pdf" in names


def test_download_zip_token_single_use():
    excel = make_excel_bytes([{"Nome": "Ana Silva", "Matricula": "001"}])
    post = client.post(
        "/api/rename-pdfs",
        data={"name_column": "Nome", "id_column": "Matricula"},
        files=[
            ("excel_file", ("f.xlsx", excel, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")),
            ("pdf_files", ("holerite_001.pdf", make_pdf_bytes(), "application/pdf")),
        ],
    )
    token = post.json()["download_token"]
    client.get(f"/api/download-zip/{token}")  # consume it
    second = client.get(f"/api/download-zip/{token}")
    assert second.status_code == 404


def test_download_zip_invalid_token():
    response = client.get("/api/download-zip/token-invalido-xyz")
    assert response.status_code == 404
