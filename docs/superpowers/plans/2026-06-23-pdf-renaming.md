# PDF Renaming Feature — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a `POST /api/rename-pdfs` endpoint that accepts an Excel employee roster and multiple PDFs, matches each PDF to an employee by matrícula or name substring, and returns a ZIP of renamed files.

**Architecture:** A new service `pdf_renamer.py` contains pure matching + ZIP logic (no I/O, easily testable). A new router `rename.py` handles HTTP concerns. The existing `main.py` registers the router. No new dependencies — `zipfile`, `unicodedata`, and `re` are stdlib; `pandas` and `openpyxl` are already installed.

**Tech Stack:** FastAPI, Python 3.10, pandas 2.2, openpyxl 3.1, zipfile (stdlib), unicodedata (stdlib).

## Global Constraints

- Run all tests from `backend/` directory: `python -m pytest tests/ -v`
- Python 3.10 — no `match` statements, use `str | None` type hints only in stdlib contexts; use `Optional` from `typing` in Pydantic schemas
- All new files live under `backend/app/`
- Follow existing code style: `asyncio.to_thread` for blocking I/O, `logger = logging.getLogger(__name__)`, snake_case
- Commit after every task with the exact commands shown

---

## File Map

| Action | File | Responsibility |
|---|---|---|
| Create | `backend/app/schemas/rename.py` | Pydantic response model |
| Create | `backend/app/services/pdf_renamer.py` | normalize, match, ZIP logic |
| Create | `backend/app/routes/rename.py` | FastAPI router (`/rename-pdfs`, `/download-zip/{token}`) |
| Create | `backend/tests/test_pdf_renamer.py` | Unit tests for service |
| Create | `backend/tests/test_rename_route.py` | Integration tests for route |
| Modify | `backend/main.py` | Register new router |

---

## Task 1: Pydantic schema + pure service logic

**Files:**
- Create: `backend/app/schemas/rename.py`
- Create: `backend/app/services/pdf_renamer.py`
- Create: `backend/tests/test_pdf_renamer.py`

**Interfaces:**
- Produces:
  - `normalize(text: str) -> str`
  - `load_employees(excel_bytes: bytes, name_col: str, id_col: str) -> list[dict]`
    - Each dict: `{"id": str, "name": str, "norm_id": str, "norm_name": str}`
  - `find_match(filename_stem: str, employees: list[dict]) -> dict | None`
  - `build_renamed_zip(pdf_files: list[tuple[str, bytes]], employees: list[dict], output_pattern: str) -> tuple[bytes, int, int, list[str]]`
    - Returns `(zip_bytes, matched_count, unmatched_count, unmatched_filenames)`
  - `RenameResponse` Pydantic model

---

- [ ] **Step 1: Write the failing tests**

Create `backend/tests/test_pdf_renamer.py`:

```python
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
```

- [ ] **Step 2: Run tests — verify they fail with ImportError**

```
cd backend
python -m pytest tests/test_pdf_renamer.py -v
```

Expected: `ImportError: cannot import name 'normalize' from 'app.services.pdf_renamer'` (module does not exist yet)

- [ ] **Step 3: Create `backend/app/schemas/rename.py`**

```python
from typing import List, Optional
from pydantic import BaseModel


class RenameResponse(BaseModel):
    status: str
    download_token: Optional[str] = None
    matched: int
    unmatched: int
    unmatched_files: List[str]
    total: int
```

- [ ] **Step 4: Create `backend/app/services/pdf_renamer.py`**

```python
import io
import re
import unicodedata
import zipfile
from typing import Optional

import pandas as pd


def normalize(text: str) -> str:
    text = unicodedata.normalize("NFD", str(text))
    text = "".join(c for c in text if unicodedata.category(c) != "Mn")
    return re.sub(r"[^a-z0-9 ]", "", text.lower())


def load_employees(excel_bytes: bytes, name_col: str, id_col: str) -> list[dict]:
    df = pd.read_excel(io.BytesIO(excel_bytes))
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
    norm_stem = normalize(filename_stem)
    for emp in employees:
        if emp["norm_id"] and emp["norm_id"] in norm_stem:
            return emp
    for emp in employees:
        if emp["norm_name"] and emp["norm_name"] in norm_stem:
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
            new_stem = output_pattern.format(id=emp["id"], name=emp["name"])
            if new_stem in used_names:
                used_names[new_stem] += 1
                new_stem = f"{new_stem} ({used_names[new_stem]})"
            else:
                used_names[new_stem] = 0
            zf.writestr(f"{new_stem}.pdf", content)
            matched += 1

    return buf.getvalue(), matched, unmatched, unmatched_files
```

- [ ] **Step 5: Run tests — verify they pass**

```
cd backend
python -m pytest tests/test_pdf_renamer.py -v
```

Expected output (all green):
```
tests/test_pdf_renamer.py::test_normalize_strips_accents PASSED
tests/test_pdf_renamer.py::test_normalize_lowercases PASSED
tests/test_pdf_renamer.py::test_normalize_removes_special_chars PASSED
tests/test_pdf_renamer.py::test_normalize_handles_numbers PASSED
tests/test_pdf_renamer.py::test_load_employees_returns_list PASSED
tests/test_pdf_renamer.py::test_load_employees_raises_on_missing_column PASSED
tests/test_pdf_renamer.py::test_find_match_by_id PASSED
tests/test_pdf_renamer.py::test_find_match_by_name PASSED
tests/test_pdf_renamer.py::test_find_match_by_name_with_accent PASSED
tests/test_pdf_renamer.py::test_find_match_returns_none_when_no_match PASSED
tests/test_pdf_renamer.py::test_find_match_id_priority_over_name PASSED
tests/test_pdf_renamer.py::test_build_renamed_zip_renames_matched PASSED
tests/test_pdf_renamer.py::test_build_renamed_zip_tracks_unmatched PASSED
tests/test_pdf_renamer.py::test_build_renamed_zip_deduplicates_names PASSED
tests/test_pdf_renamer.py::test_build_renamed_zip_custom_pattern PASSED
15 passed
```

- [ ] **Step 6: Commit**

```bash
git add backend/app/schemas/rename.py backend/app/services/pdf_renamer.py backend/tests/test_pdf_renamer.py
git commit -m "feat: adiciona schema e serviço de renomeação de PDFs"
```

---

## Task 2: FastAPI route + main.py registration

**Files:**
- Create: `backend/app/routes/rename.py`
- Create: `backend/tests/test_rename_route.py`
- Modify: `backend/main.py` (lines 10 and 61)

**Interfaces:**
- Consumes (from Task 1):
  - `load_employees(excel_bytes, name_col, id_col) -> list[dict]`
  - `build_renamed_zip(pdf_files, employees, output_pattern) -> tuple[bytes, int, int, list[str]]`
  - `RenameResponse`
- Produces:
  - `POST /api/rename-pdfs` → `RenameResponse`
  - `GET /api/download-zip/{token}` → ZIP stream (single-use token)

---

- [ ] **Step 1: Write the failing integration tests**

Create `backend/tests/test_rename_route.py`:

```python
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
```

- [ ] **Step 2: Run tests — verify they fail**

```
cd backend
python -m pytest tests/test_rename_route.py -v
```

Expected: errors like `404 Not Found` for `/api/rename-pdfs` (route not registered yet)

- [ ] **Step 3: Create `backend/app/routes/rename.py`**

```python
import asyncio
import io
import uuid
from typing import List

from fastapi import APIRouter, File, Form, HTTPException, UploadFile
from fastapi.responses import StreamingResponse

from app.schemas.rename import RenameResponse
from app.services.pdf_renamer import build_renamed_zip, load_employees

router = APIRouter()

_zip_map: dict[str, bytes] = {}


@router.post("/rename-pdfs", response_model=RenameResponse)
async def rename_pdfs(
    excel_file: UploadFile = File(...),
    pdf_files: List[UploadFile] = File(...),
    name_column: str = Form(...),
    id_column: str = Form(...),
    output_pattern: str = Form("{id} - {name}"),
):
    excel_bytes = await excel_file.read()
    try:
        employees = await asyncio.to_thread(
            load_employees, excel_bytes, name_column, id_column
        )
    except KeyError as e:
        raise HTTPException(status_code=422, detail=f"Coluna não encontrada na planilha: {e}")
    except Exception as e:
        raise HTTPException(status_code=422, detail=f"Erro ao ler planilha: {e}")

    pdf_data: list[tuple[str, bytes]] = []
    for pdf in pdf_files:
        content = await pdf.read()
        pdf_data.append((pdf.filename or "arquivo.pdf", content))

    zip_bytes, matched, unmatched, unmatched_files = await asyncio.to_thread(
        build_renamed_zip, pdf_data, employees, output_pattern
    )

    if matched == 0:
        raise HTTPException(
            status_code=400,
            detail=(
                "Nenhum PDF pôde ser associado a um funcionário. "
                "Verifique se os nomes dos arquivos contêm a matrícula ou o nome dos funcionários."
            ),
        )

    token = str(uuid.uuid4())
    _zip_map[token] = zip_bytes

    return RenameResponse(
        status="success",
        download_token=token,
        matched=matched,
        unmatched=unmatched,
        unmatched_files=unmatched_files,
        total=matched + unmatched,
    )


@router.get("/download-zip/{token}")
async def download_zip(token: str):
    zip_bytes = _zip_map.pop(token, None)
    if zip_bytes is None:
        raise HTTPException(
            status_code=404, detail="Token inválido ou ZIP já baixado."
        )
    return StreamingResponse(
        io.BytesIO(zip_bytes),
        media_type="application/zip",
        headers={"Content-Disposition": "attachment; filename=pdfs_renomeados.zip"},
    )
```

- [ ] **Step 4: Register the router in `backend/main.py`**

Add the import on line 10 (after the existing route imports):

```python
from app.routes import health, organize, profiles, rename
```

Add the router registration at the bottom of the file (after line 61):

```python
app.include_router(rename.router, prefix="/api", tags=["Renomear PDFs"])
```

The final bottom of `main.py` should look like:

```python
app.include_router(health.router, tags=["Saúde"])
app.include_router(organize.router, prefix="/api", tags=["Organizar"])
app.include_router(profiles.router, prefix="/api", tags=["Perfis"])
app.include_router(rename.router, prefix="/api", tags=["Renomear PDFs"])
```

- [ ] **Step 5: Run all tests — verify everything passes**

```
cd backend
python -m pytest tests/ -v
```

Expected: all 23 tests pass (8 existing + 15 service + 7 route tests with no failures).

- [ ] **Step 6: Commit**

```bash
git add backend/app/routes/rename.py backend/app/schemas/rename.py backend/tests/test_rename_route.py backend/main.py
git commit -m "feat: adiciona rota de renomeação de PDFs via planilha Excel"
```

---

## Self-Review Notes

**Spec coverage check:**
- ✅ `POST /api/rename-pdfs` with all form fields (excel_file, pdf_files, name_column, id_column, output_pattern)
- ✅ `GET /api/download-zip/{token}` single-use download
- ✅ Matching by matrícula (exact substring, priority) then by name (normalized)
- ✅ ZIP output with renamed files
- ✅ Response includes matched/unmatched/total/unmatched_files/download_token
- ✅ `400` when zero matches
- ✅ `422` on unreadable Excel / missing column
- ✅ Duplicate output names get index suffix
- ✅ `instruction` param listed in spec but marked optional for AI-based column detection — **not implemented** (deferred per spec's "Out of Scope" for multi-step AI matching; simple column name passing covers the use case)

**Placeholder scan:** None found.

**Type consistency:** `load_employees` returns `list[dict]`, consumed identically in `pdf_renamer.py` and `rename.py`. `build_renamed_zip` signature matches between tasks.
