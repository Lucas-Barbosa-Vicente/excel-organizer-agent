# PDF Renaming Feature — Design Spec
**Date:** 2026-06-23  
**Status:** Approved

## Problem

The Excel Organizer Agent currently only processes Excel files. Users need to rename PDF files (e.g., pay stubs, HR documents) using employee data from an Excel spreadsheet, matching each PDF to an employee by name or registration number (matrícula).

## Assumptions

1. Each PDF filename contains either the employee's matrícula or name as a substring.
2. The Excel file has at least one column with employee names and one with matrícula numbers.
3. One PDF per employee (or multiple PDFs mapping to the same employee are all renamed).
4. Output is a ZIP archive of renamed PDFs.

## Approach

**Substring matching (automatic):** For each PDF, search for the matrícula (exact) or normalized name (case-insensitive, accent-stripped) as a substring in the filename. Matrícula matching takes priority.

The AI agent (Claude) identifies which Excel columns correspond to name and matrícula when the user provides a natural language instruction or when auto-detection is needed.

## API Contract

### `POST /api/rename-pdfs`

**Form fields:**
| Field | Type | Required | Description |
|---|---|---|---|
| `excel_file` | UploadFile (.xlsx) | yes | Employee data source |
| `pdf_files` | List[UploadFile] | yes | PDFs to rename (up to 200) |
| `name_column` | str | no | Excel column for employee name |
| `id_column` | str | no | Excel column for matrícula |
| `output_pattern` | str | no | Filename pattern (default: `{id} - {name}`) |
| `instruction` | str | no | Natural language for column auto-detection |

**Response:**
```json
{
  "status": "success",
  "download_token": "<uuid>",
  "matched": 45,
  "unmatched": 3,
  "unmatched_files": ["doc_999.pdf", "arquivo_xyz.pdf", "outro.pdf"],
  "total": 48
}
```

### `GET /api/download-zip/{token}`

Returns the ZIP archive of renamed PDFs. Token is single-use.

## Matching Algorithm

```
for each PDF:
  1. Extract filename stem (no extension)
  2. Normalize: strip accents, lowercase, remove special chars
  3. For each employee row in Excel:
     a. If matrícula value found as exact substring → match
     b. Else if normalized name found as substring → match
  4. If matched: add to ZIP as "{id} - {name}.pdf"
  5. If not matched: add to unmatched list
```

Ties (multiple employees match one PDF): first match wins; logged as a warning.

## New Files

| File | Purpose |
|---|---|
| `backend/app/routes/rename.py` | FastAPI router for `/rename-pdfs` and `/download-zip/{token}` |
| `backend/app/services/pdf_renamer.py` | Matching logic, ZIP creation |
| `backend/app/schemas/rename.py` | Pydantic schemas for request/response |

## Modified Files

| File | Change |
|---|---|
| `backend/main.py` | Register new router |
| `backend/requirements.txt` | Add `python-multipart` if not present (already used), no new deps needed — `zipfile` is stdlib |

## Error Handling

- No PDFs matched at all → `400` with descriptive message
- Excel file unreadable → `422`
- PDF file exceeds size limit → skip and add to unmatched list
- Duplicate output filenames (two employees map to same name) → append index suffix

## Out of Scope

- Frontend UI changes (API-first; existing add-in or HTTP client can call it)
- Editing PDF content
- Multi-step fuzzy matching via AI on each PDF (reserved for a future enhancement)
