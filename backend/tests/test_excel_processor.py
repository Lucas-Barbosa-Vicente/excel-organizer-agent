import asyncio
import os
import tempfile
import pytest
import pandas as pd
import openpyxl

from app.schemas.organize import OrganizeParameters, SortRule, ColorRule
from app.services.excel_processor import ExcelProcessor


def make_xlsx(df: pd.DataFrame) -> str:
    tmp = tempfile.NamedTemporaryFile(suffix=".xlsx", delete=False)
    df.to_excel(tmp.name, index=False)
    tmp.close()
    return tmp.name


def run(coro):
    return asyncio.get_event_loop().run_until_complete(coro)


# ─── helpers ──────────────────────────────────────────────────────────────────

def make_processor(df: pd.DataFrame) -> ExcelProcessor:
    path = make_xlsx(df)
    proc = ExcelProcessor(path)
    run(proc.load())
    return proc


# ─── tests ────────────────────────────────────────────────────────────────────

def test_sort_single_column():
    df = pd.DataFrame({"Nome": ["Carlos", "Ana", "Bruno"], "Idade": [30, 25, 28]})
    proc = make_processor(df)
    params = OrganizeParameters(sort_by=[SortRule(column="Nome", direction="asc")])
    run(proc.apply_parameters(params))
    sheet = list(proc.dataframes.values())[0]
    assert list(sheet["Nome"]) == ["Ana", "Bruno", "Carlos"]


def test_sort_multiple_columns():
    df = pd.DataFrame({
        "Dep": ["TI", "TI", "RH"],
        "Nome": ["Carlos", "Ana", "Bruno"],
    })
    proc = make_processor(df)
    params = OrganizeParameters(sort_by=[
        SortRule(column="Dep", direction="asc"),
        SortRule(column="Nome", direction="asc"),
    ])
    run(proc.apply_parameters(params))
    sheet = list(proc.dataframes.values())[0]
    assert list(sheet["Nome"]) == ["Bruno", "Ana", "Carlos"]


def test_remove_duplicates():
    df = pd.DataFrame({"Email": ["a@x.com", "b@x.com", "a@x.com"], "Nome": ["A", "B", "A2"]})
    proc = make_processor(df)
    params = OrganizeParameters(remove_duplicates=True, duplicate_columns=["Email"], keep_duplicate="first")
    stats = run(proc.apply_parameters(params))
    assert stats["rows_after"] == 2
    assert "Duplicatas removidas: 1" in proc.transformations[0]


def test_standardize_text_upper():
    df = pd.DataFrame({"Nome": ["ana", "Bruno"]})
    proc = make_processor(df)
    params = OrganizeParameters(standardize_text={"Nome": "upper"})
    run(proc.apply_parameters(params))
    sheet = list(proc.dataframes.values())[0]
    assert list(sheet["Nome"]) == ["ANA", "BRUNO"]


def test_standardize_text_lower():
    df = pd.DataFrame({"Nome": ["ANA", "BRUNO"]})
    proc = make_processor(df)
    params = OrganizeParameters(standardize_text={"Nome": "lower"})
    run(proc.apply_parameters(params))
    sheet = list(proc.dataframes.values())[0]
    assert list(sheet["Nome"]) == ["ana", "bruno"]


def test_color_rules_applied_without_error():
    df = pd.DataFrame({"Status": ["Ativo", "Inativo", "Ativo"], "Valor": [100, 200, 150]})
    proc = make_processor(df)
    params = OrganizeParameters(
        color_rules=[ColorRule(column="Status", operator="equals", value="Inativo", color="red")]
    )
    run(proc.apply_parameters(params))
    assert any("cor" in t.lower() for t in proc.transformations)


def test_split_by_category():
    df = pd.DataFrame({
        "Dep": ["TI", "RH", "TI", "RH"],
        "Nome": ["A", "B", "C", "D"],
    })
    proc = make_processor(df)
    params = OrganizeParameters(split_by_category="Dep", keep_original_sheet=True)
    stats = run(proc.apply_parameters(params))
    assert "TI" in stats["sheets_created"]
    assert "RH" in stats["sheets_created"]


def test_large_file_performance():
    import time
    rows = 10_001
    df = pd.DataFrame({
        "Nome": [f"Nome_{i}" for i in range(rows)],
        "Valor": range(rows),
    })
    proc = make_processor(df)
    params = OrganizeParameters(
        sort_by=[SortRule(column="Nome", direction="asc")],
        remove_duplicates=True,
    )
    start = time.time()
    run(proc.apply_parameters(params))
    elapsed = time.time() - start
    assert elapsed < 30, f"Processamento demorou {elapsed:.1f}s (limite: 30s)"
