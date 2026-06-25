import io
import json
import os
import zipfile
import tempfile
from unittest.mock import patch

import pandas as pd
import pytest


# ── extrair_zip ───────────────────────────────────────────────────────────────

def test_extrair_zip_extrai_pdfs(tmp_path):
    from email_reader import extrair_zip

    zip_path = tmp_path / "holerites.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("000003 EMILCE.pdf", b"%PDF")
        zf.writestr("000007 MARCIA.pdf", b"%PDF")

    nomes = extrair_zip(str(zip_path), str(tmp_path / "HOLERITES"))

    assert "000003 EMILCE.pdf" in nomes
    assert "000007 MARCIA.pdf" in nomes
    assert (tmp_path / "HOLERITES" / "000003 EMILCE.pdf").exists()


def test_extrair_zip_cria_pasta_destino(tmp_path):
    from email_reader import extrair_zip

    zip_path = tmp_path / "h.zip"
    with zipfile.ZipFile(zip_path, "w") as zf:
        zf.writestr("doc.pdf", b"%PDF")

    destino = str(tmp_path / "nova_pasta")
    extrair_zip(str(zip_path), destino)
    assert os.path.isdir(destino)


# ── carregar_emails ───────────────────────────────────────────────────────────

def make_email_xlsx(rows: list, path: str):
    pd.DataFrame(rows).to_excel(path, index=False)


def test_carregar_emails_retorna_dict(tmp_path):
    from email_sender import carregar_emails

    planilha = str(tmp_path / "emails.xlsx")
    make_email_xlsx([
        {"Matricula": "000003", "Email": "emilce@empresa.com"},
        {"Matricula": "000007", "Email": "marcia@empresa.com"},
    ], planilha)

    resultado = carregar_emails(planilha)
    assert resultado["000003"] == "emilce@empresa.com"
    assert resultado["000007"] == "marcia@empresa.com"


def test_carregar_emails_normaliza_matricula_com_zeros(tmp_path):
    from email_sender import carregar_emails

    planilha = str(tmp_path / "emails.xlsx")
    make_email_xlsx([{"Matricula": 3, "Email": "emilce@empresa.com"}], planilha)

    resultado = carregar_emails(planilha)
    assert "000003" in resultado


# ── montar_envios ─────────────────────────────────────────────────────────────

def test_montar_envios_casa_matricula(tmp_path):
    from email_sender import montar_envios

    pasta = tmp_path / "HOLERITES"
    pasta.mkdir()
    (pasta / "A000003_Emilce_Gomes_Holerite_202601.pdf").write_bytes(b"%PDF")

    lookup = {"000003": "emilce@empresa.com"}
    envios = montar_envios(str(pasta), lookup)

    assert len(envios) == 1
    assert envios[0]["matricula"] == "000003"
    assert envios[0]["email"] == "emilce@empresa.com"
    assert "Emilce" in envios[0]["nome"]


def test_montar_envios_email_none_quando_nao_encontrado(tmp_path):
    from email_sender import montar_envios

    pasta = tmp_path / "HOLERITES"
    pasta.mkdir()
    (pasta / "A000099_Pedro_Ingro_Holerite_202601.pdf").write_bytes(b"%PDF")

    envios = montar_envios(str(pasta), {})
    assert envios[0]["email"] is None


def test_montar_envios_ignora_nao_pdf(tmp_path):
    from email_sender import montar_envios

    pasta = tmp_path / "HOLERITES"
    pasta.mkdir()
    (pasta / "A000003_Emilce_Gomes_Holerite_202601.pdf").write_bytes(b"%PDF")
    (pasta / "readme.txt").write_text("ignore")

    envios = montar_envios(str(pasta), {"000003": "e@e.com"})
    assert len(envios) == 1


# ── substituir_placeholders ───────────────────────────────────────────────────

def test_substituir_placeholders_preenche_todos():
    from email_sender import substituir_placeholders

    resultado = substituir_placeholders(
        "Olá {nome}, holerite {mes}/{ano}.", "Ana", "01", "2026"
    )
    assert resultado == "Olá Ana, holerite 01/2026."


def test_substituir_placeholders_sem_placeholder():
    from email_sender import substituir_placeholders

    resultado = substituir_placeholders("Sem nada.", "Ana", "01", "2026")
    assert resultado == "Sem nada."


# ── carregar_config ───────────────────────────────────────────────────────────

def test_carregar_config_le_json(tmp_path):
    from pipeline import carregar_config

    cfg = {"assunto_email": "Teste {mes}", "corpo_email": "Corpo"}
    path = str(tmp_path / "config.json")
    with open(path, "w") as f:
        json.dump(cfg, f)

    resultado = carregar_config(path)
    assert resultado["assunto_email"] == "Teste {mes}"


# ── exibir_conferencia ────────────────────────────────────────────────────────

def test_exibir_conferencia_retorna_true_em_s(capsys):
    from pipeline import exibir_conferencia

    envios = [
        {"nome": "Ana Silva", "matricula": "000003",
         "email": "ana@e.com", "arquivo": "/HOLERITES/A000003_Ana_Silva_Holerite_202601.pdf"},
    ]
    with patch("builtins.input", return_value="s"):
        assert exibir_conferencia(envios) is True


def test_exibir_conferencia_retorna_false_em_n(capsys):
    from pipeline import exibir_conferencia

    envios = [
        {"nome": "Ana Silva", "matricula": "000003",
         "email": "ana@e.com", "arquivo": "/HOLERITES/A000003_Ana_Silva_Holerite_202601.pdf"},
    ]
    with patch("builtins.input", return_value="n"):
        assert exibir_conferencia(envios) is False


def test_exibir_conferencia_mostra_sem_email(capsys):
    from pipeline import exibir_conferencia

    envios = [
        {"nome": "Ana Silva", "matricula": "000003",
         "email": None, "arquivo": "/HOLERITES/A000003_Ana_Silva_Holerite_202601.pdf"},
    ]
    with patch("builtins.input", return_value="n"):
        exibir_conferencia(envios)
    out = capsys.readouterr().out
    assert "SEM E-MAIL" in out
