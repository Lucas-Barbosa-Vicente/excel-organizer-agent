# Pipeline Automático de Holerites — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Automatizar o fluxo completo de holerites: baixar ZIP do Outlook → renomear PDFs → conferência → enviar cada PDF ao funcionário, tudo com duplo clique num .bat.

**Architecture:** Três módulos independentes (`email_reader.py`, `email_sender.py`, `pipeline.py`) orquestrados pelo `pipeline.py`. Outlook COM (win32com) para leitura e envio de e-mails, com imports lazy para manter testabilidade. `renomear_holerites.py` existente é reutilizado sem alteração via import de `carregar_nomes`.

**Tech Stack:** Python 3.10, pywin32 (Outlook COM), pandas 2.2, openpyxl 3.1, zipfile (stdlib), unittest.mock (stdlib).

## Global Constraints

- Python 3.10 — sem `match` statements, sem `X | None` (usar `Optional`)
- Todos os arquivos ficam na raiz do projeto (não há subpasta `src/`)
- Testes ficam em `tests/` na raiz; comando: `python -m pytest tests/ -v` da raiz do projeto
- `renomear_holerites.py` NÃO deve ser alterado
- `win32com` deve ser importado de forma lazy (dentro das funções que o usam) para permitir testes sem Outlook instalado
- Snake_case em todo o código
- Mensagens ao usuário sempre em português

---

## File Map

| Ação | Arquivo | Responsabilidade |
|---|---|---|
| Criar | `config.json` | Assunto e corpo do e-mail configuráveis |
| Criar | `email_reader.py` | Busca e-mail no Outlook, baixa e extrai ZIP |
| Criar | `email_sender.py` | Carrega e-mails de funcionários, envia PDFs via Outlook COM |
| Criar | `pipeline.py` | Orquestra os 4 passos; exibe conferência; main() |
| Criar | `RENOMEAR E ENVIAR HOLERITES.bat` | Ponto de entrada único para o fluxo completo |
| Criar | `tests/test_pure.py` | Testa todas as funções puras (sem Outlook) |
| Criar | `tests/test_outlook.py` | Testa funções COM com mocks |
| Manter | `renomear_holerites.py` | Sem alteração |
| Manter | `RENOMEAR HOLERITES.bat` | Mantido para uso avulso |

---

## Task 1: Funções puras + config.json

Cria `config.json` e todas as funções que não dependem do Outlook. Essas funções são testáveis diretamente.

**Files:**
- Create: `config.json`
- Create: `email_reader.py` (apenas `extrair_zip`)
- Create: `email_sender.py` (apenas `carregar_emails`, `montar_envios`, `substituir_placeholders`)
- Create: `pipeline.py` (apenas `carregar_config`, `exibir_conferencia`)
- Create: `tests/__init__.py`
- Create: `tests/test_pure.py`

**Interfaces produzidas:**
- `extrair_zip(caminho_zip: str, pasta_destino: str) -> list[str]`
- `carregar_emails(planilha: str) -> dict[str, str]` — `{matricula_6dig: email}`
- `montar_envios(pasta_pdfs: str, lookup_emails: dict) -> list[dict]` — cada dict: `{nome, matricula, email, arquivo}`
- `substituir_placeholders(template: str, nome: str, mes: str, ano: str) -> str`
- `carregar_config(caminho: str) -> dict`
- `exibir_conferencia(envios: list[dict]) -> bool`

---

- [ ] **Step 1: Criar `config.json`**

```json
{
  "assunto_email": "Holerite {mes}/{ano}",
  "corpo_email": "Olá {nome},\n\nSegue em anexo seu holerite referente a {mes}/{ano}.\n\nAtenciosamente,\nRH"
}
```

- [ ] **Step 2: Escrever os testes (vão falhar — módulos não existem ainda)**

Criar `tests/__init__.py` (arquivo vazio) e `tests/test_pure.py`:

```python
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

def make_email_xlsx(rows: list[dict], path: str):
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
```

- [ ] **Step 3: Rodar testes — confirmar que falham com ImportError**

```
python -m pytest tests/test_pure.py -v
```

Esperado: `ImportError: No module named 'email_reader'` (módulos não existem ainda)

- [ ] **Step 4: Criar `email_reader.py` (apenas `extrair_zip`)**

```python
import os
import zipfile


def extrair_zip(caminho_zip: str, pasta_destino: str) -> list[str]:
    os.makedirs(pasta_destino, exist_ok=True)
    with zipfile.ZipFile(caminho_zip, "r") as zf:
        zf.extractall(pasta_destino)
        return zf.namelist()


def conectar_outlook():
    import win32com.client
    outlook = win32com.client.Dispatch("Outlook.Application")
    return outlook.GetNamespace("MAPI")


def buscar_emails_holerite(namespace) -> list:
    inbox = namespace.GetDefaultFolder(6)
    items = inbox.Items
    items.Sort("[ReceivedTime]", True)
    encontrados = []
    for item in items:
        try:
            if "holerite" in item.Subject.lower():
                encontrados.append(item)
        except AttributeError:
            continue
    return encontrados


def escolher_email(emails: list):
    if len(emails) == 1:
        return emails[0]
    print(f"\n{len(emails)} e-mails com holerite encontrados:\n")
    for i, mail in enumerate(emails, 1):
        recebido = mail.ReceivedTime.strftime("%d/%m/%Y %H:%M")
        print(f"  {i}. [{recebido}] {mail.Subject}")
    while True:
        escolha = input(f"\nQual deseja usar? (1-{len(emails)}): ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(emails):
            return emails[int(escolha) - 1]
        print("Opcao invalida, tente novamente.")


def baixar_e_extrair_zip(mail_item, pasta_destino: str) -> list[str]:
    os.makedirs(pasta_destino, exist_ok=True)
    for anexo in mail_item.Attachments:
        if anexo.FileName.lower().endswith(".zip"):
            caminho_zip = os.path.join(pasta_destino, "_temp_holerites.zip")
            anexo.SaveAsFile(caminho_zip)
            nomes = extrair_zip(caminho_zip, pasta_destino)
            os.remove(caminho_zip)
            return nomes
    raise FileNotFoundError("Nenhum arquivo ZIP encontrado nos anexos do e-mail.")
```

- [ ] **Step 5: Criar `email_sender.py` (funções puras)**

```python
import os
import re
import pandas as pd


def carregar_emails(planilha: str) -> dict:
    df = pd.read_excel(planilha, dtype={"Matricula": str})
    df["Matricula"] = df["Matricula"].str.strip().str.zfill(6)
    df["Email"] = df["Email"].str.strip()
    return dict(zip(df["Matricula"], df["Email"]))


def montar_envios(pasta_pdfs: str, lookup_emails: dict) -> list:
    pdfs = sorted(f for f in os.listdir(pasta_pdfs) if f.lower().endswith(".pdf"))
    envios = []
    for pdf in pdfs:
        m = re.search(r"[A-Z](\d{6})_([^_]+)_([^_]+)_", pdf)
        if not m:
            continue
        matricula = m.group(1)
        nome = f"{m.group(2)} {m.group(3)}"
        envios.append({
            "nome": nome,
            "matricula": matricula,
            "email": lookup_emails.get(matricula),
            "arquivo": os.path.join(pasta_pdfs, pdf),
        })
    return envios


def substituir_placeholders(template: str, nome: str, mes: str, ano: str) -> str:
    return (
        template
        .replace("{nome}", nome)
        .replace("{mes}", mes)
        .replace("{ano}", ano)
    )


def enviar_holerite(outlook_app, nome: str, email: str, arquivo: str,
                    assunto: str, corpo: str):
    mail = outlook_app.CreateItem(0)
    mail.To = email
    mail.Subject = assunto
    mail.Body = corpo
    mail.Attachments.Add(os.path.abspath(arquivo))
    mail.Send()


def enviar_todos(envios: list, config: dict, outlook_app) -> tuple:
    com_email = [e for e in envios if e["email"]]
    total = len(com_email)
    enviados = 0
    falhas = []

    for envio in com_email:
        m = re.search(r"_(\d{4})(\d{2})\.pdf$", envio["arquivo"])
        ano = m.group(1) if m else "?"
        mes = m.group(2) if m else "?"

        assunto = substituir_placeholders(
            config["assunto_email"], envio["nome"], mes, ano
        )
        corpo = substituir_placeholders(
            config["corpo_email"], envio["nome"], mes, ano
        )

        try:
            enviar_holerite(
                outlook_app, envio["nome"], envio["email"],
                envio["arquivo"], assunto, corpo
            )
            enviados += 1
            print(f"  Enviando {enviados}/{total}... {envio['nome']}")
        except Exception as e:
            falhas.append(envio["nome"])
            print(f"  ERRO ao enviar para {envio['nome']}: {e}")

    return enviados, falhas
```

- [ ] **Step 6: Criar `pipeline.py` (apenas funções puras)**

```python
import json
import os
import re
import sys


BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_PDFS      = os.path.join(BASE, "HOLERITES")
PLANILHA_NOMES  = os.path.join(BASE, "Holerites renomear.xlsx")
PLANILHA_EMAILS = os.path.join(BASE, "Emails funcionarios.xlsx")
CONFIG_PATH     = os.path.join(BASE, "config.json")


def carregar_config(caminho: str) -> dict:
    with open(caminho, encoding="utf-8") as f:
        return json.load(f)


def exibir_conferencia(envios: list) -> bool:
    com_email  = [e for e in envios if e["email"]]
    sem_email  = [e for e in envios if not e["email"]]

    print(f"\n{'FUNCIONARIO':<30} {'E-MAIL':<35} ARQUIVO")
    print("-" * 100)
    for e in envios:
        email_str = e["email"] or "[SEM E-MAIL]"
        arquivo   = os.path.basename(e["arquivo"])
        print(f"{e['nome']:<30} {email_str:<35} {arquivo}")

    print(f"\n{len(com_email)} prontos para envio", end="")
    if sem_email:
        nomes = ", ".join(e["nome"] for e in sem_email)
        print(f" | {len(sem_email)} sem e-mail: {nomes}", end="")
    print()

    return input("\nConfirmar envio? (s/n): ").strip().lower() == "s"


def _executar_renomeacao(pasta_pdfs: str, lookup_nomes: dict):
    """Rename PDFs without asking for user confirmation (used by pipeline)."""
    for pdf in os.listdir(pasta_pdfs):
        if not pdf.lower().endswith(".pdf"):
            continue
        m = re.match(r"^(\d{6})", pdf)
        if m and m.group(1) in lookup_nomes:
            os.rename(
                os.path.join(pasta_pdfs, pdf),
                os.path.join(pasta_pdfs, lookup_nomes[m.group(1)] + ".pdf"),
            )
```

- [ ] **Step 7: Rodar testes — confirmar que passam**

```
python -m pytest tests/test_pure.py -v
```

Esperado: 13 testes passando.

- [ ] **Step 8: Commit**

```bash
git add config.json email_reader.py email_sender.py pipeline.py tests/__init__.py tests/test_pure.py
git commit -m "feat: adiciona funcoes puras e config do pipeline de holerites"
```

---

## Task 2: Funções Outlook COM com mocks

Completa as funções que usam win32com e testa com mocks.

**Files:**
- Test: `tests/test_outlook.py`

**Interfaces consumidas (de Task 1):**
- `extrair_zip(caminho_zip, pasta_destino) -> list[str]`
- `enviar_holerite(outlook_app, nome, email, arquivo, assunto, corpo)`
- `enviar_todos(envios, config, outlook_app) -> tuple[int, list[str]]`

**Interfaces produzidas:**
- `conectar_outlook() -> mapi_namespace` — chama `win32com.client.Dispatch`
- `buscar_emails_holerite(namespace) -> list` — filtra caixa de entrada
- `escolher_email(emails) -> mail_item` — escolha quando há múltiplos
- `baixar_e_extrair_zip(mail_item, pasta_destino) -> list[str]`

---

- [ ] **Step 1: Escrever testes COM com mocks**

Criar `tests/test_outlook.py`:

```python
import os
import zipfile
from unittest.mock import MagicMock, patch, call
import pytest


# ── buscar_emails_holerite ────────────────────────────────────────────────────

def _make_mail(subject: str):
    item = MagicMock()
    item.Subject = subject
    return item


def _make_namespace(mail_items: list):
    item_iter = MagicMock()
    item_iter.__iter__ = MagicMock(return_value=iter(mail_items))
    item_iter.Sort = MagicMock()
    folder = MagicMock()
    folder.Items = item_iter
    ns = MagicMock()
    ns.GetDefaultFolder.return_value = folder
    return ns


def test_buscar_emails_holerite_filtra_por_assunto():
    from email_reader import buscar_emails_holerite

    ns = _make_namespace([
        _make_mail("Holerite Janeiro 2026"),
        _make_mail("Reunião de equipe"),
        _make_mail("HOLERITE Fevereiro"),
    ])
    resultado = buscar_emails_holerite(ns)
    assert len(resultado) == 2


def test_buscar_emails_holerite_lista_vazia_quando_nenhum():
    from email_reader import buscar_emails_holerite

    ns = _make_namespace([_make_mail("Sem relação")])
    assert buscar_emails_holerite(ns) == []


def test_buscar_emails_holerite_ignora_item_sem_subject():
    from email_reader import buscar_emails_holerite

    item_sem_subject = MagicMock(spec=[])  # sem atributo Subject
    ns = _make_namespace([item_sem_subject, _make_mail("Holerite OK")])
    resultado = buscar_emails_holerite(ns)
    assert len(resultado) == 1


# ── escolher_email ────────────────────────────────────────────────────────────

def test_escolher_email_retorna_direto_quando_unico():
    from email_reader import escolher_email

    mail = _make_mail("Holerite")
    assert escolher_email([mail]) is mail


def test_escolher_email_pede_escolha_com_multiplos():
    from email_reader import escolher_email

    mails = [_make_mail("Holerite Jan"), _make_mail("Holerite Fev")]
    mails[0].ReceivedTime.strftime.return_value = "01/01/2026 08:00"
    mails[1].ReceivedTime.strftime.return_value = "01/02/2026 08:00"

    with patch("builtins.input", return_value="2"):
        resultado = escolher_email(mails)

    assert resultado is mails[1]


# ── baixar_e_extrair_zip ──────────────────────────────────────────────────────

def test_baixar_e_extrair_zip_salva_e_extrai(tmp_path):
    from email_reader import baixar_e_extrair_zip

    # Cria um ZIP real em memória para o mock retornar
    import io as _io
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("000003 EMILCE.pdf", b"%PDF")
    zip_bytes = buf.getvalue()

    # Mock do mail_item com um anexo .zip
    anexo = MagicMock()
    anexo.FileName = "holerites.zip"
    def salvar_zip(caminho):
        with open(caminho, "wb") as f:
            f.write(zip_bytes)
    anexo.SaveAsFile.side_effect = salvar_zip

    mail = MagicMock()
    mail.Attachments.__iter__ = MagicMock(return_value=iter([anexo]))

    pasta = str(tmp_path / "HOLERITES")
    nomes = baixar_e_extrair_zip(mail, pasta)

    assert "000003 EMILCE.pdf" in nomes
    assert os.path.exists(os.path.join(pasta, "000003 EMILCE.pdf"))
    assert not os.path.exists(os.path.join(pasta, "_temp_holerites.zip"))


def test_baixar_e_extrair_zip_erro_sem_zip(tmp_path):
    from email_reader import baixar_e_extrair_zip

    anexo = MagicMock()
    anexo.FileName = "documento.pdf"
    mail = MagicMock()
    mail.Attachments.__iter__ = MagicMock(return_value=iter([anexo]))

    with pytest.raises(FileNotFoundError):
        baixar_e_extrair_zip(mail, str(tmp_path))


# ── enviar_holerite ───────────────────────────────────────────────────────────

def test_enviar_holerite_chama_send(tmp_path):
    from email_sender import enviar_holerite

    pdf = tmp_path / "A000003_Emilce_Gomes_Holerite_202601.pdf"
    pdf.write_bytes(b"%PDF")

    outlook = MagicMock()
    mail_mock = MagicMock()
    outlook.CreateItem.return_value = mail_mock

    enviar_holerite(outlook, "Emilce Gomes", "emilce@e.com",
                    str(pdf), "Holerite 01/2026", "Corpo do e-mail")

    mail_mock.Send.assert_called_once()
    assert mail_mock.To == "emilce@e.com"
    assert mail_mock.Subject == "Holerite 01/2026"


# ── enviar_todos ──────────────────────────────────────────────────────────────

def test_enviar_todos_pula_sem_email(tmp_path):
    from email_sender import enviar_todos

    pdf = tmp_path / "A000003_Emilce_Gomes_Holerite_202601.pdf"
    pdf.write_bytes(b"%PDF")

    envios = [
        {"nome": "Emilce Gomes", "matricula": "000003",
         "email": "emilce@e.com", "arquivo": str(pdf)},
        {"nome": "Sem Email", "matricula": "000099",
         "email": None, "arquivo": str(pdf)},
    ]
    config = {
        "assunto_email": "Holerite {mes}/{ano}",
        "corpo_email": "Ola {nome}, {mes}/{ano}",
    }
    outlook = MagicMock()
    outlook.CreateItem.return_value = MagicMock()

    enviados, falhas = enviar_todos(envios, config, outlook)

    assert enviados == 1
    assert falhas == []


def test_enviar_todos_registra_falha(tmp_path):
    from email_sender import enviar_todos

    pdf = tmp_path / "A000003_Emilce_Gomes_Holerite_202601.pdf"
    pdf.write_bytes(b"%PDF")

    envios = [{"nome": "Emilce Gomes", "matricula": "000003",
               "email": "emilce@e.com", "arquivo": str(pdf)}]
    config = {
        "assunto_email": "Holerite {mes}/{ano}",
        "corpo_email": "Ola {nome}",
    }

    outlook = MagicMock()
    outlook.CreateItem.side_effect = Exception("Outlook indisponivel")

    enviados, falhas = enviar_todos(envios, config, outlook)

    assert enviados == 0
    assert "Emilce Gomes" in falhas
```

- [ ] **Step 2: Rodar testes — confirmar que passam**

```
python -m pytest tests/test_outlook.py -v
```

Esperado: 12 testes passando. (As funções COM já existem em `email_reader.py` e `email_sender.py` criados na Task 1.)

- [ ] **Step 3: Commit**

```bash
git add tests/test_outlook.py
git commit -m "feat: adiciona testes das funcoes Outlook COM com mocks"
```

---

## Task 3: pipeline.py main() + .bat

Completa o `main()` de `pipeline.py` e cria o novo `.bat`.

**Files:**
- Modify: `pipeline.py` — adicionar `main()` ao final
- Create: `RENOMEAR E ENVIAR HOLERITES.bat`

**Interfaces consumidas (de Tasks 1 e 2):**
- `conectar_outlook() -> namespace`
- `buscar_emails_holerite(namespace) -> list`
- `escolher_email(emails) -> mail_item`
- `baixar_e_extrair_zip(mail_item, pasta) -> list[str]`
- `carregar_nomes(planilha) -> dict` — de `renomear_holerites.py`
- `_executar_renomeacao(pasta, lookup) -> None` — de `pipeline.py`
- `carregar_emails(planilha) -> dict`
- `montar_envios(pasta, lookup) -> list`
- `exibir_conferencia(envios) -> bool`
- `enviar_todos(envios, config, outlook_app) -> tuple`

---

- [ ] **Step 1: Adicionar `main()` ao final de `pipeline.py`**

Adicionar ao final do arquivo `pipeline.py` (após `_executar_renomeacao`):

```python
def main():
    # Validar arquivos necessários
    for caminho, label in [
        (PLANILHA_NOMES, "Holerites renomear.xlsx"),
        (PLANILHA_EMAILS, "Emails funcionarios.xlsx"),
        (CONFIG_PATH, "config.json"),
    ]:
        if not os.path.exists(caminho):
            print(f"\nERRO: {label} nao encontrado em:\n  {caminho}")
            sys.exit(1)

    config = carregar_config(CONFIG_PATH)

    # Passo 1: Baixar ZIP do Outlook
    print("\n[1/4] Conectando ao Outlook e buscando e-mail...")
    try:
        from email_reader import (
            conectar_outlook,
            buscar_emails_holerite,
            escolher_email,
            baixar_e_extrair_zip,
        )
    except ImportError:
        print("ERRO: pywin32 nao instalado. Execute: pip install pywin32")
        sys.exit(1)

    try:
        ns = conectar_outlook()
    except Exception:
        print("ERRO: Nao foi possivel conectar ao Outlook.")
        print("Verifique se o Outlook esta aberto e tente novamente.")
        sys.exit(1)

    emails = buscar_emails_holerite(ns)
    if not emails:
        print("Nenhum e-mail com holerite encontrado na caixa de entrada.")
        sys.exit(0)

    email_escolhido = escolher_email(emails)
    print(f"  E-mail selecionado: {email_escolhido.Subject}")

    print("\n[2/4] Baixando e extraindo ZIP...")
    try:
        baixar_e_extrair_zip(email_escolhido, PASTA_PDFS)
    except FileNotFoundError as e:
        print(f"ERRO: {e}")
        sys.exit(1)

    # Passo 2: Renomear PDFs
    print("\n[3/4] Renomeando PDFs...")
    from renomear_holerites import carregar_nomes
    lookup_nomes = carregar_nomes(PLANILHA_NOMES)
    _executar_renomeacao(PASTA_PDFS, lookup_nomes)
    print("  PDFs renomeados.")

    # Passo 3: Conferência
    print("\n[4/4] Preparando conferencia...")
    from email_sender import carregar_emails, montar_envios, enviar_todos
    lookup_emails = carregar_emails(PLANILHA_EMAILS)
    envios = montar_envios(PASTA_PDFS, lookup_emails)

    if not exibir_conferencia(envios):
        print("\nCancelado. Nenhum e-mail foi enviado.")
        sys.exit(0)

    # Passo 4: Enviar
    print("\nEnviando holerites...")
    import win32com.client
    outlook_app = win32com.client.Dispatch("Outlook.Application")
    enviados, falhas = enviar_todos(envios, config, outlook_app)

    print(f"\n{enviados} e-mail(s) enviado(s) com sucesso.")
    if falhas:
        nomes = ", ".join(falhas)
        print(f"Nao enviados ({len(falhas)}): {nomes}")


if __name__ == "__main__":
    main()
```

- [ ] **Step 2: Criar `RENOMEAR E ENVIAR HOLERITES.bat`**

```bat
@echo off
chcp 65001 > nul
title Envio de Holerites
color 0A

echo.
echo ============================================
echo      ENVIO AUTOMATICO DE HOLERITES
echo ============================================
echo.

python --version > nul 2>&1
if errorlevel 1 (
    color 0C
    echo ERRO: Python nao encontrado no computador.
    echo Instale em: https://www.python.org/downloads/
    echo Marque "Add Python to PATH" durante a instalacao.
    echo.
    pause
    exit /b 1
)

echo Verificando dependencias...
python -c "import win32com, pandas, openpyxl" > nul 2>&1
if errorlevel 1 (
    echo Instalando dependencias, aguarde...
    pip install pywin32 pandas openpyxl --quiet
    echo.
)

for %%F in ("%~dp0Emails funcionarios.xlsx") do (
    if not exist "%%F" (
        color 0C
        echo ERRO: "Emails funcionarios.xlsx" nao encontrado.
        echo Coloque o arquivo na mesma pasta deste programa.
        echo.
        pause
        exit /b 1
    )
)

for %%F in ("%~dp0Holerites renomear.xlsx") do (
    if not exist "%%F" (
        color 0C
        echo ERRO: "Holerites renomear.xlsx" nao encontrado.
        echo Coloque o arquivo na mesma pasta deste programa.
        echo.
        pause
        exit /b 1
    )
)

echo Iniciando pipeline...
echo.

python "%~dp0pipeline.py"

echo.
if errorlevel 1 (
    color 0C
    echo Ocorreu um erro. Verifique as mensagens acima.
) else (
    color 0A
    echo Processo concluido.
)

echo.
pause
```

- [ ] **Step 3: Rodar todos os testes — confirmar que continuam passando**

```
python -m pytest tests/ -v
```

Esperado: 25 testes passando (13 puros + 12 COM mocks).

- [ ] **Step 4: Teste manual de smoke test (documentado)**

Sem Outlook real disponível, verificar que o pipeline falha de forma amigável:

```
python pipeline.py
```

Esperado (se Outlook fechado ou não instalado):
```
[1/4] Conectando ao Outlook e buscando e-mail...
ERRO: Nao foi possivel conectar ao Outlook.
Verifique se o Outlook esta aberto e tente novamente.
```

- [ ] **Step 5: Commit**

```bash
git add pipeline.py "RENOMEAR E ENVIAR HOLERITES.bat"
git commit -m "feat: adiciona pipeline completo e bat de envio de holerites"
git push origin master
```

---

## Self-Review

**Spec coverage:**
- ✅ Passo 1 — email_reader.py: `conectar_outlook`, `buscar_emails_holerite`, `escolher_email`, `baixar_e_extrair_zip`
- ✅ Passo 2 — `renomear_holerites.py` reutilizado via `carregar_nomes` + `_executar_renomeacao` (sem input())
- ✅ Passo 3 — `exibir_conferencia` em pipeline.py: tabela + contagens + s/n
- ✅ Passo 4 — `email_sender.py`: `enviar_todos` com progresso e relatório
- ✅ Erro: Outlook fechado → mensagem clara, sys.exit(1)
- ✅ Erro: nenhum e-mail encontrado → mensagem, sys.exit(0)
- ✅ Erro: mais de um e-mail → `escolher_email` pede número
- ✅ Erro: funcionário sem e-mail → pulado, listado no final
- ✅ Erro: falha de envio individual → continua os outros, lista no final
- ✅ config.json com `assunto_email` e `corpo_email` com placeholders `{nome}`, `{mes}`, `{ano}`
- ✅ `renomear_holerites.py` — sem alteração
- ✅ `RENOMEAR HOLERITES.bat` — mantido

**Placeholder scan:** Nenhum encontrado.

**Type consistency:** `montar_envios` → `list[dict]` com chaves `{nome, matricula, email, arquivo}` usadas consistentemente em `exibir_conferencia`, `enviar_todos`.
