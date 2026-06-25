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

    item_sem_subject = MagicMock(spec=[])
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

    import io as _io
    buf = _io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("000003 EMILCE.pdf", b"%PDF")
    zip_bytes = buf.getvalue()

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
