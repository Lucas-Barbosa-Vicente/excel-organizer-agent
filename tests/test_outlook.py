import io
import os
import zipfile
from unittest.mock import MagicMock, patch
import email as email_lib
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
import pytest


# ── buscar_emails_holerite ────────────────────────────────────────────────────

def _make_imap(ids_bytes: bytes, msg: email_lib.message.Message):
    imap = MagicMock()
    imap.search.return_value = ("OK", [ids_bytes])
    imap.fetch.return_value = ("OK", [(b"data", msg.as_bytes())])
    return imap


def test_buscar_emails_holerite_retorna_emails():
    from email_reader import buscar_emails_holerite

    msg = email_lib.message_from_string("Subject: Holerite Janeiro 2026\n\n")
    imap = _make_imap(b"1", msg)

    resultado = buscar_emails_holerite(imap)

    imap.select.assert_called_once_with("INBOX")
    assert len(resultado) == 1
    assert resultado[0][1].get("Subject") == "Holerite Janeiro 2026"


def test_buscar_emails_holerite_retorna_vazio():
    from email_reader import buscar_emails_holerite

    imap = MagicMock()
    imap.search.return_value = ("OK", [b""])

    resultado = buscar_emails_holerite(imap)
    assert resultado == []


def test_buscar_emails_holerite_multiplos():
    from email_reader import buscar_emails_holerite

    msg1 = email_lib.message_from_string("Subject: Holerite Jan\n\n")
    msg2 = email_lib.message_from_string("Subject: Holerite Fev\n\n")

    imap = MagicMock()
    imap.search.return_value = ("OK", [b"1 2"])
    msgs = {b"2": msg1, b"1": msg2}  # reversed iteration: 2 first, then 1
    imap.fetch.side_effect = lambda eid, fmt: ("OK", [(b"data", msgs[eid].as_bytes())])

    resultado = buscar_emails_holerite(imap)
    assert len(resultado) == 2


# ── escolher_email ────────────────────────────────────────────────────────────

def test_escolher_email_retorna_direto_quando_unico():
    from email_reader import escolher_email

    msg = email_lib.message_from_string("Subject: Holerite\n\n")
    item = (b"1", msg)
    assert escolher_email([item]) is item


def test_escolher_email_pede_escolha_com_multiplos():
    from email_reader import escolher_email

    msg1 = email_lib.message_from_string(
        "Subject: Holerite Jan\nDate: Thu, 01 Jan 2026 08:00:00 +0000\n\n"
    )
    msg2 = email_lib.message_from_string(
        "Subject: Holerite Fev\nDate: Sun, 01 Feb 2026 08:00:00 +0000\n\n"
    )
    emails = [(b"1", msg1), (b"2", msg2)]

    with patch("builtins.input", return_value="2"):
        resultado = escolher_email(emails)

    assert resultado is emails[1]


# ── baixar_e_extrair_zip ──────────────────────────────────────────────────────

def test_baixar_e_extrair_zip_salva_e_extrai(tmp_path):
    from email_reader import baixar_e_extrair_zip

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w") as zf:
        zf.writestr("000003 EMILCE.pdf", b"%PDF")
    zip_bytes = buf.getvalue()

    msg_mime = MIMEMultipart()
    part = MIMEApplication(zip_bytes, Name="holerites.zip")
    part["Content-Disposition"] = 'attachment; filename="holerites.zip"'
    msg_mime.attach(part)

    msg = email_lib.message_from_bytes(msg_mime.as_bytes())
    mail_item = (b"1", msg)

    pasta = str(tmp_path / "HOLERITES")
    nomes = baixar_e_extrair_zip(mail_item, pasta)

    assert "000003 EMILCE.pdf" in nomes
    assert os.path.exists(os.path.join(pasta, "000003 EMILCE.pdf"))
    assert not os.path.exists(os.path.join(pasta, "_temp_holerites.zip"))


def test_baixar_e_extrair_zip_erro_sem_zip(tmp_path):
    from email_reader import baixar_e_extrair_zip

    msg = email_lib.message_from_string("Subject: Holerite\n\nSem anexo.")
    mail_item = (b"1", msg)

    with pytest.raises(FileNotFoundError):
        baixar_e_extrair_zip(mail_item, str(tmp_path))


# ── enviar_holerite ───────────────────────────────────────────────────────────

def test_enviar_holerite_chama_send(tmp_path):
    from email_sender import enviar_holerite

    pdf = tmp_path / "A000003_Emilce_Gomes_Holerite_202601.pdf"
    pdf.write_bytes(b"%PDF")

    smtp = MagicMock()

    enviar_holerite(smtp, "rh@empresa.com", "Emilce Gomes", "emilce@e.com",
                    str(pdf), "Holerite 01/2026", "Corpo do e-mail")

    smtp.sendmail.assert_called_once()
    remetente, destinatario, _ = smtp.sendmail.call_args[0]
    assert remetente == "rh@empresa.com"
    assert destinatario == "emilce@e.com"


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
    smtp = MagicMock()

    enviados, falhas = enviar_todos(envios, config, smtp, "rh@empresa.com")

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
    smtp = MagicMock()
    smtp.sendmail.side_effect = Exception("SMTP error")

    enviados, falhas = enviar_todos(envios, config, smtp, "rh@empresa.com")

    assert enviados == 0
    assert "Emilce Gomes" in falhas
