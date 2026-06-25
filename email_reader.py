import imaplib
import email as _email
import os
import zipfile
from email.header import decode_header


def extrair_zip(caminho_zip: str, pasta_destino: str) -> list:
    os.makedirs(pasta_destino, exist_ok=True)
    with zipfile.ZipFile(caminho_zip, "r") as zf:
        zf.extractall(pasta_destino)
        return zf.namelist()


def _decodificar_assunto(msg) -> str:
    partes = decode_header(msg.get("Subject", ""))
    resultado = []
    for parte, charset in partes:
        if isinstance(parte, bytes):
            resultado.append(parte.decode(charset or "utf-8", errors="replace"))
        else:
            resultado.append(str(parte))
    return "".join(resultado)


def conectar_imap(usuario: str, senha: str,
                  servidor: str = "imap-mail.outlook.com") -> imaplib.IMAP4_SSL:
    conn = imaplib.IMAP4_SSL(servidor, 993)
    conn.login(usuario, senha)
    return conn


def buscar_emails_holerite(imap_conn) -> list:
    imap_conn.select("INBOX")
    _, data = imap_conn.search(None, 'SUBJECT "holerite"')
    ids = data[0].split() if data[0] else []
    encontrados = []
    for eid in reversed(ids):
        _, raw = imap_conn.fetch(eid, "(RFC822)")
        msg = _email.message_from_bytes(raw[0][1])
        encontrados.append((eid, msg))
    return encontrados


def escolher_email(emails: list):
    if len(emails) == 1:
        return emails[0]
    print(f"\n{len(emails)} e-mails com holerite encontrados:\n")
    for i, (eid, msg) in enumerate(emails, 1):
        assunto = _decodificar_assunto(msg)
        data = msg.get("Date", "")
        print(f"  {i}. [{data}] {assunto}")
    while True:
        escolha = input(f"\nQual deseja usar? (1-{len(emails)}): ").strip()
        if escolha.isdigit() and 1 <= int(escolha) <= len(emails):
            return emails[int(escolha) - 1]
        print("Opcao invalida, tente novamente.")


def baixar_e_extrair_zip(mail_item: tuple, pasta_destino: str) -> list:
    eid, msg = mail_item
    os.makedirs(pasta_destino, exist_ok=True)
    for part in msg.walk():
        filename = part.get_filename()
        if filename and filename.lower().endswith(".zip"):
            caminho_zip = os.path.join(pasta_destino, "_temp_holerites.zip")
            with open(caminho_zip, "wb") as f:
                f.write(part.get_payload(decode=True))
            nomes = extrair_zip(caminho_zip, pasta_destino)
            os.remove(caminho_zip)
            return nomes
    raise FileNotFoundError("Nenhum arquivo ZIP encontrado nos anexos do e-mail.")
