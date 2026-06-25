import html
import os
import re

import pandas as pd


_EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def _email_valido(endereco: str) -> bool:
    return bool(endereco and _EMAIL_RE.match(endereco))


def carregar_emails(planilha: str) -> dict:
    raw = pd.read_excel(planilha, header=None, dtype=str)
    header_row = 0
    for i, row in raw.iterrows():
        if any("matricula" in str(v).lower() for v in row.values):
            header_row = i
            break

    df = pd.read_excel(planilha, header=header_row, dtype=str)
    df.columns = [str(c).strip().lower() for c in df.columns]
    df = df.dropna(subset=["matricula"])
    df["matricula"] = df["matricula"].str.strip().str.lstrip("AaSs").str.zfill(6)
    df["email"] = df["email"].str.strip()

    resultado = {}
    for _, row in df.iterrows():
        email = row["email"]
        if _email_valido(email):
            resultado[row["matricula"]] = email
        else:
            resultado[row["matricula"]] = None
    return resultado


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


def conectar_sendgrid(api_key: str):
    from sendgrid import SendGridAPIClient
    return SendGridAPIClient(api_key)


def enviar_holerite_sendgrid(sg_client, remetente: str, nome: str, email_dest: str,
                              arquivo: str, assunto: str, corpo: str):
    import base64
    from sendgrid.helpers.mail import (
        Mail, Attachment, FileContent, FileName, FileType, Disposition,
    )

    message = Mail(
        from_email=remetente,
        to_emails=email_dest,
        subject=assunto,
        html_content=html.escape(corpo).replace("\n", "<br>"),
    )

    with open(arquivo, "rb") as f:
        encoded = base64.b64encode(f.read()).decode()

    attachment = Attachment(
        FileContent(encoded),
        FileName(os.path.basename(arquivo)),
        FileType("application/pdf"),
        Disposition("attachment"),
    )
    message.attachment = attachment
    sg_client.send(message)


def enviar_todos_sendgrid(envios: list, config: dict, sg_client, remetente: str) -> tuple:
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
            enviar_holerite_sendgrid(
                sg_client, remetente, envio["nome"], envio["email"],
                envio["arquivo"], assunto, corpo
            )
            enviados += 1
            print(f"  Enviando {enviados}/{total}... {envio['nome']}")
        except Exception as e:
            falhas.append(envio["nome"])
            print(f"  ERRO ao enviar para {envio['nome']}: {e}")

    return enviados, falhas


def conectar_smtp(usuario: str, senha: str,
                  servidor: str = "smtp-mail.outlook.com") -> object:
    import smtplib
    smtp = smtplib.SMTP(servidor, 587)
    smtp.ehlo()
    smtp.starttls()
    smtp.login(usuario, senha)
    return smtp


def enviar_holerite(smtp_conn, remetente: str, nome: str, email_dest: str,
                    arquivo: str, assunto: str, corpo: str):
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.application import MIMEApplication

    msg = MIMEMultipart()
    msg["From"] = remetente
    msg["To"] = email_dest
    msg["Subject"] = assunto
    msg.attach(MIMEText(corpo, "plain", "utf-8"))

    with open(arquivo, "rb") as f:
        part = MIMEApplication(f.read(), Name=os.path.basename(arquivo))
    part["Content-Disposition"] = f'attachment; filename="{os.path.basename(arquivo)}"'
    msg.attach(part)

    smtp_conn.sendmail(remetente, email_dest, msg.as_string())


def enviar_todos(envios: list, config: dict, smtp_conn, remetente: str) -> tuple:
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
                smtp_conn, remetente, envio["nome"], envio["email"],
                envio["arquivo"], assunto, corpo
            )
            enviados += 1
            print(f"  Enviando {enviados}/{total}... {envio['nome']}")
        except Exception as e:
            falhas.append(envio["nome"])
            print(f"  ERRO ao enviar para {envio['nome']}: {e}")

    return enviados, falhas
