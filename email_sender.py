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
