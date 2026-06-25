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
    com_email = [e for e in envios if e["email"]]
    sem_email = [e for e in envios if not e["email"]]

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
    for pdf in os.listdir(pasta_pdfs):
        if not pdf.lower().endswith(".pdf"):
            continue
        m = re.match(r"^(\d{6})", pdf)
        if m and m.group(1) in lookup_nomes:
            os.rename(
                os.path.join(pasta_pdfs, pdf),
                os.path.join(pasta_pdfs, lookup_nomes[m.group(1)] + ".pdf"),
            )


def main():
    for caminho, label in [
        (PLANILHA_NOMES, "Holerites renomear.xlsx"),
        (PLANILHA_EMAILS, "Emails funcionarios.xlsx"),
        (CONFIG_PATH, "config.json"),
    ]:
        if not os.path.exists(caminho):
            print(f"\nERRO: {label} nao encontrado em:\n  {caminho}")
            sys.exit(1)

    config = carregar_config(CONFIG_PATH)

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

    print("\n[3/4] Renomeando PDFs...")
    from renomear_holerites import carregar_nomes, pedir_competencia
    print("\nQual a competencia dos holerites?")
    competencia = pedir_competencia()
    lookup_nomes = carregar_nomes(PLANILHA_NOMES, competencia)
    _executar_renomeacao(PASTA_PDFS, lookup_nomes)
    print("  PDFs renomeados.")

    print("\n[4/4] Preparando conferencia...")
    from email_sender import carregar_emails, montar_envios, enviar_todos
    lookup_emails = carregar_emails(PLANILHA_EMAILS)
    envios = montar_envios(PASTA_PDFS, lookup_emails)

    if not exibir_conferencia(envios):
        print("\nCancelado. Nenhum e-mail foi enviado.")
        sys.exit(0)

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
