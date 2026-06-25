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


def _encontrar_zip(pasta: str):
    zips = [f for f in os.listdir(pasta) if f.lower().endswith(".zip")]
    if len(zips) == 1:
        return os.path.join(pasta, zips[0])
    if len(zips) > 1:
        print(f"\n{len(zips)} ZIPs encontrados na pasta:")
        for i, z in enumerate(zips, 1):
            print(f"  {i}. {z}")
        while True:
            escolha = input(f"Qual deseja usar? (1-{len(zips)}): ").strip()
            if escolha.isdigit() and 1 <= int(escolha) <= len(zips):
                return os.path.join(pasta, zips[int(escolha) - 1])
    return None


def carregar_credenciais() -> tuple:
    env_path = os.path.join(BASE, ".env")
    credenciais = {}
    if os.path.exists(env_path):
        with open(env_path, encoding="utf-8") as f:
            for linha in f:
                linha = linha.strip()
                if "=" in linha and not linha.startswith("#"):
                    chave, valor = linha.split("=", 1)
                    credenciais[chave.strip()] = valor.strip()
    return credenciais.get("EMAIL_USUARIO", ""), credenciais.get("EMAIL_SENHA", "")


def main():
    ENV_PATH = os.path.join(BASE, ".env")

    for caminho, label in [
        (PLANILHA_NOMES, "Holerites renomear.xlsx"),
        (PLANILHA_EMAILS, "Emails funcionarios.xlsx"),
        (CONFIG_PATH, "config.json"),
        (ENV_PATH, ".env"),
    ]:
        if not os.path.exists(caminho):
            print(f"\nERRO: {label} nao encontrado em:\n  {caminho}")
            sys.exit(1)

    config = carregar_config(CONFIG_PATH)
    usuario, senha = carregar_credenciais()

    if not usuario or not senha:
        print("\nERRO: EMAIL_USUARIO ou EMAIL_SENHA nao configurados no arquivo .env")
        sys.exit(1)

    print("\n[1/4] Localizando ZIP dos holerites...")
    caminho_zip = _encontrar_zip(BASE)
    if caminho_zip:
        print(f"  ZIP encontrado: {os.path.basename(caminho_zip)}")
    else:
        print("  Nenhum ZIP encontrado na pasta do projeto.")
        print("  Baixe o ZIP do e-mail e cole o caminho abaixo.")
        caminho_zip = input("  Caminho do arquivo ZIP: ").strip().strip('"')
        if not os.path.exists(caminho_zip):
            print(f"ERRO: Arquivo nao encontrado:\n  {caminho_zip}")
            sys.exit(1)

    print("\n[2/4] Extraindo PDFs do ZIP...")
    from email_reader import extrair_zip
    nomes = extrair_zip(caminho_zip, PASTA_PDFS)
    print(f"  {len(nomes)} arquivo(s) extraido(s).")

    print("\n[3/4] Renomeando PDFs...")
    from renomear_holerites import carregar_nomes, pedir_competencia
    print("\nQual a competencia dos holerites?")
    competencia = pedir_competencia()
    lookup_nomes = carregar_nomes(PLANILHA_NOMES, competencia)
    _executar_renomeacao(PASTA_PDFS, lookup_nomes)
    print("  PDFs renomeados.")

    print("\n[4/4] Preparando conferencia...")
    from email_sender import carregar_emails, montar_envios, enviar_todos, conectar_smtp
    lookup_emails = carregar_emails(PLANILHA_EMAILS)
    envios = montar_envios(PASTA_PDFS, lookup_emails)

    if not exibir_conferencia(envios):
        print("\nCancelado. Nenhum e-mail foi enviado.")
        sys.exit(0)

    print("\nConectando ao servidor SMTP...")
    try:
        smtp = conectar_smtp(usuario, senha)
    except Exception as e:
        print(f"ERRO: Nao foi possivel conectar ao servidor SMTP: {e}")
        sys.exit(1)

    print("Enviando holerites...")
    try:
        enviados, falhas = enviar_todos(envios, config, smtp, usuario)
    finally:
        smtp.quit()

    print(f"\n{enviados} e-mail(s) enviado(s) com sucesso.")
    if falhas:
        nomes = ", ".join(falhas)
        print(f"Nao enviados ({len(falhas)}): {nomes}")


if __name__ == "__main__":
    main()
