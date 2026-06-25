import os
import zipfile


def extrair_zip(caminho_zip: str, pasta_destino: str) -> list:
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


def baixar_e_extrair_zip(mail_item, pasta_destino: str) -> list:
    os.makedirs(pasta_destino, exist_ok=True)
    for anexo in mail_item.Attachments:
        if anexo.FileName.lower().endswith(".zip"):
            caminho_zip = os.path.join(pasta_destino, "_temp_holerites.zip")
            anexo.SaveAsFile(caminho_zip)
            nomes = extrair_zip(caminho_zip, pasta_destino)
            os.remove(caminho_zip)
            return nomes
    raise FileNotFoundError("Nenhum arquivo ZIP encontrado nos anexos do e-mail.")
