import re
import os
import sys
import pandas as pd

# Caminhos relativos à pasta onde este script está
PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_PDFS = os.path.join(PASTA_BASE, "HOLERITES")
PLANILHA   = os.path.join(PASTA_BASE, "Holerites renomear.xlsx")


def carregar_nomes(planilha: str) -> dict:
    df = pd.read_excel(planilha)
    nomes = []
    for col in [df.columns[0], df.columns[2]]:
        nomes += df[col].dropna().tolist()

    lookup = {}
    for nome in nomes:
        m = re.search(r"[A-Z](\d{6})_", str(nome))
        if m:
            lookup[m.group(1)] = str(nome)
    return lookup


def renomear(pasta: str, lookup: dict):
    pdfs = [f for f in os.listdir(pasta) if f.lower().endswith(".pdf")]

    renomeados = []
    sem_match  = []

    for pdf in pdfs:
        m = re.match(r"^(\d{6})", pdf)
        if m and m.group(1) in lookup:
            renomeados.append((pdf, lookup[m.group(1)] + ".pdf"))
        else:
            sem_match.append(pdf)

    if not renomeados:
        print("\nNenhum PDF encontrado para renomear.")
        print("Verifique se os PDFs estão na pasta HOLERITES.")
        return

    print(f"\nPrevia — {len(renomeados)} arquivo(s) serao renomeados:\n")
    for orig, novo in renomeados[:5]:
        print(f"  {orig}")
        print(f"  -> {novo}\n")
    if len(renomeados) > 5:
        print(f"  ... e mais {len(renomeados) - 5} arquivo(s)\n")

    if sem_match:
        print(f"Sem correspondencia na planilha ({len(sem_match)} arquivo(s)):")
        for f in sem_match:
            print(f"  - {f}")
        print()

    confirma = input("Confirmar renomeacao? (s/n): ").strip().lower()
    if confirma != "s":
        print("\nCancelado. Nenhum arquivo foi alterado.")
        return

    for orig, novo in renomeados:
        os.rename(
            os.path.join(pasta, orig),
            os.path.join(pasta, novo),
        )

    print(f"\n{len(renomeados)} arquivo(s) renomeados com sucesso!")


if __name__ == "__main__":
    erros = []

    if not os.path.exists(PLANILHA):
        erros.append(f"Planilha nao encontrada: {PLANILHA}")
    if not os.path.exists(PASTA_PDFS):
        erros.append(f"Pasta de PDFs nao encontrada: {PASTA_PDFS}")

    if erros:
        print("\nERRO — Verifique os itens abaixo:")
        for e in erros:
            print(f"  - {e}")
        sys.exit(1)

    try:
        lookup = carregar_nomes(PLANILHA)
        renomear(PASTA_PDFS, lookup)
    except Exception as e:
        print(f"\nErro inesperado: {e}")
        sys.exit(1)
