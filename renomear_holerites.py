import re
import os
import pandas as pd

# ─── CONFIGURAÇÃO ─────────────────────────────────────────────────────────────

PASTA_PDFS  = r"C:\Users\Lucas\Documents\HOLERITES"
PLANILHA    = r"C:\Users\Lucas\Documents\Holerites renomear.xlsx"

# ──────────────────────────────────────────────────────────────────────────────


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
        print("Nenhum PDF encontrado para renomear.")
        return

    print(f"\nPrévia — {len(renomeados)} arquivo(s) serão renomeados:\n")
    for orig, novo in renomeados[:5]:
        print(f"  {orig}")
        print(f"  → {novo}\n")
    if len(renomeados) > 5:
        print(f"  ... e mais {len(renomeados) - 5} arquivo(s)\n")

    if sem_match:
        print(f"Sem correspondência ({len(sem_match)}): {sem_match}\n")

    confirma = input("Confirmar renomeação? (s/n): ").strip().lower()
    if confirma != "s":
        print("Cancelado.")
        return

    for orig, novo in renomeados:
        os.rename(
            os.path.join(pasta, orig),
            os.path.join(pasta, novo),
        )

    print(f"\n✓ {len(renomeados)} arquivo(s) renomeados com sucesso.")


if __name__ == "__main__":
    lookup = carregar_nomes(PLANILHA)
    renomear(PASTA_PDFS, lookup)
