import re
import os
import sys
import pandas as pd

# Caminhos relativos à pasta onde este script está
PASTA_BASE = os.path.dirname(os.path.abspath(__file__))
PASTA_PDFS = os.path.join(PASTA_BASE, "HOLERITES")
PLANILHA   = os.path.join(PASTA_BASE, "Holerites renomear.xlsx")


def pedir_competencia() -> str:
    """Pede o ano e mês ao usuário e valida o formato."""
    while True:
        competencia = input("Digite o ano e mês de competência (ex: 202606): ").strip()
        if re.fullmatch(r"\d{6}", competencia):
            return competencia
        print("  Formato inválido. Digite 6 dígitos: ano + mês, ex: 202606\n")


def carregar_nomes(planilha: str, competencia: str) -> dict:
    df = pd.read_excel(planilha)
    nomes = []
    for col in [df.columns[0], df.columns[2]]:
        nomes += df[col].dropna().tolist()

    lookup = {}
    for nome in nomes:
        m = re.search(r"[A-Z](\d{6})_", str(nome))
        if m:
            # Substitui o ano/mês que vier na planilha pelo informado pelo usuário
            novo_nome = re.sub(r"_\d{6}$", f"_{competencia}", str(nome))
            lookup[m.group(1)] = novo_nome
    return lookup


def renomear(pasta: str, lookup: dict):
    todos = os.listdir(pasta)
    pdfs = [f for f in todos if f.lower().endswith(".pdf")]

    # ── DIAGNÓSTICO ──────────────────────────────────────────────────────────
    print(f"\n[DIAGNÓSTICO]")
    print(f"  Pasta HOLERITES : {pasta}")
    print(f"  Arquivos totais : {len(todos)}")
    print(f"  PDFs encontrados: {len(pdfs)}")

    if not pdfs:
        print("\n  ATENÇÃO: Nenhum PDF encontrado na pasta HOLERITES!")
        print("  Verifique se os arquivos PDF estão dentro da pasta HOLERITES.")
        if todos:
            print(f"\n  Arquivos que existem na pasta:")
            for f in todos[:10]:
                print(f"    - {f}")
        input("\nPressione ENTER para sair...")
        return

    print(f"\n  Primeiros PDFs encontrados:")
    for f in pdfs[:5]:
        print(f"    - {f}")
    print()
    # ─────────────────────────────────────────────────────────────────────────

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
        if sem_match:
            print(f"\nArquivos sem correspondência ({len(sem_match)}):")
            for f in sem_match[:10]:
                m = re.match(r"^(\d{6})", f)
                if m:
                    print(f"  - {f}  (chave '{m.group(1)}' não encontrada na planilha)")
                else:
                    print(f"  - {f}  (não começa com 6 dígitos)")
        input("\nPressione ENTER para sair...")
        return

    print(f"Prévia — {len(renomeados)} arquivo(s) serão renomeados:\n")
    for orig, novo in renomeados[:5]:
        print(f"  {orig}")
        print(f"  -> {novo}\n")
    if len(renomeados) > 5:
        print(f"  ... e mais {len(renomeados) - 5} arquivo(s)\n")

    if sem_match:
        print(f"Sem correspondência na planilha ({len(sem_match)} arquivo(s)):")
        for f in sem_match:
            print(f"  - {f}")
        print()

    confirma = input("Confirmar renomeação? (s/n): ").strip().lower()
    if confirma != "s":
        print("\nCancelado. Nenhum arquivo foi alterado.")
        input("\nPressione ENTER para sair...")
        return

    pasta_real = os.path.realpath(pasta)
    for orig, novo in renomeados:
        destino = os.path.realpath(os.path.join(pasta, novo))
        if not destino.startswith(pasta_real + os.sep):
            print(f"  IGNORADO (caminho inválido): {novo}")
            continue
        os.rename(
            os.path.join(pasta, orig),
            destino,
        )

    print(f"\n{len(renomeados)} arquivo(s) renomeados com sucesso!")
    input("\nPressione ENTER para sair...")


if __name__ == "__main__":
    erros = []

    if not os.path.exists(PLANILHA):
        erros.append(f"Planilha não encontrada: {PLANILHA}")
    if not os.path.exists(PASTA_PDFS):
        erros.append(f"Pasta de PDFs não encontrada: {PASTA_PDFS}")

    if erros:
        print("\nERRO — Verifique os itens abaixo:")
        for e in erros:
            print(f"  - {e}")
        input("\nPressione ENTER para sair...")
        sys.exit(1)

    try:
        print("\nQual a competência dos holerites?")
        competencia = pedir_competencia()
        print()
        lookup = carregar_nomes(PLANILHA, competencia)
        renomear(PASTA_PDFS, lookup)
    except Exception as e:
        print(f"\nErro inesperado: {e}")
        import traceback
        traceback.print_exc()
        input("\nPressione ENTER para sair...")
        sys.exit(1)
