import pandas as pd
from pathlib import Path
import subprocess
import os

REPO = Path(__file__).parent
UPLOAD = REPO / "upload"
DATA = REPO / "data"

DATA.mkdir(exist_ok=True)

ARQUIVO = UPLOAD / "CONVENIOSAIDA.xlsx"

if not ARQUIVO.exists():
    raise FileNotFoundError(f"Arquivo não encontrado: {ARQUIVO}")

print("Lendo arquivo...")

abas = pd.read_excel(
    ARQUIVO,
    sheet_name=None,
    header=1
)

print("Abas encontradas:")
print(list(abas.keys()))

arquivos_gerados = 0

for nome_aba, df in abas.items():

    print(f"Processando aba: {nome_aba}")

    # remove primeira coluna
    df = df.iloc[:, 1:]

    # remove linhas vazias
    df = df.dropna(how="all")

    # remove colunas vazias
    df = df.dropna(axis=1, how="all")

    # tenta encontrar o ano no nome da aba
    ano = "".join(filter(str.isdigit, str(nome_aba)))

    if ano not in ["2022", "2023", "2024", "2025", "2026"]:
        print(f"Aba ignorada (ano não identificado): {nome_aba}")
        continue

    nome_saida = f"empenho{ano}.xlsx"

    caminho_saida = DATA / nome_saida

    df.to_excel(
        caminho_saida,
        index=False
    )

    print(f"Gerado: {caminho_saida}")

    arquivos_gerados += 1

if arquivos_gerados > 0:

    os.remove(ARQUIVO)

    print("Arquivo original removido.")

    resultado = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=REPO,
        capture_output=True,
        text=True
    )

    if resultado.stdout.strip():

        subprocess.run(
            ["git", "add", "data"],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            [
                "git",
                "commit",
                "-m",
                "Atualização automática portal empenho"
            ],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            ["git", "push"],
            cwd=REPO,
            check=True
        )

        print("GitHub atualizado com sucesso.")

    else:
        print("Nenhuma alteração encontrada.")

else:
    print("Nenhum arquivo foi gerado.")