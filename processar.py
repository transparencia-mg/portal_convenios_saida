import pandas as pd
from pathlib import Path
import subprocess
import os
import re

REPO = Path(__file__).parent
UPLOAD = REPO / "upload"
DATA = REPO / "upload"

DATA.mkdir(exist_ok=True)

ARQUIVO_DEPARA = REPO / "de_para.xlsx"

if not ARQUIVO_DEPARA.exists():
    raise FileNotFoundError(
        f"Arquivo de_para não encontrado: {ARQUIVO_DEPARA}"
    )

print("Carregando de_para...")

depara = pd.read_excel(
    ARQUIVO_DEPARA,
    dtype=str
)

depara.columns = depara.columns.str.strip()

depara = depara[
    [
        "CODIGO_ORGAO",
        "ORGAO_SIGLA",
        "DESCRICAO_ORGAO"
    ]
].copy()

depara["CODIGO_ORGAO"] = (
    depara["CODIGO_ORGAO"]
    .astype(str)
    .str.strip()
)

depara = depara.drop_duplicates(
    subset=["CODIGO_ORGAO"]
)

arquivos_gerados = 0

# =====================================================
# PROCESSA VW_V2_CONVENIO.xlsx
# =====================================================

arquivo_convenios = UPLOAD / "VW_V2_CONVENIO.xlsx"

if arquivo_convenios.exists():

    print("\nProcessando VW_V2_CONVENIO.xlsx...")

    df = pd.read_excel(
        arquivo_convenios,
        dtype=str
    )

    df.columns = df.columns.str.strip()

    # Remove colunas desnecessárias
    colunas_remover = [
        "tempo_analise_tecnica_continuo",
        "tempo_analisejuridica_continuo",
        "tempo_a_encaminhador_continuo"
    ]

    for coluna in colunas_remover:
        if coluna in df.columns:
            df = df.drop(columns=[coluna])

    # Filtra status
    if "STATUS_CONVENIO" in df.columns:

        df["STATUS_CONVENIO"] = (
            df["STATUS_CONVENIO"]
            .astype(str)
            .str.strip()
            .str.upper()
        )

        df = df[
            df["STATUS_CONVENIO"].isin(
                [
                    "VIGENTE",
                    "ENCERRADO"
                ]
            )
        ]

    # Atualiza descrição órgão
    if "DESCRICAO_ORGAO" in df.columns:
        df = df.drop(columns=["DESCRICAO_ORGAO"])

    if "CODIGO_ORGAO" in df.columns:

        df["CODIGO_ORGAO"] = (
            df["CODIGO_ORGAO"]
            .astype(str)
            .str.strip()
        )

        df = df.merge(
            depara[
                [
                    "CODIGO_ORGAO",
                    "DESCRICAO_ORGAO"
                ]
            ],
            on="CODIGO_ORGAO",
            how="left"
        )

    if (
        "ORGAO_SIGLA" in df.columns
        and "DESCRICAO_ORGAO" in df.columns
    ):

        cols = list(df.columns)

        cols.remove("DESCRICAO_ORGAO")

        pos = cols.index("ORGAO_SIGLA") + 1

        cols.insert(
            pos,
            "DESCRICAO_ORGAO"
        )

        df = df[cols]

    caminho_saida = UPLOAD / "convenios_saida.xlsx"

    df.to_excel(
        caminho_saida,
        index=False
    )

    print(f"Gerado: {caminho_saida}")

    try:
        os.remove(arquivo_convenios)
    except:
        pass

    arquivos_gerados += 1

# =====================================================
# PROCESSA CONVENIOSAIDAAAA.xlsx
# =====================================================

arquivos_saida = list(
    UPLOAD.glob("CONVENIOSAIDA*.xlsx")
)

for arquivo in arquivos_saida:

    nome = arquivo.stem

    match = re.search(r"(\d{4})", nome)

    if not match:
        print(
            f"Ano não encontrado em {arquivo.name}"
        )
        continue

    ano = match.group(1)

    print(f"\nProcessando {arquivo.name}")

    abas = pd.read_excel(
        arquivo,
        sheet_name=None,
        header=1
    )

    # ===============================
    # ABA PAGAMENTO
    # ===============================

    aba_pagamento = None

    for nome_aba in abas.keys():

        if nome_aba.lower().strip() == "pagamento":
            aba_pagamento = nome_aba
            break

    if aba_pagamento:

        df_pagamento = abas[aba_pagamento]

        df_pagamento = df_pagamento.iloc[:, 1:]

        df_pagamento = df_pagamento.dropna(
            how="all"
        )

        df_pagamento = df_pagamento.dropna(
            axis=1,
            how="all"
        )

        saida = UPLOAD / f"pagamento{ano}.xlsx"

        df_pagamento.to_excel(
            saida,
            index=False
        )

        print(f"Gerado: {saida}")

        arquivos_gerados += 1

    # ===============================
    # ABA PAGAMENTORP
    # ===============================

    aba_pagamentorp = None

    for nome_aba in abas.keys():

        if nome_aba.lower().strip() == "pagamentorp":
            aba_pagamentorp = nome_aba
            break

    if aba_pagamentorp:

        df_rp = abas[aba_pagamentorp]

        df_rp = df_rp.iloc[:, 1:]

        df_rp = df_rp.dropna(
            how="all"
        )

        df_rp = df_rp.dropna(
            axis=1,
            how="all"
        )

        saida = UPLOAD / f"pagamentorp{ano}.xlsx"

        df_rp.to_excel(
            saida,
            index=False
        )

        print(f"Gerado: {saida}")

        arquivos_gerados += 1

    try:
        os.remove(arquivo)
    except:
        pass

# =====================================================
# GIT
# =====================================================

if arquivos_gerados > 0:

    resultado = subprocess.run(
        ["git", "status", "--porcelain"],
        cwd=REPO,
        capture_output=True,
        text=True
    )

    if resultado.stdout.strip():

        subprocess.run(
            ["git", "add", "."],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            [
                "git",
                "commit",
                "-m",
                "Atualização automática portal convênios"
            ],
            cwd=REPO,
            check=True
        )

        subprocess.run(
            ["git", "push"],
            cwd=REPO,
            check=True
        )

        print(
            "\nGitHub atualizado com sucesso."
        )

    else:

        print(
            "\nNenhuma alteração encontrada."
        )

else:

    print(
        "\nNenhum arquivo foi processado."
    )