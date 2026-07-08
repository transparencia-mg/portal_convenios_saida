#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
gerar_pagamentos_filtrados.py
==============================

O QUE ESTE SCRIPT FAZ
----------------------
Le o arquivo "convenios_saida.xlsx" (a lista de convenios de saida) e usa a
coluna NUMERO_SIAFI dos convenios com STATUS_CONVENIO "Vigente" ou
"Encerrado" para filtrar os arquivos brutos de pagamento:

    pagamentoAAAA.csv       (ex.: pagamento2022.csv ... pagamento2026.csv)
    pagamentorpAAAA.csv     (ex.: pagamentorp2022.csv ... pagamentorp2026.csv)

Esses arquivos brutos trazem pagamentos de TODOS os contratos/convenios do
estado (centenas de milhares de linhas), mas o dashboard só usa as linhas
cujo SIAFI pertence a um convenio de saida valido (normalmente 1-3% do
total). Este script gera copias enxutas desses arquivos, mantendo apenas as
linhas necessarias -- o mesmo layout de colunas, só que bem menores
(tipicamente de ~250 MB para poucos MB no total).

COMO FUNCIONA EM QUALQUER COMPUTADOR (não só no seu)
-----------------------------------------------------
O script NÃO tem nenhum caminho fixo tipo "G:\\Meu Drive\\...". Por padrão,
ele espera esta estrutura de pastas, relativa a onde o script está salvo:

    portal_convenios_saida/            <- coloque o script aqui
        upload/
            convenios_saida.xlsx
            pagamento2022.csv ... pagamento2026.csv
            pagamentorp2022.csv ... pagamentorp2026.csv

Ou seja:

    1. Copie este arquivo .py para a pasta "portal_convenios_saida"
       (na raiz, ao lado da pasta "upload").
    2. Deixe o convenios_saida.xlsx e todos os pagamentoAAAA.csv /
       pagamentorpAAAA.csv dentro da subpasta "upload".
    3. Rode o script (veja "COMO RODAR" abaixo).
    4. Ele SUBSTITUI cada arquivo pela sua versão filtrada, mesmo nome,
       na própria pasta "upload/" (o arquivo bruto original é perdido --
       só sobra a versão já filtrada).

Como isso é sempre relativo à pasta do script (não a um caminho fixo tipo
"G:\\"), funciona igual em qualquer computador, mesmo que a letra do drive
ou o caminho até "portal_convenios_saida" seja diferente.

Se a sua estrutura de pastas for outra, dá para apontar os caminhos
manualmente por parâmetro -- veja os exemplos abaixo.

REQUISITOS (instalar uma vez por computador)
---------------------------------------------
    pip install pandas openpyxl

COMO RODAR
----------
Uso mais simples (estrutura padrão "portal_convenios_saida/upload/..."):

    python gerar_pagamentos_filtrados.py

Apontando pastas manualmente (caso sua estrutura seja diferente):

    python gerar_pagamentos_filtrados.py --convenios "caminho\\convenios_saida.xlsx" --entrada "caminho\\da\\pasta" --saida "caminho\\de\\saida"
"""

import argparse
import glob
import os
import re
import sys
import unicodedata

import pandas as pd

STATUS_EXIBIDOS = ("vigente", "encerrado")


def normalizar_coluna(nome):
    """Normaliza um nome de coluna para comparação, tirando acentos, espaços
    e maiúsculas/minúsculas. Assim "ContratoConvênio Saída", "Contrato
    Convenio Saida" e "contratoconvenio_saida" são todos reconhecidos como a
    mesma coluna, sem depender de o arquivo seguir um padrão exato."""
    nome = str(nome).strip()
    nome = unicodedata.normalize("NFKD", nome).encode("ascii", "ignore").decode("ascii")
    nome = nome.lower()
    nome = re.sub(r"\s+", "_", nome)
    nome = re.sub(r"[^a-z0-9_]", "", nome)
    return nome


def achar_coluna(colunas, alvo_normalizado):
    """Procura, entre os nomes de coluna originais, aquele cuja versão
    normalizada bate com alvo_normalizado. Retorna o nome ORIGINAL da
    coluna (preservando acentuação/maiúsculas do arquivo), ou None."""
    for c in colunas:
        if normalizar_coluna(c) == alvo_normalizado:
            return c
    return None


def caminho_padrao(nome_arquivo):
    """Resolve um caminho relativo à pasta onde este script está salvo,
    não à pasta de onde o comando foi executado -- assim funciona igual em
    qualquer computador, independentemente de onde o .py foi colocado."""
    pasta_do_script = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(pasta_do_script, nome_arquivo)


def carregar_siafis_validos(caminho_xlsx):
    if not os.path.exists(caminho_xlsx):
        raise SystemExit(
            f'\nERRO: não encontrei o arquivo de convênios em:\n  {caminho_xlsx}\n'
            f'Confira se o nome/local está certo ou use --convenios "caminho".'
        )

    df = pd.read_excel(caminho_xlsx, dtype=str)

    col_status = achar_coluna(df.columns, "status_convenio")
    col_siafi = achar_coluna(df.columns, "numero_siafi")
    if not col_status or not col_siafi:
        raise SystemExit(
            "\nERRO: a planilha de convênios precisa ter as colunas "
            "STATUS_CONVENIO e NUMERO_SIAFI. Colunas encontradas:\n  "
            + ", ".join(df.columns)
        )

    status = df[col_status].fillna("").str.strip().str.lower()
    mask = status.apply(lambda s: any(alvo in s for alvo in STATUS_EXIBIDOS))

    siafis = (
        df.loc[mask, col_siafi]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.replace(r"[,.]0$", "", regex=True)  # remove ",0" / ".0" no final
    )
    validos = set(s for s in siafis if s)
    return validos


def limpar_siafi(valor):
    v = "" if valor is None else str(valor).strip()
    if not v:
        return v
    # remove sufixo decimal tipo "9503579,0" ou "9503579.0"
    for sep in (",", "."):
        if sep in v:
            v = v.split(sep)[0]
    return v


def _ler_bruto_em_pedacos(caminho_entrada):
    """Le o arquivo bruto (pagamentoAAAA / pagamentorpAAAA) em pedaços,
    como texto puro (sem conversão numérica/data), funcionando tanto para
    .csv quanto para .xlsx. Para .xlsx não há leitura em chunks nativa no
    pandas, então lemos inteiro de uma vez (arquivos já são só algumas
    dezenas de milhares de linhas, tranquilo para o openpyxl)."""
    ext = os.path.splitext(caminho_entrada)[1].lower()
    if ext == ".csv":
        for pedaco in pd.read_csv(
            caminho_entrada,
            sep=";",
            dtype=str,
            encoding="utf-8-sig",
            chunksize=50_000,
            keep_default_na=False,
        ):
            yield pedaco
    elif ext in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(caminho_entrada, dtype=str, keep_default_na=False)
        yield df
    else:
        raise SystemExit(f"\nERRO: extensão não suportada: {caminho_entrada}")


def filtrar_arquivo_csv(caminho_entrada, caminho_saida_xlsx, siafis_validos):
    # dtype=str + keep_default_na garantem que NENHUM valor é reinterpretado
    # como número: "2521,12" continua exatamente "2521,12" (nunca vira
    # "251212" ou "2521.12"). Os cabeçalhos originais das colunas (com
    # acento/espaço/maiúscula, ex.: "ContratoConvênio Saída") são mantidos
    # como estão -- só usamos a versão normalizada para IDENTIFICAR qual
    # coluna é a do SIAFI, sem renomear nada no arquivo de saída.
    total = 0
    partes = []
    colunas = None
    col_siafi_original = None
    for pedaco in _ler_bruto_em_pedacos(caminho_entrada):
        if col_siafi_original is None:
            col_siafi_original = achar_coluna(pedaco.columns, "contratoconvenio_saida")
            if not col_siafi_original:
                raise SystemExit(
                    f'\nERRO: {caminho_entrada} não tem uma coluna equivalente a '
                    f'"ContratoConvênio Saída". Colunas encontradas:\n  '
                    + ", ".join(str(c) for c in pedaco.columns)
                )
        colunas = list(pedaco.columns)
        total += len(pedaco)
        siafi_limpo = pedaco[col_siafi_original].map(limpar_siafi)
        filtrado = pedaco[siafi_limpo.isin(siafis_validos)]
        if len(filtrado):
            partes.append(filtrado)

    resultado = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame(columns=colunas or [])
    mantidas = len(resultado)

    # Grava em .xlsx com todas as colunas como TEXTO (não como número/data),
    # para o Excel nunca reinterpretar valores tipo "2521,12" -- evita
    # qualquer perda/alteração de casas decimais, zeros à esquerda em
    # CNPJ/código, etc.
    with pd.ExcelWriter(caminho_saida_xlsx, engine="openpyxl") as writer:
        resultado.to_excel(writer, index=False, sheet_name="dados")
        ws = writer.sheets["dados"]
        for col_idx in range(1, len(resultado.columns) + 1):
            for row in ws.iter_rows(min_col=col_idx, max_col=col_idx, min_row=2):
                for cell in row:
                    cell.number_format = "@"  # formato "Texto" no Excel

    return total, mantidas


def main():
    parser = argparse.ArgumentParser(
        description="Gera versões filtradas (bem menores) dos CSVs de pagamento, "
        "usando convenios_saida.xlsx como referência de SIAFIs válidos."
    )
    parser.add_argument(
        "--convenios",
        default=caminho_padrao(os.path.join("upload", "convenios_saida.xlsx")),
        help="Caminho do arquivo convenios_saida.xlsx (padrão: pasta 'upload' ao lado do script)",
    )
    parser.add_argument(
        "--entrada",
        default=caminho_padrao("upload"),
        help="Pasta onde estão os pagamentoAAAA.csv / pagamentorpAAAA.csv (padrão: pasta 'upload' ao lado do script)",
    )
    parser.add_argument(
        "--saida",
        default=None,
        help='Pasta onde salvar os arquivos filtrados (padrão: a MESMA pasta de --entrada, '
             "sobrescrevendo os arquivos brutos originais com a versão filtrada, mesmo nome)",
    )
    args = parser.parse_args()

    pasta_entrada = os.path.abspath(args.entrada)
    pasta_saida = os.path.abspath(args.saida) if args.saida else pasta_entrada
    os.makedirs(pasta_saida, exist_ok=True)
    sobrescrevendo = (pasta_saida == pasta_entrada)

    print(f"Convênios:      {args.convenios}")
    print(f"Pasta entrada:  {pasta_entrada}")
    print(f"Pasta saída:    {pasta_saida}")
    if sobrescrevendo:
        print("ATENÇÃO: pasta de saída = pasta de entrada -> os arquivos brutos originais")
        print("         serão SUBSTITUÍDOS pela versão filtrada (mesmo nome).\n")
    else:
        print()

    siafis_validos = carregar_siafis_validos(args.convenios)
    print(f"SIAFIs válidos (Vigente/Encerrado): {len(siafis_validos)}\n")

    # Descobre automaticamente todos os arquivos pagamentoAAAA e
    # pagamentorpAAAA (em .xlsx ou .csv) presentes na pasta de entrada --
    # não depende de uma lista fixa de anos nem de uma extensão fixa.
    caminho_convenios_abs = os.path.abspath(args.convenios)
    todos = set()
    for padrao in ("pagamento*.xlsx", "pagamento*.xls", "pagamento*.csv"):
        for caminho in glob.glob(os.path.join(pasta_entrada, padrao)):
            if os.path.abspath(caminho) != caminho_convenios_abs:
                todos.add(caminho)
    todos_arquivos = sorted(todos)
    if not todos_arquivos:
        # Diagnóstico: mostra o que EXISTE na pasta, pra facilitar achar o
        # motivo (nome diferente, extensão diferente, pasta errada, etc.)
        try:
            conteudo = os.listdir(pasta_entrada)
        except FileNotFoundError:
            conteudo = None
        msg = f"\nERRO: nenhum arquivo pagamento* (xlsx/csv) encontrado em:\n  {pasta_entrada}\n"
        if conteudo is None:
            msg += "(essa pasta nem existe -- confira o caminho)\n"
        elif not conteudo:
            msg += "(a pasta existe mas está vazia)\n"
        else:
            msg += "Conteúdo encontrado nessa pasta:\n  " + "\n  ".join(sorted(conteudo)) + "\n"
        msg += 'Use --entrada "caminho da pasta" para apontar onde estão os arquivos.'
        raise SystemExit(msg)

    total_antes = total_depois = 0
    for caminho in todos_arquivos:
        nome_base = os.path.splitext(os.path.basename(caminho))[0]
        saida = os.path.join(pasta_saida, nome_base + ".xlsx")
        tam_antes = os.path.getsize(caminho)
        total, mantidas = filtrar_arquivo_csv(caminho, saida, siafis_validos)
        tam_depois = os.path.getsize(saida)
        total_antes += tam_antes
        total_depois += tam_depois
        print(
            f"{os.path.basename(caminho)} -> {os.path.basename(saida)}: {total} -> {mantidas} linhas | "
            f"{tam_antes/1024/1024:.1f}MB -> {tam_depois/1024:.0f}KB"
        )

    if total_depois:
        print(
            f"\nTOTAL: {total_antes/1024/1024:.1f}MB -> {total_depois/1024/1024:.2f}MB "
            f"({total_antes/total_depois:.0f}x menor)"
        )
    print(f"\nPronto! Arquivos filtrados salvos em:\n  {pasta_saida}")
    print("Suba o conteúdo dessa pasta para o repositório (mesmo local de sempre).")


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        sys.exit(1)
