import subprocess
import sys
from pathlib import Path

# Pasta onde este atualizar.py está salvo (raiz do projeto, ex.:
# portal_convenios_saida/). Os demais scripts ficam na subpasta "scripts/".
REPO = Path(__file__).resolve().parent
SCRIPTS_DIR = REPO / "scripts"

SCRIPTS = [
    "processar.py",
    "gerar_pagamentos_filtrados.py",
    "publicar.py",
]

for script in SCRIPTS:
    caminho_script = SCRIPTS_DIR / script

    if not caminho_script.exists():
        print(f"\nERRO: não encontrei {caminho_script}. Confira se ele está dentro de 'scripts/'.")
        sys.exit(1)

    print(f"\n{'=' * 60}\nExecutando {script}\n{'=' * 60}")
    # sys.executable garante que usa o mesmo Python (e mesmo ambiente/venv)
    # que está rodando o atualizar.py, em vez de depender do "python" do PATH.
    resultado = subprocess.run([sys.executable, str(caminho_script)])
    if resultado.returncode != 0:
        print(f"\nERRO: {script} falhou (código {resultado.returncode}). Pipeline interrompido.")
        sys.exit(resultado.returncode)

print("\nPipeline concluído com sucesso.")
