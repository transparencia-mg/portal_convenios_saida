import subprocess
import sys

SCRIPTS = [
    "processar.py",
    "gerar_pagamentos_filtrados.py",
    "publicar.py",
]

for script in SCRIPTS:
    print(f"\n{'=' * 60}\nExecutando {script}\n{'=' * 60}")
    resultado = subprocess.run(["python", script])
    if resultado.returncode != 0:
        print(f"\nERRO: {script} falhou (código {resultado.returncode}). Pipeline interrompido.")
        sys.exit(resultado.returncode)

print("\nPipeline concluído com sucesso.")
