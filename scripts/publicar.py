import subprocess
from pathlib import Path

# Este script agora fica em portal_convenios_saida/scripts/publicar.py.
# O repositório git (a pasta com o .git) é a RAIZ do projeto, pasta pai de
# "scripts/" -- por isso o git add/commit/push precisa rodar de lá, senão
# ele só enxergaria mudanças dentro de scripts/ e ignoraria o que muda em
# upload/.
REPO = Path(__file__).resolve().parent.parent

print("Verificando alterações...")

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

    print("\nGitHub atualizado com sucesso.")

else:

    print("\nNenhuma alteração encontrada.")
