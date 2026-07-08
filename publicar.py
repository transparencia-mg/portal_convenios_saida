import subprocess
from pathlib import Path

REPO = Path(__file__).parent

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
