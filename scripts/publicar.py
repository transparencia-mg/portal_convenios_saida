import subprocess

comandos = [
    ["git", "add", "."],
    ["git", "commit", "-m", "Atualização automática dos dados"],
    ["git", "push"]
]

for cmd in comandos:
    subprocess.run(cmd, check=True)

print("GitHub atualizado com sucesso.")