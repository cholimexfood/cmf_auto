import subprocess

with open('requirements.txt', 'r') as f:
    for line in f:
        package = line.split('==')[0].strip()
        subprocess.run(['pip', 'install', '--upgrade', package], check=True)
