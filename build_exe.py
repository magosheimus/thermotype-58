"""
Script para criar executável do ThermoType 58
"""

import PyInstaller.__main__
import os

# Diretório atual
current_dir = os.path.dirname(os.path.abspath(__file__))

# Configuração do PyInstaller
PyInstaller.__main__.run([
    'main.py',
    '--onefile',
    '--windowed',
    '--name=ThermoType 58',
    '--clean',
    '--noconfirm',
    '--add-data=printer.ico;.',
    '--icon=printer.ico',
])

print("\n" + "="*50)
print("Executável criado com sucesso!")
print("Localização: dist/ThermoType 58.exe")
print("="*50)
