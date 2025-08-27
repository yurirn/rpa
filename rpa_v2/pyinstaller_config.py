#!/usr/bin/env python3
"""
Configuração avançada para PyInstaller
Use este arquivo para personalizar ainda mais o build
"""

import PyInstaller.__main__

# Configurações do executável
PyInstaller.__main__.run([
    'main.py',                           # Arquivo principal
    '--onefile',                         # Arquivo único
    '--windowed',                        # Sem console (GUI)
    '--name=Sistema_RPA',                # Nome do executável
    '--distpath=dist',                   # Pasta de saída
    '--workpath=build',                  # Pasta de trabalho
    '--specpath=.',                      # Pasta do arquivo .spec
    
    # Incluir arquivos de dados
    '--add-data=config.json;.',
    '--add-data=modules.json;.',
    '--add-data=src;src',
    '--add-data=xml;xml',
    
    # Incluir imports ocultos
    '--hidden-import=selenium',
    '--hidden-import=webdriver_manager',
    '--hidden-import=requests',
    '--hidden-import=pandas',
    '--hidden-import=openpyxl',
    '--hidden-import=tkinter',
    '--hidden-import=tkinter.ttk',
    '--hidden-import=tkinter.messagebox',
    '--hidden-import=tkinter.filedialog',
    
    # Otimizações
    '--optimize=2',                      # Otimização máxima
    '--strip',                           # Remover símbolos de debug
    
    # Configurações específicas do Windows
    '--win-private-assemblies',
    '--win-no-prefer-redirects',
    
    # Limpar cache
    '--clean',
    
    # Logs detalhados
    '--log-level=INFO',
]) 