#!/usr/bin/env python3
"""
Script para gerar executável (.exe) do Sistema RPA
Execute: python build_exe.py
"""

import os
import sys
import subprocess
import shutil

def install_pyinstaller():
    """Instala o PyInstaller se não estiver disponível"""
    try:
        import PyInstaller
        print("✓ PyInstaller já está instalado")
    except ImportError:
        print("Instalando PyInstaller...")
        subprocess.check_call([sys.executable, "-m", "pip", "install", "pyinstaller"])
        print("✓ PyInstaller instalado com sucesso")

def create_spec_file():
    """Cria arquivo .spec personalizado para o PyInstaller"""
    spec_content = '''# -*- mode: python ; coding: utf-8 -*-

block_cipher = None

a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=[],
    datas=[
        ('config.json', '.'),
        ('modules.json', '.'),
        ('src/', 'src/'),
        ('xml/', 'xml/'),
    ],
    hiddenimports=[
        # Selenium completo
        'selenium',
        'selenium.webdriver',
        'selenium.webdriver.common',
        'selenium.webdriver.common.by',
        'selenium.webdriver.common.keys',
        'selenium.webdriver.common.action_chains',
        'selenium.webdriver.common.desired_capabilities',
        'selenium.webdriver.common.proxy',
        'selenium.webdriver.common.service',
        'selenium.webdriver.common.utils',
        'selenium.webdriver.chrome',
        'selenium.webdriver.chrome.service',
        'selenium.webdriver.chrome.options',
        'selenium.webdriver.firefox',
        'selenium.webdriver.firefox.service',
        'selenium.webdriver.firefox.options',
        'selenium.webdriver.edge',
        'selenium.webdriver.edge.service',
        'selenium.webdriver.edge.options',
        'selenium.webdriver.support',
        'selenium.webdriver.support.ui',
        'selenium.webdriver.support.expected_conditions',
        'selenium.webdriver.support.select',
        'selenium.webdriver.support.wait',
        
        # WebDriver Manager
        'webdriver_manager',
        'webdriver_manager.chrome',
        'webdriver_manager.firefox',
        'webdriver_manager.edge',
        'webdriver_manager.core',
        'webdriver_manager.core.driver',
        'webdriver_manager.core.manager',
        'webdriver_manager.core.os_manager',
        'webdriver_manager.core.utils',
        
        # Outras dependências
        'requests',
        'pandas',
        'openpyxl',
        'tkinter',
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'json',
        'threading',
        'datetime',
        'os',
        'sys',
        'time',
        're',
        'csv',
        'xml',
        'xml.etree',
        'xml.etree.ElementTree',
    ],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='Sistema_RPA',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # False para aplicação GUI sem console
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=None,  # Adicione um ícone aqui se desejar
)
'''
    
    with open('Sistema_RPA.spec', 'w', encoding='utf-8') as f:
        f.write(spec_content)
    
    print("✓ Arquivo .spec criado: Sistema_RPA.spec")

def copy_required_files():
    """Copia arquivos necessários para a pasta dist"""
    print("Copiando arquivos necessários...")
    
    dist_dir = 'dist'
    if not os.path.exists(dist_dir):
        print("❌ Pasta dist não encontrada")
        return False
    
    # Lista de arquivos e pastas necessários
    required_files = [
        'config.json',
        'modules.json'
    ]
    
    required_dirs = [
        'src',
        'xml'
    ]
    
    # Copiar arquivos
    for file in required_files:
        if os.path.exists(file):
            shutil.copy2(file, os.path.join(dist_dir, file))
            print(f"✓ Copiado: {file}")
        else:
            print(f"⚠️  Arquivo não encontrado: {file}")
    
    # Copiar pastas
    for dir_name in required_dirs:
        if os.path.exists(dir_name):
            dest_dir = os.path.join(dist_dir, dir_name)
            if os.path.exists(dest_dir):
                shutil.rmtree(dest_dir)
            shutil.copytree(dir_name, dest_dir)
            print(f"✓ Copiado: {dir_name}/")
        else:
            print(f"⚠️  Pasta não encontrada: {dir_name}")
    
    return True

def build_exe():
    """Gera o executável usando PyInstaller"""
    print("Gerando executável...")
    
    # Usar o arquivo .spec personalizado
    cmd = [
        sys.executable, "-m", "PyInstaller",
        "--clean",
        "Sistema_RPA.spec"
    ]
    
    try:
        subprocess.check_call(cmd)
        print("✓ Executável gerado com sucesso!")
        print("📁 Localização: dist/Sistema_RPA.exe")
        
        # Copiar arquivos necessários
        if copy_required_files():
            print("✓ Todos os arquivos necessários foram copiados")
        else:
            print("⚠️  Alguns arquivos não puderam ser copiados")
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Erro ao gerar executável: {e}")
        return False
    
    return True

def create_installer_script():
    """Cria um script de instalação simples"""
    installer_content = '''@echo off
echo ========================================
echo    Instalador do Sistema RPA
echo ========================================
echo.

REM Criar diretório de instalação
set INSTALL_DIR=%USERPROFILE%\\Desktop\\Sistema_RPA
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Instalando em: %INSTALL_DIR%
echo.

REM Copiar arquivos
copy "dist\\Sistema_RPA.exe" "%INSTALL_DIR%\\"
copy "dist\\config.json" "%INSTALL_DIR%\\"
copy "dist\\modules.json" "%INSTALL_DIR%\\"
xcopy "dist\\src" "%INSTALL_DIR%\\src\\" /E /I /Y
xcopy "dist\\xml" "%INSTALL_DIR%\\xml\\" /E /I /Y

REM Criar atalho na área de trabalho
echo @echo off > "%USERPROFILE%\\Desktop\\Sistema RPA.bat"
echo cd /d "%INSTALL_DIR%" >> "%USERPROFILE%\\Desktop\\Sistema RPA.bat"
echo start Sistema_RPA.exe >> "%USERPROFILE%\\Desktop\\Sistema RPA.bat"

echo.
echo ========================================
echo    Instalação concluída!
echo ========================================
echo.
echo O sistema foi instalado em: %INSTALL_DIR%
echo Um atalho foi criado na área de trabalho.
echo.
pause
'''
    
    with open('instalar_sistema.bat', 'w', encoding='utf-8') as f:
        f.write(installer_content)
    
    print("✓ Script de instalação criado: instalar_sistema.bat")

def main():
    """Função principal"""
    print("🔧 Gerador de Executável - Sistema RPA")
    print("=" * 50)
    
    # Verificar se estamos no diretório correto
    if not os.path.exists('main.py'):
        print("❌ Erro: Execute este script no diretório raiz do projeto")
        return
    
    # Instalar PyInstaller
    install_pyinstaller()
    
    # Criar arquivo .spec
    create_spec_file()
    
    # Gerar executável
    if build_exe():
        # Criar script de instalação
        create_installer_script()
        
        print("\n🎉 Processo concluído com sucesso!")
        print("\n📋 Próximos passos:")
        print("1. O executável está em: dist/Sistema_RPA.exe")
        print("2. Execute 'instalar_sistema.bat' para instalar o sistema")
        print("3. Ou copie manualmente a pasta 'dist' para o local desejado")
        
        # Abrir pasta dist
        if os.path.exists('dist'):
            try:
                os.startfile('dist')
            except:
                print("📁 Abra manualmente a pasta 'dist' para ver o executável")
    else:
        print("\n❌ Falha na geração do executável")

if __name__ == "__main__":
    main() 