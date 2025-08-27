@echo off
echo ========================================
echo    Instalador do Sistema RPA
echo ========================================
echo.

REM Criar diretório de instalação
set INSTALL_DIR=%USERPROFILE%\Desktop\Sistema_RPA
if not exist "%INSTALL_DIR%" mkdir "%INSTALL_DIR%"

echo Instalando em: %INSTALL_DIR%
echo.

REM Copiar arquivos
copy "dist\Sistema_RPA.exe" "%INSTALL_DIR%\"
copy "dist\config.json" "%INSTALL_DIR%\"
copy "dist\modules.json" "%INSTALL_DIR%\"
xcopy "dist\src" "%INSTALL_DIR%\src\" /E /I /Y
xcopy "dist\xml" "%INSTALL_DIR%\xml\" /E /I /Y

REM Criar atalho na área de trabalho
echo @echo off > "%USERPROFILE%\Desktop\Sistema RPA.bat"
echo cd /d "%INSTALL_DIR%" >> "%USERPROFILE%\Desktop\Sistema RPA.bat"
echo start Sistema_RPA.exe >> "%USERPROFILE%\Desktop\Sistema RPA.bat"

echo.
echo ========================================
echo    Instalação concluída!
echo ========================================
echo.
echo O sistema foi instalado em: %INSTALL_DIR%
echo Um atalho foi criado na área de trabalho.
echo.
pause
