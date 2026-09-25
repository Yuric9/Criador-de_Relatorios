@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo   Criador de Relatorios - Build Windows
echo ============================================
echo.

where py >nul 2>&1
if errorlevel 1 (
  echo ERRO: Python nao foi encontrado.
  echo Instale Python 3.12+ e tente novamente.
  pause
  exit /b 1
)

py -m pip install --upgrade pip
if errorlevel 1 goto :error

py -m pip install -r requirements.txt
if errorlevel 1 goto :error

py -m pip install pyinstaller
if errorlevel 1 goto :error

if exist build rmdir /s /q build
if exist dist rmdir /s /q dist
if exist RelatorioFotografico.spec del /q RelatorioFotografico.spec

py -m PyInstaller --onefile --noconsole --clean --name RelatorioFotografico relatorio_fotografico.py
if errorlevel 1 goto :error

echo.
echo ============================================
echo   BUILD CONCLUIDO
 echo   dist\RelatorioFotografico.exe
 echo ============================================
echo.
pause
exit /b 0

:error
echo.
echo ============================================
echo   ERRO DURANTE O BUILD
 echo ============================================
echo.
pause
exit /b 1
