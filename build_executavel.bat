@echo off
setlocal

set "PYTHON=.venv\Scripts\python.exe"
set "APP=validador_glebas_app.py"
set "NOME_EXE=ValidadorGlebas"
set "ICONE=assets\validador_glebas_icon.ico"
set "VERSAO=version_info.txt"

echo.
echo  ============================================================
echo   Validador de Glebas - SICOR ^| Gerando executavel...
echo  ============================================================
echo.

if not exist "%PYTHON%" (
    echo  ERRO: ambiente virtual nao encontrado em "%PYTHON%".
    echo  Crie a .venv antes de gerar o executavel.
    echo.
    pause
    exit /b 1
)

if not exist "%APP%" (
    echo  ERRO: arquivo principal "%APP%" nao encontrado.
    echo.
    pause
    exit /b 1
)

if not exist "%ICONE%" (
    echo  ERRO: icone "%ICONE%" nao encontrado.
    echo.
    pause
    exit /b 1
)

if not exist "%VERSAO%" (
    echo  ERRO: arquivo de versao "%VERSAO%" nao encontrado.
    echo.
    pause
    exit /b 1
)

echo [1/3] Instalando dependencias de build na .venv...
"%PYTHON%" -m pip install pyinstaller
if errorlevel 1 goto :erro
echo.

echo [2/3] Compilando com PyInstaller...
"%PYTHON%" -m PyInstaller ^
  --noconfirm ^
  --clean ^
  --onefile ^
  --windowed ^
  --name "%NOME_EXE%" ^
  --icon "%ICONE%" ^
  --version-file "%VERSAO%" ^
  --hidden-import=pandas ^
  --hidden-import=openpyxl ^
  --hidden-import=xlrd ^
  --hidden-import=customtkinter ^
  --collect-all customtkinter ^
  "%APP%"
if errorlevel 1 goto :erro
echo.

echo [3/3] Verificando resultado...
if exist "dist\%NOME_EXE%.exe" (
    echo.
    echo  =============================================
    echo   SUCESSO! Executavel gerado em:
    echo   %cd%\dist\%NOME_EXE%.exe
    echo  =============================================
    echo.
    pause
    exit /b 0
)

:erro
echo.
echo  =============================================
echo   ERRO: nao foi possivel gerar o executavel.
echo   Revise as mensagens exibidas acima.
echo  =============================================
echo.
pause
exit /b 1
