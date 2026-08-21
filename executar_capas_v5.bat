@echo off
setlocal
cd /d "%~dp0"

echo ==============================================================
echo       COLETOR DE CAPAS v5 - BUSCA GERAL NA WEB
echo ==============================================================
echo.

where py >nul 2>&1
if %errorlevel%==0 (
    set "PY=py"
) else (
    where python >nul 2>&1
    if %errorlevel%==0 (
        set "PY=python"
    ) else (
        echo Python nao encontrado.
        echo Instale em https://www.python.org/downloads/
        echo Marque "Add Python to PATH".
        pause
        exit /b 1
    )
)

echo Instalando bibliotecas...
%PY% -m pip install --upgrade requests openpyxl beautifulsoup4
if %errorlevel% neq 0 (
    echo Falha ao instalar as bibliotecas.
    pause
    exit /b 1
)

echo.
echo A v5 vai PRESERVAR as capas encontradas na v4.
echo Ela vai pesquisar principalmente as que faltaram.
echo A nova camada faz busca geral na web e abre as paginas
echo encontradas para localizar a imagem da capa.
echo.
echo Deixe a janela aberta ate terminar.
echo.

%PY% baixar_capas_v5.py

echo.
pause
