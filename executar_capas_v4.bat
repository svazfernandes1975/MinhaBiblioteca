@echo off
setlocal
cd /d "%~dp0"

echo ==============================================================
echo       COLETOR DE CAPAS v4 - VARIACOES DE TITULO
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
%PY% -m pip install --upgrade requests openpyxl
if %errorlevel% neq 0 (
    echo Falha ao instalar bibliotecas.
    pause
    exit /b 1
)

echo.
echo A v4 vai testar varias versoes de cada titulo.
echo Exemplo:
echo   "Os cinco porquinhos (Hercule Poirot, #25)"
echo   sera pesquisado tambem como "Os cinco porquinhos".
echo.
echo Iniciando...
echo.

%PY% baixar_capas_v4.py

echo.
pause
