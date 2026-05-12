@echo off
REM ════════════════════════════════════════════════════════
REM   eSuit — Lanzador rapido (Windows)
REM   Requiere: Python 3.10+ y las dependencias de requirements.txt
REM   instaladas (pip install -r requirements.txt).
REM ════════════════════════════════════════════════════════
title eSuit - Calculo electrico profesional
cd /d "%~dp0"

REM Verificar que Python esta disponible
python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo [ERROR] Python no esta instalado o no esta en PATH.
    echo Instala Python 3.10+ desde https://www.python.org/downloads/
    echo y asegurate de marcar "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)

REM Verificar que Streamlit esta instalado
python -c "import streamlit" >nul 2>&1
if errorlevel 1 (
    echo.
    echo [INFO] Instalando dependencias por primera vez...
    python -m pip install --upgrade pip
    python -m pip install -r requirements.txt
    if errorlevel 1 (
        echo.
        echo [ERROR] No se pudieron instalar las dependencias.
        pause
        exit /b 1
    )
)

REM Arrancar la app
echo.
echo Iniciando eSuit...
python launcher.py

REM Si llegamos aqui, el server termino. Mantener ventana abierta.
echo.
echo La aplicacion se ha cerrado.
pause
