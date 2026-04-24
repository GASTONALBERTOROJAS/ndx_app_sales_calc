@echo off
REM ============================================================
REM  NDX Sales Calculation Tool - Lanzador rapido
REM  Requiere Python instalado en el sistema
REM ============================================================
cd /d "%~dp0"

REM Intentar con entorno virtual primero
if exist ".venv\Scripts\python.exe" (
    echo Usando entorno virtual...
    ".venv\Scripts\python.exe" src\main.py
) else (
    REM Usar Python del sistema
    python --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: Python no encontrado en el sistema.
        echo.
        echo Opciones:
        echo   1. Instala Python desde https://www.python.org/downloads/
        echo   2. O usa el ejecutable NDX_Sales_Calculation.exe si tienes la version compilada
        pause
        exit /b 1
    )
    python src\main.py
)
