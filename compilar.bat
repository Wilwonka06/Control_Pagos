@echo off
REM =====================================================
REM COMPILADOR SIMPLE - CONTROL DE PAGOS GCO v3.0
REM Sin ejecutar automaticamente el .exe al final
REM =====================================================

echo.
echo ========================================
echo   COMPILADOR CONTROL DE PAGOS GCO v3.0
echo ========================================
echo.

REM Verificar archivo
if not exist "control_pagos_semana.py" (
    echo ERROR: No se encuentra control_pagos_semana.py
    echo.
    pause
    exit /b 1
)

REM Verificar Python
echo [1/5] Verificando Python...
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python no esta instalado
    pause
    exit /b 1
)
echo OK
echo.

REM Instalar dependencias
echo [2/5] Verificando dependencias...
pip install pyinstaller pywin32 pandas openpyxl tkcalendar --quiet
echo OK
echo.

REM Limpiar archivos anteriores
echo [3/5] Limpiando compilaciones anteriores...
if exist "build" rmdir /s /q "build" 2>nul
if exist "dist" rmdir /s /q "dist" 2>nul
if exist "*.spec" del /q "*.spec" 2>nul
echo OK
echo.

REM Seleccionar modo
echo [4/5] Modo de compilacion:
echo.
echo 1 = CON CONSOLA (para debug)
echo 2 = SIN CONSOLA (version final)
echo.
set /p opcion="Opcion (1 o 2): "

if "%opcion%"=="1" (
    set "modo=--console"
) else (
    set "modo=--noconsole"
)

echo.
echo [5/5] Compilando...
echo.

REM Compilar
pyinstaller --onefile %modo% ^
    --icon=icon.ico ^
    --name="Control de Pagos" ^
    --add-data "icon.ico;." ^
    --hidden-import=win32com ^
    --hidden-import=win32com.client ^
    --hidden-import=win32com.client.gencache ^
    --hidden-import=win32com.client.CLSIDToClass ^
    --hidden-import=pythoncom ^
    --hidden-import=pywintypes ^
    --hidden-import=win32timezone ^
    --hidden-import=openpyxl ^
    --hidden-import=openpyxl.styles ^
    --hidden-import=openpyxl.styles.fonts ^
    --hidden-import=openpyxl.styles.borders ^
    --hidden-import=openpyxl.styles.alignment ^
    --hidden-import=openpyxl.cell ^
    --hidden-import=openpyxl.utils ^
    --hidden-import=openpyxl.utils.dataframe ^
    --hidden-import=pandas ^
    --hidden-import=tkcalendar ^
    --hidden-import=babel.numbers ^
    --collect-all win32com ^
    --collect-all tkcalendar ^
    control_pagos_semana.py

if errorlevel 1 (
    echo.
    echo ERROR: Compilacion fallida
    echo.
    pause
    exit /b 1
)

echo.
echo ========================================
echo   COMPILACION EXITOSA
echo ========================================
echo.
echo Ejecutable creado en:
echo   %cd%\dist\Control de Pagos.exe
echo.
echo Funcionalidades v3.0:
echo   + Bucle de repeticion (no se cierra)
echo   + Actualizacion automatica de datos
echo   + RefreshAll + CalculateFull
echo.
echo Presiona cualquier tecla para abrir la carpeta...
pause >nul

explorer dist