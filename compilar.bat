@echo off
REM =====================================================
REM COMPILADOR MODULAR - CONTROL DE PAGOS GCO v2.4.1
REM 3 Archivos: Principal + Semanal + Mensual
REM =====================================================

echo.
echo ==========================================
echo   COMPILADOR CONTROL DE PAGOS GCO v2.4.1
echo   SISTEMA MODULAR (3 archivos)
echo ==========================================
echo.

REM Verificar archivos principales
if not exist "main.py" (
    echo ERROR: No se encuentra main.py
    echo.
    pause
    exit /b 1
)

if not exist "scripts\proceso_semanal.py" (
    echo ERROR: No se encuentra scripts\proceso_semanal.py
    echo.
    pause
    exit /b 1
)

if not exist "scripts\proceso_mensual.py" (
    echo ERROR: No se encuentra scripts\proceso_mensual.py
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
echo [5/5] Compilando sistema modular...
echo.

REM Definir ruta de destino final
set "DEPLOY_PATH=O:\Finanzas\Info Bancos\Pagos Internacionales\PROYECCION PAGOS SEMANAL Y MENSUAL"

REM Compilar - Archivo principal que importa los módulos
pyinstaller --onedir %modo% ^
    --icon=icon.ico ^
    --name="Proyecciones RPA" ^
    --distpath "%DEPLOY_PATH%" ^
    --add-data "icon.ico;." ^
    --add-data "scripts\proceso_semanal.py;scripts" ^
    --add-data "scripts\proceso_mensual.py;scripts" ^
    --hidden-import=scripts.proceso_semanal ^
    --hidden-import=scripts.proceso_mensual ^
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
    main.py

if errorlevel 1 (
    echo.
    echo ERROR: Compilacion fallida
    echo.
    pause
    exit /b 1
)

echo.
echo ==========================================
echo   COMPILACION EXITOSA
echo ==========================================
echo.
echo Ejecutable creado en:
echo   %cd%\dist\Control de Pagos GCO.exe
echo.
echo ESTRUCTURA MODULAR:
echo   main.py  (Interfaz + Coordinacion)
echo   proceso_semanal.py          (Logica semanal)
echo   proceso_mensual.py          (Logica mensual)
echo.
echo Funcionalidades v2.4.1:
echo   + Selector de tipo de proyeccion
echo   + Proyeccion Semanal
echo   + Proyeccion Mensual
echo   + Bucle de repeticion
echo   + Codigo modular y mantenible
echo.
echo Presiona cualquier tecla para abrir la carpeta...
pause >nul

explorer "%DEPLOY_PATH%"