@echo off
REM =====================================================
REM SCRIPT DE COMPILACION - CONTROL DE PAGOS GCO v3.0
REM =====================================================

echo.
echo ========================================
echo   COMPILADOR DE CONTROL DE PAGOS GCO
echo   VERSION 3.0  
echo ========================================
echo.

REM =====================================================
REM PASO 0: Verificar que estamos en la carpeta correcta
REM =====================================================
echo [0/7] Verificando entorno...

if not exist "control_pagos_v1.py" (
    echo.
    echo ERROR: No se encuentra control_pagos_v1.py
    echo.
    echo Asegurate de ejecutar este .bat desde la carpeta donde esta:
    echo   - control_pagos_v1.py
    echo   - icon.ico
    echo.
    pause
    exit /b 1
)

if not exist "icon.ico" (
    echo.
    echo ADVERTENCIA: No se encuentra icon.ico
    echo El ejecutable se creara sin icono personalizado
    echo.
    timeout /t 3 >nul
)

echo OK: Archivos encontrados
echo.

REM =====================================================
REM PASO 1: Verificar Python
REM =====================================================
echo [1/7] Verificando Python...

python --version >nul 2>&1
if errorlevel 1 (
    echo.
    echo ERROR: Python no esta instalado o no esta en el PATH
    echo.
    echo Descarga Python desde: https://www.python.org/downloads/
    echo IMPORTANTE: Durante la instalacion marca "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

for /f "tokens=*" %%i in ('python --version') do set PYTHON_VERSION=%%i
echo OK: %PYTHON_VERSION% instalado
echo.

REM =====================================================
REM PASO 2: Verificar e instalar dependencias
REM =====================================================
echo [2/7] Verificando dependencias...
echo.

REM PyInstaller
python -c "import PyInstaller" 2>nul
if errorlevel 1 (
    echo Instalando PyInstaller...
    pip install pyinstaller
    if errorlevel 1 (
        echo ERROR: No se pudo instalar PyInstaller
        pause
        exit /b 1
    )
)
echo   - PyInstaller: OK

REM pywin32
python -c "import win32com.client" 2>nul
if errorlevel 1 (
    echo Instalando pywin32...
    pip install pywin32
    if errorlevel 1 (
        echo ERROR: No se pudo instalar pywin32
        pause
        exit /b 1
    )
    
    REM Ejecutar post-install de pywin32
    echo Configurando pywin32...
    python Scripts\pywin32_postinstall.py -install 2>nul
)
echo   - pywin32: OK

REM pandas
python -c "import pandas" 2>nul
if errorlevel 1 (
    echo Instalando pandas...
    pip install pandas
    if errorlevel 1 (
        echo ERROR: No se pudo instalar pandas
        pause
        exit /b 1
    )
)
echo   - pandas: OK

REM openpyxl
python -c "import openpyxl" 2>nul
if errorlevel 1 (
    echo Instalando openpyxl...
    pip install openpyxl
    if errorlevel 1 (
        echo ERROR: No se pudo instalar openpyxl
        pause
        exit /b 1
    )
)
echo   - openpyxl: OK

REM tkcalendar
python -c "import tkcalendar" 2>nul
if errorlevel 1 (
    echo Instalando tkcalendar...
    pip install tkcalendar
    if errorlevel 1 (
        echo ERROR: No se pudo instalar tkcalendar
        pause
        exit /b 1
    )
)
echo   - tkcalendar: OK

echo.
echo OK: Todas las dependencias instaladas
echo.

REM =====================================================
REM PASO 3: Limpiar compilaciones anteriores
REM =====================================================
echo [3/7] Limpiando archivos temporales...

if exist "build" (
    echo Eliminando carpeta build...
    rmdir /s /q "build" 2>nul
)

if exist "dist" (
    echo Eliminando carpeta dist...
    rmdir /s /q "dist" 2>nul
)

if exist "Control de Pagos.spec" (
    echo Eliminando .spec anterior...
    del /q "Control de Pagos.spec" 2>nul
)

if exist "__pycache__" (
    echo Eliminando cache de Python...
    rmdir /s /q "__pycache__" 2>nul
)

echo OK: Archivos temporales eliminados
echo.

REM =====================================================
REM PASO 4: Seleccionar modo de compilacion
REM =====================================================
echo [4/7] Seleccione el modo de compilacion:
echo.
echo 1. CON CONSOLA (recomendado para PRIMERA VEZ / debugging)
echo    - Muestra ventana de consola con mensajes de error
echo    - Util para identificar problemas
echo.
echo 2. SIN CONSOLA (version FINAL, solo interfaz grafica)
echo    - No muestra consola
echo    - Mas profesional pero no muestra errores
echo.
set /p opcion="Ingrese opcion (1 o 2): "
echo.

if "%opcion%"=="1" (
    set "modo=--console"
    set "nombre_modo=CON CONSOLA (DEBUG)"
    echo IMPORTANTE: Se mostrara una ventana de consola al ejecutar
    echo             Esto es NORMAL y ayuda a ver errores
) else if "%opcion%"=="2" (
    set "modo=--noconsole"
    set "nombre_modo=SIN CONSOLA (FINAL)"
    echo Solo se mostrara la interfaz grafica
) else (
    echo ERROR: Opcion invalida
    pause
    exit /b 1
)

echo.
echo Modo seleccionado: %nombre_modo%
echo.

REM =====================================================
REM PASO 5: Compilar con PyInstaller
REM =====================================================
echo [5/7] Compilando aplicacion...
echo.
echo Esto puede tomar varios minutos...
echo No cierres esta ventana hasta que termine.
echo.

REM Comando de compilación con TODOS los hidden-imports necesarios
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
    --hidden-import=openpyxl.cell ^
    --hidden-import=openpyxl.utils ^
    --hidden-import=pandas ^
    --hidden-import=tkcalendar ^
    --hidden-import=babel.numbers ^
    --collect-all win32com ^
    --collect-all tkcalendar ^
    control_pagos_v1.py

if errorlevel 1 (
    echo.
    echo ========================================
    echo ERROR: La compilacion FALLO
    echo ========================================
    echo.
    echo Revise los mensajes de error arriba.
    echo.
    echo SOLUCIONES COMUNES:
    echo 1. Cierra cualquier archivo Excel que este abierto
    echo 2. Desactiva temporalmente el antivirus
    echo 3. Ejecuta este .bat como Administrador
    echo 4. Verifica que no haya errores en control_pagos_v1.py
    echo.
    pause
    exit /b 1
)

echo.
echo OK: Compilacion completada
echo.

REM =====================================================
REM PASO 6: Verificar que el ejecutable se creo
REM =====================================================
echo [6/7] Verificando ejecutable...

if not exist "dist\Control de Pagos.exe" (
    echo.
    echo ERROR: El ejecutable no se genero
    echo.
    echo Revisa los mensajes de error arriba
    echo.
    pause
    exit /b 1
)

for %%I in ("dist\Control de Pagos.exe") do set "TAMANO=%%~zI"
set /a TAMANO_MB=%TAMANO% / 1048576

echo OK: Ejecutable creado
echo    Ubicacion: %cd%\dist\Control de Pagos.exe
echo    Tamano: %TAMANO_MB% MB
echo.

REM =====================================================
REM PASO 7: Prueba del ejecutable
REM =====================================================
echo [7/7] Prueba del ejecutable...
echo.

if "%modo%"=="--console" (
    echo IMPORTANTE: 
    echo El ejecutable se va a abrir AHORA para probarlo.
    echo.
    echo Como compilaste con CONSOLA visible:
    echo   - Veras una ventana negra (consola) Y la interfaz grafica
    echo   - Esto es NORMAL
    echo   - Revisa la consola si hay errores
    echo.
    echo Si TODO funciona bien:
    echo   1. Cierra el programa
    echo   2. Vuelve a compilar con opcion 2 (SIN CONSOLA)
    echo   3. Asi tendras la version final sin consola
    echo.
) else (
    echo IMPORTANTE:
    echo El ejecutable se va a abrir AHORA para probarlo.
    echo.
    echo Como compilaste SIN CONSOLA:
    echo   - Solo veras la interfaz grafica
    echo   - NO veras mensajes de error si algo falla
    echo.
    echo Si el programa NO ARRANCA o se cierra inmediatamente:
    echo   1. Vuelve a compilar con opcion 1 (CON CONSOLA)
    echo   2. Revisa los errores en la consola
    echo   3. Comparte esos errores para recibir ayuda
    echo.
)

pause

REM Ejecutar el .exe
cd dist
start Control de Pagos.exe
cd ..

echo.
echo ========================================
echo   COMPILACION EXITOSA
echo ========================================
echo.
echo El ejecutable esta en:
echo   %cd%\dist\Control de Pagos.exe
echo.
echo PROXIMOS PASOS:
echo.

if "%modo%"=="--console" (
    echo 1. Prueba el programa completamente
    echo 2. Si funciona bien, vuelve a compilar con opcion 2
    echo    para eliminar la ventana de consola
    echo 3. Distribuye el .exe final
) else (
    echo 1. Si el programa funciono correctamente:
    echo    - Puedes distribuir este .exe
    echo    - Copia dist\Control de Pagos.exe a donde lo necesites
    echo.
    echo 2. Si hubo problemas:
    echo    - Vuelve a compilar con opcion 1 (CON CONSOLA)
    echo    - Revisa los errores que aparezcan
    echo    - Comparte los errores para recibir ayuda
)

echo.
echo [BONUS] Abriendo carpeta de destino...
timeout /t 2 >nul
explorer dist

echo.
echo PROCESO COMPLETADO.
echo.
pause