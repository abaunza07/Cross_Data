@echo off
setlocal enabledelayedexpansion
pushd "%~dp0"

echo ============================================================
echo  Carpeta actual: %cd%
echo  (aqui deberian aparecer al final las carpetas build y dist)
echo ============================================================
echo.

REM ------------------------------------------------------------
REM 1) Verificar que Python esta instalado y accesible
REM ------------------------------------------------------------
echo [1/4] Verificando Python...
where python >nul 2>nul
if %errorlevel% neq 0 (
    echo.
    echo ERROR: "python" no se reconoce como comando.
    echo Prueba abrir una consola CMD y escribir:  python --version
    echo Si tampoco funciona ahi, Python no esta instalado o no esta
    echo agregado al PATH de Windows. Instalalo desde python.org
    echo marcando la casilla "Add Python to PATH" durante la instalacion.
    echo.
    pause
    exit /b 1
)
python --version
echo.

REM ------------------------------------------------------------
REM 2) Verificar que los archivos necesarios existen aqui
REM ------------------------------------------------------------
echo [2/4] Verificando archivos...
if not exist "cruce_fuid.py" (
    echo ERROR: no encuentro "cruce_fuid.py" en esta carpeta.
    echo Asegurate de que compilar_exe.bat y cruce_fuid.py
    echo esten en la MISMA carpeta.
    echo.
    pause
    exit /b 1
)
echo   OK: cruce_fuid.py encontrado.
echo.

REM ------------------------------------------------------------
REM La version sale de la mas alta entre cruce_fuid.version y
REM los .exe ya creados (cruce_fuid_v5.exe cuenta como 5.0).
REM El usuario solo elige menor o mayor. No se reescribe un .exe
REM que ya exista.
REM ------------------------------------------------------------
set "ULT_MAYOR=0"
set "ULT_MENOR=0"
if exist "cruce_fuid.version" (
    set /p VERSION_GUARDADA=<cruce_fuid.version
    call :tomar_version "!VERSION_GUARDADA!"
)
if exist "dist\cruce_fuid_*.exe" (
    for %%F in (dist\cruce_fuid_*.exe) do (
        set "NUM=%%~nF"
        set "NUM=!NUM:cruce_fuid_=!"
        if /i "!NUM:~0,1!"=="v" set "NUM=!NUM:~1!"
        call :tomar_version "!NUM!"
    )
)
set /a "MENOR_SIG=ULT_MENOR+1"
set /a "MAYOR_SIG=ULT_MAYOR+1"
echo Version actual: !ULT_MAYOR!.!ULT_MENOR!
echo.
echo  1^) Cambio menor  -^> !ULT_MAYOR!.!MENOR_SIG!
echo  2^) Cambio mayor  -^> !MAYOR_SIG!.0
echo.
:pedir_tipo
set "TIPO="
set /p TIPO=Elige 1 o 2: 
if "!TIPO!"=="1" (
    set "NUEVA_VERSION=!ULT_MAYOR!.!MENOR_SIG!"
) else if "!TIPO!"=="2" (
    set "NUEVA_VERSION=!MAYOR_SIG!.0"
) else (
    echo Escribe 1 o 2.
    goto pedir_tipo
)
set "NOMBRE_EXE=cruce_fuid_!NUEVA_VERSION!"
if exist "dist\!NOMBRE_EXE!.exe" (
    echo.
    echo ERROR: ya existe dist\!NOMBRE_EXE!.exe
    echo No se crea otra vez la misma version.
    echo.
    pause
    exit /b 1
)
echo.
echo El ejecutable se llamara: dist\!NOMBRE_EXE!.exe
echo.

REM ------------------------------------------------------------
REM 3) Instalar dependencias (usando "python -m" para evitar
REM    problemas de PATH con pip/pyinstaller sueltos)
REM ------------------------------------------------------------
echo [3/4] Instalando dependencias (pyinstaller, openpyxl, pypdf)...
python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller openpyxl pypdf
if %errorlevel% neq 0 (
    echo.
    echo ERROR: fallo la instalacion de dependencias. Revisa el mensaje
    echo de error de arriba ^(problema de internet, permisos, proxy, etc^).
    echo.
    pause
    exit /b 1
)
echo.

REM ------------------------------------------------------------
REM 4) Compilar usando "python -m PyInstaller" (mas confiable que
REM    el comando "pyinstaller" suelto, que a veces no queda en PATH)
REM ------------------------------------------------------------
echo [4/4] Compilando !NOMBRE_EXE!.exe ...
python -m PyInstaller --onefile --console --name "!NOMBRE_EXE!" --clean cruce_fuid.py
if %errorlevel% neq 0 (
    echo.
    echo ERROR: la compilacion fallo. Copia el mensaje de error de
    echo arriba y enviamelo para revisarlo.
    echo.
    pause
    exit /b 1
)

echo.
echo ============================================================
if exist "dist\!NOMBRE_EXE!.exe" (
    >cruce_fuid.version echo !NUEVA_VERSION!
    echo  LISTO. El ejecutable quedo en:
    echo  %cd%\dist\!NOMBRE_EXE!.exe
    echo  Version registrada: !NUEVA_VERSION!
) else (
    echo  ADVERTENCIA: el proceso termino sin error pero no encuentro
    echo  dist\!NOMBRE_EXE!.exe. Revisa el mensaje de arriba.
)
echo ============================================================
pause
exit /b 0

:tomar_version
set "TM="
set "TN="
for /f "tokens=1,2 delims=." %%A in ("%~1") do (
    set "TM=%%A"
    set "TN=%%B"
)
if not defined TM goto :eof
if not defined TN set "TN=0"
set /a "TM_NUM=TM" 2>nul
if errorlevel 1 goto :eof
set /a "TN_NUM=TN" 2>nul
if errorlevel 1 goto :eof
set /a "PUNTOS_NUEVO=TM_NUM*1000+TN_NUM"
set /a "PUNTOS_VIEJO=ULT_MAYOR*1000+ULT_MENOR"
if !PUNTOS_NUEVO! GTR !PUNTOS_VIEJO! (
    set "ULT_MAYOR=!TM_NUM!"
    set "ULT_MENOR=!TN_NUM!"
)
goto :eof
