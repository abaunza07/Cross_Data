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
if not exist "cruce_fuid_v3.py" (
    echo ERROR: no encuentro "cruce_fuid_v3.py" en esta carpeta.
    echo Asegurate de que compilar_exe.bat y cruce_fuid_v3.py
    echo esten en la MISMA carpeta.
    echo.
    pause
    exit /b 1
)
echo   OK: cruce_fuid_v3.py encontrado.
echo.

REM ------------------------------------------------------------
REM 3) Instalar dependencias (usando "python -m" para evitar
REM    problemas de PATH con pip/pyinstaller sueltos)
REM ------------------------------------------------------------
echo [3/4] Instalando dependencias (pyinstaller, openpyxl)...
python -m pip install --upgrade pip
python -m pip install --upgrade pyinstaller openpyxl
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
echo [4/4] Compilando cruce_fuid_v3.exe ...
python -m PyInstaller --onefile --console --name cruce_fuid_v3 --clean cruce_fuid_v3.py
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
if exist "dist\cruce_fuid_v3.exe" (
    echo  LISTO. El ejecutable quedo en:
    echo  %cd%\dist\cruce_fuid_v3.exe
) else (
    echo  ADVERTENCIA: el proceso termino sin error pero no encuentro
    echo  dist\cruce_fuid_servidor.exe. Revisa el mensaje de arriba.
)
echo ============================================================
pause
