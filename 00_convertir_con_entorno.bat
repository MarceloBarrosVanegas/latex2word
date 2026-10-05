@echo off
setlocal
cd /d "%~dp0"
chcp 65001 >nul
cls

echo ============================================
echo  CONVERSOR LaTeX --^> Word (con entorno virtual)
echo ============================================
echo.

:: Verificar que el entorno virtual existe
if not exist ".venv\Scripts\python.exe" (
    echo [ERROR] No se encontro el entorno virtual.
    echo Ejecuta primero: 00_setup_entorno.bat
    pause
    exit /b 1
)

:: Comprobar el ejecutable, no solamente la carpeta del entorno.
set "CONVERTER_PYTHON=%~dp0.venv\Scripts\python.exe"
"%CONVERTER_PYTHON%" -c "import docx, lxml, PIL" >nul 2>&1
if errorlevel 1 (
    echo [ERROR] El entorno local no funciona o le faltan dependencias.
    echo Ejecuta 00_setup_entorno.bat para repararlo en este equipo.
    pause
    exit /b 1
)
call 00_CONVERTIR.bat %*
exit /b %errorlevel%
