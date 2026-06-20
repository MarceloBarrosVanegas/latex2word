@echo off
chcp 65001 >nul
cls

echo ============================================
echo  CONVERSOR LaTeX --^> Word (con entorno virtual)
echo ============================================
echo.

:: Verificar que el entorno virtual existe
if not exist "venv\Scripts\activate.bat" (
    echo [ERROR] No se encontro el entorno virtual.
    echo Ejecuta primero: 00_setup_entorno.bat
    pause
    exit /b 1
)

:: Activar entorno virtual y ejecutar el convertidor original
call venv\Scripts\activate.bat
call 00_CONVERTIR.bat
