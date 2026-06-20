@echo off
chcp 65001 >nul
cls

echo ============================================
echo  CONFIGURACION DEL ENTORNO - LaTeX a Word
echo ============================================
echo.

:: Verificar que Python esta disponible
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python no esta en el PATH.
    echo Instala Python desde https://www.python.org/downloads/
    echo Asegurate de marcar "Add Python to PATH" durante la instalacion.
    pause
    exit /b 1
)

:: Crear entorno virtual si no existe
if not exist "venv" (
    echo Creando entorno virtual "venv"...
    python -m venv venv
    if %errorlevel% neq 0 (
        echo [ERROR] No se pudo crear el entorno virtual.
        pause
        exit /b 1
    )
) else (
    echo El entorno virtual "venv" ya existe.
)

:: Activar entorno e instalar/actualizar dependencias
echo.
echo Instalando dependencias en el entorno virtual...
call venv\Scripts\activate.bat
python -m pip install --upgrade pip
pip install -r requirements.txt

if %errorlevel% neq 0 (
    echo.
    echo [ERROR] Fallo la instalacion de dependencias.
    pause
    exit /b 1
)

:: Verificar Pandoc
echo.
echo Verificando Pandoc...
pandoc --version >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo [ADVERTENCIA] Pandoc NO se encontro en el PATH.
    echo El conversor lo necesita para las formulas matematicas.
    echo.
    
    :: Intentar instalar con winget si esta disponible
    winget --version >nul 2>&1
    if %errorlevel% equ 0 (
        echo Intentando instalar Pandoc con winget...
        winget install --id JohnMacFarlane.Pandoc -e --source winget --accept-package-agreements --accept-source-agreements
    ) else (
        echo No se encontro winget. Abriendo pagina de descarga de Pandoc...
        start https://pandoc.org/installing.html
        echo Por favor instala Pandoc manualmente y vuelve a ejecutar este script.
    )
) else (
    echo [OK] Pandoc detectado correctamente.
)

echo.
echo ============================================
echo  CONFIGURACION COMPLETADA
echo ============================================
echo.
echo Para convertir usando este entorno, ejecuta:
echo    00_convertir_con_entorno.bat
echo.
pause
