@echo off
setlocal enabledelayedexpansion
cd /d "%~dp0"
chcp 65001 >nul
cls

:: ============================================================
:: CONVERSOR LaTeX a Word - Galapagos Water Project
:: ============================================================
:: TODO SE HACE EN CARPETA TEMPORAL - Solo el .docx final se guarda
:: ============================================================

:: ------------------------------------------------------------
:: CONFIGURA AQUI LA RUTA COMPLETA DE TU ARCHIVO .tex
:: ------------------------------------------------------------
SET TEX_PATH=C:\Users\Alienware\OneDrive\00_OFC\12_documentos_ofc\02_ISP\02_operational_accompaniment\ToR_operational_accompaniment_en_v4.tex
:: ------------------------------------------------------------

:: CONFIGURA AQUI LA RUTA DE UNA PLANTILLA .docx (OPCIONAL)
:: Dejala en blanco si no quieres usar plantilla.
:: Ejemplo: SET TEMPLATE_PATH=C:\Users\Usuario\Documentos\plantilla.docx
:: ------------------------------------------------------------
:: Plantilla usada anteriormente, disponible si quieres volver a configurarla:
:: 
SET TEMPLATE_PATH=
::SET TEMPLATE_PATH=C:\Users\Alienware\OneDrive\00_OFC\12_documentos_ofc\02_ISP\01_linea_base\01_procurment\03_invitations\00_word\00_formato.docx
:: ------------------------------------------------------------

:: Argumentos opcionales: entrada.tex salida.docx plantilla.docx
:: Usa "-" como tercer argumento para conservar el formato del LaTeX.
if not "%~1"=="" set "TEX_PATH=%~1"
if "%~3"=="-" (
    set "TEMPLATE_PATH="
) else (
    if not "%~3"=="" set "TEMPLATE_PATH=%~3"
)

:: Extraer solo el nombre del archivo (sin ruta ni extension)
FOR %%F IN ("%TEX_PATH%") DO SET TEX_FILE=%%~nF
SET OUTPUT_DOCX=%TEX_FILE%.docx
if not "%~2"=="" set "OUTPUT_DOCX=%~2"

echo ============================================
echo  CONVERSOR LaTeX --^> Word
echo ============================================
echo.
echo Archivo entrada: %TEX_PATH%
echo Archivo salida:  %OUTPUT_DOCX%
if defined TEMPLATE_PATH (
    if exist "%TEMPLATE_PATH%" (
        echo Plantilla:       %TEMPLATE_PATH%
    ) else (
        echo Plantilla:       no encontrada, se usara formato por defecto
    )
) else (
    echo Plantilla:       no configurada
)
echo.

:: Verificar que existe el archivo .tex
if not exist "%TEX_PATH%" (
    echo [ERROR] No se encuentra: %TEX_PATH%
    echo.
    pause
    exit /b 1
)

:: Usar el entorno local cuando esta configurado y sea valido para ESTE equipo.
:: Un .venv sincronizado por OneDrive desde otro equipo apunta en pyvenv.cfg
:: a un 'home' que no existe aqui; en ese caso se ignora y se prueba otro entorno.
if not defined CONVERTER_PYTHON (
    for %%V in (.venv venv) do (
        if not defined CONVERTER_PYTHON (
            if exist "%%~V\Scripts\python.exe" (
                set "VENV_HOME="
                for /f "usebackq tokens=1,2 delims== " %%A in ("%%~V\pyvenv.cfg") do (
                    if /i "%%~A"=="home" if not defined VENV_HOME set "VENV_HOME=%%~B"
                )
                if defined VENV_HOME (
                    if exist "!VENV_HOME!\python.exe" (
                        set "CONVERTER_PYTHON=%~dp0%%~V\Scripts\python.exe"
                    )
                )
            )
        )
    )
)
if not defined CONVERTER_PYTHON set "CONVERTER_PYTHON=python"

:: Ejecutar conversion (todo en carpeta temporal)
:: Si TEMPLATE_PATH esta definida y existe, se pasa como tercer argumento.
if defined TEMPLATE_PATH (
    if exist "%TEMPLATE_PATH%" (
        "%CONVERTER_PYTHON%" latex_to_docx.py "%TEX_PATH%" "%OUTPUT_DOCX%" "%TEMPLATE_PATH%"
    ) else (
        "%CONVERTER_PYTHON%" latex_to_docx.py "%TEX_PATH%" "%OUTPUT_DOCX%"
    )
) else (
    "%CONVERTER_PYTHON%" latex_to_docx.py "%TEX_PATH%" "%OUTPUT_DOCX%"
)

if errorlevel 1 (
    echo.
    echo [ERROR] La conversion fallo
    pause
    exit /b 1
)

echo.
echo ============================================
echo  [OK] Documento generado!
echo ============================================
echo.
echo ARCHIVO: %OUTPUT_DOCX%
echo CARPETA: %CD%
echo.
echo ============================================
echo  ACTUALIZAR CAMPOS EN WORD
echo ============================================
echo.
echo 1. Abre el documento en Word
echo.
echo 2. Presiona Ctrl+A (selecciona TODO)
echo.
echo 3. Presiona F9 (actualiza campos)
echo.
echo 4. Si hay un indice, selecciona "Update entire table"
echo    (Actualizar tabla completa)
echo.
echo 5. Haz clic en OK
echo.
echo 6. Guarda el documento (Ctrl+S)
echo.
echo    La proxima vez que abras el archivo, los indices
echo    ya apareceran actualizados sin ningun mensaje.
echo.
echo    Los indices solo aparecen cuando el LaTeX los solicita.
echo.
echo ============================================
pause
