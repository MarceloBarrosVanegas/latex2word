@echo off
chcp 65001 >nul
cls

:: ============================================================
:: CONVERSOR LaTeX a Word - Galapagos Water Project
:: ============================================================
:: TODO SE HACE EN CARPETA TEMPORAL - Solo el .docx final se guarda
:: ============================================================

:: ------------------------------------------------------------
:: CONFIGURA AQUI LA RUTA COMPLETA DE TU ARCHIVO .tex
:: Ejemplo: SET TEX_PATH=C:\Users\Usuario\Documentos\mi_archivo.tex
:: ------------------------------------------------------------
SET TEX_PATH=C:\Users\Alienware\OneDrive\PROYECTOS\07_elecaustro_tuni\02_Ejecucion\FASE_1\00_geotenica\informe.tex
:: ------------------------------------------------------------

:: Extraer solo el nombre del archivo (sin ruta ni extension)
FOR %%F IN ("%TEX_PATH%") DO SET TEX_FILE=%%~nF
SET OUTPUT_DOCX=%TEX_FILE%.docx

echo ============================================
echo  CONVERSOR LaTeX --^> Word
echo ============================================
echo.
echo Archivo entrada: %TEX_PATH%
echo Archivo salida:  %OUTPUT_DOCX%
echo.

:: Verificar que existe el archivo .tex
if not exist "%TEX_PATH%" (
    echo [ERROR] No se encuentra: %TEX_PATH%
    echo.
    pause
    exit /b 1
)

:: Ejecutar conversion (todo en carpeta temporal)
python latex_to_docx.py "%TEX_PATH%" "%OUTPUT_DOCX%"

if %errorlevel% neq 0 (
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
echo  COMO ACTIVAR LOS INDICES (IMPORTANTE!)
echo ============================================
echo.
echo 1. Abre el documento en Word
echo.
echo 2. Presiona Ctrl+A (selecciona TODO)
echo.
echo 3. Presiona F9 (actualiza campos)
echo.
echo 4. Selecciona "Update entire table"
echo    (Actualizar tabla completa)
echo.
echo 5. Haz clic en OK
echo.
echo 6. Guarda el documento (Ctrl+S)
echo.
echo    La proxima vez que abras el archivo, los indices
echo    ya apareceran actualizados sin ningun mensaje.
echo.
echo    Nota: hay 3 indices: Contenidos, Figuras y Tablas
echo.
echo ============================================
pause
