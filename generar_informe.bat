@echo off
title Generando Informe de Automatizacion - Word
echo =======================================================
echo   Generando Informe de Efectividad de Automatizacion
echo =======================================================
echo.
echo Procesando archivos de retiros y generando graficos...
python "%~dp0generar_informe_docx.py"
echo.
if %errorlevel% equ 0 (
    echo =======================================================
    echo   [+] ¡Informe de Word generado exitosamente!
    echo   Guardado en: Informe_Analisis_Efectividad_Automatizacion.docx
    echo =======================================================
) else (
    echo =======================================================
    echo   [-] Hubo un error al generar el informe.
    echo   Asegurate de que no tengas el archivo de Word abierto.
    echo =======================================================
)
echo.
pause
