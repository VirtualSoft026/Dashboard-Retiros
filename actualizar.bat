@echo off
echo ======================================================================
echo INICIANDO ACTUALIZACION DE REPORTES DE RETIROS DESDE MICROSTRATEGY
echo ======================================================================
echo.
python "%~dp0test_mstr_connection.py"
echo.
if %errorlevel% neq 0 (
    echo ======================================================================
    echo [-] HUBO UN ERROR DURANTE LA ACTUALIZACION. NO SE SUBIRAN CAMBIOS.
    echo ======================================================================
    pause
    exit /b %errorlevel%
)

echo ======================================================================
echo SUBIENDO CAMBIOS A GITHUB PAGES
echo ======================================================================
git add index.html Dashboard_Efectividad.html Reporte_Efectividad_Automatizacion.xlsx
git diff --cached --quiet
if %errorlevel% neq 0 (
    git commit -m "Actualizacion automatica dashboard %date% %time%"
    git push origin main
    echo.
    echo ======================================================================
    echo PROCESO COMPLETADO EXITOSAMENTE. NUEVOS CAMBIOS SUBIDOS A LA WEB.
    echo ======================================================================
) else (
    echo [+] El Dashboard web ya se encuentra al dia con la ultima informacion.
    echo.
    echo ======================================================================
    echo PROCESO COMPLETADO EXITOSAMENTE. ENLACE WEB AL DIA.
    echo ======================================================================
)
pause
