@echo off
echo ============================================================
echo SUBIENDO CAMBIOS DE RETIROS A GITHUB PAGES
echo ============================================================
git add index.html Dashboard_Efectividad.html Reporte_Efectividad_Automatizacion.xlsx
git diff --cached --quiet
if %errorlevel% neq 0 (
    git commit -m "Actualizacion automatica dashboard %date% %time%"
    git push origin main
    echo ============================================================
    echo DASHBOARD ACTUALIZADO EN GITHUB PAGES CON EXITO
    echo ============================================================
) else (
    echo [+] El Dashboard web ya se encuentra al dia en GitHub Pages.
    echo ============================================================
    echo NO HAY CAMBIOS PENDIENTES. ENLACE WEB AL DIA.
    echo ============================================================
)
pause
