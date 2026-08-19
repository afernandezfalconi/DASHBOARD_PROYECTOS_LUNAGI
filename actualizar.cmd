@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

echo.
echo ============================================================
echo   AUDITORIA DE LOTES - LUNA GI
echo   Leyendo AUDITORIA_LOTES.xlsx y regenerando el dashboard...
echo ============================================================
echo.

echo   *** OJO ***
echo   La fuente de verdad hoy es el GOOGLE SHEET, no la copia local.
echo   Para publicar usa el Sheet:  menu Dashboard -^> Publicar ahora
echo.
echo   Seguir por aqui PISA lo que publicaste desde el Sheet con la
echo   copia local, que puede estar vieja. Es solo un respaldo.
echo.
choice /C SN /N /M "   Continuar de todos modos? (S/N): "
if errorlevel 2 (
  echo.
  echo   Cancelado. No se toco nada.
  exit /b 0
)
echo.

python scripts\auditar.py --silencio
if errorlevel 1 (
  echo.
  echo *** ERROR al generar. No se publico nada. ***
  pause
  exit /b 1
)

echo.
git add -A
git diff --cached --quiet
if not errorlevel 1 (
  echo   No hubo cambios: el dashboard ya estaba al dia.
  pause
  exit /b 0
)

for /f "tokens=1-3 delims=/ " %%a in ("%date%") do set HOY=%%c-%%b-%%a
git commit -m "Actualiza auditoria de lotes (%HOY%)"
if errorlevel 1 (
  echo *** ERROR al hacer commit. ***
  pause
  exit /b 1
)

git push
if errorlevel 1 (
  echo.
  echo *** ERROR al publicar. Revisa tu conexion o el acceso a GitHub. ***
  pause
  exit /b 1
)

echo.
echo ============================================================
echo   LISTO. El dashboard vivo ya tiene los cambios:
echo   https://afernandezfalconi.github.io/DASHBOARD_PROYECTOS_LUNAGI/
echo.
echo   (GitHub tarda ~1 minuto en publicar. Si no ves el cambio,
echo    recarga con Ctrl+F5.)
echo ============================================================
echo.
pause
