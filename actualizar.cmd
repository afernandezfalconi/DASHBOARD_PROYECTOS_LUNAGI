@echo off
chcp 65001 >nul
setlocal
cd /d "%~dp0"
set PYTHONIOENCODING=utf-8

echo.
echo ============================================================
echo   AUDITORIA DE LOTES - LUNA GI
echo   Leyendo las bases y regenerando el dashboard...
echo ============================================================
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
