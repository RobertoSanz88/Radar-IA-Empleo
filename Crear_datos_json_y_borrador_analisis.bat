@echo off
chcp 65001 >nul
set PYTHONIOENCODING=utf-8
cd /d "%~dp0"
title Radar IA-Empleo - Paso 1: datos JSON y borrador del analisis

echo ==========================================================================
echo   RADAR IA-EMPLEO  -  PASO 1: crear datos JSON y borrador del análisis
echo ==========================================================================
echo.
echo   Este script usa el fichero:
echo       "Datos radar SEPE original.xlsx"       (el Excel mensual del SEPE)
echo.
echo   y genera dos ficheros en esta carpeta:
echo       datos_radar.json        las cifras que carga la Web
echo       analisis_borrador.md    borrador del análisis, escrito por reglas
echo.
echo   IMPORTANTE: el borrador NO es el texto final. El comentarista humano de
echo   los datos, opcionalmente ayudado por IA, debe revisarlo y guardarlo como
echo       analisis_final.md       el texto que carga la Web
echo   Este paso NUNCA toca analisis_final.md: si ya existe (por ejemplo, el del
echo   mes anterior), sigue ahí hasta que lo sustituyas por el nuevo.
echo.
echo --------------------------------------------------------------------------

set "PY=python"
%PY% --version >nul 2>&1
if errorlevel 1 set "PY=py"
%PY% --version >nul 2>&1
if errorlevel 1 goto sinpython

if exist "Datos radar SEPE original.xlsx" goto hayexcel
if exist "Datos_radar_SEPE_original.xlsx" goto hayexcel
goto sinexcel

:hayexcel
%PY% Crear_datos_json_y_borrador_analisis.py
if errorlevel 1 goto errorscript

echo.
echo --------------------------------------------------------------------------
echo   HECHO. Ficheros generados en esta carpeta:
for %%F in (datos_radar.json analisis_borrador.md) do if exist "%%F" echo       [OK] %%F
echo.
echo   SIGUIENTES PASOS:
echo     1. Abre analisis_borrador.md, revísalo (o pásaselo a Claude) y
echo        guárdalo como analisis_final.md
echo     2. Ejecuta:  python comprobar_analisis_final.py
echo        (avisa si el periodo o las cifras del texto no cuadran con los datos)
echo     3. Si quieres un HTML suelto con todo dentro, ejecuta:
echo        python embeber_datos.py
echo        Para la Web, sube datos_radar.json y analisis_final.md.
echo --------------------------------------------------------------------------
goto fin

:sinpython
echo.
echo   ERROR: no encuentro Python (ni "python" ni "py") en este equipo.
goto fin

:sinexcel
echo.
echo   ERROR: no encuentro "Datos radar SEPE original.xlsx" en esta carpeta:
echo       %~dp0
echo   Copia ahí el Excel mensual y vuelve a ejecutar este fichero.
goto fin

:errorscript
echo.
echo   ERROR: el script ha fallado (mira el mensaje de arriba).
echo   Si no lo entiendes, copia el mensaje completo y pásaselo a Claude.

:fin
echo.
pause
