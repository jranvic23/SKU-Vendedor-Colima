@echo off
cd /d "%~dp0"
title Corte semanal SKU
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0Generar_Corte_Semanal.ps1"
set "codigo=%ERRORLEVEL%"
echo.
if not "%codigo%"=="0" (
  echo El corte semanal no se genero. Revisa el mensaje anterior.
)
pause
exit /b %codigo%
