@echo off
chcp 65001 >nul
setlocal
rem Registra Sermon Check como complemento de Word para este usuario.
set "ID=b5f0c0ab-418b-4d02-b193-c7889874740f"
set "DEST=%LOCALAPPDATA%\SermonCheck"

if not exist "%~dp0manifest.xml" (
  echo No se encuentra manifest.xml junto a este archivo.
  pause
  exit /b 1
)
if not exist "%DEST%" mkdir "%DEST%"
copy /y "%~dp0manifest.xml" "%DEST%\manifest.xml" >nul
if errorlevel 1 (
  echo No se pudo copiar el manifiesto a %DEST%.
  pause
  exit /b 1
)
reg add "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ID%" /t REG_SZ /d "%DEST%\manifest.xml" /f >nul
if errorlevel 1 (
  echo No se pudo registrar el complemento.
  pause
  exit /b 1
)

echo.
echo Sermon Check quedo instalado.
echo.
echo  1. Cierra Word por completo y vuelve a abrirlo.
echo  2. En la pestana Inicio aparecera el boton "Sermon Check".
echo.
echo Si no aparece: Inicio ^> Complementos ^> Mas complementos ^> pestana
echo "Mis complementos" ^> "Complementos de desarrollador".
echo.
pause
