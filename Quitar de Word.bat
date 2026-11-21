@echo off
chcp 65001 >nul
setlocal
rem Quita Sermon Check de Word para este usuario.
set "ID=b5f0c0ab-418b-4d02-b193-c7889874740f"
reg delete "HKCU\Software\Microsoft\Office\16.0\WEF\Developer" /v "%ID%" /f >nul 2>&1
if exist "%LOCALAPPDATA%\SermonCheck\manifest.xml" del "%LOCALAPPDATA%\SermonCheck\manifest.xml"
echo.
echo Sermon Check se quito de Word. Cierra Word y vuelve a abrirlo.
echo.
pause
