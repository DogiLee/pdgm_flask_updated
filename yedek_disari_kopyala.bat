@echo off
REM ============================================================
REM  PDGM - Gecelik yedek kopyalama
REM  Task Scheduler ile her gece calistirin.
REM  HEDEF yolunu kendi ag paylasiminizla degistirin.
REM  Robocopy: 0-7 basari/uyari, 8+ hata.
REM ============================================================

set KAYNAK=%~dp0data
set HEDEF=\\dosyasunucu\yedek\pdgm

if not exist "%HEDEF%" (
    echo HATA: Hedef erisilemiyor: %HEDEF%
    exit /b 1
)

for /f %%I in ('powershell -NoProfile -Command "Get-Date -Format yyyyMMdd"') do set BUGUN=%%I

robocopy "%KAYNAK%" "%HEDEF%\%BUGUN%" /E /R:2 /W:5 /NFL /NDL /LOG+:"%HEDEF%\robocopy.log"
set RC=%ERRORLEVEL%

powershell -NoProfile -Command ^
  "Get-ChildItem -Path '%HEDEF%' -Directory | Where-Object { $_.Name -match '^\d{8}$' -and $_.LastWriteTime -lt (Get-Date).AddDays(-60) } | Remove-Item -Recurse -Force"

if %RC% GEQ 8 (
    echo HATA: robocopy basarisiz, kod=%RC%
    exit /b %RC%
)

exit /b 0
