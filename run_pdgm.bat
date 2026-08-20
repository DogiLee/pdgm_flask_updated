@echo off
REM PDGM sunucusunu baslatir.
REM Task Scheduler: "Run only when user is logged on" (Excel COM icin zorunlu)
REM On failure restart: her 1 dk, en fazla 3 kez.

cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
    echo HATA: .venv bulunamadi. Once: python -m venv .venv ^& .venv\Scripts\pip install -r requirements.txt
    exit /b 1
)

.venv\Scripts\python.exe app.py
exit /b %ERRORLEVEL%
