@echo off
setlocal enabledelayedexpansion
chcp 65001 > nul

:: ============================================================
::  AT614 Configuration Editor - Script di rilascio
:: ============================================================

set "PROJECT_ROOT=%~dp0"
set "SPEC_FILE=%PROJECT_ROOT%at614_editor.spec"
set "DIST_DIR=%PROJECT_ROOT%dist\at614-editor"
set "BUILD_DIR=%PROJECT_ROOT%build"

cd /d "%PROJECT_ROOT%"

:: Versione via Python (piu' affidabile di parsing del file)
for /f %%v in ('python -c "from at614_editor import __version__; print(__version__)"') do set "VERSION=%%v"
if "%VERSION%"=="" set "VERSION=unknown"

echo.
echo ============================================================
echo   AT614 Configuration Editor - Rilascio v%VERSION%
echo ============================================================
echo.

:: Verifica Python
python --version > nul 2>&1
if errorlevel 1 (
    echo [ERRORE] Python non trovato nel PATH.
    pause & exit /b 1
)
for /f "tokens=*" %%p in ('python --version') do echo   Python:       %%p

:: Verifica PyInstaller
python -m PyInstaller --version > nul 2>&1
if errorlevel 1 (
    echo [ERRORE] PyInstaller non trovato. Esegui: pip install pyinstaller
    pause & exit /b 1
)
for /f "tokens=*" %%p in ('python -m PyInstaller --version') do echo   PyInstaller:  %%p

:: Verifica spec
if not exist "%SPEC_FILE%" (
    echo [ERRORE] File spec non trovato: %SPEC_FILE%
    pause & exit /b 1
)
echo.

:: ── Passo 1: pulizia ─────────────────────────────────────────
echo [1/3] Pulizia cartelle build e dist precedenti...
if exist "%BUILD_DIR%" rmdir /s /q "%BUILD_DIR%"
if exist "%DIST_DIR%"  rmdir /s /q "%DIST_DIR%"
echo       OK.
echo.

:: ── Passo 2: build ───────────────────────────────────────────
echo [2/3] Build in corso (puo' richiedere 1-3 minuti)...
echo.
python -m PyInstaller "%SPEC_FILE%" --noconfirm
if errorlevel 1 (
    echo.
    echo [ERRORE] Build fallita. Controlla i messaggi sopra.
    pause & exit /b 1
)

:: ── Passo 3: ZIP ─────────────────────────────────────────────
echo.
echo [3/3] Creazione archivio ZIP...

set "ZIP_NAME=at614-editor-v%VERSION%.zip"
set "ZIP_PATH=%PROJECT_ROOT%dist\%ZIP_NAME%"

if exist "%ZIP_PATH%" del /q "%ZIP_PATH%"

:: tar.exe (nativo Windows 10+) gestisce correttamente i DLL di sistema gia' in memoria
tar -a -c -f "%ZIP_PATH%" -C "%PROJECT_ROOT%dist" "at614-editor"
if errorlevel 1 (
    echo       [AVVISO] ZIP non creato.
    echo       La cartella dist\at614-editor\ e' comunque pronta per la distribuzione.
) else (
    echo       Creato: dist\%ZIP_NAME%
)

:: ── Riepilogo ────────────────────────────────────────────────
for /f %%s in ('powershell -NoProfile -Command "[math]::Round((Get-ChildItem '%DIST_DIR%' -Recurse | Measure-Object Length -Sum).Sum / 1MB, 1)"') do set "SIZE_MB=%%s"
for /f %%n in ('powershell -NoProfile -Command "(Get-ChildItem '%DIST_DIR%' -Recurse -File).Count"') do set "FILE_COUNT=%%n"

echo.
echo ============================================================
echo   BUILD COMPLETATA  -  v%VERSION%
echo ============================================================
echo   Exe      : dist\at614-editor\at614-editor.exe
echo   Cartella : dist\at614-editor\  (%FILE_COUNT% file, ~%SIZE_MB% MB)
if exist "%ZIP_PATH%" echo   ZIP      : dist\%ZIP_NAME%
echo ------------------------------------------------------------
echo   Per distribuire: copia l'intera cartella dist\at614-editor\
echo   oppure consegna lo ZIP.
echo ============================================================
echo.
pause
endlocal
