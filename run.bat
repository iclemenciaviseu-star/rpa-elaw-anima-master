@echo off
chcp 65001 >nul
setlocal EnableDelayedExpansion
cd /d "%~dp0"

echo.
echo ====================================================
echo   RPA eLaw Anima - Launcher
echo ====================================================
echo.

echo [1/4] Verificando Python...
where python >nul 2>nul
if errorlevel 1 (
    echo [ERRO] Python nao encontrado no PATH.
    echo Reinstale de https://www.python.org/downloads/ marcando "Add Python to PATH"
    pause
    exit /b 1
)

for /f "tokens=2" %%v in ('python --version 2^>^&1') do set PYVER=%%v
echo    Encontrado: Python !PYVER!

for /f "tokens=1,2 delims=." %%a in ("!PYVER!") do (
    set PYMAJOR=%%a
    set PYMINOR=%%b
)
if !PYMAJOR! LSS 3 goto pyoldver
if !PYMAJOR! EQU 3 if !PYMINOR! LSS 10 goto pyoldver
goto pyok
:pyoldver
echo [ERRO] Python !PYVER! e muito antigo. Precisamos de 3.10+.
pause
exit /b 1
:pyok

echo [2/4] Verificando ambiente virtual...
if not exist ".venv\Scripts\python.exe" (
    echo    Criando ambiente virtual...
    python -m venv .venv
    if errorlevel 1 ( echo [ERRO] Falha ao criar venv. & pause & exit /b 1 )
    echo    OK.
) else (
    echo    OK.
)

call ".venv\Scripts\activate.bat"

echo [3/4] Verificando dependencias...
if not exist ".venv\.installed" (
    echo    Instalando pacotes Python...
    python -m pip install --upgrade pip --quiet
    python -m pip install -r requirements.txt
    if errorlevel 1 ( echo [ERRO] Falha na instalacao. & pause & exit /b 1 )
    echo    Baixando o Chromium do Playwright...
    python -m playwright install chromium
    if errorlevel 1 ( echo [ERRO] Falha ao baixar o Chromium. & pause & exit /b 1 )
    type nul > ".venv\.installed"
    echo    [OK] Instalacao concluida!
) else (
    echo    OK.
)

echo [4/4] Iniciando servidor local...
echo.
echo ====================================================
echo   Servidor: http://127.0.0.1:8765
echo   Mantenha esta janela aberta durante o uso.
echo ====================================================
echo.

timeout /t 2 /nobreak >nul
start "" "http://127.0.0.1:8765"

python server.py

echo.
echo ====================================================
echo   O servidor encerrou.
echo ====================================================
pause
endlocal
