@echo off
title Atualizando Dashboard Logistico...

echo.
echo ==================================================
echo    DASHBOARD LOGISTICO - IAB e VERBUM
echo    Atualizando dados e gerando dashboard...
echo ==================================================
echo.

cd /d "%~dp0etl"

echo [1/2] Processando Quadro de Envios...
echo.
python etl_runner.py
if %errorlevel% neq 0 (
    echo.
    echo ==================================================
    echo    ERRO no ETL - verifique o erro acima
    echo ==================================================
    echo.
    pause
    exit /b 1
)

echo.
echo [2/2] Gerando dashboard.html atualizado...
echo.
python gerar_dashboard.py
if %errorlevel% neq 0 (
    echo.
    echo ==================================================
    echo    ERRO ao gerar dashboard - verifique acima
    echo ==================================================
    echo.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo    DASHBOARD ATUALIZADO COM SUCESSO!
echo    Abrindo no navegador...
echo ==================================================
echo.

start "" "%~dp0dashboard.html"

echo Pressione qualquer tecla para fechar...
pause >nul
