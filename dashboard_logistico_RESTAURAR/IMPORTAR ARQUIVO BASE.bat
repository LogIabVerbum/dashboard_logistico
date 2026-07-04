@echo off
title Importando Arquivo Base para o Quadro de Envios...

echo.
echo ==================================================
echo    IMPORTACAO ARQUIVO BASE - QUADRO DE ENVIOS
echo    %date% %time%
echo ==================================================
echo.

set PASTA=%~dp0..

echo [1/2] Verificando arquivos...

if not exist "%PASTA%\arquivo_base.xls" (
    echo.
    echo ERRO: arquivo_base.xls nao encontrado em:
    echo %PASTA%
    echo.
    echo Verifique se o arquivo foi exportado do ERP e esta na pasta correta.
    pause
    exit /b 1
)

if not exist "%PASTA%\Quadro de Envios.xlsx" (
    echo.
    echo ERRO: Quadro de Envios.xlsx nao encontrado em:
    echo %PASTA%
    pause
    exit /b 1
)

echo     Arquivos encontrados.

echo.
echo [2/2] Importando novas NFs para o Quadro de Envios...
echo     (Backup automatico sera gerado antes de qualquer alteracao)
echo.

cd /d "%~dp0etl"
python importar_base.py

if %errorlevel% neq 0 (
    echo.
    echo ==================================================
    echo    ERRO na importacao - verifique o erro acima
    echo ==================================================
    echo.
    pause
    exit /b 1
)

echo.
echo ==================================================
echo    Importacao concluida!
echo    Confira o Quadro de Envios, faca os ajustes
echo    e depois rode o ATUALIZAR DASHBOARD.bat
echo ==================================================
echo.
echo Pressione qualquer tecla para fechar...
pause >nul
