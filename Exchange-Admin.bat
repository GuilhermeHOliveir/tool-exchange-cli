@echo off
setlocal
title Exchange Online - Administracao

if not exist "%~dp0Exchange-Admin.ps1" (
    echo ERRO: Exchange-Admin.ps1 nao encontrado na pasta deste .bat.
    pause
    exit /b 1
)

pushd "%~dp0"
if errorlevel 1 (
    echo ERRO: nao foi possivel acessar a pasta do programa.
    pause
    exit /b 1
)

"%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%~dp0Exchange-Admin.ps1" %*
set "ExitCode=%ERRORLEVEL%"
popd

if not "%ExitCode%"=="0" (
    echo.
    echo ERRO: o programa encerrou com codigo %ExitCode%.
    pause
)

exit /b %ExitCode%
