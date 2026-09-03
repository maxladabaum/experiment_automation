@echo off
setlocal
title Experiment Automation Launcher

set "HERE=%~dp0"
set "APPDIR="

if exist "%HERE%main.py" set "APPDIR=%HERE%"
if not defined APPDIR if exist "%HERE%experiment_automation\main.py" set "APPDIR=%HERE%experiment_automation\"

if not defined APPDIR (
    echo [ERROR] Could not find the experiment_automation app folder.
    echo.
    echo Put this launcher either:
    echo   1. inside the experiment_automation folder, or
    echo   2. one folder above experiment_automation.
    echo.
    pause
    exit /b 1
)

set "GIT_BASH_BIN="

if exist "%ProgramFiles%\Git\bin\bash.exe" set "GIT_BASH_BIN=%ProgramFiles%\Git\bin\bash.exe"
if not defined GIT_BASH_BIN if exist "%LocalAppData%\Programs\Git\bin\bash.exe" set "GIT_BASH_BIN=%LocalAppData%\Programs\Git\bin\bash.exe"

if not defined GIT_BASH_BIN (
    echo [ERROR] Could not find Git Bash.
    echo.
    echo Install Git for Windows, then double-click this launcher again.
    echo.
    pause
    exit /b 1
)

set "GUI_PYTHON=%LocalAppData%\ExperimentAutomation\venvs\gui32\Scripts\python.exe"
set "ANALYSIS_PYTHON=%LocalAppData%\ExperimentAutomation\venvs\analysis64\Scripts\python.exe"

if not exist "%GUI_PYTHON%" (
    echo [ERROR] The machine-local 32-bit GUI environment was not found:
    echo   %GUI_PYTHON%
    echo.
    pause
    exit /b 1
)

if not exist "%ANALYSIS_PYTHON%" (
    echo [ERROR] The machine-local 64-bit analysis environment was not found:
    echo   %ANALYSIS_PYTHON%
    echo.
    pause
    exit /b 1
)

start "Experiment Automation" "%GIT_BASH_BIN%" --login -i -c "cd '%APPDIR:\=/%' && export EA_BO_ANALYSIS_PYTHON='%ANALYSIS_PYTHON:\=/%' && echo '========================================' && echo '[INFO] Experiment Automation Launcher' && echo '[INFO] App folder: %APPDIR:\=/%' && echo '[INFO] Git Bash: %GIT_BASH_BIN:\=/%' && echo '[INFO] GUI Python: %GUI_PYTHON:\=/%' && echo '[INFO] Analysis Python: %ANALYSIS_PYTHON:\=/%' && echo '[INFO] Launching app in 2 seconds...' && echo '========================================' && sleep 2 && '%GUI_PYTHON:\=/%' -X faulthandler -m main; ec=$?; if [ $ec -ne 0 ]; then echo; echo 'Launcher detected an error (exit' $ec '). Press Enter to close...'; read -r; fi"

exit /b 0
