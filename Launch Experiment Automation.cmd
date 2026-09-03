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

set "GUI_PYTHON="
call :use_gui_python_if_ready "%EA_GUI_PYTHON%"
call :use_gui_python_if_ready "%LocalAppData%\ExperimentAutomation\venvs\gui32\Scripts\python.exe"
call :use_gui_python_if_ready "%APPDIR%venv_gui\Scripts\python.exe"
call :use_gui_python_if_ready "%USERPROFILE%\Documents\GitHub\experiment_automation\venv_gui\Scripts\python.exe"

if not exist "%GUI_PYTHON%" (
    echo [ERROR] No compatible GUI Python environment was found.
    echo Set EA_GUI_PYTHON or install dependencies in the machine-local GUI environment.
    echo.
    pause
    exit /b 1
)

set "ANALYSIS_PYTHON=%EA_BO_ANALYSIS_PYTHON%"
if not exist "%ANALYSIS_PYTHON%" set "ANALYSIS_PYTHON=%LocalAppData%\ExperimentAutomation\venvs\analysis64\Scripts\python.exe"
if not exist "%ANALYSIS_PYTHON%" set "ANALYSIS_PYTHON=%USERPROFILE%\anaconda3\envs\ea-bo-analysis\python.exe"

set "ANALYSIS_EXPORT="
set "ANALYSIS_INFO=[WARN] Machine-local 64-bit analysis environment not found; using configured analysis fallback."
if exist "%ANALYSIS_PYTHON%" (
    set "ANALYSIS_EXPORT=export EA_BO_ANALYSIS_PYTHON='%ANALYSIS_PYTHON:\=/%' && "
    set "ANALYSIS_INFO=[INFO] Analysis Python: %ANALYSIS_PYTHON:\=/%"
)

if not exist "%ANALYSIS_PYTHON%" (
    echo [WARN] The optional machine-local 64-bit analysis environment was not found:
    echo   %ANALYSIS_PYTHON%
)

start "Experiment Automation" "%GIT_BASH_BIN%" --login -i -c "cd '%APPDIR:\=/%' && %ANALYSIS_EXPORT%echo '========================================' && echo '[INFO] Experiment Automation Launcher' && echo '[INFO] App folder: %APPDIR:\=/%' && echo '[INFO] Git Bash: %GIT_BASH_BIN:\=/%' && echo '[INFO] GUI Python: %GUI_PYTHON:\=/%' && echo '%ANALYSIS_INFO%' && echo '[INFO] Launching app in 2 seconds...' && echo '========================================' && sleep 2 && '%GUI_PYTHON:\=/%' -X faulthandler -m main; ec=$?; if [ $ec -ne 0 ]; then echo; echo 'Launcher detected an error (exit' $ec '). Press Enter to close...'; read -r; fi"

exit /b 0

:use_gui_python_if_ready
if defined GUI_PYTHON exit /b 0
if "%~1"=="" exit /b 0
if not exist "%~1" exit /b 0
pushd "%APPDIR%"
"%~1" -c "import tkinter, serial, pandas, matplotlib, numpy, main" >nul 2>&1
set "PYTHON_CHECK_ERROR=%ERRORLEVEL%"
popd
if not "%PYTHON_CHECK_ERROR%"=="0" exit /b 0
set "GUI_PYTHON=%~1"
exit /b 0
