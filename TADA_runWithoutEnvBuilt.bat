:: @ECHO OFF
:: %~dp0 : current dir
set DIR_PATH=%~dp0
set PYTHON_REL_PATH=TADApyEnv
set PYTHON_PATH=%DIR_PATH%%PYTHON_REL_PATH%

:: set env variables
set PATH=%PYTHON_PATH%;%PYTHON_PATH%\Scripts;%PYTHON_PATH%\DLLS;%PYTHON_PATH%\libs;

cd /d %~dp0
::.\TADApyEnv\Scripts\activate
cd src
cd code
streamlit run TADA_interface.py
cd ..
pause