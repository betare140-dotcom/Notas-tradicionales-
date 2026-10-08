@echo off
cd /d "%~dp0"
where py >nul 2>nul
if errorlevel 1 goto usepython
py -3 -m venv .venv
goto setup
:usepython
python -m venv .venv
:setup
if not exist .venv\Scripts\python.exe goto error
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto error
.venv\Scripts\python.exe -m streamlit run app.py
pause
exit /b
:error
echo No se pudo iniciar. Instala Python 3.11 o superior y revisa tu conexion.
pause
