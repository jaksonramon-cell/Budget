@echo off
rem Abre el Presupuestador CIREC. La primera vez instala lo necesario (requiere Python 3.11+).
cd /d "%~dp0"
if not exist .venv (
    echo Instalando por primera vez...
    py -m venv .venv || python -m venv .venv
    .venv\Scripts\python -m pip install --upgrade pip
    .venv\Scripts\python -m pip install -r requirements.txt
)
.venv\Scripts\python app.py
