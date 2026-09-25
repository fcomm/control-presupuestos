@echo off
cd /d "%~dp0"
python actualizar_listas_sat.py >> "%USERPROFILE%\.smi\listas_sat.log" 2>&1
