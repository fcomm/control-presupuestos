@echo off
cd /d "%~dp0"
python actualizar_listas_sat.py >> "%USERPROFILE%\.smi\listas_sat.log" 2>&1
python actualizar_tipo_cambio.py >> "%USERPROFILE%\.smi\tipo_cambio.log" 2>&1