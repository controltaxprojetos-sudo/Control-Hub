@echo off
chcp 65001 >nul
cd /d "%~dp0"
where py >nul 2>nul && (set PY=py) || (set PY=python)
echo Instalando componentes (so na primeira vez)...
%PY% -m pip install --quiet pymupdf openpyxl rapidocr-onnxruntime pillow numpy tzdata
%PY% robo_inventario_imobilizado.py
pause
