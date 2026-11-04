@echo off
cd /d "%~dp0"
python atualizar_order_fill.py %*
echo.
pause
