@echo off
cd /d "%~dp0"
where py >nul 2>nul && (py atualizar_order_fill.py %*) || (python atualizar_order_fill.py %*)
echo.
pause
