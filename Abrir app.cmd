@echo off
rem Abre a app neste computador (le as pastas Compras/Vendas e grava os Excel em Despesas).
cd /d "%~dp0"
python servidor_local.py --abrir
pause
