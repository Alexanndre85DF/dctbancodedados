@echo off
chcp 65001 >nul
cd /d "%~dp0\.."
echo.
echo Abrindo a consulta neste computador...
start "" "http://127.0.0.1:8765/"
echo.
echo O endereço para a rede aparece abaixo (http://SEU-IP:8765/).
echo Este computador precisa ficar ligado, com esta janela aberta.
echo Se o Windows perguntar sobre firewall, escolha "Permitir".
echo.
python consulta\servidor.py
pause
