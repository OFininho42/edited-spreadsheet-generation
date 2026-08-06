@echo off
title Executor de Script Python

:: Navega até a pasta onde o script Python está salvo
cd /d "C:\PROJETOS\Planilha"

:: Executa o script usando o Python instalado
python "principalarchive.py"

:: Impede que o terminal feche imediatamente após terminar a execução
pause