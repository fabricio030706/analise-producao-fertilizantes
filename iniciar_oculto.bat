@echo off

cd /d "C:\Users\pcm.arc_sudeste\Desktop\Automacao_de_Dados"

start "" pythonw runner.py > log.txt 2>&1

start "" "C:\Users\pcm.arc_sudeste\Desktop\Automacao_de_Dados\dashboard_paradas.html"

start "" "C:\Users\pcm.arc_sudeste\Desktop\Automacao_de_Dados\dashboard_fertipar_consolidado.html"
