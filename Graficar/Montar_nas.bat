@echo off
wsl.exe -- bash -lc "sudo mount -t cifs '//10.10.71.137/Gestion_Administrativa_SolucionTecnologica/ETL-Lenovo/L8 Entregables internos/Fase intermedia/Controles' /mnt/red -o username=oabaunza,uid=$(id -u),gid=$(id -g)"
pause