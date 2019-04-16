@echo off
set /p opc=(1)USB a PC o (2)PC a USB 
if %opc% == 1 goto imp
if %opc% == 2 goto exp

:imp
echo.
echo Importar
imp scott/tiger fromuser=scott touser=scott file=tablas.DMP
echo.
echo.
set /p repetir=(1)Repetir o (2)Salir 
if %repetir%==1 goto imp
exit

:exp
echo.
echo Exportar
exp scott/tiger file=tablas tables=tcursos,tusuarios,tinstructores,tturnos,tgrupos,trecordarusuario,tgruposculminados
echo.
echo.
set /p repetir=(1)Repetir o (2)Salir 
if %repetir%==1 goto exp
exit