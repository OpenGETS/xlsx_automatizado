#!/bin/sh
today=`date '+%Y-%m-%d'`
echo 'PASO 1';
filename="/var/www/html/semilla_automatizado/avance_semilla_x_$today.xlsx"
echo 'PASO 2';
chmod u+x $filename
echo 'PASO 3';
echo 'Reporte diario' | mail -a $filename -s jmacias@opengets.cl "Resumen avance semilla al $today"
echo 'PASO 4';
