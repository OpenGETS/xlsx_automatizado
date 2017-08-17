<?php
/** Error reporting */
error_reporting(E_ALL);
ini_set('display_errors', TRUE);
ini_set('display_startup_errors', TRUE);
date_default_timezone_set('America/Santiago');
header('Content-type: application/vnd.ms-excel');//esta es la principal
header("Content-Disposition: attachment; filename=archivo.xls");
header("Pragma: no-cache");
header("Expires: 0");


require 'PHPExcel.php';
define('EOL',(PHP_SAPI == 'cli') ? PHP_EOL : '<br />');


$log = "log.txt";
$time = time();
$fecha = date("d-m-Y H:i:s", $time);
file_put_contents($log,"Hora de Inicio ".$fecha);

$connect = mysql_connect('localhost','auditordb','Auditor_2016');
mysql_select_db('auditordb');



if(!$connect){
    echo 'no me pude conectar';
}


$diaActual = date('Ymd');
$inicioMes = date('Ym').'01';

$fechas = compact('diaActual','inicioMes');

$query = "CALL sp_lis_avance_semilla('')";

//$sql = $conexion->query($query);

//$resultado = $conexion->query($query);
$resultado = mysql_query($query);

//$data = mysql_fetch_array($resultado);

while ($row = mysql_fetch_row($resultado)) {
            $clientes[] = $row;
            print_r($row);
}
mysql_close();

$ea = new PHPExcel();
$ea->getProperties()
    ->setCreator('OpenGETS')
    ->setTitle('Reporte del avance de semilla')
    ->setLastModifiedBy('Joaquín Macías Cáceres');
$ews = $ea->getSheet(0);
$ews->setTitle('Reporte de Avance Semilla');
$ews->setCellValue('A1', '# TBN');
$ews->setCellValue('B1', 'DESCRIPCION');
$ews->setCellValue('C1', 'LOCALIDAD');
$ews->setCellValue('D1', 'PROMEDIO_DIA');
$ews->setCellValue('E1', 'ROL_MINIMO');
$ews->setCellValue('F1', 'ROL_MAXIMO');
$ews->setCellValue('G1', 'MIN_FECHA_INGRESO_PJ');
$ews->setCellValue('H1', 'MAX_FECHA_INGRESO_PJ');
$ews->setCellValue('I1', 'ROL_MINIMO_SEMILLA_INV');
$ews->setCellValue('J1', 'MIN_FECHA_INGRESO_PJ_INV');
$ews->setCellValue('K1', 'FECHA_CONSULTA');
$ews->setCellValue('L1', 'ROLES_INGRESADOS');
$ews->setCellValue('M1', 'ROLES_INGRESADOS_SEM_INV');
$ews->setCellValue('N1', 'ROLES_EN_PROCESO');
$ews->setCellValue('O1', 'ROLES_AUDITADOS');

$ews->fromArray($clientes, ' ', 'A2');

echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($ea, 'Excel2007');
//$objWriter->setPreCalculateFormulas(true);

//$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
$objWriter->save('avance_semilla_x_'.date('Y-m-d').'.xlsx', __FILE__);

$callEndTime = microtime(true);

$callTime = $callEndTime - $callStartTime;

$log = "log.txt";
$time = time();
$fecha = date("d-m-Y H:i:s", $time);
file_put_contents($log,"Hora de Termino ".$fecha);
