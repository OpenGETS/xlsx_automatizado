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

$connect = mysqli_connect("138.121.170.143", "root", "", "normalizacion_beco");

if(!$connect){
    echo 'no me pude conectar';
}


$diaActual = date('Ymd');
$inicioMes = date('Ym').'01';

$fechas = compact('diaActual','inicioMes');

// REPORTE A NORMALIZAR - 20170301
$query = "CALL FiltroTablaCapaTotal99_xlsx('INFOEMX_SAL_DIR_20170301.TXT');";
mysqli_set_charset($connect, 'utf8');

$resultado = mysqli_query($connect,$query);
//$resultado->set_charset('utf8');

while ($row = mysqli_fetch_row($resultado)) {
            $clientes[] = $row;
            //print_r($row);
}
mysqli_close($connect);

$ea = new PHPExcel();
$ea->getProperties()
    ->setCreator('OpenGETS')
    ->setTitle('Reporte - A normalizar - 20170301')
    ->setLastModifiedBy('Joaquín Macías Cáceres');
$ews = $ea->getSheet(0);
$ews->setTitle('Reporte_a_normalizar_20170301');
$ews->setCellValue('A1', 'RUT CLIENTE');
$ews->setCellValue('B1', 'NOMBRE_EJEC');
$ews->setCellValue('C1', 'ZONA_EJEC');
$ews->setCellValue('D1', 'CENTRO_EJEC');
$ews->setCellValue('E1', 'DIR');
$ews->setCellValue('F1', 'CALLE');
$ews->setCellValue('G1', 'NUMERO');
$ews->setCellValue('H1', 'COD_COMUNA');
$ews->setCellValue('I1', 'DESC_COMUNA');
$ews->setCellValue('J1', 'COD_CIUDAD');
$ews->setCellValue('K1', 'COD_REGION');
$ews->setCellValue('L1', 'DESC_REGION');
$ews->setCellValue('M1', 'LATITUD');
$ews->setCellValue('N1', 'LONGITUD');
$ews->setCellValue('O1', 'SECTOR');
$ews->setCellValue('P1', 'OBS');

$style = array(
        'alignment' => array(
            'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
        )
    );

$ews->getDefaultStyle()->applyFromArray($style);

$ews->fromArray($clientes, ' ', 'A2');

echo date('H:i:s') , " Write to Excel2007 format" , EOL;
$callStartTime = microtime(true);

$objWriter = PHPExcel_IOFactory::createWriter($ea, 'Excel2007');
//$objWriter->setPreCalculateFormulas(true);

//$objWriter->save(str_replace('.php', '.xlsx', __FILE__));
//$objWriter->save('NORM_20170618.xlsx', __FILE__);
$objWriter->save('Reporte_a_normalizar_20170301_'.date('Y-m-d').'.xlsx', __FILE__);

$callEndTime = microtime(true);

$callTime = $callEndTime - $callStartTime;

$log = "log.txt";
$time = time();
$fecha = date("d-m-Y H:i:s", $time);
file_put_contents($log,"Hora de Termino ".$fecha);
