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

$connect = mysqli_connect("138.121.170.143", "root", "", "normalizacion_beco_web");

if(!$connect){
    echo 'no me pude conectar';
}


$diaActual = date('Ymd');
$inicioMes = date('Ym').'01';

$fechas = compact('diaActual','inicioMes');

// REPORTE A NORMALIZAR - 20170301
$query = "CALL FiltroTablaCapaTotal99_xlsx('INFOEMX_SAL_DIR_20170706.TXT');";
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
    ->setTitle('Reporte - A normalizar - 20170706')
    ->setLastModifiedBy('Joaquín Macías Cáceres');
$ews = $ea->getSheet(0);
$ews->setTitle('Reporte_a_normalizar_20170706');
$ews->setCellValue('A1', 'RUT CLIENTE');
$ews->setCellValue('B1', 'ID_COR');
$ews->setCellValue('C1', 'ID_DIRECCION');
$ews->setCellValue('D1', 'NOMBRE_EJEC');
$ews->setCellValue('E1', 'ZONA_EJEC');
$ews->setCellValue('F1', 'CENTRO_EJEC');
$ews->setCellValue('G1', 'TIPO');
$ews->setCellValue('H1', 'DIR');
$ews->setCellValue('I1', 'CALLE');
$ews->setCellValue('J1', 'NUMERO');
$ews->setCellValue('K1', 'COD_COMUNA');
$ews->setCellValue('L1', 'DESC_COMUNA');
$ews->setCellValue('M1', 'COD_CIUDAD');
$ews->setCellValue('N1', 'COD_REGION');
$ews->setCellValue('O1', 'DESC_REGION');
$ews->setCellValue('P1', 'LATITUD');
$ews->setCellValue('Q1', 'LONGITUD');
$ews->setCellValue('R1', 'SECTOR');
$ews->setCellValue('S1', 'OBS');
$ews->setCellValue('T1', 'RESULTADO');

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
$objWriter->save('ANORMALIZAR20170706.xlsx', __FILE__);

$callEndTime = microtime(true);

$callTime = $callEndTime - $callStartTime;

$log = "log.txt";
$time = time();
$fecha = date("d-m-Y H:i:s", $time);
file_put_contents($log,"Hora de Termino ".$fecha);
