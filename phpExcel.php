<?php
require_once 'PHPExcel/Classes/PHPExcel.php';
$nombreArchivo = "prueba1";

$archivo = "modeloReporte.xlsx";
$inputFileType = PHPExcel_IOFactory::identify($archivo);
$objReader = PHPExcel_IOFactory::createReader($inputFileType);
$objPHPExcel = $objReader->load($archivo);

//cambiar por idPacienteHidden2
$idPaciente = $_POST['idPacienteHidden2'];
setlocale(LC_ALL, "es_ES");


$border_style_red = array('borders' => array('allborders' => array('style' =>
PHPExcel_Style_Border::BORDER_THICK, 'color' => array('argb' => 'FF0000'),)));

$border_style_green = array('borders' => array('allborders' => array('style' =>
PHPExcel_Style_Border::BORDER_THICK, 'color' => array('argb' => '008000'),)));

$border_style_blue = array('borders' => array('allborders' => array('style' =>
PHPExcel_Style_Border::BORDER_THICK, 'color' => array('argb' => '0000FF'),)));



$connection = mysqli_connect('localhost', 'root', '', 'dbgaceta');

if (!$connection) {
    exit("database connection error");
}


$dataPaciente = mysqli_query($connection, "SELECT * FROM paciente WHERE idPaciente=$idPaciente");
$dataIngreso = mysqli_query($connection, "SELECT *, DATEDIFF(CURDATE(), fechaHora) AS DE FROM ingreso WHERE idPaciente=$idPaciente AND estatusIngreso=1");
$dataGaceta = mysqli_query($connection, "SELECT * 
        FROM gaceta G 
        INNER JOIN paciente P ON G.idPaciente = P.idPaciente
        INNER JOIN codigo_dh DH ON P.idCodigoDh = DH.idCodigo_dh
        INNER JOIN municipio M ON P.idMunicipio = M.idMunicipio
        INNER JOIN ingreso I ON I.idPaciente = P.idPaciente
        INNER JOIN cama C ON G.idCama = C.idCama
        WHERE G.idPaciente = $idPaciente AND G.estatusGaceta = 1 AND I.estatusIngreso = 1");
$dataEnfermeria = mysqli_query($connection, "SELECT * FROM enfermeria WHERE idPaciente=$idPaciente");
$dataEnfermeriaB = mysqli_query($connection, "SELECT * FROM enfermeriab WHERE idPaciente=$idPaciente");

while ($itemPaciente = mysqli_fetch_array($dataPaciente)) {

    if ($itemPaciente['sexo'] == 'F') {
        $objPHPExcel->getActiveSheet()->setCellValue('N7', '(F) X');
    }
    if ($itemPaciente['sexo'] == 'M') {
        $objPHPExcel->getActiveSheet()->setCellValue('P7', '(M) X');
    }

    $nombreArchivo = $itemPaciente['nombrePac'];

    $edad = utf8_encode($itemPaciente['edad']);
    $objPHPExcel->getActiveSheet()->setCellValue('R7', 'EDAD:' . strtoupper($edad));

    $nombrePaciente = utf8_encode($itemPaciente['apellidoPatPac']) . ' ' . utf8_encode($itemPaciente['apellidoMatPac']) . ' ' . utf8_encode($itemPaciente['nombrePac']);
    $objPHPExcel->getActiveSheet()->setCellValue('A7', 'NOMBRE: ' . strtoupper($nombrePaciente));
    date_default_timezone_set('America/Mexico_City');
    setlocale(LC_TIME, "spanish");



    $hoy = date("d M Y");

    $objPHPExcel->getActiveSheet()->setCellValue('K5', 'FECHA: ' . strtoupper($hoy));
}


while ($itemIngreso = mysqli_fetch_array($dataIngreso)) {

    $objPHPExcel->getActiveSheet()->setCellValue('H8', 'DIAS DE ESTANCIA: ' . STRTOUPPER($itemIngreso['DE'] . ' DIAS'));

    $objPHPExcel->getActiveSheet()->setCellValue('A9', 'DIAGNOSTICO MEDICO: ' . STRTOUPPER($itemIngreso['dxIngreso']));

    $fecha = $itemIngreso['fechaHora'];
    $theDate    = new DateTime($fecha);
    $stringDate = $theDate->format(' d M Y');

    $objPHPExcel->getActiveSheet()->setCellValue('A8', 'FECHA DE INGRESO: ' . STRTOUPPER($stringDate));
}


while ($itemGaceta = mysqli_fetch_array($dataGaceta)) {

    $objPHPExcel->getActiveSheet()->setCellValue('U5', '# DE CAMA_____ ' . STRTOUPPER($itemGaceta['numCama'] . '____'));
}


while ($itemEnfermeria = mysqli_fetch_array($dataEnfermeria)) {



    $prueba1 =  $itemEnfermeria['dataPaciente'];
    $prueba2 =  $itemEnfermeria['dataPacienteSistemica'];
    $prueba3 =  $itemEnfermeria['dataPacienteIngresos'];
    $prueba4 =  $itemEnfermeria['dataPacienteEgresos'];
    $prueba5 =  $itemEnfermeria['dataPacienteMedicamentos'];
    $prueba6 =  $itemEnfermeria['dataPacienteCateteres'];
    $prueba7 =  $itemEnfermeria['dataPacienteRespiratorios'];
    $prueba8 =  $itemEnfermeria['dataPacienteCultivos'];



    // Returns an object (The top level item in the JSON string is a JSON dictionary)
    // $json_string = '{"name": "Jeff", "age": 20, "active": true, "colors": ["red", "blue"]}';

    $object = json_decode($prueba1);
    $object2 = json_decode($prueba2);
    $object3 = json_decode($prueba3);
    $object4 = json_decode($prueba4);
    $object5 = json_decode($prueba5);
    $object6 = json_decode($prueba6);
    $object7 = json_decode($prueba7);
    $object8 = json_decode($prueba8);

    $objPHPExcel->getActiveSheet()->setCellValue('A10', 'PESO: ' . strtoupper($object->peso));
    $objPHPExcel->getActiveSheet()->setCellValue('A11', 'DIETA: ' . strtoupper($object->dieta));
    $objPHPExcel->getActiveSheet()->setCellValue('H10', 'TALLA: ' . strtoupper($object->talla));
    $objPHPExcel->getActiveSheet()->setCellValue('N10', 'SUPERF. CORPORAL: ' . strtoupper($object->superficieCorporal));
    $objPHPExcel->getActiveSheet()->setCellValue('X10', 'GPO. Y RH SANGUINEO: ' . strtoupper($object->grupoSanguineo));
    $objPHPExcel->getActiveSheet()->setCellValue('P8', 'CIRUGIA REALIZADA: ' . STRTOUPPER(utf8_encode($object->cirugiaRealizada)));

    $objPHPExcel->getActiveSheet()->setCellValue('H13', STRTOUPPER(utf8_encode($object->conciencia->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I13', STRTOUPPER(utf8_encode($object->conciencia->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J13', STRTOUPPER(utf8_encode($object->conciencia->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K13', STRTOUPPER(utf8_encode($object->conciencia->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L13', STRTOUPPER(utf8_encode($object->conciencia->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M13', STRTOUPPER(utf8_encode($object->conciencia->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N13', STRTOUPPER(utf8_encode($object->conciencia->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O13', STRTOUPPER(utf8_encode($object->conciencia->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P13', STRTOUPPER(utf8_encode($object->conciencia->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q13', STRTOUPPER(utf8_encode($object->conciencia->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R13', STRTOUPPER(utf8_encode($object->conciencia->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S13', STRTOUPPER(utf8_encode($object->conciencia->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T13', STRTOUPPER(utf8_encode($object->conciencia->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U13', STRTOUPPER(utf8_encode($object->conciencia->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V13', STRTOUPPER(utf8_encode($object->conciencia->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W13', STRTOUPPER(utf8_encode($object->conciencia->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X13', STRTOUPPER(utf8_encode($object->conciencia->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y13', STRTOUPPER(utf8_encode($object->conciencia->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z13', STRTOUPPER(utf8_encode($object->conciencia->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA13', STRTOUPPER(utf8_encode($object->conciencia->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB13', STRTOUPPER(utf8_encode($object->conciencia->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC13', STRTOUPPER(utf8_encode($object->conciencia->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD13', STRTOUPPER(utf8_encode($object->conciencia->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE13', STRTOUPPER(utf8_encode($object->conciencia->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H14', STRTOUPPER(utf8_encode($object->piel->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I14', STRTOUPPER(utf8_encode($object->piel->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J14', STRTOUPPER(utf8_encode($object->piel->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K14', STRTOUPPER(utf8_encode($object->piel->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L14', STRTOUPPER(utf8_encode($object->piel->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M14', STRTOUPPER(utf8_encode($object->piel->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N14', STRTOUPPER(utf8_encode($object->piel->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O14', STRTOUPPER(utf8_encode($object->piel->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P14', STRTOUPPER(utf8_encode($object->piel->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q14', STRTOUPPER(utf8_encode($object->piel->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R14', STRTOUPPER(utf8_encode($object->piel->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S14', STRTOUPPER(utf8_encode($object->piel->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T14', STRTOUPPER(utf8_encode($object->piel->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U14', STRTOUPPER(utf8_encode($object->piel->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V14', STRTOUPPER(utf8_encode($object->piel->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W14', STRTOUPPER(utf8_encode($object->piel->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X14', STRTOUPPER(utf8_encode($object->piel->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y14', STRTOUPPER(utf8_encode($object->piel->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z14', STRTOUPPER(utf8_encode($object->piel->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA14', STRTOUPPER(utf8_encode($object->piel->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB14', STRTOUPPER(utf8_encode($object->piel->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC14', STRTOUPPER(utf8_encode($object->piel->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD14', STRTOUPPER(utf8_encode($object->piel->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE14', STRTOUPPER(utf8_encode($object->piel->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H15', STRTOUPPER(utf8_encode($object->palidezIctericia->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I15', STRTOUPPER(utf8_encode($object->palidezIctericia->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J15', STRTOUPPER(utf8_encode($object->palidezIctericia->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K15', STRTOUPPER(utf8_encode($object->palidezIctericia->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L15', STRTOUPPER(utf8_encode($object->palidezIctericia->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M15', STRTOUPPER(utf8_encode($object->palidezIctericia->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N15', STRTOUPPER(utf8_encode($object->palidezIctericia->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O15', STRTOUPPER(utf8_encode($object->palidezIctericia->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P15', STRTOUPPER(utf8_encode($object->palidezIctericia->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q15', STRTOUPPER(utf8_encode($object->palidezIctericia->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R15', STRTOUPPER(utf8_encode($object->palidezIctericia->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S15', STRTOUPPER(utf8_encode($object->palidezIctericia->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T15', STRTOUPPER(utf8_encode($object->palidezIctericia->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U15', STRTOUPPER(utf8_encode($object->palidezIctericia->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V15', STRTOUPPER(utf8_encode($object->palidezIctericia->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W15', STRTOUPPER(utf8_encode($object->palidezIctericia->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X15', STRTOUPPER(utf8_encode($object->palidezIctericia->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y15', STRTOUPPER(utf8_encode($object->palidezIctericia->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z15', STRTOUPPER(utf8_encode($object->palidezIctericia->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA15', STRTOUPPER(utf8_encode($object->palidezIctericia->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB15', STRTOUPPER(utf8_encode($object->palidezIctericia->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC15', STRTOUPPER(utf8_encode($object->palidezIctericia->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD15', STRTOUPPER(utf8_encode($object->palidezIctericia->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE15', STRTOUPPER(utf8_encode($object->palidezIctericia->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H16', STRTOUPPER(utf8_encode($object->cianosis->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I16', STRTOUPPER(utf8_encode($object->cianosis->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J16', STRTOUPPER(utf8_encode($object->cianosis->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K16', STRTOUPPER(utf8_encode($object->cianosis->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L16', STRTOUPPER(utf8_encode($object->cianosis->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M16', STRTOUPPER(utf8_encode($object->cianosis->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N16', STRTOUPPER(utf8_encode($object->cianosis->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O16', STRTOUPPER(utf8_encode($object->cianosis->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P16', STRTOUPPER(utf8_encode($object->cianosis->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q16', STRTOUPPER(utf8_encode($object->cianosis->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R16', STRTOUPPER(utf8_encode($object->cianosis->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S16', STRTOUPPER(utf8_encode($object->cianosis->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T16', STRTOUPPER(utf8_encode($object->cianosis->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U16', STRTOUPPER(utf8_encode($object->cianosis->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V16', STRTOUPPER(utf8_encode($object->cianosis->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W16', STRTOUPPER(utf8_encode($object->cianosis->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X16', STRTOUPPER(utf8_encode($object->cianosis->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y16', STRTOUPPER(utf8_encode($object->cianosis->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z16', STRTOUPPER(utf8_encode($object->cianosis->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA16', STRTOUPPER(utf8_encode($object->cianosis->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB16', STRTOUPPER(utf8_encode($object->cianosis->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC16', STRTOUPPER(utf8_encode($object->cianosis->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD16', STRTOUPPER(utf8_encode($object->cianosis->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE16', STRTOUPPER(utf8_encode($object->cianosis->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE17', STRTOUPPER(utf8_encode($object->llenadoCapilar->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H18', STRTOUPPER(utf8_encode($object->dolor->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I18', STRTOUPPER(utf8_encode($object->dolor->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J18', STRTOUPPER(utf8_encode($object->dolor->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K18', STRTOUPPER(utf8_encode($object->dolor->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L18', STRTOUPPER(utf8_encode($object->dolor->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M18', STRTOUPPER(utf8_encode($object->dolor->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N18', STRTOUPPER(utf8_encode($object->dolor->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O18', STRTOUPPER(utf8_encode($object->dolor->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P18', STRTOUPPER(utf8_encode($object->dolor->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q18', STRTOUPPER(utf8_encode($object->dolor->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R18', STRTOUPPER(utf8_encode($object->dolor->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S18', STRTOUPPER(utf8_encode($object->dolor->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T18', STRTOUPPER(utf8_encode($object->dolor->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U18', STRTOUPPER(utf8_encode($object->dolor->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V18', STRTOUPPER(utf8_encode($object->dolor->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W18', STRTOUPPER(utf8_encode($object->dolor->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X18', STRTOUPPER(utf8_encode($object->dolor->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y18', STRTOUPPER(utf8_encode($object->dolor->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z18', STRTOUPPER(utf8_encode($object->dolor->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA18', STRTOUPPER(utf8_encode($object->dolor->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB18', STRTOUPPER(utf8_encode($object->dolor->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC18', STRTOUPPER(utf8_encode($object->dolor->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD18', STRTOUPPER(utf8_encode($object->dolor->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE18', STRTOUPPER(utf8_encode($object->dolor->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE19', STRTOUPPER(utf8_encode($object->temperaturaAxilar->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE20', STRTOUPPER(utf8_encode($object->frecuenciaCardiaca->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE21', STRTOUPPER(utf8_encode($object->ritmoCardiaco->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE22', STRTOUPPER(utf8_encode($object->frecuenciaRespiratoria->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE24', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H25', STRTOUPPER(utf8_encode($object2->spo->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I25', STRTOUPPER(utf8_encode($object2->spo->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J25', STRTOUPPER(utf8_encode($object2->spo->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K25', STRTOUPPER(utf8_encode($object2->spo->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L25', STRTOUPPER(utf8_encode($object2->spo->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M25', STRTOUPPER(utf8_encode($object2->spo->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N25', STRTOUPPER(utf8_encode($object2->spo->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O25', STRTOUPPER(utf8_encode($object2->spo->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P25', STRTOUPPER(utf8_encode($object2->spo->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q25', STRTOUPPER(utf8_encode($object2->spo->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R25', STRTOUPPER(utf8_encode($object2->spo->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S25', STRTOUPPER(utf8_encode($object2->spo->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T25', STRTOUPPER(utf8_encode($object2->spo->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U25', STRTOUPPER(utf8_encode($object2->spo->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V25', STRTOUPPER(utf8_encode($object2->spo->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W25', STRTOUPPER(utf8_encode($object2->spo->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X25', STRTOUPPER(utf8_encode($object2->spo->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y25', STRTOUPPER(utf8_encode($object2->spo->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z25', STRTOUPPER(utf8_encode($object2->spo->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA25', STRTOUPPER(utf8_encode($object2->spo->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB25', STRTOUPPER(utf8_encode($object2->spo->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC25', STRTOUPPER(utf8_encode($object2->spo->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD25', STRTOUPPER(utf8_encode($object2->spo->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE25', STRTOUPPER(utf8_encode($object2->spo->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE26', STRTOUPPER(utf8_encode($object2->presionVenosaCentral->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE27', STRTOUPPER(utf8_encode($object2->presionArterialSistolica->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE28', STRTOUPPER(utf8_encode($object2->presionArterialDiastolica->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE29', STRTOUPPER(utf8_encode($object2->presionArterialMedia->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE30', STRTOUPPER(utf8_encode($object2->presionArterialPulmonarMedia->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE31', STRTOUPPER(utf8_encode($object2->presionCapilarPulmonar->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE32', STRTOUPPER(utf8_encode($object2->presionIntracraneal->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE33', STRTOUPPER(utf8_encode($object2->glucemiaCapilar->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('H34', STRTOUPPER(utf8_encode($object2->insulina->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('I34', STRTOUPPER(utf8_encode($object2->insulina->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('J34', STRTOUPPER(utf8_encode($object2->insulina->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('K34', STRTOUPPER(utf8_encode($object2->insulina->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('L34', STRTOUPPER(utf8_encode($object2->insulina->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('M34', STRTOUPPER(utf8_encode($object2->insulina->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('N34', STRTOUPPER(utf8_encode($object2->insulina->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('O34', STRTOUPPER(utf8_encode($object2->insulina->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('P34', STRTOUPPER(utf8_encode($object2->insulina->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('Q34', STRTOUPPER(utf8_encode($object2->insulina->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('R34', STRTOUPPER(utf8_encode($object2->insulina->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('S34', STRTOUPPER(utf8_encode($object2->insulina->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('T34', STRTOUPPER(utf8_encode($object2->insulina->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('U34', STRTOUPPER(utf8_encode($object2->insulina->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('V34', STRTOUPPER(utf8_encode($object2->insulina->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('W34', STRTOUPPER(utf8_encode($object2->insulina->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('X34', STRTOUPPER(utf8_encode($object2->insulina->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y34', STRTOUPPER(utf8_encode($object2->insulina->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z34', STRTOUPPER(utf8_encode($object2->insulina->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA34', STRTOUPPER(utf8_encode($object2->insulina->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AB34', STRTOUPPER(utf8_encode($object2->insulina->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AC34', STRTOUPPER(utf8_encode($object2->insulina->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD34', STRTOUPPER(utf8_encode($object2->insulina->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE34', STRTOUPPER(utf8_encode($object2->insulina->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D39', STRTOUPPER(utf8_encode($object3->ingreso1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S39', STRTOUPPER(utf8_encode($object3->ingreso1->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U39', STRTOUPPER(utf8_encode($object3->ingreso1->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V39', STRTOUPPER(utf8_encode($object3->ingreso1->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W39', STRTOUPPER(utf8_encode($object3->ingreso1->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X39', STRTOUPPER(utf8_encode($object3->ingreso1->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y39', STRTOUPPER(utf8_encode($object3->ingreso1->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z39', STRTOUPPER(utf8_encode($object3->ingreso1->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA39', STRTOUPPER(utf8_encode($object3->ingreso1->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD39', STRTOUPPER(utf8_encode($object3->ingreso1->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE39', STRTOUPPER(utf8_encode($object3->ingreso1->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF39', STRTOUPPER(utf8_encode($object3->ingreso1->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG39', STRTOUPPER(utf8_encode($object3->ingreso1->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH39', STRTOUPPER(utf8_encode($object3->ingreso1->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI39', STRTOUPPER(utf8_encode($object3->ingreso1->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL39', STRTOUPPER(utf8_encode($object3->ingreso1->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM39', STRTOUPPER(utf8_encode($object3->ingreso1->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN39', STRTOUPPER(utf8_encode($object3->ingreso1->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO39', STRTOUPPER(utf8_encode($object3->ingreso1->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP39', STRTOUPPER(utf8_encode($object3->ingreso1->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ39', STRTOUPPER(utf8_encode($object3->ingreso1->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR39', STRTOUPPER(utf8_encode($object3->ingreso1->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS39', STRTOUPPER(utf8_encode($object3->ingreso1->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT39', STRTOUPPER(utf8_encode($object3->ingreso1->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU39', STRTOUPPER(utf8_encode($object3->ingreso1->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV39', STRTOUPPER(utf8_encode($object3->ingreso1->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D40', STRTOUPPER(utf8_encode($object3->ingreso2->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S40', STRTOUPPER(utf8_encode($object3->ingreso2->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U40', STRTOUPPER(utf8_encode($object3->ingreso2->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V40', STRTOUPPER(utf8_encode($object3->ingreso2->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W40', STRTOUPPER(utf8_encode($object3->ingreso2->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X40', STRTOUPPER(utf8_encode($object3->ingreso2->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y40', STRTOUPPER(utf8_encode($object3->ingreso2->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z40', STRTOUPPER(utf8_encode($object3->ingreso2->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA40', STRTOUPPER(utf8_encode($object3->ingreso2->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD40', STRTOUPPER(utf8_encode($object3->ingreso2->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE40', STRTOUPPER(utf8_encode($object3->ingreso2->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF40', STRTOUPPER(utf8_encode($object3->ingreso2->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG40', STRTOUPPER(utf8_encode($object3->ingreso2->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH40', STRTOUPPER(utf8_encode($object3->ingreso2->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI40', STRTOUPPER(utf8_encode($object3->ingreso2->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL40', STRTOUPPER(utf8_encode($object3->ingreso2->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM40', STRTOUPPER(utf8_encode($object3->ingreso2->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN40', STRTOUPPER(utf8_encode($object3->ingreso2->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO40', STRTOUPPER(utf8_encode($object3->ingreso2->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP40', STRTOUPPER(utf8_encode($object3->ingreso2->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ40', STRTOUPPER(utf8_encode($object3->ingreso2->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR40', STRTOUPPER(utf8_encode($object3->ingreso2->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS40', STRTOUPPER(utf8_encode($object3->ingreso2->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT40', STRTOUPPER(utf8_encode($object3->ingreso2->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU40', STRTOUPPER(utf8_encode($object3->ingreso2->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV40', STRTOUPPER(utf8_encode($object3->ingreso2->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D41', STRTOUPPER(utf8_encode($object3->ingreso3->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S41', STRTOUPPER(utf8_encode($object3->ingreso3->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U41', STRTOUPPER(utf8_encode($object3->ingreso3->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V41', STRTOUPPER(utf8_encode($object3->ingreso3->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W41', STRTOUPPER(utf8_encode($object3->ingreso3->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X41', STRTOUPPER(utf8_encode($object3->ingreso3->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y41', STRTOUPPER(utf8_encode($object3->ingreso3->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z41', STRTOUPPER(utf8_encode($object3->ingreso3->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA41', STRTOUPPER(utf8_encode($object3->ingreso3->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD41', STRTOUPPER(utf8_encode($object3->ingreso3->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE41', STRTOUPPER(utf8_encode($object3->ingreso3->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF41', STRTOUPPER(utf8_encode($object3->ingreso3->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG41', STRTOUPPER(utf8_encode($object3->ingreso3->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH41', STRTOUPPER(utf8_encode($object3->ingreso3->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI41', STRTOUPPER(utf8_encode($object3->ingreso3->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL41', STRTOUPPER(utf8_encode($object3->ingreso3->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM41', STRTOUPPER(utf8_encode($object3->ingreso3->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN41', STRTOUPPER(utf8_encode($object3->ingreso3->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO41', STRTOUPPER(utf8_encode($object3->ingreso3->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP41', STRTOUPPER(utf8_encode($object3->ingreso3->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ41', STRTOUPPER(utf8_encode($object3->ingreso3->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR41', STRTOUPPER(utf8_encode($object3->ingreso3->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS41', STRTOUPPER(utf8_encode($object3->ingreso3->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT41', STRTOUPPER(utf8_encode($object3->ingreso3->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU41', STRTOUPPER(utf8_encode($object3->ingreso3->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV41', STRTOUPPER(utf8_encode($object3->ingreso3->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D42', STRTOUPPER(utf8_encode($object3->ingreso4->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S42', STRTOUPPER(utf8_encode($object3->ingreso4->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U42', STRTOUPPER(utf8_encode($object3->ingreso4->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V42', STRTOUPPER(utf8_encode($object3->ingreso4->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W42', STRTOUPPER(utf8_encode($object3->ingreso4->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X42', STRTOUPPER(utf8_encode($object3->ingreso4->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y42', STRTOUPPER(utf8_encode($object3->ingreso4->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z42', STRTOUPPER(utf8_encode($object3->ingreso4->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA42', STRTOUPPER(utf8_encode($object3->ingreso4->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD42', STRTOUPPER(utf8_encode($object3->ingreso4->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE42', STRTOUPPER(utf8_encode($object3->ingreso4->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF42', STRTOUPPER(utf8_encode($object3->ingreso4->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG42', STRTOUPPER(utf8_encode($object3->ingreso4->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH42', STRTOUPPER(utf8_encode($object3->ingreso4->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI42', STRTOUPPER(utf8_encode($object3->ingreso4->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL42', STRTOUPPER(utf8_encode($object3->ingreso4->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM42', STRTOUPPER(utf8_encode($object3->ingreso4->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN42', STRTOUPPER(utf8_encode($object3->ingreso4->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO42', STRTOUPPER(utf8_encode($object3->ingreso4->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP42', STRTOUPPER(utf8_encode($object3->ingreso4->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ42', STRTOUPPER(utf8_encode($object3->ingreso4->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR42', STRTOUPPER(utf8_encode($object3->ingreso4->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS42', STRTOUPPER(utf8_encode($object3->ingreso4->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT42', STRTOUPPER(utf8_encode($object3->ingreso4->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU42', STRTOUPPER(utf8_encode($object3->ingreso4->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV42', STRTOUPPER(utf8_encode($object3->ingreso4->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D43', STRTOUPPER(utf8_encode($object3->ingreso5->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S43', STRTOUPPER(utf8_encode($object3->ingreso5->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U43', STRTOUPPER(utf8_encode($object3->ingreso5->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V43', STRTOUPPER(utf8_encode($object3->ingreso5->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W43', STRTOUPPER(utf8_encode($object3->ingreso5->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X43', STRTOUPPER(utf8_encode($object3->ingreso5->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y43', STRTOUPPER(utf8_encode($object3->ingreso5->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z43', STRTOUPPER(utf8_encode($object3->ingreso5->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA43', STRTOUPPER(utf8_encode($object3->ingreso5->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD43', STRTOUPPER(utf8_encode($object3->ingreso5->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE43', STRTOUPPER(utf8_encode($object3->ingreso5->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF43', STRTOUPPER(utf8_encode($object3->ingreso5->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG43', STRTOUPPER(utf8_encode($object3->ingreso5->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH43', STRTOUPPER(utf8_encode($object3->ingreso5->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI43', STRTOUPPER(utf8_encode($object3->ingreso5->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL43', STRTOUPPER(utf8_encode($object3->ingreso5->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM43', STRTOUPPER(utf8_encode($object3->ingreso5->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN43', STRTOUPPER(utf8_encode($object3->ingreso5->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO43', STRTOUPPER(utf8_encode($object3->ingreso5->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP43', STRTOUPPER(utf8_encode($object3->ingreso5->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ43', STRTOUPPER(utf8_encode($object3->ingreso5->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR43', STRTOUPPER(utf8_encode($object3->ingreso5->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS43', STRTOUPPER(utf8_encode($object3->ingreso5->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT43', STRTOUPPER(utf8_encode($object3->ingreso5->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU43', STRTOUPPER(utf8_encode($object3->ingreso5->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV43', STRTOUPPER(utf8_encode($object3->ingreso5->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D344', STRTOUPPER(utf8_encode($object3->ingreso6->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S344', STRTOUPPER(utf8_encode($object3->ingreso6->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U344', STRTOUPPER(utf8_encode($object3->ingreso6->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V344', STRTOUPPER(utf8_encode($object3->ingreso6->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W344', STRTOUPPER(utf8_encode($object3->ingreso6->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X344', STRTOUPPER(utf8_encode($object3->ingreso6->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y344', STRTOUPPER(utf8_encode($object3->ingreso6->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z344', STRTOUPPER(utf8_encode($object3->ingreso6->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA44', STRTOUPPER(utf8_encode($object3->ingreso6->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD44', STRTOUPPER(utf8_encode($object3->ingreso6->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE44', STRTOUPPER(utf8_encode($object3->ingreso6->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF44', STRTOUPPER(utf8_encode($object3->ingreso6->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG44', STRTOUPPER(utf8_encode($object3->ingreso6->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH44', STRTOUPPER(utf8_encode($object3->ingreso6->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI44', STRTOUPPER(utf8_encode($object3->ingreso6->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL44', STRTOUPPER(utf8_encode($object3->ingreso6->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM44', STRTOUPPER(utf8_encode($object3->ingreso6->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN44', STRTOUPPER(utf8_encode($object3->ingreso6->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO44', STRTOUPPER(utf8_encode($object3->ingreso6->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP44', STRTOUPPER(utf8_encode($object3->ingreso6->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ44', STRTOUPPER(utf8_encode($object3->ingreso6->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR44', STRTOUPPER(utf8_encode($object3->ingreso6->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS44', STRTOUPPER(utf8_encode($object3->ingreso6->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT44', STRTOUPPER(utf8_encode($object3->ingreso6->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU44', STRTOUPPER(utf8_encode($object3->ingreso6->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV44', STRTOUPPER(utf8_encode($object3->ingreso6->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D45', STRTOUPPER(utf8_encode($object3->ingreso7->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S45', STRTOUPPER(utf8_encode($object3->ingreso7->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U45', STRTOUPPER(utf8_encode($object3->ingreso7->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V45', STRTOUPPER(utf8_encode($object3->ingreso7->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W45', STRTOUPPER(utf8_encode($object3->ingreso7->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X45', STRTOUPPER(utf8_encode($object3->ingreso7->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y45', STRTOUPPER(utf8_encode($object3->ingreso7->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z45', STRTOUPPER(utf8_encode($object3->ingreso7->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA45', STRTOUPPER(utf8_encode($object3->ingreso7->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD45', STRTOUPPER(utf8_encode($object3->ingreso7->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE45', STRTOUPPER(utf8_encode($object3->ingreso7->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF45', STRTOUPPER(utf8_encode($object3->ingreso7->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG45', STRTOUPPER(utf8_encode($object3->ingreso7->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH45', STRTOUPPER(utf8_encode($object3->ingreso7->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI45', STRTOUPPER(utf8_encode($object3->ingreso7->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL45', STRTOUPPER(utf8_encode($object3->ingreso7->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM45', STRTOUPPER(utf8_encode($object3->ingreso7->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN45', STRTOUPPER(utf8_encode($object3->ingreso7->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO45', STRTOUPPER(utf8_encode($object3->ingreso7->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP45', STRTOUPPER(utf8_encode($object3->ingreso7->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ45', STRTOUPPER(utf8_encode($object3->ingreso7->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR45', STRTOUPPER(utf8_encode($object3->ingreso7->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS45', STRTOUPPER(utf8_encode($object3->ingreso7->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT45', STRTOUPPER(utf8_encode($object3->ingreso7->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU45', STRTOUPPER(utf8_encode($object3->ingreso7->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV45', STRTOUPPER(utf8_encode($object3->ingreso7->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D46', STRTOUPPER(utf8_encode($object3->ingreso8->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S46', STRTOUPPER(utf8_encode($object3->ingreso8->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U46', STRTOUPPER(utf8_encode($object3->ingreso8->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V46', STRTOUPPER(utf8_encode($object3->ingreso8->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W46', STRTOUPPER(utf8_encode($object3->ingreso8->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X46', STRTOUPPER(utf8_encode($object3->ingreso8->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y46', STRTOUPPER(utf8_encode($object3->ingreso8->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z46', STRTOUPPER(utf8_encode($object3->ingreso8->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA46', STRTOUPPER(utf8_encode($object3->ingreso8->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD46', STRTOUPPER(utf8_encode($object3->ingreso8->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE46', STRTOUPPER(utf8_encode($object3->ingreso8->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF46', STRTOUPPER(utf8_encode($object3->ingreso8->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG46', STRTOUPPER(utf8_encode($object3->ingreso8->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH46', STRTOUPPER(utf8_encode($object3->ingreso8->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI46', STRTOUPPER(utf8_encode($object3->ingreso8->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL46', STRTOUPPER(utf8_encode($object3->ingreso8->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM46', STRTOUPPER(utf8_encode($object3->ingreso8->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN46', STRTOUPPER(utf8_encode($object3->ingreso8->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO46', STRTOUPPER(utf8_encode($object3->ingreso8->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP46', STRTOUPPER(utf8_encode($object3->ingreso8->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ46', STRTOUPPER(utf8_encode($object3->ingreso8->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR46', STRTOUPPER(utf8_encode($object3->ingreso8->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS46', STRTOUPPER(utf8_encode($object3->ingreso8->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT46', STRTOUPPER(utf8_encode($object3->ingreso8->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU46', STRTOUPPER(utf8_encode($object3->ingreso8->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV46', STRTOUPPER(utf8_encode($object3->ingreso8->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D47', STRTOUPPER(utf8_encode($object3->ingreso9->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S47', STRTOUPPER(utf8_encode($object3->ingreso9->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U47', STRTOUPPER(utf8_encode($object3->ingreso9->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V47', STRTOUPPER(utf8_encode($object3->ingreso9->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W47', STRTOUPPER(utf8_encode($object3->ingreso9->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X47', STRTOUPPER(utf8_encode($object3->ingreso9->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y47', STRTOUPPER(utf8_encode($object3->ingreso9->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z47', STRTOUPPER(utf8_encode($object3->ingreso9->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA47', STRTOUPPER(utf8_encode($object3->ingreso9->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD47', STRTOUPPER(utf8_encode($object3->ingreso9->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE47', STRTOUPPER(utf8_encode($object3->ingreso9->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF47', STRTOUPPER(utf8_encode($object3->ingreso9->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG47', STRTOUPPER(utf8_encode($object3->ingreso9->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH47', STRTOUPPER(utf8_encode($object3->ingreso9->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI47', STRTOUPPER(utf8_encode($object3->ingreso9->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL47', STRTOUPPER(utf8_encode($object3->ingreso9->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM47', STRTOUPPER(utf8_encode($object3->ingreso9->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN47', STRTOUPPER(utf8_encode($object3->ingreso9->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO47', STRTOUPPER(utf8_encode($object3->ingreso9->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP47', STRTOUPPER(utf8_encode($object3->ingreso9->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ47', STRTOUPPER(utf8_encode($object3->ingreso9->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR47', STRTOUPPER(utf8_encode($object3->ingreso9->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS47', STRTOUPPER(utf8_encode($object3->ingreso9->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT47', STRTOUPPER(utf8_encode($object3->ingreso9->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU47', STRTOUPPER(utf8_encode($object3->ingreso9->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV47', STRTOUPPER(utf8_encode($object3->ingreso9->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D48', STRTOUPPER(utf8_encode($object3->ingreso10->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S48', STRTOUPPER(utf8_encode($object3->ingreso10->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U48', STRTOUPPER(utf8_encode($object3->ingreso10->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V48', STRTOUPPER(utf8_encode($object3->ingreso10->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W48', STRTOUPPER(utf8_encode($object3->ingreso10->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X48', STRTOUPPER(utf8_encode($object3->ingreso10->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y48', STRTOUPPER(utf8_encode($object3->ingreso10->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z48', STRTOUPPER(utf8_encode($object3->ingreso10->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA48', STRTOUPPER(utf8_encode($object3->ingreso10->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD48', STRTOUPPER(utf8_encode($object3->ingreso10->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE48', STRTOUPPER(utf8_encode($object3->ingreso10->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF48', STRTOUPPER(utf8_encode($object3->ingreso10->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG48', STRTOUPPER(utf8_encode($object3->ingreso10->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH48', STRTOUPPER(utf8_encode($object3->ingreso10->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI48', STRTOUPPER(utf8_encode($object3->ingreso10->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL48', STRTOUPPER(utf8_encode($object3->ingreso10->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM48', STRTOUPPER(utf8_encode($object3->ingreso10->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN48', STRTOUPPER(utf8_encode($object3->ingreso10->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO48', STRTOUPPER(utf8_encode($object3->ingreso10->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP48', STRTOUPPER(utf8_encode($object3->ingreso10->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ48', STRTOUPPER(utf8_encode($object3->ingreso10->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR48', STRTOUPPER(utf8_encode($object3->ingreso10->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS48', STRTOUPPER(utf8_encode($object3->ingreso10->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT48', STRTOUPPER(utf8_encode($object3->ingreso10->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU48', STRTOUPPER(utf8_encode($object3->ingreso10->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV48', STRTOUPPER(utf8_encode($object3->ingreso10->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D49', STRTOUPPER(utf8_encode($object3->ingreso11->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S49', STRTOUPPER(utf8_encode($object3->ingreso11->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U49', STRTOUPPER(utf8_encode($object3->ingreso11->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V49', STRTOUPPER(utf8_encode($object3->ingreso11->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W49', STRTOUPPER(utf8_encode($object3->ingreso11->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X49', STRTOUPPER(utf8_encode($object3->ingreso11->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y49', STRTOUPPER(utf8_encode($object3->ingreso11->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z49', STRTOUPPER(utf8_encode($object3->ingreso11->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA49', STRTOUPPER(utf8_encode($object3->ingreso11->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD49', STRTOUPPER(utf8_encode($object3->ingreso11->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE49', STRTOUPPER(utf8_encode($object3->ingreso11->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF49', STRTOUPPER(utf8_encode($object3->ingreso11->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG49', STRTOUPPER(utf8_encode($object3->ingreso11->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH49', STRTOUPPER(utf8_encode($object3->ingreso11->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI49', STRTOUPPER(utf8_encode($object3->ingreso11->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL49', STRTOUPPER(utf8_encode($object3->ingreso11->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM49', STRTOUPPER(utf8_encode($object3->ingreso11->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN49', STRTOUPPER(utf8_encode($object3->ingreso11->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO49', STRTOUPPER(utf8_encode($object3->ingreso11->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP49', STRTOUPPER(utf8_encode($object3->ingreso11->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ49', STRTOUPPER(utf8_encode($object3->ingreso11->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR49', STRTOUPPER(utf8_encode($object3->ingreso11->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS49', STRTOUPPER(utf8_encode($object3->ingreso11->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT49', STRTOUPPER(utf8_encode($object3->ingreso11->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU49', STRTOUPPER(utf8_encode($object3->ingreso11->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV49', STRTOUPPER(utf8_encode($object3->ingreso11->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('S50', STRTOUPPER(utf8_encode($object3->ingreso12->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('D50', STRTOUPPER(utf8_encode($object3->ingreso12->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('U50', STRTOUPPER(utf8_encode($object3->ingreso12->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V50', STRTOUPPER(utf8_encode($object3->ingreso12->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W50', STRTOUPPER(utf8_encode($object3->ingreso12->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X50', STRTOUPPER(utf8_encode($object3->ingreso12->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y50', STRTOUPPER(utf8_encode($object3->ingreso12->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z50', STRTOUPPER(utf8_encode($object3->ingreso12->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA50', STRTOUPPER(utf8_encode($object3->ingreso12->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD50', STRTOUPPER(utf8_encode($object3->ingreso12->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE50', STRTOUPPER(utf8_encode($object3->ingreso12->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF50', STRTOUPPER(utf8_encode($object3->ingreso12->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG50', STRTOUPPER(utf8_encode($object3->ingreso12->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH50', STRTOUPPER(utf8_encode($object3->ingreso12->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI50', STRTOUPPER(utf8_encode($object3->ingreso12->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL50', STRTOUPPER(utf8_encode($object3->ingreso12->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM50', STRTOUPPER(utf8_encode($object3->ingreso12->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN50', STRTOUPPER(utf8_encode($object3->ingreso12->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO50', STRTOUPPER(utf8_encode($object3->ingreso12->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP50', STRTOUPPER(utf8_encode($object3->ingreso12->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ50', STRTOUPPER(utf8_encode($object3->ingreso12->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR50', STRTOUPPER(utf8_encode($object3->ingreso12->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS50', STRTOUPPER(utf8_encode($object3->ingreso12->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT50', STRTOUPPER(utf8_encode($object3->ingreso12->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU50', STRTOUPPER(utf8_encode($object3->ingreso12->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV50', STRTOUPPER(utf8_encode($object3->ingreso12->{7})));


    $objPHPExcel->getActiveSheet()->setCellValue('D51', STRTOUPPER(utf8_encode($object3->ingreso13->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S51', STRTOUPPER(utf8_encode($object3->ingreso13->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U51', STRTOUPPER(utf8_encode($object3->ingreso13->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V51', STRTOUPPER(utf8_encode($object3->ingreso13->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W51', STRTOUPPER(utf8_encode($object3->ingreso13->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X51', STRTOUPPER(utf8_encode($object3->ingreso13->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y51', STRTOUPPER(utf8_encode($object3->ingreso13->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z51', STRTOUPPER(utf8_encode($object3->ingreso13->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA51', STRTOUPPER(utf8_encode($object3->ingreso13->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD51', STRTOUPPER(utf8_encode($object3->ingreso13->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE51', STRTOUPPER(utf8_encode($object3->ingreso13->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF51', STRTOUPPER(utf8_encode($object3->ingreso13->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG51', STRTOUPPER(utf8_encode($object3->ingreso13->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH51', STRTOUPPER(utf8_encode($object3->ingreso13->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI51', STRTOUPPER(utf8_encode($object3->ingreso13->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL51', STRTOUPPER(utf8_encode($object3->ingreso13->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM51', STRTOUPPER(utf8_encode($object3->ingreso13->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN51', STRTOUPPER(utf8_encode($object3->ingreso13->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO51', STRTOUPPER(utf8_encode($object3->ingreso13->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP51', STRTOUPPER(utf8_encode($object3->ingreso13->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ51', STRTOUPPER(utf8_encode($object3->ingreso13->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR51', STRTOUPPER(utf8_encode($object3->ingreso13->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS51', STRTOUPPER(utf8_encode($object3->ingreso13->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT51', STRTOUPPER(utf8_encode($object3->ingreso13->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU51', STRTOUPPER(utf8_encode($object3->ingreso13->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV51', STRTOUPPER(utf8_encode($object3->ingreso13->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D52', STRTOUPPER(utf8_encode($object3->ingreso14->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S52', STRTOUPPER(utf8_encode($object3->ingreso14->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U52', STRTOUPPER(utf8_encode($object3->ingreso14->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V52', STRTOUPPER(utf8_encode($object3->ingreso14->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W52', STRTOUPPER(utf8_encode($object3->ingreso14->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X52', STRTOUPPER(utf8_encode($object3->ingreso14->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y52', STRTOUPPER(utf8_encode($object3->ingreso14->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z52', STRTOUPPER(utf8_encode($object3->ingreso14->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA52', STRTOUPPER(utf8_encode($object3->ingreso14->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD52', STRTOUPPER(utf8_encode($object3->ingreso14->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE52', STRTOUPPER(utf8_encode($object3->ingreso14->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF52', STRTOUPPER(utf8_encode($object3->ingreso14->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG52', STRTOUPPER(utf8_encode($object3->ingreso14->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH52', STRTOUPPER(utf8_encode($object3->ingreso14->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI52', STRTOUPPER(utf8_encode($object3->ingreso14->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL52', STRTOUPPER(utf8_encode($object3->ingreso14->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM52', STRTOUPPER(utf8_encode($object3->ingreso14->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN52', STRTOUPPER(utf8_encode($object3->ingreso14->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO52', STRTOUPPER(utf8_encode($object3->ingreso14->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP52', STRTOUPPER(utf8_encode($object3->ingreso14->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ52', STRTOUPPER(utf8_encode($object3->ingreso14->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR52', STRTOUPPER(utf8_encode($object3->ingreso14->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS52', STRTOUPPER(utf8_encode($object3->ingreso14->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT52', STRTOUPPER(utf8_encode($object3->ingreso14->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU52', STRTOUPPER(utf8_encode($object3->ingreso14->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV52', STRTOUPPER(utf8_encode($object3->ingreso14->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('D53', STRTOUPPER(utf8_encode($object3->ingreso15->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('S53', STRTOUPPER(utf8_encode($object3->ingreso15->cantidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('U53', STRTOUPPER(utf8_encode($object3->ingreso15->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V53', STRTOUPPER(utf8_encode($object3->ingreso15->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W53', STRTOUPPER(utf8_encode($object3->ingreso15->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X53', STRTOUPPER(utf8_encode($object3->ingreso15->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y53', STRTOUPPER(utf8_encode($object3->ingreso15->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z53', STRTOUPPER(utf8_encode($object3->ingreso15->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA53', STRTOUPPER(utf8_encode($object3->ingreso15->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD53', STRTOUPPER(utf8_encode($object3->ingreso15->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE53', STRTOUPPER(utf8_encode($object3->ingreso15->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF53', STRTOUPPER(utf8_encode($object3->ingreso15->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG53', STRTOUPPER(utf8_encode($object3->ingreso15->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH53', STRTOUPPER(utf8_encode($object3->ingreso15->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI53', STRTOUPPER(utf8_encode($object3->ingreso15->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL53', STRTOUPPER(utf8_encode($object3->ingreso15->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM53', STRTOUPPER(utf8_encode($object3->ingreso15->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN53', STRTOUPPER(utf8_encode($object3->ingreso15->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO53', STRTOUPPER(utf8_encode($object3->ingreso15->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP53', STRTOUPPER(utf8_encode($object3->ingreso15->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ53', STRTOUPPER(utf8_encode($object3->ingreso15->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR53', STRTOUPPER(utf8_encode($object3->ingreso15->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS53', STRTOUPPER(utf8_encode($object3->ingreso15->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT53', STRTOUPPER(utf8_encode($object3->ingreso15->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU53', STRTOUPPER(utf8_encode($object3->ingreso15->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV53', STRTOUPPER(utf8_encode($object3->ingreso15->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U56', STRTOUPPER(utf8_encode($object4->diuresis->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V56', STRTOUPPER(utf8_encode($object4->diuresis->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W56', STRTOUPPER(utf8_encode($object4->diuresis->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X56', STRTOUPPER(utf8_encode($object4->diuresis->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y56', STRTOUPPER(utf8_encode($object4->diuresis->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z56', STRTOUPPER(utf8_encode($object4->diuresis->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA56', STRTOUPPER(utf8_encode($object4->diuresis->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD56', STRTOUPPER(utf8_encode($object4->diuresis->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE56', STRTOUPPER(utf8_encode($object4->diuresis->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF56', STRTOUPPER(utf8_encode($object4->diuresis->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG56', STRTOUPPER(utf8_encode($object4->diuresis->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH56', STRTOUPPER(utf8_encode($object4->diuresis->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI56', STRTOUPPER(utf8_encode($object4->diuresis->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL56', STRTOUPPER(utf8_encode($object4->diuresis->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM56', STRTOUPPER(utf8_encode($object4->diuresis->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN56', STRTOUPPER(utf8_encode($object4->diuresis->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO56', STRTOUPPER(utf8_encode($object4->diuresis->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP56', STRTOUPPER(utf8_encode($object4->diuresis->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ56', STRTOUPPER(utf8_encode($object4->diuresis->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR56', STRTOUPPER(utf8_encode($object4->diuresis->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS56', STRTOUPPER(utf8_encode($object4->diuresis->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT56', STRTOUPPER(utf8_encode($object4->diuresis->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU56', STRTOUPPER(utf8_encode($object4->diuresis->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV56', STRTOUPPER(utf8_encode($object4->diuresis->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U57', STRTOUPPER(utf8_encode($object4->evacuaciones->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V57', STRTOUPPER(utf8_encode($object4->evacuaciones->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W57', STRTOUPPER(utf8_encode($object4->evacuaciones->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X57', STRTOUPPER(utf8_encode($object4->evacuaciones->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y57', STRTOUPPER(utf8_encode($object4->evacuaciones->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z57', STRTOUPPER(utf8_encode($object4->evacuaciones->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA57', STRTOUPPER(utf8_encode($object4->evacuaciones->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD57', STRTOUPPER(utf8_encode($object4->evacuaciones->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE57', STRTOUPPER(utf8_encode($object4->evacuaciones->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF57', STRTOUPPER(utf8_encode($object4->evacuaciones->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG57', STRTOUPPER(utf8_encode($object4->evacuaciones->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH57', STRTOUPPER(utf8_encode($object4->evacuaciones->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI57', STRTOUPPER(utf8_encode($object4->evacuaciones->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL57', STRTOUPPER(utf8_encode($object4->evacuaciones->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM57', STRTOUPPER(utf8_encode($object4->evacuaciones->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN57', STRTOUPPER(utf8_encode($object4->evacuaciones->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO57', STRTOUPPER(utf8_encode($object4->evacuaciones->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP57', STRTOUPPER(utf8_encode($object4->evacuaciones->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ57', STRTOUPPER(utf8_encode($object4->evacuaciones->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR57', STRTOUPPER(utf8_encode($object4->evacuaciones->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS57', STRTOUPPER(utf8_encode($object4->evacuaciones->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT57', STRTOUPPER(utf8_encode($object4->evacuaciones->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU57', STRTOUPPER(utf8_encode($object4->evacuaciones->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV57', STRTOUPPER(utf8_encode($object4->evacuaciones->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U58', STRTOUPPER(utf8_encode($object4->ostomias->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V58', STRTOUPPER(utf8_encode($object4->ostomias->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W58', STRTOUPPER(utf8_encode($object4->ostomias->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X58', STRTOUPPER(utf8_encode($object4->ostomias->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y58', STRTOUPPER(utf8_encode($object4->ostomias->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z58', STRTOUPPER(utf8_encode($object4->ostomias->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA58', STRTOUPPER(utf8_encode($object4->ostomias->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD58', STRTOUPPER(utf8_encode($object4->ostomias->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE58', STRTOUPPER(utf8_encode($object4->ostomias->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF58', STRTOUPPER(utf8_encode($object4->ostomias->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG58', STRTOUPPER(utf8_encode($object4->ostomias->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH58', STRTOUPPER(utf8_encode($object4->ostomias->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI58', STRTOUPPER(utf8_encode($object4->ostomias->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL58', STRTOUPPER(utf8_encode($object4->ostomias->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM58', STRTOUPPER(utf8_encode($object4->ostomias->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN58', STRTOUPPER(utf8_encode($object4->ostomias->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO58', STRTOUPPER(utf8_encode($object4->ostomias->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP58', STRTOUPPER(utf8_encode($object4->ostomias->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ58', STRTOUPPER(utf8_encode($object4->ostomias->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR58', STRTOUPPER(utf8_encode($object4->ostomias->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS58', STRTOUPPER(utf8_encode($object4->ostomias->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT58', STRTOUPPER(utf8_encode($object4->ostomias->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU58', STRTOUPPER(utf8_encode($object4->ostomias->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV58', STRTOUPPER(utf8_encode($object4->ostomias->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U59', STRTOUPPER(utf8_encode($object4->vomito->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V59', STRTOUPPER(utf8_encode($object4->vomito->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W59', STRTOUPPER(utf8_encode($object4->vomito->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X59', STRTOUPPER(utf8_encode($object4->vomito->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y59', STRTOUPPER(utf8_encode($object4->vomito->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z59', STRTOUPPER(utf8_encode($object4->vomito->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA59', STRTOUPPER(utf8_encode($object4->vomito->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD59', STRTOUPPER(utf8_encode($object4->vomito->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE59', STRTOUPPER(utf8_encode($object4->vomito->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF59', STRTOUPPER(utf8_encode($object4->vomito->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG59', STRTOUPPER(utf8_encode($object4->vomito->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH59', STRTOUPPER(utf8_encode($object4->vomito->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI59', STRTOUPPER(utf8_encode($object4->vomito->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL59', STRTOUPPER(utf8_encode($object4->vomito->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM59', STRTOUPPER(utf8_encode($object4->vomito->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN59', STRTOUPPER(utf8_encode($object4->vomito->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO59', STRTOUPPER(utf8_encode($object4->vomito->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP59', STRTOUPPER(utf8_encode($object4->vomito->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ59', STRTOUPPER(utf8_encode($object4->vomito->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR59', STRTOUPPER(utf8_encode($object4->vomito->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS59', STRTOUPPER(utf8_encode($object4->vomito->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT59', STRTOUPPER(utf8_encode($object4->vomito->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU59', STRTOUPPER(utf8_encode($object4->vomito->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV59', STRTOUPPER(utf8_encode($object4->vomito->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV60', STRTOUPPER(utf8_encode($object4->sondaNasogastrica->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U61', STRTOUPPER(utf8_encode($object4->sondaPleural->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V61', STRTOUPPER(utf8_encode($object4->sondaPleural->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W61', STRTOUPPER(utf8_encode($object4->sondaPleural->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X61', STRTOUPPER(utf8_encode($object4->sondaPleural->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y61', STRTOUPPER(utf8_encode($object4->sondaPleural->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z61', STRTOUPPER(utf8_encode($object4->sondaPleural->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA61', STRTOUPPER(utf8_encode($object4->sondaPleural->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD61', STRTOUPPER(utf8_encode($object4->sondaPleural->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE61', STRTOUPPER(utf8_encode($object4->sondaPleural->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF61', STRTOUPPER(utf8_encode($object4->sondaPleural->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG61', STRTOUPPER(utf8_encode($object4->sondaPleural->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH61', STRTOUPPER(utf8_encode($object4->sondaPleural->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI61', STRTOUPPER(utf8_encode($object4->sondaPleural->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL61', STRTOUPPER(utf8_encode($object4->sondaPleural->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM61', STRTOUPPER(utf8_encode($object4->sondaPleural->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN61', STRTOUPPER(utf8_encode($object4->sondaPleural->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO61', STRTOUPPER(utf8_encode($object4->sondaPleural->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP61', STRTOUPPER(utf8_encode($object4->sondaPleural->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ61', STRTOUPPER(utf8_encode($object4->sondaPleural->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR61', STRTOUPPER(utf8_encode($object4->sondaPleural->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS61', STRTOUPPER(utf8_encode($object4->sondaPleural->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT61', STRTOUPPER(utf8_encode($object4->sondaPleural->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU61', STRTOUPPER(utf8_encode($object4->sondaPleural->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV61', STRTOUPPER(utf8_encode($object4->sondaPleural->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U62', STRTOUPPER(utf8_encode($object4->penRose->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V62', STRTOUPPER(utf8_encode($object4->penRose->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W62', STRTOUPPER(utf8_encode($object4->penRose->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X62', STRTOUPPER(utf8_encode($object4->penRose->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y62', STRTOUPPER(utf8_encode($object4->penRose->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z62', STRTOUPPER(utf8_encode($object4->penRose->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA62', STRTOUPPER(utf8_encode($object4->penRose->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD62', STRTOUPPER(utf8_encode($object4->penRose->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE62', STRTOUPPER(utf8_encode($object4->penRose->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF62', STRTOUPPER(utf8_encode($object4->penRose->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG62', STRTOUPPER(utf8_encode($object4->penRose->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH62', STRTOUPPER(utf8_encode($object4->penRose->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI62', STRTOUPPER(utf8_encode($object4->penRose->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL62', STRTOUPPER(utf8_encode($object4->penRose->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM62', STRTOUPPER(utf8_encode($object4->penRose->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN62', STRTOUPPER(utf8_encode($object4->penRose->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO62', STRTOUPPER(utf8_encode($object4->penRose->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP62', STRTOUPPER(utf8_encode($object4->penRose->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ62', STRTOUPPER(utf8_encode($object4->penRose->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR62', STRTOUPPER(utf8_encode($object4->penRose->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS62', STRTOUPPER(utf8_encode($object4->penRose->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT62', STRTOUPPER(utf8_encode($object4->penRose->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU62', STRTOUPPER(utf8_encode($object4->penRose->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV62', STRTOUPPER(utf8_encode($object4->penRose->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('U64', STRTOUPPER(utf8_encode($object4->otrosLabs->{8})));
    $objPHPExcel->getActiveSheet()->setCellValue('V64', STRTOUPPER(utf8_encode($object4->otrosLabs->{9})));
    $objPHPExcel->getActiveSheet()->setCellValue('W64', STRTOUPPER(utf8_encode($object4->otrosLabs->{10})));
    $objPHPExcel->getActiveSheet()->setCellValue('X64', STRTOUPPER(utf8_encode($object4->otrosLabs->{11})));
    $objPHPExcel->getActiveSheet()->setCellValue('Y64', STRTOUPPER(utf8_encode($object4->otrosLabs->{12})));
    $objPHPExcel->getActiveSheet()->setCellValue('Z64', STRTOUPPER(utf8_encode($object4->otrosLabs->{13})));
    $objPHPExcel->getActiveSheet()->setCellValue('AA64', STRTOUPPER(utf8_encode($object4->diuresis->{14})));
    $objPHPExcel->getActiveSheet()->setCellValue('AD64', STRTOUPPER(utf8_encode($object4->otrosLabs->{15})));
    $objPHPExcel->getActiveSheet()->setCellValue('AE64', STRTOUPPER(utf8_encode($object4->otrosLabs->{16})));
    $objPHPExcel->getActiveSheet()->setCellValue('AF64', STRTOUPPER(utf8_encode($object4->otrosLabs->{17})));
    $objPHPExcel->getActiveSheet()->setCellValue('AG64', STRTOUPPER(utf8_encode($object4->otrosLabs->{18})));
    $objPHPExcel->getActiveSheet()->setCellValue('AH64', STRTOUPPER(utf8_encode($object4->otrosLabs->{19})));
    $objPHPExcel->getActiveSheet()->setCellValue('AI64', STRTOUPPER(utf8_encode($object4->otrosLabs->{20})));
    $objPHPExcel->getActiveSheet()->setCellValue('AL64', STRTOUPPER(utf8_encode($object4->otrosLabs->{21})));
    $objPHPExcel->getActiveSheet()->setCellValue('AM64', STRTOUPPER(utf8_encode($object4->otrosLabs->{22})));
    $objPHPExcel->getActiveSheet()->setCellValue('AN64', STRTOUPPER(utf8_encode($object4->otrosLabs->{23})));
    $objPHPExcel->getActiveSheet()->setCellValue('AO64', STRTOUPPER(utf8_encode($object4->otrosLabs->{24})));
    $objPHPExcel->getActiveSheet()->setCellValue('AP64', STRTOUPPER(utf8_encode($object4->otrosLabs->{1})));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ64', STRTOUPPER(utf8_encode($object4->otrosLabs->{2})));
    $objPHPExcel->getActiveSheet()->setCellValue('AR64', STRTOUPPER(utf8_encode($object4->otrosLabs->{3})));
    $objPHPExcel->getActiveSheet()->setCellValue('AS64', STRTOUPPER(utf8_encode($object4->otrosLabs->{4})));
    $objPHPExcel->getActiveSheet()->setCellValue('AT64', STRTOUPPER(utf8_encode($object4->otrosLabs->{5})));
    $objPHPExcel->getActiveSheet()->setCellValue('AU64', STRTOUPPER(utf8_encode($object4->otrosLabs->{6})));
    $objPHPExcel->getActiveSheet()->setCellValue('AV64', STRTOUPPER(utf8_encode($object4->otrosLabs->{7})));

    $objPHPExcel->getActiveSheet()->setCellValue('AG8', STRTOUPPER(utf8_encode($object5->medicamento1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL8', STRTOUPPER(utf8_encode($object5->medicamento1->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO8', STRTOUPPER(utf8_encode($object5->medicamento1->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ8', STRTOUPPER(utf8_encode($object5->medicamento1->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS8', STRTOUPPER(utf8_encode($object5->medicamento1->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT8', STRTOUPPER(utf8_encode($object5->medicamento1->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU8', STRTOUPPER(utf8_encode($object5->medicamento1->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV8', STRTOUPPER(utf8_encode($object5->medicamento1->observaciones)));

    if ($object5->medicamento1->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS8")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento1->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT8")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento1->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU8")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG9', STRTOUPPER(utf8_encode($object5->medicamento2->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL9', STRTOUPPER(utf8_encode($object5->medicamento2->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO9', STRTOUPPER(utf8_encode($object5->medicamento2->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ9', STRTOUPPER(utf8_encode($object5->medicamento2->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS9', STRTOUPPER(utf8_encode($object5->medicamento2->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT9', STRTOUPPER(utf8_encode($object5->medicamento2->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU9', STRTOUPPER(utf8_encode($object5->medicamento2->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV9', STRTOUPPER(utf8_encode($object5->medicamento2->observaciones)));

    if ($object5->medicamento2->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS9")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento2->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT9")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento2->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU9")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG10', STRTOUPPER(utf8_encode($object5->medicamento3->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL10', STRTOUPPER(utf8_encode($object5->medicamento3->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO10', STRTOUPPER(utf8_encode($object5->medicamento3->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ10', STRTOUPPER(utf8_encode($object5->medicamento3->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS10', STRTOUPPER(utf8_encode($object5->medicamento3->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT10', STRTOUPPER(utf8_encode($object5->medicamento3->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU10', STRTOUPPER(utf8_encode($object5->medicamento3->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV10', STRTOUPPER(utf8_encode($object5->medicamento3->observaciones)));

    if ($object5->medicamento3->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS10")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento3->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS10")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento3->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS10")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG11', STRTOUPPER(utf8_encode($object5->medicamento4->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL11', STRTOUPPER(utf8_encode($object5->medicamento4->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO11', STRTOUPPER(utf8_encode($object5->medicamento4->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ11', STRTOUPPER(utf8_encode($object5->medicamento4->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS11', STRTOUPPER(utf8_encode($object5->medicamento4->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT11', STRTOUPPER(utf8_encode($object5->medicamento4->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU11', STRTOUPPER(utf8_encode($object5->medicamento4->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV11', STRTOUPPER(utf8_encode($object5->medicamento4->observaciones)));

    if ($object5->medicamento4->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS11")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento4->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT11")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento4->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU11")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG12', STRTOUPPER(utf8_encode($object5->medicamento5->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL12', STRTOUPPER(utf8_encode($object5->medicamento5->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO12', STRTOUPPER(utf8_encode($object5->medicamento5->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ12', STRTOUPPER(utf8_encode($object5->medicamento5->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS12', STRTOUPPER(utf8_encode($object5->medicamento5->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT12', STRTOUPPER(utf8_encode($object5->medicamento5->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU12', STRTOUPPER(utf8_encode($object5->medicamento5->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV12', STRTOUPPER(utf8_encode($object5->medicamento5->observaciones)));

    if ($object5->medicamento5->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS12")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento5->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT12")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento5->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU12")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG13', STRTOUPPER(utf8_encode($object5->medicamento6->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL13', STRTOUPPER(utf8_encode($object5->medicamento6->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO13', STRTOUPPER(utf8_encode($object5->medicamento6->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ13', STRTOUPPER(utf8_encode($object5->medicamento6->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS13', STRTOUPPER(utf8_encode($object5->medicamento6->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT13', STRTOUPPER(utf8_encode($object5->medicamento6->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU13', STRTOUPPER(utf8_encode($object5->medicamento6->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV13', STRTOUPPER(utf8_encode($object5->medicamento6->observaciones)));

    if ($object5->medicamento6->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS13")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento6->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT13")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento6->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU13")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG14', STRTOUPPER(utf8_encode($object5->medicamento7->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL14', STRTOUPPER(utf8_encode($object5->medicamento7->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO14', STRTOUPPER(utf8_encode($object5->medicamento7->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ14', STRTOUPPER(utf8_encode($object5->medicamento7->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS14', STRTOUPPER(utf8_encode($object5->medicamento7->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT14', STRTOUPPER(utf8_encode($object5->medicamento7->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU14', STRTOUPPER(utf8_encode($object5->medicamento7->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV14', STRTOUPPER(utf8_encode($object5->medicamento7->observaciones)));

    if ($object5->medicamento7->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS14")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento7->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT14")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento7->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU14")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG8', STRTOUPPER(utf8_encode($object5->medicamento1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL8', STRTOUPPER(utf8_encode($object5->medicamento1->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO8', STRTOUPPER(utf8_encode($object5->medicamento1->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ8', STRTOUPPER(utf8_encode($object5->medicamento1->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS8', STRTOUPPER(utf8_encode($object5->medicamento1->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT8', STRTOUPPER(utf8_encode($object5->medicamento1->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU8', STRTOUPPER(utf8_encode($object5->medicamento1->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV8', STRTOUPPER(utf8_encode($object5->medicamento1->observaciones)));

    if ($object5->medicamento1->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS8")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento1->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT8")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento1->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU8")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG15', STRTOUPPER(utf8_encode($object5->medicamento8->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL15', STRTOUPPER(utf8_encode($object5->medicamento8->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO15', STRTOUPPER(utf8_encode($object5->medicamento8->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ15', STRTOUPPER(utf8_encode($object5->medicamento8->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS15', STRTOUPPER(utf8_encode($object5->medicamento8->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT15', STRTOUPPER(utf8_encode($object5->medicamento8->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU15', STRTOUPPER(utf8_encode($object5->medicamento8->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV15', STRTOUPPER(utf8_encode($object5->medicamento8->observaciones)));

    if ($object5->medicamento8->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS15")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento8->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT15")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento8->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU15")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG16', STRTOUPPER(utf8_encode($object5->medicamento9->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL16', STRTOUPPER(utf8_encode($object5->medicamento9->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO16', STRTOUPPER(utf8_encode($object5->medicamento9->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ16', STRTOUPPER(utf8_encode($object5->medicamento9->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS16', STRTOUPPER(utf8_encode($object5->medicamento9->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT16', STRTOUPPER(utf8_encode($object5->medicamento9->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU16', STRTOUPPER(utf8_encode($object5->medicamento9->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV16', STRTOUPPER(utf8_encode($object5->medicamento9->observaciones)));

    if ($object5->medicamento9->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS16")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento9->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT16")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento9->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU16")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG17', STRTOUPPER(utf8_encode($object5->medicamento10->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL17', STRTOUPPER(utf8_encode($object5->medicamento10->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO17', STRTOUPPER(utf8_encode($object5->medicamento10->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ17', STRTOUPPER(utf8_encode($object5->medicamento10->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS17', STRTOUPPER(utf8_encode($object5->medicamento10->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT17', STRTOUPPER(utf8_encode($object5->medicamento10->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU17', STRTOUPPER(utf8_encode($object5->medicamento10->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV17', STRTOUPPER(utf8_encode($object5->medicamento10->observaciones)));

    if ($object5->medicamento10->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS17")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento10->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT17")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento10->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU17")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG18', STRTOUPPER(utf8_encode($object5->medicamento11->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL18', STRTOUPPER(utf8_encode($object5->medicamento11->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO18', STRTOUPPER(utf8_encode($object5->medicamento11->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ18', STRTOUPPER(utf8_encode($object5->medicamento11->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS18', STRTOUPPER(utf8_encode($object5->medicamento11->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT18', STRTOUPPER(utf8_encode($object5->medicamento11->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU18', STRTOUPPER(utf8_encode($object5->medicamento11->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV18', STRTOUPPER(utf8_encode($object5->medicamento11->observaciones)));

    if ($object5->medicamento11->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS18")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento11->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT18")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento11->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU18")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG19', STRTOUPPER(utf8_encode($object5->medicamento12->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL19', STRTOUPPER(utf8_encode($object5->medicamento12->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO19', STRTOUPPER(utf8_encode($object5->medicamento12->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ19', STRTOUPPER(utf8_encode($object5->medicamento12->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS19', STRTOUPPER(utf8_encode($object5->medicamento12->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT19', STRTOUPPER(utf8_encode($object5->medicamento12->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU19', STRTOUPPER(utf8_encode($object5->medicamento12->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV19', STRTOUPPER(utf8_encode($object5->medicamento12->observaciones)));

    if ($object5->medicamento12->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS19")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento12->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT19")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento12->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU19")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG20', STRTOUPPER(utf8_encode($object5->medicamento13->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL20', STRTOUPPER(utf8_encode($object5->medicamento13->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO20', STRTOUPPER(utf8_encode($object5->medicamento13->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ20', STRTOUPPER(utf8_encode($object5->medicamento13->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS20', STRTOUPPER(utf8_encode($object5->medicamento13->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT20', STRTOUPPER(utf8_encode($object5->medicamento13->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU20', STRTOUPPER(utf8_encode($object5->medicamento13->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV20', STRTOUPPER(utf8_encode($object5->medicamento13->observaciones)));

    if ($object5->medicamento13->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS20")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento13->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT20")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento13->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU20")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG21', STRTOUPPER(utf8_encode($object5->medicamento14->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL21', STRTOUPPER(utf8_encode($object5->medicamento14->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO21', STRTOUPPER(utf8_encode($object5->medicamento14->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ21', STRTOUPPER(utf8_encode($object5->medicamento14->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS21', STRTOUPPER(utf8_encode($object5->medicamento14->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT21', STRTOUPPER(utf8_encode($object5->medicamento14->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU21', STRTOUPPER(utf8_encode($object5->medicamento14->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV21', STRTOUPPER(utf8_encode($object5->medicamento14->observaciones)));

    if ($object5->medicamento14->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS21")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento14->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT21")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento14->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU21")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG22', STRTOUPPER(utf8_encode($object5->medicamento15->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL22', STRTOUPPER(utf8_encode($object5->medicamento15->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO22', STRTOUPPER(utf8_encode($object5->medicamento15->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ22', STRTOUPPER(utf8_encode($object5->medicamento15->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS22', STRTOUPPER(utf8_encode($object5->medicamento15->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT22', STRTOUPPER(utf8_encode($object5->medicamento15->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU22', STRTOUPPER(utf8_encode($object5->medicamento15->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV22', STRTOUPPER(utf8_encode($object5->medicamento15->observaciones)));

    if ($object5->medicamento15->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS22")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento15->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT22")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento15->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU22")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG23', STRTOUPPER(utf8_encode($object5->medicamento16->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL23', STRTOUPPER(utf8_encode($object5->medicamento16->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO23', STRTOUPPER(utf8_encode($object5->medicamento16->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ23', STRTOUPPER(utf8_encode($object5->medicamento16->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS23', STRTOUPPER(utf8_encode($object5->medicamento16->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT23', STRTOUPPER(utf8_encode($object5->medicamento16->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU23', STRTOUPPER(utf8_encode($object5->medicamento16->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV23', STRTOUPPER(utf8_encode($object5->medicamento16->observaciones)));

    if ($object5->medicamento16->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS23")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento16->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT23")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento16->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU23")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG24', STRTOUPPER(utf8_encode($object5->medicamento17->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL24', STRTOUPPER(utf8_encode($object5->medicamento17->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO24', STRTOUPPER(utf8_encode($object5->medicamento17->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ24', STRTOUPPER(utf8_encode($object5->medicamento17->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS24', STRTOUPPER(utf8_encode($object5->medicamento17->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT24', STRTOUPPER(utf8_encode($object5->medicamento17->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU24', STRTOUPPER(utf8_encode($object5->medicamento17->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV24', STRTOUPPER(utf8_encode($object5->medicamento17->observaciones)));

    if ($object5->medicamento17->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS24")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento17->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT24")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento17->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU24")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG25', STRTOUPPER(utf8_encode($object5->medicamento18->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL25', STRTOUPPER(utf8_encode($object5->medicamento18->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO25', STRTOUPPER(utf8_encode($object5->medicamento18->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ25', STRTOUPPER(utf8_encode($object5->medicamento18->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS25', STRTOUPPER(utf8_encode($object5->medicamento18->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT25', STRTOUPPER(utf8_encode($object5->medicamento18->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU25', STRTOUPPER(utf8_encode($object5->medicamento18->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV25', STRTOUPPER(utf8_encode($object5->medicamento18->observaciones)));

    if ($object5->medicamento18->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS25")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento18->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT25")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento18->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU25")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG26', STRTOUPPER(utf8_encode($object5->medicamento19->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL26', STRTOUPPER(utf8_encode($object5->medicamento19->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO26', STRTOUPPER(utf8_encode($object5->medicamento19->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ26', STRTOUPPER(utf8_encode($object5->medicamento19->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS26', STRTOUPPER(utf8_encode($object5->medicamento19->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT26', STRTOUPPER(utf8_encode($object5->medicamento19->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU26', STRTOUPPER(utf8_encode($object5->medicamento19->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV26', STRTOUPPER(utf8_encode($object5->medicamento19->observaciones)));

    if ($object5->medicamento19->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS25")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento19->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT25")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento19->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU25")->applyFromArray($border_style_red);
    }

    $objPHPExcel->getActiveSheet()->setCellValue('AG26', STRTOUPPER(utf8_encode($object5->medicamento20->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('AL26', STRTOUPPER(utf8_encode($object5->medicamento20->dosis)));
    $objPHPExcel->getActiveSheet()->setCellValue('AO26', STRTOUPPER(utf8_encode($object5->medicamento20->via)));
    $objPHPExcel->getActiveSheet()->setCellValue('AQ26', STRTOUPPER(utf8_encode($object5->medicamento20->frecuencia)));
    $objPHPExcel->getActiveSheet()->setCellValue('AS26', STRTOUPPER(utf8_encode($object5->medicamento20->horarioMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT26', STRTOUPPER(utf8_encode($object5->medicamento20->horarioVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU26', STRTOUPPER(utf8_encode($object5->medicamento20->horarioNocturno)));
    $objPHPExcel->getActiveSheet()->setCellValue('AV26', STRTOUPPER(utf8_encode($object5->medicamento20->observaciones)));

    if ($object5->medicamento20->horarioMatutinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AS26")->applyFromArray($border_style_blue);
    }
    if ($object5->medicamento20->horarioVespertinoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AT26")->applyFromArray($border_style_green);
    }
    if ($object5->medicamento20->horarioNocturnoChecked == "checked") {
        $sheet = $objPHPExcel->getActiveSheet();
        $sheet->getStyle("AU26")->applyFromArray($border_style_red);
    }


    $objPHPExcel->setActiveSheetIndex(0);

    $objPHPExcel->getActiveSheet()->setCellValue('A3', STRTOUPPER(utf8_encode($object6->cateter1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E3', STRTOUPPER(utf8_encode($object6->cateter1->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H3', STRTOUPPER(utf8_encode($object6->cateter1->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N3', STRTOUPPER(utf8_encode($object6->cateter1->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S3', STRTOUPPER(utf8_encode($object6->cateter1->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A4', STRTOUPPER(utf8_encode($object6->cateter2->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E4', STRTOUPPER(utf8_encode($object6->cateter2->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H4', STRTOUPPER(utf8_encode($object6->cateter2->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N4', STRTOUPPER(utf8_encode($object6->cateter2->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S4', STRTOUPPER(utf8_encode($object6->cateter2->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A5', STRTOUPPER(utf8_encode($object6->cateter3->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E5', STRTOUPPER(utf8_encode($object6->cateter3->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H5', STRTOUPPER(utf8_encode($object6->cateter3->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N5', STRTOUPPER(utf8_encode($object6->cateter3->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S5', STRTOUPPER(utf8_encode($object6->cateter3->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A6', STRTOUPPER(utf8_encode($object6->cateter4->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E6', STRTOUPPER(utf8_encode($object6->cateter4->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H6', STRTOUPPER(utf8_encode($object6->cateter4->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N6', STRTOUPPER(utf8_encode($object6->cateter4->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S6', STRTOUPPER(utf8_encode($object6->cateter4->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A7', STRTOUPPER(utf8_encode($object6->cateter5->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E7', STRTOUPPER(utf8_encode($object6->cateter5->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H7', STRTOUPPER(utf8_encode($object6->cateter5->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N7', STRTOUPPER(utf8_encode($object6->cateter5->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S7', STRTOUPPER(utf8_encode($object6->cateter5->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A8', STRTOUPPER(utf8_encode($object6->cateter6->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E8', STRTOUPPER(utf8_encode($object6->cateter6->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H8', STRTOUPPER(utf8_encode($object6->cateter6->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N8', STRTOUPPER(utf8_encode($object6->cateter6->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S8', STRTOUPPER(utf8_encode($object6->cateter6->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A10', STRTOUPPER(utf8_encode($object6->sonda1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E10', STRTOUPPER(utf8_encode($object6->sonda1->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H10', STRTOUPPER(utf8_encode($object6->sonda1->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N10', STRTOUPPER(utf8_encode($object6->sonda1->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S10', STRTOUPPER(utf8_encode($object6->sonda1->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A11', STRTOUPPER(utf8_encode($object6->sonda2->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E11', STRTOUPPER(utf8_encode($object6->sonda2->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H11', STRTOUPPER(utf8_encode($object6->sonda2->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N11', STRTOUPPER(utf8_encode($object6->sonda2->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S11', STRTOUPPER(utf8_encode($object6->sonda2->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A12', STRTOUPPER(utf8_encode($object6->sonda2->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E12', STRTOUPPER(utf8_encode($object6->sonda2->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H12', STRTOUPPER(utf8_encode($object6->sonda2->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N12', STRTOUPPER(utf8_encode($object6->sonda2->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S12', STRTOUPPER(utf8_encode($object6->sonda2->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A14', STRTOUPPER(utf8_encode($object6->ostomia1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E14', STRTOUPPER(utf8_encode($object6->ostomia1->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H14', STRTOUPPER(utf8_encode($object6->ostomia1->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N14', STRTOUPPER(utf8_encode($object6->ostomia1->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S14', STRTOUPPER(utf8_encode($object6->ostomia1->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A15', STRTOUPPER(utf8_encode($object6->ostomia2->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E15', STRTOUPPER(utf8_encode($object6->ostomia2->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H15', STRTOUPPER(utf8_encode($object6->ostomia2->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N15', STRTOUPPER(utf8_encode($object6->ostomia2->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S15', STRTOUPPER(utf8_encode($object6->ostomia2->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A16', STRTOUPPER(utf8_encode($object6->ostomia3->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E16', STRTOUPPER(utf8_encode($object6->ostomia3->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H16', STRTOUPPER(utf8_encode($object6->ostomia3->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N16', STRTOUPPER(utf8_encode($object6->ostomia3->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S16', STRTOUPPER(utf8_encode($object6->ostomia3->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('A18', STRTOUPPER(utf8_encode($object6->drenes1->nombre)));
    $objPHPExcel->getActiveSheet()->setCellValue('E18', STRTOUPPER(utf8_encode($object6->drenes1->calibre)));
    $objPHPExcel->getActiveSheet()->setCellValue('H18', STRTOUPPER(utf8_encode($object6->drenes1->fechaCuracion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N18', STRTOUPPER(utf8_encode($object6->drenes1->instalacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('S18', STRTOUPPER(utf8_encode($object6->drenes1->retiro)));

    $objPHPExcel->getActiveSheet()->setCellValue('F36', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->hora)));
    $objPHPExcel->getActiveSheet()->setCellValue('J36', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->hora)));
    $objPHPExcel->getActiveSheet()->setCellValue('N36', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->hora)));
    $objPHPExcel->getActiveSheet()->setCellValue('R36', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->hora)));

    $objPHPExcel->getActiveSheet()->setCellValue('F37', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->mVentilacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('J37', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->mVentilacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('N37', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->mVentilacion)));
    $objPHPExcel->getActiveSheet()->setCellValue('R37', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->mVentilacion)));

    $objPHPExcel->getActiveSheet()->setCellValue('F38', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->vc)));
    $objPHPExcel->getActiveSheet()->setCellValue('J38', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->vc)));
    $objPHPExcel->getActiveSheet()->setCellValue('N38', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->vc)));
    $objPHPExcel->getActiveSheet()->setCellValue('R38', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->vc)));

    $objPHPExcel->getActiveSheet()->setCellValue('F39', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->fr)));
    $objPHPExcel->getActiveSheet()->setCellValue('J39', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->fr)));
    $objPHPExcel->getActiveSheet()->setCellValue('N39', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->fr)));
    $objPHPExcel->getActiveSheet()->setCellValue('R39', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->fr)));

    $objPHPExcel->getActiveSheet()->setCellValue('F40', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->peep)));
    $objPHPExcel->getActiveSheet()->setCellValue('J40', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->peep)));
    $objPHPExcel->getActiveSheet()->setCellValue('N40', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->peep)));
    $objPHPExcel->getActiveSheet()->setCellValue('R40', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->peep)));

    $objPHPExcel->getActiveSheet()->setCellValue('F41', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->fio)));
    $objPHPExcel->getActiveSheet()->setCellValue('J41', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->fio)));
    $objPHPExcel->getActiveSheet()->setCellValue('N41', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->fio)));
    $objPHPExcel->getActiveSheet()->setCellValue('R41', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->fio)));

    $objPHPExcel->getActiveSheet()->setCellValue('F42', STRTOUPPER(utf8_encode($object7->parametroRespiratorio1->ie)));
    $objPHPExcel->getActiveSheet()->setCellValue('J42', STRTOUPPER(utf8_encode($object7->parametroRespiratorio2->ie)));
    $objPHPExcel->getActiveSheet()->setCellValue('N42', STRTOUPPER(utf8_encode($object7->parametroRespiratorio3->ie)));
    $objPHPExcel->getActiveSheet()->setCellValue('R42', STRTOUPPER(utf8_encode($object7->parametroRespiratorio4->ie)));

    $objPHPExcel->getActiveSheet()->setCellValue('F44', STRTOUPPER(utf8_encode($object7->gasometria1->ph)));
    $objPHPExcel->getActiveSheet()->setCellValue('J44', STRTOUPPER(utf8_encode($object7->gasometria2->ph)));
    $objPHPExcel->getActiveSheet()->setCellValue('N44', STRTOUPPER(utf8_encode($object7->gasometria3->ph)));
    $objPHPExcel->getActiveSheet()->setCellValue('R44', STRTOUPPER(utf8_encode($object7->gasometria4->ph)));

    $objPHPExcel->getActiveSheet()->setCellValue('F45', STRTOUPPER(utf8_encode($object7->gasometria1->pao)));
    $objPHPExcel->getActiveSheet()->setCellValue('J45', STRTOUPPER(utf8_encode($object7->gasometria2->pao)));
    $objPHPExcel->getActiveSheet()->setCellValue('N45', STRTOUPPER(utf8_encode($object7->gasometria3->pao)));
    $objPHPExcel->getActiveSheet()->setCellValue('R45', STRTOUPPER(utf8_encode($object7->gasometria4->pao)));

    $objPHPExcel->getActiveSheet()->setCellValue('F46', STRTOUPPER(utf8_encode($object7->gasometria1->pvo)));
    $objPHPExcel->getActiveSheet()->setCellValue('J46', STRTOUPPER(utf8_encode($object7->gasometria2->pvo)));
    $objPHPExcel->getActiveSheet()->setCellValue('N46', STRTOUPPER(utf8_encode($object7->gasometria3->pvo)));
    $objPHPExcel->getActiveSheet()->setCellValue('R46', STRTOUPPER(utf8_encode($object7->gasometria4->pvo)));

    $objPHPExcel->getActiveSheet()->setCellValue('F47', STRTOUPPER(utf8_encode($object7->gasometria1->co)));
    $objPHPExcel->getActiveSheet()->setCellValue('J47', STRTOUPPER(utf8_encode($object7->gasometria2->co)));
    $objPHPExcel->getActiveSheet()->setCellValue('N47', STRTOUPPER(utf8_encode($object7->gasometria3->co)));
    $objPHPExcel->getActiveSheet()->setCellValue('R47', STRTOUPPER(utf8_encode($object7->gasometria4->co)));

    $objPHPExcel->getActiveSheet()->setCellValue('F48', STRTOUPPER(utf8_encode($object7->gasometria1->sao)));
    $objPHPExcel->getActiveSheet()->setCellValue('J48', STRTOUPPER(utf8_encode($object7->gasometria2->sao)));
    $objPHPExcel->getActiveSheet()->setCellValue('N48', STRTOUPPER(utf8_encode($object7->gasometria3->sao)));
    $objPHPExcel->getActiveSheet()->setCellValue('R48', STRTOUPPER(utf8_encode($object7->gasometria4->sao)));

    $objPHPExcel->getActiveSheet()->setCellValue('F49', STRTOUPPER(utf8_encode($object7->gasometria1->hb)));
    $objPHPExcel->getActiveSheet()->setCellValue('J49', STRTOUPPER(utf8_encode($object7->gasometria2->hb)));
    $objPHPExcel->getActiveSheet()->setCellValue('N49', STRTOUPPER(utf8_encode($object7->gasometria3->hb)));
    $objPHPExcel->getActiveSheet()->setCellValue('R49', STRTOUPPER(utf8_encode($object7->gasometria4->hb)));

    $objPHPExcel->getActiveSheet()->setCellValue('F50', STRTOUPPER(utf8_encode($object7->gasometria1->hco)));
    $objPHPExcel->getActiveSheet()->setCellValue('J50', STRTOUPPER(utf8_encode($object7->gasometria2->hco)));
    $objPHPExcel->getActiveSheet()->setCellValue('N50', STRTOUPPER(utf8_encode($object7->gasometria3->hco)));
    $objPHPExcel->getActiveSheet()->setCellValue('R50', STRTOUPPER(utf8_encode($object7->gasometria4->hco)));

    $objPHPExcel->getActiveSheet()->setCellValue('AE2', STRTOUPPER(utf8_encode($object8->secresion->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AI2', STRTOUPPER(utf8_encode($object8->secresion->germenes)));
    $objPHPExcel->getActiveSheet()->setCellValue('AN2', STRTOUPPER(utf8_encode($object8->secresion->sensibilidad)));

    $objPHPExcel->getActiveSheet()->setCellValue('AE3', STRTOUPPER(utf8_encode($object8->urocultivo->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AI3', STRTOUPPER(utf8_encode($object8->urocultivo->germenes)));
    $objPHPExcel->getActiveSheet()->setCellValue('AN3', STRTOUPPER(utf8_encode($object8->urocultivo->sensibilidad)));

    $objPHPExcel->getActiveSheet()->setCellValue('AE4', STRTOUPPER(utf8_encode($object8->hemocultivo->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AI4', STRTOUPPER(utf8_encode($object8->hemocultivo->germenes)));
    $objPHPExcel->getActiveSheet()->setCellValue('AN4', STRTOUPPER(utf8_encode($object8->hemocultivo->sensibilidad)));

    $objPHPExcel->getActiveSheet()->setCellValue('AE6', STRTOUPPER(utf8_encode($object8->otrosCultivos1->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AI6', STRTOUPPER(utf8_encode($object8->otrosCultivos1->germenes)));
    $objPHPExcel->getActiveSheet()->setCellValue('AN6', STRTOUPPER(utf8_encode($object8->otrosCultivos1->sensibilidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('W6', STRTOUPPER(utf8_encode($object8->otrosCultivos1->nombre)));


    $objPHPExcel->getActiveSheet()->setCellValue('AE7', STRTOUPPER(utf8_encode($object8->otrosCultivos2->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AI7', STRTOUPPER(utf8_encode($object8->otrosCultivos2->germenes)));
    $objPHPExcel->getActiveSheet()->setCellValue('AN7', STRTOUPPER(utf8_encode($object8->otrosCultivos2->sensibilidad)));
    $objPHPExcel->getActiveSheet()->setCellValue('W7', STRTOUPPER(utf8_encode($object8->otrosCultivos2->nombre)));
}

while ($itemEnfermeriaB = mysqli_fetch_array($dataEnfermeriaB)) {


    $objPHPExcel->setActiveSheetIndex(0);
    $prueba9 =  $itemEnfermeriaB['dataPacienteEstudios'];
    $prueba10 =  $itemEnfermeriaB['dataPacienteObservaciones'];
    $prueba11 =  $itemEnfermeriaB['dataPacienteEscalas'];
    $prueba12 =  $itemEnfermeriaB['dataPacientePlanEnfermeria'];
    $prueba13 =  $itemEnfermeriaB['dataPacienteFirmas'];

    $object9 = json_decode($prueba9);
    $object10 = json_decode($prueba10);
    $object11 = json_decode($prueba11);
    $object12 = json_decode($prueba12);
    $object13 = json_decode($prueba13);





    $objPHPExcel->getActiveSheet()->setCellValue('W11', STRTOUPPER(utf8_encode($object9->estudio1->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE11', STRTOUPPER(utf8_encode($object9->estudio1->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W12', STRTOUPPER(utf8_encode($object9->estudio2->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE12', STRTOUPPER(utf8_encode($object9->estudio2->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W13', STRTOUPPER(utf8_encode($object9->estudio3->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE13', STRTOUPPER(utf8_encode($object9->estudio3->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W14', STRTOUPPER(utf8_encode($object9->estudio4->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE14', STRTOUPPER(utf8_encode($object9->estudio4->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W15', STRTOUPPER(utf8_encode($object9->estudio5->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE15', STRTOUPPER(utf8_encode($object9->estudio5->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W16', STRTOUPPER(utf8_encode($object9->estudio6->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE16', STRTOUPPER(utf8_encode($object9->estudio6->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W17', STRTOUPPER(utf8_encode($object9->estudio7->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE17', STRTOUPPER(utf8_encode($object9->estudio7->nombre)));

    $objPHPExcel->getActiveSheet()->setCellValue('W17', STRTOUPPER(utf8_encode($object9->estudio7->fecha)));
    $objPHPExcel->getActiveSheet()->setCellValue('AE17', STRTOUPPER(utf8_encode($object9->estudio7->nombre)));

    $rowArray = explode(",", $object10->observaciones->matutino);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'W21'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $rowArray = explode(",", $object10->observaciones->vespertino);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'W28'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $rowArray = explode(",", $object10->observaciones->nocturno);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'W35'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );


    $objPHPExcel->getActiveSheet()->setCellValue('AK51', "TOTAL:  " . STRTOUPPER(utf8_encode($object11->escalaGlasgow->Ocular)));
    $objPHPExcel->getActiveSheet()->setCellValue('AR51', STRTOUPPER(utf8_encode($object11->escalaGlasgow->Motriz)));
    $objPHPExcel->getActiveSheet()->setCellValue('AY51', STRTOUPPER(utf8_encode($object11->escalaGlasgow->Verbal)));

    $objPHPExcel->getActiveSheet()->setCellValue('AR61', STRTOUPPER(utf8_encode($object11->escalaRamsay->Valoracion)));

    $objPHPExcel->getActiveSheet()->setCellValue('AT1', STRTOUPPER(utf8_encode($object11->escalaLundBrower->cabezaPosterior)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT2', STRTOUPPER(utf8_encode($object11->escalaLundBrower->torsoPosterior)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT3', STRTOUPPER(utf8_encode($object11->escalaLundBrower->brazoIzqPosterior)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT4', STRTOUPPER(utf8_encode($object11->escalaLundBrower->brazoDerPosterior)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT5', STRTOUPPER(utf8_encode($object11->escalaLundBrower->piernaIzqPosterior)));
    $objPHPExcel->getActiveSheet()->setCellValue('AT6', STRTOUPPER(utf8_encode($object11->escalaLundBrower->piernaDerPosterior)));

    $objPHPExcel->getActiveSheet()->setCellValue('AU1', STRTOUPPER(utf8_encode($object11->escalaLundBrower->cabezaFrontal)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU2', STRTOUPPER(utf8_encode($object11->escalaLundBrower->torsoFrontal)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU3', STRTOUPPER(utf8_encode($object11->escalaLundBrower->brazoIzqFrontal)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU4', STRTOUPPER(utf8_encode($object11->escalaLundBrower->brazoDerFrontal)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU5', STRTOUPPER(utf8_encode($object11->escalaLundBrower->piernaIzqFrontal)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU6', STRTOUPPER(utf8_encode($object11->escalaLundBrower->piernaDerFrontal)));
    $objPHPExcel->getActiveSheet()->setCellValue('AU7', STRTOUPPER(utf8_encode($object11->escalaLundBrower->genitalesFrontal)));

    $objPHPExcel->getActiveSheet()->setCellValue('BH4', STRTOUPPER(utf8_encode($object12->planEnfermeria1->fecha1)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL4', STRTOUPPER(utf8_encode($object12->planEnfermeria1->diagnostico1)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP4', STRTOUPPER(utf8_encode($object12->planEnfermeria1->evaluacion1)));

    $rowArray = explode(",", $object12->planEnfermeria1->intervenciones1);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA4'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH7', STRTOUPPER(utf8_encode($object12->planEnfermeria2->fecha2)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL7', STRTOUPPER(utf8_encode($object12->planEnfermeria2->diagnostico2)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP7', STRTOUPPER(utf8_encode($object12->planEnfermeria2->evaluacion2)));

    $rowArray = explode(",", $object12->planEnfermeria2->intervenciones2);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA7'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH10', STRTOUPPER(utf8_encode($object12->planEnfermeria3->fecha3)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL10', STRTOUPPER(utf8_encode($object12->planEnfermeria3->diagnostico3)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP10', STRTOUPPER(utf8_encode($object12->planEnfermeria3->evaluacion3)));

    $rowArray = explode(",", $object12->planEnfermeria3->intervenciones3);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA10'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH13', STRTOUPPER(utf8_encode($object12->planEnfermeria4->fecha4)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL13', STRTOUPPER(utf8_encode($object12->planEnfermeria4->diagnostico4)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP13', STRTOUPPER(utf8_encode($object12->planEnfermeria4->evaluacion4)));

    $rowArray = explode(",", $object12->planEnfermeria4->intervenciones4);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA13'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH16', STRTOUPPER(utf8_encode($object12->planEnfermeria5->fecha5)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL16', STRTOUPPER(utf8_encode($object12->planEnfermeria5->diagnostico5)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP16', STRTOUPPER(utf8_encode($object12->planEnfermeria5->evaluacion5)));

    $rowArray = explode(",", $object12->planEnfermeria5->intervenciones5);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA16'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH19', STRTOUPPER(utf8_encode($object12->planEnfermeria6->fecha6)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL19', STRTOUPPER(utf8_encode($object12->planEnfermeria6->diagnostico6)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP19', STRTOUPPER(utf8_encode($object12->planEnfermeria6->evaluacion6)));

    $rowArray = explode(",", $object12->planEnfermeria6->intervenciones6);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA19'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH22', STRTOUPPER(utf8_encode($object12->planEnfermeria7->fecha7)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL22', STRTOUPPER(utf8_encode($object12->planEnfermeria7->diagnostico7)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP22', STRTOUPPER(utf8_encode($object12->planEnfermeria7->evaluacion7)));

    $rowArray = explode(",", $object12->planEnfermeria7->intervenciones7);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA22'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH25', STRTOUPPER(utf8_encode($object12->planEnfermeria8->fecha8)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL25', STRTOUPPER(utf8_encode($object12->planEnfermeria8->diagnostico8)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP25', STRTOUPPER(utf8_encode($object12->planEnfermeria8->evaluacion8)));

    $rowArray = explode(",", $object12->planEnfermeria8->intervenciones8);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA25'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH28', STRTOUPPER(utf8_encode($object12->planEnfermeria9->fecha9)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL28', STRTOUPPER(utf8_encode($object12->planEnfermeria9->diagnostico9)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP28', STRTOUPPER(utf8_encode($object12->planEnfermeria9->evaluacion9)));

    $rowArray = explode(",", $object12->planEnfermeria9->intervenciones9);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA28'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH30', STRTOUPPER(utf8_encode($object12->planEnfermeria10->fecha10)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL30', STRTOUPPER(utf8_encode($object12->planEnfermeria10->diagnostico10)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP30', STRTOUPPER(utf8_encode($object12->planEnfermeria10->evaluacion10)));

    $rowArray = explode(",", $object12->planEnfermeria10->intervenciones10);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA30'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH33', STRTOUPPER(utf8_encode($object12->planEnfermeria11->fecha11)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL33', STRTOUPPER(utf8_encode($object12->planEnfermeria11->diagnostico11)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP33', STRTOUPPER(utf8_encode($object12->planEnfermeria11->evaluacion11)));

    $rowArray = explode(",", $object12->planEnfermeria11->intervenciones11);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA33'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH36', STRTOUPPER(utf8_encode($object12->planEnfermeria12->fecha12)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL36', STRTOUPPER(utf8_encode($object12->planEnfermeria12->diagnostico12)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP36', STRTOUPPER(utf8_encode($object12->planEnfermeria12->evaluacion12)));

    $rowArray = explode(",", $object12->planEnfermeria12->intervenciones12);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA36'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH39', STRTOUPPER(utf8_encode($object12->planEnfermeria13->fecha13)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL39', STRTOUPPER(utf8_encode($object12->planEnfermeria13->diagnostico13)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP39', STRTOUPPER(utf8_encode($object12->planEnfermeria13->evaluacion13)));

    $rowArray = explode(",", $object12->planEnfermeria13->intervenciones13);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA39'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH42', STRTOUPPER(utf8_encode($object12->planEnfermeria14->fecha14)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL42', STRTOUPPER(utf8_encode($object12->planEnfermeria14->diagnostico14)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP42', STRTOUPPER(utf8_encode($object12->planEnfermeria14->evaluacion14)));

    $rowArray = explode(",", $object12->planEnfermeria14->intervenciones14);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA42'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH45', STRTOUPPER(utf8_encode($object12->planEnfermeria15->fecha15)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL45', STRTOUPPER(utf8_encode($object12->planEnfermeria15->diagnostico15)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP45', STRTOUPPER(utf8_encode($object12->planEnfermeria15->evaluacion15)));

    $rowArray = explode(",", $object12->planEnfermeria15->intervenciones15);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA45'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH48', STRTOUPPER(utf8_encode($object12->planEnfermeria16->fecha16)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL48', STRTOUPPER(utf8_encode($object12->planEnfermeria16->diagnostico16)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP48', STRTOUPPER(utf8_encode($object12->planEnfermeria16->evaluacion16)));

    $rowArray = explode(",", $object12->planEnfermeria16->intervenciones16);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA48'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH51', STRTOUPPER(utf8_encode($object12->planEnfermeria17->fecha17)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL51', STRTOUPPER(utf8_encode($object12->planEnfermeria17->diagnostico17)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP51', STRTOUPPER(utf8_encode($object12->planEnfermeria17->evaluacion17)));

    $rowArray = explode(",", $object12->planEnfermeria17->intervenciones17);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA51'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH54', STRTOUPPER(utf8_encode($object12->planEnfermeria18->fecha18)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL54', STRTOUPPER(utf8_encode($object12->planEnfermeria18->diagnostico18)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP54', STRTOUPPER(utf8_encode($object12->planEnfermeria18->evaluacion18)));

    $rowArray = explode(",", $object12->planEnfermeria18->intervenciones18);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA54'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH57', STRTOUPPER(utf8_encode($object12->planEnfermeria19->fecha19)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL57', STRTOUPPER(utf8_encode($object12->planEnfermeria19->diagnostico19)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP57', STRTOUPPER(utf8_encode($object12->planEnfermeria19->evaluacion19)));

    $rowArray = explode(",", $object12->planEnfermeria19->intervenciones19);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA57'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('BH60', STRTOUPPER(utf8_encode($object12->planEnfermeria20->fecha20)));
    $objPHPExcel->getActiveSheet()->setCellValue('BL60', STRTOUPPER(utf8_encode($object12->planEnfermeria20->diagnostico20)));
    $objPHPExcel->getActiveSheet()->setCellValue('CP60', STRTOUPPER(utf8_encode($object12->planEnfermeria20->evaluacion20)));

    $rowArray = explode(",", $object12->planEnfermeria20->intervenciones20);

    $columnArray = array_chunk($rowArray, 1);
    $objPHPExcel->getActiveSheet()
        ->fromArray(
            $columnArray,   // The data to set
            NULL,           // Array values with this value will not be set
            'CA60'            // Top left coordinate of the worksheet range where
            //    we want to set these values (default is A1)
        );

    $objPHPExcel->getActiveSheet()->setCellValue('I54', STRTOUPPER(utf8_encode($object13->firmas->enfermeroMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('I57', STRTOUPPER(utf8_encode($object13->firmas->enfermeroVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('I60', STRTOUPPER(utf8_encode($object13->firmas->enfermeroNocturno)));

    $objPHPExcel->getActiveSheet()->setCellValue('U54', STRTOUPPER(utf8_encode($object13->firmas->jefeMatutino)));
    $objPHPExcel->getActiveSheet()->setCellValue('U57', STRTOUPPER(utf8_encode($object13->firmas->jefeVespertino)));
    $objPHPExcel->getActiveSheet()->setCellValue('U60', STRTOUPPER(utf8_encode($object13->firmas->jefeNocturno)));
}



// redirect output to client browser
//name

header('Content-Type: application/vnd.openxmlformats-officedocument.objPHPExcelml.sheet');
header('Content-Disposition: attachment;filename="' . $nombrePaciente . '.xlsx"');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
