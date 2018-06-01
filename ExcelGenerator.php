<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xls;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

if(!empty($_POST['coinSymbolFrom'])){
    $coinSymbolFrom = $_POST['coinSymbolFrom'];
    $coinSymbolTo = $_POST['coinSymbolTo'];
    $data = getData($coinSymbolFrom, $coinSymbolTo);
    writeToSpreadSheet($data, $coinSymbolFrom, $coinSymbolTo);
}

/**
 * @throws \PhpOffice\PhpSpreadsheet\Exception
 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
 */
function writeToSpreadSheet($data , $coinSymbolFrom, $coinSymbolTo){
    $spreadsheet = new Spreadsheet();
    $sheet1 = new Worksheet($spreadsheet, 'Sells');
    $sheet2 = new Worksheet($spreadsheet, 'Buys');

// Data writing

    $spreadsheet->addSheet($sheet1,0);
    $spreadsheet->addSheet($sheet2,1);

    $sheets = array("Asks" => $spreadsheet->getSheet(0), "Bids" => $spreadsheet->getSheet(1));

    $asks = $data["asks"];
    $bids = $data["bids"];

    $cellWidth = 20;

    $sheets['Asks']->getColumnDimension('A')->setWidth($cellWidth - 2);
    $sheets['Asks']->getColumnDimension('B')->setWidth($cellWidth);
    $sheets['Asks']->getColumnDimension('C')->setWidth($cellWidth - 4);
    $sheets['Asks']->getColumnDimension('D')->setWidth($cellWidth - 4);

    $sheets['Bids']->getColumnDimension('A')->setWidth($cellWidth - 2);
    $sheets['Bids']->getColumnDimension('B')->setWidth($cellWidth);
    $sheets['Bids']->getColumnDimension('C')->setWidth($cellWidth - 4);
    $sheets['Bids']->getColumnDimension('D')->setWidth($cellWidth - 4);

    $sheets['Asks']->getStyle("A1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Asks']->getCell("A1")->setValue(setValue($asks[0][0]));
    $sheets['Asks']->getStyle('A1')->getAlignment()->setWrapText(true);

    $sheets['Asks']->getStyle("B1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Asks']->getCell("B1")->setValue(setValue($asks[0][1]));
    $sheets['Asks']->getStyle('B1')->getAlignment()->setWrapText(true);

    $sheets['Asks']->getStyle("C1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Asks']->getCell("C1")->setValue(setValue($asks[0][0]) * setValue($asks[0][1]));
    $sheets['Asks']->getStyle('C1')->getAlignment()->setWrapText(true);

    $sheets['Asks']->getStyle("D1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Asks']->getCell("D1")->setValue(setValue($asks[0][0]) * setValue($asks[0][1]));
    $sheets['Asks']->getStyle('D1')->getAlignment()->setWrapText(true);

    $sheets['Bids']->getStyle("A1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Bids']->getCell("A1")->setValue(setValue($bids[0][0]));
    $sheets['Bids']->getStyle('A1')->getAlignment()->setWrapText(true);

    $sheets['Bids']->getStyle("B1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Bids']->getCell("B1")->setValue(setValue($bids[0][1]));
    $sheets['Bids']->getStyle('B1')->getAlignment()->setWrapText(true);

    $sheets['Bids']->getStyle("C1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Bids']->getCell("C1")->setValue(setValue($bids[0][0]) * setValue($bids[0][1]));
    $sheets['Bids']->getStyle('C1')->getAlignment()->setWrapText(true);

    $sheets['Bids']->getStyle("D1")->getNumberFormat()->setFormatCode('0.00000000');
    $sheets['Bids']->getCell("D1")->setValue(setValue($bids[0][0]) * setValue($bids[0][1]));
    $sheets['Bids']->getStyle('D1')->getAlignment()->setWrapText(true);

    for($i = 2; $i <= count($asks); $i++){
        $sheets['Asks']->getStyle("A".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $sheets['Asks']->getCell("A".$i)->setValue(setValue($asks[$i - 1][0]));
        $sheets['Asks']->getStyle("A".$i)->getAlignment()->setWrapText(true);

        $sheets['Asks']->getStyle("B".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $sheets['Asks']->getCell("B".$i)->setValue(setValue($asks[$i - 1][1]));
        $sheets['Asks']->getStyle("B".$i)->getAlignment()->setWrapText(true);

        $sheets['Asks']->getStyle("C".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $sheets['Asks']->getCell("C".$i)->setValue(setValue($asks[$i - 1][0]) * setValue($asks[$i - 1][1]));
        $sheets['Asks']->getStyle("C".$i)->getAlignment()->setWrapText(true);

        $sheets['Asks']->getStyle("D".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $value = $sheets['Asks']->getCell("D".($i - 1))->getValue() + setValue($asks[$i - 1][0]) * setValue($asks[$i - 1][1]);
        $sheets['Asks']->getCell("D".$i)->setValue($value);
        $sheets['Asks']->getStyle("D".$i)->getAlignment()->setWrapText(true);
    }

    for($i = 2; $i <= count($bids); $i++){
        $sheets['Bids']->getStyle("A".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $sheets['Bids']->getCell("A".$i)->setValue(setValue($bids[$i - 1][0]));
        $sheets['Bids']->getStyle("A".$i)->getAlignment()->setWrapText(true);

        $sheets['Bids']->getStyle("B".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $sheets['Bids']->getCell("B".$i)->setValue(setValue($bids[$i - 1][1]));
        $sheets['Bids']->getStyle("B".$i)->getAlignment()->setWrapText(true);

        $sheets['Bids']->getStyle("C".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $sheets['Bids']->getCell("C".$i)->setValue(setValue($bids[$i - 1][0]) * setValue($bids[$i - 1][1]));
        $sheets['Bids']->getStyle("C".$i)->getAlignment()->setWrapText(true);

        $sheets['Bids']->getStyle("D".$i)->getNumberFormat()->setFormatCode('0.00000000');
        $value = $sheets['Bids']->getCell("D".($i - 1))->getValue() + setValue($bids[$i - 1][0]) * setValue($bids[$i - 1][1]);
        $sheets['Bids']->getCell("D".$i)->setValue($value);
        $sheets['Bids']->getStyle("D".$i)->getAlignment()->setWrapText(true);
    }


    $sheetIndex = $spreadsheet->getIndex(
        $spreadsheet->getSheetByName('Worksheet')
    );
    $spreadsheet->removeSheetByIndex($sheetIndex);
    $writer = new Xls($spreadsheet);

    $t=time();
    $date = date("Y-m-d",$t);

    $filename = "$coinSymbolTo-$coinSymbolFrom $date $t";
    header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    header('Content-Disposition: attachment;filename="'. $filename .'.xls"'); /*-- $filename is  xsl filename ---*/
    header('Cache-Control: max-age=0');

    $writer->save('php://output');

}

function setValue($value){
    return number_format( (float) $value, 8, '.', '');
}

function getData($coinSymbolFrom, $coinSymbolTo){

    $results = array();
    $curl = curl_init();

    curl_setopt_array($curl, array(
        CURLOPT_URL => "https://poloniex.com/public?command=returnOrderBook&currencyPair=$coinSymbolTo".'_'."$coinSymbolFrom&depth=9000",
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_ENCODING => "",
        CURLOPT_MAXREDIRS => 10,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
        CURLOPT_CUSTOMREQUEST => "GET",
        CURLOPT_HTTPHEADER => array(
            "Cache-Control: no-cache",
            "Postman-Token: db11ee92-e5be-43ae-b6b4-5c6319bf6570"
        ),
    ));

    $response = curl_exec($curl);
    $err = curl_error($curl);

    curl_close($curl);

    if ($err) {
        echo "cURL Error #:" . $err;
    } else {
        $results = json_decode($response, true);
    }

    return $results;
}
