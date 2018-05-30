<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

$data = getData("ETH");
//echo json_encode($data);
//writeData($data);

writeToSpreadSheet($data);

/**
 * @throws \PhpOffice\PhpSpreadsheet\Exception
 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
 */
function writeToSpreadSheet($data , $coinSymbol){
    $spreadsheet = new Spreadsheet();
    $sheet1 = new Worksheet($spreadsheet, 'Asks');
    $sheet2 = new Worksheet($spreadsheet, 'Bids');

// Data writing

    $spreadsheet->addSheet($sheet1,0);
    $spreadsheet->addSheet($sheet2,1);

    $sheets = array("Asks" => $spreadsheet->getSheet(0), "Bids" => $spreadsheet->getSheet(1));

    foreach ($sheets as $sheet){
        $sheet->getCell('A1')->setValue("Price");
        $sheet->getStyle('A1')->getAlignment()->setWrapText(true);
        $sheet->getCell('B1')->setValue($coinSymbol);
        $sheet->getStyle('B1')->getAlignment()->setWrapText(true);
        $sheet->getCell('C1')->setValue("BTC");
        $sheet->getStyle('C1')->getAlignment()->setWrapText(true);
        $sheet->getCell('D1')->setValue("Sum(BTC)");
        $sheet->getStyle('D1')->getAlignment()->setWrapText(true);
    }

    $asks = $data["asks"];
    $bids = $data["bids"];


    $sheets['Asks']->getCell("A2")->setValue($asks[0][0]);
    $sheets['Asks']->getStyle("A2")->getAlignment()->setWrapText(true);
    $sheets['Asks']->getCell("B2")->setValue($asks[0][1]);
    $sheets['Asks']->getStyle("B2")->getAlignment()->setWrapText(true);
    $sheets['Asks']->getCell("C2")->setValue($asks[0][0] * $asks[0][1]);
    $sheets['Asks']->getStyle("C2")->getAlignment()->setWrapText(true);
    $sheets['Asks']->getCell("D2")->setValue($asks[0][0] * $asks[0][1]);
    $sheets['Asks']->getStyle("D2")->getAlignment()->setWrapText(true);

    $sheets['Bids']->getCell("A2")->setValue($bids[0][0]);
    $sheets['Bids']->getStyle("A2")->getAlignment()->setWrapText(true);
    $sheets['Bids']->getCell("B2")->setValue($bids[0][1]);
    $sheets['Bids']->getStyle("B2")->getAlignment()->setWrapText(true);
    $sheets['Bids']->getCell("C2")->setValue($bids[0][0] * $bids[0][1]);
    $sheets['Bids']->getStyle("C2")->getAlignment()->setWrapText(true);
    $sheets['Bids']->getCell("D2")->setValue($bids[0][0] * $bids[0][1]);
    $sheets['Bids']->getStyle("D2")->getAlignment()->setWrapText(true);

    for($i = 3; $i < count($asks); $i++){
        $sheets['Asks']->getCell("A".$i)->setValue($asks[$i - 2][0]);
        $sheets['Asks']->getStyle("A".$i)->getAlignment()->setWrapText(true);
        $sheets['Asks']->getCell("B".$i)->setValue($asks[$i - 2][1]);
        $sheets['Asks']->getStyle("B".$i)->getAlignment()->setWrapText(true);
        $sheets['Asks']->getCell("C".$i)->setValue($asks[$i - 2][0] * $asks[$i - 2][1]);
        $sheets['Asks']->getStyle("C".$i)->getAlignment()->setWrapText(true);

        $value = $sheets['Asks']->getCell("D".($i - 1))->getValue() + $asks[$i - 2][0] * $asks[$i - 2][1];
        $sheets['Asks']->getCell("D".$i)->setValue($value);
        $sheets['Asks']->getStyle("D".$i)->getAlignment()->setWrapText(true);
    }

    for($i = 3; $i < count($bids); $i++){
        $sheets['Bids']->getCell("A".$i)->setValue($bids[$i - 2][0]);
        $sheets['Bids']->getStyle("A".$i)->getAlignment()->setWrapText(true);
        $sheets['Bids']->getCell("B".$i)->setValue($bids[$i - 2][1]);
        $sheets['Bids']->getStyle("B".$i)->getAlignment()->setWrapText(true);
        $sheets['Bids']->getCell("C".$i)->setValue($bids[$i - 2][0] * $bids[$i - 2][1]);
        $sheets['Bids']->getStyle("C".$i)->getAlignment()->setWrapText(true);

        $value = $sheets['Asks']->getCell("D".($i - 1))->getValue() + $bids[$i - 2][0] * $bids[$i - 2][1];
        $sheets['Bids']->getCell("D".$i)->setValue($value);
        $sheets['Bids']->getStyle("D".$i)->getAlignment()->setWrapText(true);
    }


    $sheetIndex = $spreadsheet->getIndex(
        $spreadsheet->getSheetByName('Worksheet')
    );
    $spreadsheet->removeSheetByIndex($sheetIndex);
    $writer = new Xlsx($spreadsheet);

    $filename = "$coinSymbol-BTC";
    header('Content-Disposition: attachment;filename="'. $filename .'.xls"'); /*-- $filename is  xsl filename ---*/
    header('Cache-Control: max-age=0');

    $writer->save('php://output');

}

function getData($coinSymbol){

    $results = array();
    $curl = curl_init();

    curl_setopt_array($curl, array(
        CURLOPT_URL => "https://poloniex.com/public?command=returnOrderBook&currencyPair=BTC_$coinSymbol&depth=5000",
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