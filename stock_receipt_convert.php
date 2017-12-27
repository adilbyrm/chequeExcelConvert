<?php
ini_set('display_errors', 'on');
require_once 'PHPExcel/IOFactory.php';
$objPHPExcel = PHPExcel_IOFactory::load('stock_receipts.xlsx');
$rows = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$currentTime = Date("c", time());

$stockReceiptsArr = [];
foreach ($rows as $key => $row){ 
	if ($key == 1) {
		continue;
	}

	if ($row['G'] == '0') {
		$depotID = '3';
		$depotName = 'İSTOÇ';
	} elseif ($row['G'] == '2' || $row['G'] == '10' || $row['G'] == '12' || $row['G'] == '13' || $row['G'] == '14' || $row['G'] == '17' || $row['G'] == '18') {
		$depotID = '1';
		$depotName = 'FABRİKA';
	} elseif ($row['G'] == '3') {
		$depotID = '2';
		$depotName = 'ANTALYA';
	} elseif ($row['G'] == '7' || $row['G'] == '8') {
		$depotID = '5';
		$depotName = 'YEDEK PARÇA DEPOSU';
	} elseif ($row['G'] == '19') {
		$depotID = '6';
		$depotName = 'OĞUZ TÜRKİSTANLI ARAÇ';
	} elseif ($row['G'] == '20') {
		$depotID = '10';
		$depotName = 'ÖKKEŞ KAYA ARAÇ';
	} elseif ($row['G'] == '21') {
		$depotID = '7';
		$depotName = 'TARIK DÖNBAK ARAÇ';
	} elseif ($row['G'] == '22') {
		$depotID = '12';
		$depotName = 'FATİH OĞUZ ARAÇ';
	} elseif ($row['G'] == '25') {
		$depotID = '9';
		$depotName = 'GAZANFER ABDULLAH ARAÇ';
	} elseif ($row['G'] == '26') {
		$depotID = '11';
		$depotName = 'FATİH ŞİMŞEK ARAÇ';
	} elseif ($row['G'] == '27') {
		$depotID = '8';
		$depotName = 'TUGAY YILDIRIM ARAÇ';
	} elseif ($row['G'] == '11') {
		$depotID = '13';
		$depotName = '34 PVR 10 ANTALYA ARAÇ';
	} else {
		continue;
 	}

	$arr = [	
		'barcode' => $row['A'],
	    'depotAmount' => $row['C'],
	    'depotID' => $depotID,
	    'depotName' => $depotName
	];

	$stockReceiptsArr[$arr['depotName']][] = $arr;
}

// var_dump($stockReceiptsArr);exit;

$i = 0;
$k = 0;

foreach ($stockReceiptsArr as $key1 => $depots) {
	$path = $key1;
	$k++;
	$totalAmount = 0;
	$cloneDepots = $depots;
	foreach ($cloneDepots as $value) {
		$totalAmount += $value['depotAmount'];
	}


	$xml = '';
	foreach ($depots as $key2 => $stocks) {
		$i++;

		$xml = "<StockReceiptStocks>\n";
		$xml .= "<RowID>1</RowID>\n";
		$xml .= "<RowAddDateTime>{$currentTime}</RowAddDateTime>\n";
		$xml .= "<RowAddUserNo>0</RowAddUserNo>\n";
		$xml .= "<RowEditDateTime>{$currentTime}</RowEditDateTime>\n";
		$xml .= "<RowEditUserNo>0</RowEditUserNo>\n";
		$xml .= "<ID>1</ID>\n";
		$xml .= "<ReceiptType>0</ReceiptType>\n";
		$xml .= "<Time>{$currentTime}</Time>\n";
		$xml .= "<DepotID>".$stocks['depotID']."</DepotID>\n";
		$xml .= "<TargetDepotID>0</TargetDepotID>\n";
		$xml .= "<StockCode>".$stocks['barcode']."</StockCode>\n";
		$xml .= "<Number></Number>\n";
		$xml .= "<UnitName>Adet</UnitName>\n";
		$xml .= "<Amount>".$stocks['depotAmount']."</Amount>\n";
		$xml .= "<DepotAmount>0</DepotAmount>\n";
		$xml .= "<TargetDepotAmount>0</TargetDepotAmount>\n";
		$xml .= "<Status>1</Status>\n";
		$xml .= "</StockReceiptStocks>\n";

		$stockReceipts = "<StockReceipts>\n";
		$stockReceipts .= "<RowID>1</RowID>\n";
		$stockReceipts .= "<RowAddDateTime>{$currentTime}</RowAddDateTime>\n";
		$stockReceipts .= "<RowAddUserNo>0</RowAddUserNo>\n";
		$stockReceipts .= "<RowEditDateTime>{$currentTime}</RowEditDateTime>\n";
		$stockReceipts .= "<RowEditUserNo>0</RowEditUserNo>\n";
		$stockReceipts .= "<ID>1</ID>\n";
		$stockReceipts .= "<ReceiptNo>{$k}</ReceiptNo>\n";
		$stockReceipts .= "<ReceiptType>0</ReceiptType>\n";
		$stockReceipts .= "<Time>".Date("c", time())."</Time>\n";
		$stockReceipts .= "<DepotID>".$stocks['depotID']."</DepotID>\n";
		$stockReceipts .= "<DepotName>".$stocks['depotName']."</DepotName>\n";
		$stockReceipts .= "<TargetDepotID>0</TargetDepotID>\n";
		$stockReceipts .= "<Status>1</Status>\n";
		$stockReceipts .= "<Explanation></Explanation>\n";
		$stockReceipts .= "<TotalAmount>".$stocks['depotAmount']."</TotalAmount>\n";
		$stockReceipts .= "<SettingID>1</SettingID>\n";
		$stockReceipts .= "</StockReceipts>\n";

		$output = '<StockReceipts>' . $stockReceipts . $xml . '</StockReceipts>';

		if (!is_dir('stock_receipt_export/' . $path . '/')) {
            mkdir('stock_receipt_export/' . $path, 0755, true);
        }
        $fileName = str_replace('/', '_', $path . '/');
		$file = fopen('stock_receipt_export/' . $path . '/'. $path . $i . ".xml", "w");
		fwrite($file, $output);
		fclose($file);

	}

	

	// $file = fopen('stock_receipt_export/StockReceipts_' . $path . ".xml", "w");
	// fwrite($file, $output);
	// fclose($file);
}