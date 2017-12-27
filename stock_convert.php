<?php
ini_set('display_errors', 'on');
require_once 'PHPExcel/IOFactory.php';
$objPHPExcel = PHPExcel_IOFactory::load('ekinler_quick_shop.xlsx');
$rows = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$currentTime = Date("c", time());

$stockReceiptsArr = [];
foreach ($rows as $key => $row){ 
	if ($key == 1) {
		continue;
	}

	$arr = [	
		'code' => $row['A'],
	    'name' => $row['B'],
	    'name2' => $row['C'],
	    'specialCode' => $row['D'],
	    'cardType' => $row['E'],
	    'trackingType' => $row['G']
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