<?php
ini_set('display_errors', 'on');
require_once 'PHPExcel/IOFactory.php';
$objPHPExcel = PHPExcel_IOFactory::load('list.xlsx');
$rows = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);

$currentTime = Date("c", time());
$time1 = Date("c", strtotime(' -1 day'));

$chequeBondReceiptsArr = [];
foreach ($rows as $key => $row){ 
	if ($key == 1) {
		continue;
	}
	$arr = [	
		'chequeOrBond' => $row['A'],
	    'bank' => $row['B'],
	    'bankNo' => $row['D'],
	    'branch' => $row['E'],
	    'branchNo' => $row['G'],
	    'chequeNo' => $row['I'],
	    'debtor' => $row['J'], // keşidedi
	    'maturityDate' => $row['L'],  // vadesi // MaturityDate
	    'price' => (double)str_replace(',', '', $row['M']),
	    'currencyCode' => $row['N'],
	    'currencyNo' => $row['O'],
	    'specialCode' => $row['P'],
	    'accountCode' => (int)$row['R'],
	    'accountName' => $row['T'],
	    'lastStatus' => $row['W'], // status=0
	    'lastStatusCode' => $row['Y'],
	    'updateType' => $row['AA'],
	    'updateTypeCode' => $row['AC'],
	    'chequeAccountNumber' => $row['AD'],
	    'paymentPlace' => $row['AF'],
	    'own' => $row['AH'],
	    'ownCode' => $row['AJ'],
	    'currentCurrencyCode' => $row['AK'],
	    'currenctCurrencyNo' => $row['AN'],
	    'time' => $time1,
	    'pursuitNo' => $row['AO'],
	    'depotID' => $row['AQ']
	];
	if ($arr['accountCode'] < 1) {
		continue;
	}
	$chequeBondReceiptsArr[$arr['lastStatus']][$arr['currencyCode']][$arr['updateType']][$arr['accountCode']][] = $arr;
}

// var_dump($chequeBondReceiptsArr);exit;

$i = 0;
$k = 0;
foreach ($chequeBondReceiptsArr as $key1 => $cheque) {
	
	foreach ($cheque as $key2 => $cheque1) {
		foreach ($cheque1 as $key3 => $cheque2) {
			$path = $key1. '/' .$key2. '/'. $key3 .'/';
			foreach ($cheque2 as $key4 => $cheque3) {
				$i++;
				$chequeBondReceiptChequeBonds = '';
				$chequesBonds = '';
				$totalPrice = 0.00;
				$totalMaturityDate = 0;
				$maturityTotalPrice = 0;
				foreach ($cheque3 as $key5 => $chequeDetail){
					$k++;
					$totalPrice += $chequeDetail['price'];
					$time = $chequeDetail['time'];
					$maturityDate = $chequeDetail['maturityDate'];
					$maturityDate = Date("c", strtotime($maturityDate));
					$dateDiff = strtotime($maturityDate) - strtotime($time);
					$maturityDayCount = floor($dateDiff / (60 * 60 * 24));
					$own = $chequeDetail['ownCode'] == '0' ? 'false' : 'true';
					$currencyNo = $chequeDetail['currencyNo'];
					$accountCode = $chequeDetail['accountCode'];
					$accountName = $chequeDetail['accountName'];
					$updateTypeCode = $chequeDetail['updateTypeCode'];
					$totalMaturityDate += (int)strtotime($maturityDate);
					$maturityTotalPrice += $maturityDayCount * $chequeDetail['price'];
					$depotID = $chequeDetail['depotID'] == '0' ? '3' : ($chequeDetail['depotID'] == '2' ? '1' : ($chequeDetail['depotID'] == '3' ? '2' : '1'));
					$depotName = $depotID == '3' ? 'İSTOÇ' : ($depotID == '1' ? 'FABRİKA' : ($depotID == '2' ? 'ANTALYA' : '')) ;
					$count = count($cheque3); 

					if ($chequeDetail['currencyNo'] == '1') {
						$currencyPrice = 1;
						$currencyCode = 'TL';
					} elseif ($chequeDetail['currencyNo'] == '2') {
						$currencyPrice = 3.7710;
						$currencyCode = 'USD';
					} else {
						$currencyPrice = 4.0506;
						$currencyCode = 'EUR';
					}

					if ($chequeDetail['currenctCurrencyNo'] == '1') {
						$currentCurrencyPrice = 1;
						$currentCurrencyNo = '1';
						$currentCurrencyCode = 'TL';
					} elseif ($chequeDetail['currenctCurrencyNo'] == '2') {
						$currentCurrencyPrice = 3.7710;
						$currentCurrencyNo = '2';
						$currentCurrencyCode = 'USD';
					} else {
						$currentCurrencyPrice = 4.0506;
						$currentCurrencyNo = '5';
						$currentCurrencyCode = 'EUR';
					}
					
					$chequeBondReceiptChequeBonds .= "<ChequeBondReceiptChequesBonds>";
					$chequeBondReceiptChequeBonds .= "<RowID>{$k}</RowID>";
					$chequeBondReceiptChequeBonds .= "<RowAddDateTime>{$currentTime}</RowAddDateTime>";
					$chequeBondReceiptChequeBonds .= "<RowAddUserNo>1</RowAddUserNo>";
					$chequeBondReceiptChequeBonds .= "<RowEditDateTime>{$currentTime}</RowEditDateTime>";
					$chequeBondReceiptChequeBonds .= "<RowEditUserNo>1</RowEditUserNo>";
					$chequeBondReceiptChequeBonds .= "<ReceiptID>{$i}</ReceiptID>";
					$chequeBondReceiptChequeBonds .= "<ID>{$k}</ID>";
					$chequeBondReceiptChequeBonds .= "<ChequeBondID>{$k}</ChequeBondID>";
					$chequeBondReceiptChequeBonds .= "<Time>{$time}</Time>";
					$chequeBondReceiptChequeBonds .= "<Expense>0</Expense>";
					$chequeBondReceiptChequeBonds .= "<Price>{$chequeDetail['price']}</Price>";
					$chequeBondReceiptChequeBonds .= "<NetPrice>{$chequeDetail['price']}</NetPrice>";
					$chequeBondReceiptChequeBonds .= "<CurrencyNo>{$chequeDetail['currencyNo']}</CurrencyNo>";
					$chequeBondReceiptChequeBonds .= "<CurrencyPrice>{$currencyPrice}</CurrencyPrice>";
					$chequeBondReceiptChequeBonds .= "<MaturityDayCount>{$maturityDayCount}</MaturityDayCount>";
					$chequeBondReceiptChequeBonds .= "<MaturityDate>{$maturityDate}</MaturityDate>";
					$chequeBondReceiptChequeBonds .= "<MaturityPrice>". $maturityDayCount * $chequeDetail['price'] ."</MaturityPrice>";
					$chequeBondReceiptChequeBonds .= "</ChequeBondReceiptChequesBonds>";

					$chequesBonds .= "<ChequesBonds>";
					$chequesBonds .= "<RowID>{$k}</RowID>";
					$chequesBonds .= "<RowAddDateTime>{$currentTime}</RowAddDateTime>";
					$chequesBonds .= "<RowAddUserNo>1</RowAddUserNo>";
					$chequesBonds .= "<RowEditDateTime>{$currentTime}</RowEditDateTime>";
					$chequesBonds .= "<RowEditUserNo>1</RowEditUserNo>";
					$chequesBonds .= "<ID>{$k}</ID>";
					$chequesBonds .= "<Type>{$chequeDetail['chequeOrBond']}</Type>";
					$chequesBonds .= "<PursuitNo>{$chequeDetail['pursuitNo']}</PursuitNo>";
					$chequesBonds .= "<SpecialCode>{$chequeDetail['specialCode']}</SpecialCode>";
					$chequesBonds .= "<ArrangeDate>{$currentTime}</ArrangeDate>";
					$chequesBonds .= "<MaturityDate>{$maturityDate}</MaturityDate>";
					$chequesBonds .= "<CurrencyNo>{$currencyNo}</CurrencyNo>";
					$chequesBonds .= "<CurrencyCode>{$currencyCode}</CurrencyCode>";
					$chequesBonds .= "<Price>{$chequeDetail['price']}</Price>";
					$chequesBonds .= "<Status>0</Status>";
					$chequesBonds .= "<EndorsementType>0</EndorsementType>";
					$chequesBonds .= "<Own>{$own}</Own>";
					$chequesBonds .= "<Number>{$chequeDetail['chequeNo']}</Number>";
					$chequesBonds .= "<DebtorName>{$chequeDetail['debtor']}</DebtorName>";
					$chequesBonds .= "<PaymentPlace>{$chequeDetail['paymentPlace']}</PaymentPlace>";
					if ($chequeDetail['bankNo'] > 0) {
						$chequesBonds .= "<BankCode>{$chequeDetail['bankNo']}</BankCode>";
						$chequesBonds .= "<BankName>{$chequeDetail['bank']}</BankName>";
						$chequesBonds .= "<BranchCode>{$chequeDetail['branchNo']}</BranchCode>";
						$chequesBonds .= "<BranchName>{$chequeDetail['branch']}</BranchName>";
					}
					$chequesBonds .= "<AccountNumber>{$chequeDetail['chequeAccountNumber']}</AccountNumber>";
					$chequesBonds .= "<ReceiptID>0</ReceiptID>";
					$chequesBonds .= "<DepotID>{$depotID}</DepotID>";
					$chequesBonds .= "<DepotName>{$depotName}</DepotName>";
					$chequesBonds .= "</ChequesBonds>";
				}
				// $zeroCount = 10 - strlen($i);
				// $receiptNo = str_repeat('0', $zeroCount) . $i;
				$receiptNo = $i;
				$averageMaturity = date('c', $totalMaturityDate/$count);
				$averageDay = (($totalMaturityDate/$count) - strtotime($time)) / (60 * 60 * 24);
				$averageDay = floor($averageDay);

				$chequeBondReceipts = "<ChequeBondReceipts>";
				$chequeBondReceipts .= "<RowID>{$i}</RowID>";
				$chequeBondReceipts .= "<RowAddDateTime>{$time}</RowAddDateTime>";
				$chequeBondReceipts .= "<RowAddUserNo>1</RowAddUserNo>";
				$chequeBondReceipts .= "<RowEditDateTime>{$time}</RowEditDateTime>";
				$chequeBondReceipts .= "<RowEditUserNo>1</RowEditUserNo>";
				$chequeBondReceipts .= "<ID>{$i}</ID>";
				$chequeBondReceipts .= "<ReceiptNo>{$receiptNo}</ReceiptNo>";
				$chequeBondReceipts .= "<ReceiptType>0</ReceiptType>";
				$chequeBondReceipts .= "<Time>{$time}</Time>";
				$chequeBondReceipts .= "<CurrencyNo>{$currencyNo}</CurrencyNo>";
				$chequeBondReceipts .= "<CurrencyCode>{$currencyCode}</CurrencyCode>";
				$chequeBondReceipts .= "<CurrencyPrice>{$currencyPrice}</CurrencyPrice>";
				$chequeBondReceipts .= "<DepotID>{$depotID}</DepotID>";
				$chequeBondReceipts .= "<DepotName>{$depotName}</DepotName>";
				$chequeBondReceipts .= "<AccountCode>{$accountCode}</AccountCode>";
				$chequeBondReceipts .= "<AccountName>{$accountName}</AccountName>";
				$chequeBondReceipts .= "<BalanceCurrencyNo>{$currentCurrencyNo}</BalanceCurrencyNo>";
				$chequeBondReceipts .= "<BalanceCurrencyCode>{$currentCurrencyCode}</BalanceCurrencyCode>";
				$chequeBondReceipts .= "<BalanceCurrencyPrice>{$currentCurrencyPrice}</BalanceCurrencyPrice>";
				$chequeBondReceipts .= "<CanSelectCurrentAccount>true</CanSelectCurrentAccount>";
				$chequeBondReceipts .= "<UpdateType>{$updateTypeCode}</UpdateType>";
				$chequeBondReceipts .= "<AverageDay>{$averageDay}</AverageDay>";
				$chequeBondReceipts .= "<AverageMaturity>{$averageMaturity}</AverageMaturity>";
				$chequeBondReceipts .= "<Remainder>0</Remainder>";
				$chequeBondReceipts .= "<Count>{$count}</Count>";
				$chequeBondReceipts .= "<TotalPrice>{$totalPrice}</TotalPrice>";
				$chequeBondReceipts .= "<TotalExpense>0</TotalExpense>";
				$chequeBondReceipts .= "<NetTotalPrice>{$totalPrice}</NetTotalPrice>";
				$chequeBondReceipts .= "<AccountNetTotalPrice>{$totalPrice}</AccountNetTotalPrice>";
				$chequeBondReceipts .= "<MaturityTotalPrice>{$maturityTotalPrice}</MaturityTotalPrice>";
				$chequeBondReceipts .= "<Explanation></Explanation>";
				$chequeBondReceipts .= "<SettingID>0</SettingID>";
				$chequeBondReceipts .= "<Status>1</Status>";
				$chequeBondReceipts .= "</ChequeBondReceipts>";

				$output = "<ChequeBondReceipts>\n";
				$output .= $chequeBondReceipts . $chequeBondReceiptChequeBonds . $chequesBonds;
				$output .= "</ChequeBondReceipts>\n";

				if (!is_dir('export/' . $path)) {
		            mkdir('export/' . $path, 0755, true);
		        }
		        $fileName = str_replace('/', '_', $path);
				$file = fopen('export/' . $path . $fileName . $accountCode . ".xml", "w");
				fwrite($file, $output);
				fclose($file);
			}
		}
	}
}