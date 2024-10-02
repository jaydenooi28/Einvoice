<?php
require 'vendor/autoload.php';
include('../connection/dbcon.php');
include('api/checkState.php');
include ('FTP.php');
include ('generateToLocal.php');

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Csv;


$templatePath = "template/CNDN_Template.csv";
$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();



$sql = "
WITH Address AS (
  SELECT 
    D.t_bpid AS BP, 
    t_ln01 AS RegName, 
    A.t_cadr AS AdCode, 
    A.t_dsca AS City, 
    A.t_ccty AS Country, 
    D.t_telp AS Tel, 
    t_ln02 + ', ' AS Add1, 
    t_ln03 AS Add2, 
    t_ln04 + t_ln05  AS Add3 
	,A.t_nama as OTname
  FROM 
    erp.dbo.ttccom130800 A WITH (NOLOCK) 
    LEFT JOIN erp.dbo.ttccom100800 D WITH (NOLOCK) ON D.t_cadr = A.t_cadr 
	where D.t_prst = 2

), oneTime AS(
		select 
		t_cadr as AdCode,t_ln01 as OTname, t_dsca  as City,t_ccty as Country,t_telp AS Tel, 
		       t_ln02 + ', ' AS Add1
		, t_ln03  as Add2, 
		t_ln04+t_ln05  AS Add3
		from  erp.dbo.ttccom130800  WITH (NOLOCK) 
		
),
InvoiceInfo AS (
  SELECT 
  'ARCNDN' AS DocumentType,
  case when  A.t_tran IN ('SCN','S10') then 'Credit Note' else 'Debit Note' end as Type,
    '800/' + CONVERT(VARCHAR(50), A.t_tran) + '/' + CONVERT(VARCHAR(50), A.t_idoc) AS InvoiceID, 
      FORMAT(DATEADD(HOUR, 8, A.t_idat), 'MM/dd/yyyy HH:mm:ss') AS DocumentDate, 
    A.t_itbp AS CusCode, 
       case when A.t_itbp= 'OT0000001' then oneTime.OTname else  CUS.RegName end AS CusRegName,
    case when A.t_itbp= 'OT0000001' then oneTime.Add1 else   CUS.Add1 end AS CusAddress1,
    case when A.t_itbp= 'OT0000001' then oneTime.Add2 else  CUS.Add2 end AS CusAddress2,
     case when A.t_itbp= 'OT0000001' then oneTime.Add3 else CUS.Add3 end AS CusAddress3,
    case when A.t_itbp= 'OT0000001' then oneTime.Country else CUS.Country end AS Country, 
		 case when A.t_itbp= 'OT0000001' then oneTime.City else CUS.City end AS City, 
	 	 case when A.t_itbp= 'OT0000001' then oneTime.Tel else CUS.Tel end AS Tel, 
    A.t_ccur AS Currency, 
    A.t_rate_1 AS CurrencyRate, 
    F.t_orno  AS SoRef, 
    C.t_corn as CusPo,
    B.t_shpf AS DoRef, 
     case when A.t_itbp = 'MA0000007' then '60 Days' else  I.t_dsca end AS Terms,
    A.t_amti AS InvoiceTotalAmount,  
	 '' AS PartNo, 
		case when ltrim(D.t_item) in ('Z05','Z07') then '027' else '022' end as Classification,

	case when  A.t_tran in( 'SCN','SDN') then H.t_refa else G.t_dsca end as ItemDescription,
    D.t_dqua AS [InvoiceQty], 
    Case when F.t_cups = '/p2' then D.t_pric/100 else D.t_pric end AS UnitPrice, 
    06 AS TaxType, 
    0 AS TaxRate, 
    0 AS TaxAmount, 
    0 AS TaxPrice, 
    D.t_amti AS ItemAmt, 
    	D.t_amti *A.t_rate_1 as AmtInMyr,
	case when B.t_slsf ='' then J.t_cuni else D.t_cuqs end as OrderUOM,
	--case when A.t_tran = 'S10' ,
    
	   case when A.t_itbp= 'OT0000001' then oneTime.OTname else   Ship.RegName  end AS ShipReceiptName,
      case when A.t_itbp= 'OT0000001' then oneTime.Add1 else Ship.Add1 end AS ShipAddress1, 
     case when A.t_itbp= 'OT0000001' then oneTime.Add2 else Ship.Add2 end AS ShipAddress2, 
    case when A.t_itbp= 'OT0000001' then oneTime.Add3 else  Ship.Add3 end AS ShipAddress3, 
     case when A.t_itbp= 'OT0000001' then oneTime.Country else  Ship.Country end AS ShipCountry
      ,(DATEADD(HOUR, 8, A.t_idat)) as ddd
    ,case when A.t_tran = 'S10' then F.t_corp else H.t_msid end as EInvRefNo,B.t_slsf 
  FROM 
    erp.dbo.tcisli305800 A WITH (NOLOCK) 
    LEFT JOIN erp.dbo.tcisli200800 B WITH (NOLOCK) ON B.t_brid = A.t_brid 
    LEFT JOIN erp.dbo.ttdsls400800 C WITH (NOLOCK) ON C.t_orno = B.t_slsf 
    LEFT JOIN erp.dbo.tcisli310800 D WITH (NOLOCK) ON D.t_idoc = A.t_idoc AND D.t_tran = A.t_tran 
    LEFT JOIN erp.dbo.ttdsls401800 F WITH (NOLOCK) ON F.t_orno = C.t_orno AND F.t_invn = A.t_idoc AND F.t_pono = D.t_pono 
    LEFT JOIN erp.dbo.ttcibd001800 G WITH (NOLOCK) ON G.t_item = F.t_item 
    LEFT JOIN Address Ship ON Ship.BP = D.t_stbp
    LEFT JOIN Address CUS ON CUS.BP = D.t_ofbp
	
    LEFT join erp.dbo.tcisli220800 H WITH (NOLOCK) on A.t_msid = H.t_msid
	LEFT join erp.dbo.tcisli225800 J WITH (NOLOCK) on A.t_msid = J.t_msid and J.t_msln = D.t_pono 
	left join oneTime on  oneTime.AdCode = A.t_itoa
	left  join erp.dbo.ttcmcs013800 I WITH (NOLOCK) on I.t_cpay = A.t_cpay 
    WHERE A.t_tran IN ('SCN', 'SDN','S10') 
	--and DATEPART(YEAR,(DATEADD(HOUR, 8, A.t_idat))) = 2024 AND DATEPART(MONTH,(DATEADD(HOUR, 8, A.t_idat))) = 8


		) 
SELECT 
* 
FROM 
  InvoiceInfo

" ;


// echo date('l') ;
if (date('l') === 'Monday') {

  $sql .="
  WHERE ddd BETWEEN DATEADD(HOUR, 16, CAST(CAST(GETDATE() AS DATE) AS DATETIME) - 3) AND GETDATE();
";
} else {
$sql .="
WHERE ddd BETWEEN DATEADD(HOUR, 16, CAST(CAST(GETDATE() AS DATE) AS DATETIME) - 1) AND GETDATE();
";
}



// $sql .="
// 	 WHERE ddd BETWEEN DATEADD(HOUR, 16, CAST(CAST(GETDATE() AS DATE) AS DATETIME) - 1) AND GETDATE();
// ";


$result = sqlsrv_query($conn, $sql);
if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}
$invoiceIdCounts = array();
$rowNumber = 1;
while ($row = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC)) {
    // Increment row number
    $rowNumber++;
      // Track the occurrence of the InvoiceID
      $invoiceId = $row['InvoiceID'];
      if (!isset($invoiceIdCounts[$invoiceId])) {
          $invoiceIdCounts[$invoiceId] = 0;
      }
      $invoiceIdCounts[$invoiceId]++;
      
      $count = $invoiceIdCounts[$invoiceId];
      $suffix = '';
      if ($count > 100) {
          // Calculate suffix dynamically
          $suffixIndex = ceil($count / 100) - 1;
          $suffix = '-' . $suffixIndex;
      }
  
      // Append suffix to InvoiceID if needed
      $modifiedInvoiceId = $invoiceId . $suffix;
    // Set values in specific CSV columns
    $sheet->setCellValue('A' . $rowNumber, $row['DocumentType']);
$sheet->setCellValue('B' . $rowNumber, $row['Type']);
$sheet->setCellValue('C' . $rowNumber, $row['InvoiceID']);  // New column
$sheet->setCellValue('D' . $rowNumber, $row['DocumentDate']);
$sheet->setCellValue('E' . $rowNumber, $row['CusCode']);
$sheet->setCellValue('L' . $rowNumber, $row['CusRegName']);
$sheet->setCellValue('N' . $rowNumber, $row['CusAddress1']);
$sheet->setCellValue('O' . $rowNumber, $row['CusAddress2']);
$sheet->setCellValue('P' . $rowNumber, $row['CusAddress3']);
$sheet->setCellValue('T' . $rowNumber, $row['Country']);
$sheet->setCellValue('X' . $rowNumber, $row['Currency']);
$sheet->setCellValue('Y' . $rowNumber, $row['CurrencyRate']);
$sheet->setCellValue('Z' . $rowNumber, $row['CusPo']);
$sheet->setCellValue('AA' . $rowNumber, $row['SoRef']);
$sheet->setCellValue('AB' . $rowNumber, $row['DoRef']);
$sheet->setCellValue('AC' . $rowNumber, $row['Terms']);
$sheet->setCellValue('AD' . $rowNumber, $row['InvoiceTotalAmount']);
$sheet->setCellValue('AF' . $rowNumber, $row['PartNo']);
$sheet->setCellValue('AG' . $rowNumber, $row['Classification']);
$sheet->setCellValue('AH' . $rowNumber, $row['ItemDescription']);
$sheet->setCellValue('AI' . $rowNumber, $row['InvoiceQty']);
$sheet->setCellValue('AJ' . $rowNumber, $row['OrderUOM']);
$sheet->setCellValue('AK' . $rowNumber, $row['UnitPrice']);
$sheet->setCellValue('AN' . $rowNumber, $row['TaxType']);
$sheet->setCellValue('AO' . $rowNumber, $row['TaxRate']);
$sheet->setCellValue('AP' . $rowNumber, $row['TaxAmount']);
$sheet->setCellValue('AQ' . $rowNumber, $row['TaxPrice']);
$sheet->setCellValue('AR' . $rowNumber, $row['ItemAmt']);
$sheet->setCellValue('AS' . $rowNumber, $row['ShipReceiptName']);
$sheet->setCellValue('AT' . $rowNumber, $row['ShipAddress1']);
$sheet->setCellValue('AU' . $rowNumber, $row['ShipAddress2']);
$sheet->setCellValue('AV' . $rowNumber, $row['ShipAddress3']);
$sheet->setCellValue('BC' . $rowNumber, $row['ShipCountry']);
$sheet->setCellValue('BL' . $rowNumber, $row['EInvRefNo']);
$sheet->setCellValue('BM' . $rowNumber, $row['AmtInMyr']); 
    
$city = $row['City'];
$stateCode = getStateCode($city);
$sheet->setCellValue('R' . $rowNumber, $row['City']); 
$sheet->setCellValue('S' . $rowNumber, $stateCode);

$tel = $row['Tel'];
$cleanTel = empty($tel) ? '999' : standardizePhoneNumber($tel);
$sheet->setCellValue('U' . $rowNumber, $cleanTel);
}


sqlsrv_free_stmt($result);
sqlsrv_close($conn);


$memoryStream = createMemoryStream($spreadsheet);
// Live = 1
// Test = 0
$isLive = 0;

// Local = L
// FTP = F
$location = 'F';


$outputFolder = 'ARCNDN';

if ($location ==='F'){
// Upload the file to FTP
 uploadToFtp($isLive,$memoryStream, $outputFolder);
}else{
// Generate to Local
generateToLocal($outputFolder,$spreadsheet);
}




?>
