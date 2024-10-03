<!-- C:\Share\eInvoice\Import\TestEnv\ErrorFiles -->
<?php
require 'vendor/autoload.php';
include('../connection/dbcon.php');
include('api/checkState.php');
include ('FTP.php');
include ('generateToLocal.php');

use PhpOffice\PhpSpreadsheet\IOFactory;


// Define the invoice type
$invoiceType = 'SaleInvoice';

$templatePath = "template/{$invoiceType}_Template.csv";
$spreadsheet = IOFactory::load($templatePath);

// Get the active sheet
$sheet = $spreadsheet->getActiveSheet();



$sql = "

WITH RankedRecords AS (
  SELECT *,  
         ROW_NUMBER() OVER (PARTITION BY t_bpid ORDER BY t_exdt DESC) AS rank
  FROM erp.dbo.ttctax400800 b WITH (NOLOCK)

),
Address AS (
SELECT 
  D.t_bpid AS BP, 
b.t_fovn as TIN,
D.t_beid as BRN,
  D.t_nama AS RegName, 
  A.t_cadr AS AdCode, 
  A.t_dsca AS City, 
 A.t_ccty AS Country, 
  D.t_telp AS Tel, 
  t_ln02  AS Add1, 
  t_ln03 AS Add2, 
  t_ln04 + t_ln05 AS Add3 
,A.t_nama as OTname
FROM 
  erp.dbo.ttccom130800 A WITH (NOLOCK) 
  LEFT JOIN erp.dbo.ttccom100800 D WITH (NOLOCK) ON D.t_cadr = A.t_cadr 
left join RankedRecords b on D.t_bpid = b.t_bpid and rank =1
		where D.t_prst = 2
), OneTime AS(
		select 
		t_cadr as AdCode,t_ln01 as OTname, t_dsca  as City,t_ccty as Country,t_telp as Tel,
		       t_ln02  AS Add1
		,t_ln03  as Add2, 
		t_ln04+t_ln05  AS Add3
		from  erp.dbo.ttccom130800  WITH (NOLOCK) 
		
),
InvoiceInfo AS (
  SELECT 
    'SInvoice' AS DocumentType, 
    '800' + '/' + CONVERT(VARCHAR(50), A.t_tran) + '/' + CONVERT(VARCHAR(50), A.t_idoc) AS InvoiceID, 
    FORMAT(DATEADD(HOUR, 8, A.t_idat), 'MM/dd/yyyy HH:mm:ss') AS DocumentDate, 
    A.t_itbp AS CusCode, 
    	 	  case when A.t_itbp= 'OT0000001' then 'EI00000000010'
	   when CUS.Country != 'MYS' then 'EI00000000010'
	  else  CUS.TIN end AS TIN,
	    case when A.t_itbp= 'OT0000001' then 'NA' else  CUS.BRN end AS BRN,
       case when A.t_itbp= 'OT0000001' then OneTime.OTname else  CUS.RegName end AS CusRegName,
    case when A.t_itbp= 'OT0000001' then OneTime.Add1 else   CUS.Add1 end AS CusAddress1,
    case when A.t_itbp= 'OT0000001' then OneTime.Add2 else  CUS.Add2 end AS CusAddress2,
     case when A.t_itbp= 'OT0000001' then OneTime.Add3 else CUS.Add3 end AS CusAddress3,
    case when A.t_itbp= 'OT0000001' then OneTime.Country else CUS.Country end AS Country, 
	 case when A.t_itbp= 'OT0000001' then OneTime.City else CUS.City end AS City, 
	 	 case when A.t_itbp= 'OT0000001' then OneTime.Tel else CUS.Tel end AS Tel, 
    A.t_ccur AS Currency, 
    A.t_rate_1 AS CurrencyRate, 
    B.t_slsf AS SoRef, 
    C.t_corn as CusPo,
    B.t_shpf AS DoRef, 
    I.t_dsca AS Terms, 
    A.t_amti AS InvoiceTotalAmount, 
   '' AS PartNo, 
	case when ltrim(D.t_item) in ('Z05','Z07') then '027' else '022' end as Classification,
	--case when  A.t_tran = 'S02' then H.t_refb else G.t_dsca end as ItemDescription,
	case when  A.t_tran = 'S02' then J.t_desc else G.t_dsca end as ItemDescription,
	case when  A.t_tran = 'S01' then F.t_cuqs else J.t_cuni end as OrderUOM,
    D.t_dqua AS [InvoiceQty], 
	Case when F.t_cups = '/p2' then D.t_pric/100 else D.t_pric end AS UnitPrice, 
    '06' AS TaxType,  
    0 AS TaxRate, 
    0 AS TaxAmount, 
    0 AS TaxPrice, 
     D.t_amti AS ItemAmt, 
	   case when A.t_itbp= 'OT0000001' then OneTime.OTname else   Ship.RegName  end AS ShipReceiptName,
      case when A.t_itbp= 'OT0000001' then OneTime.Add1 else Ship.Add1 end AS ShipAddress1, 
     case when A.t_itbp= 'OT0000001' then OneTime.Add2 else Ship.Add2 end AS ShipAddress2, 
    case when A.t_itbp= 'OT0000001' then OneTime.Add3 else  Ship.Add3 end AS ShipAddress3, 
     case when A.t_itbp= 'OT0000001' then OneTime.Country else  Ship.Country end AS ShipCountry
    ,(DATEADD(HOUR, 8, A.t_idat)) as ddd
	--,LTRIM(RTRIM(SUBSTRING(CAST(tt.t_text AS CHAR(25)), 9, 12))) AS CustomForm1
	,D.t_amth_1 as AmtInMyr


  FROM 
    erp.dbo.tcisli305800 A WITH (NOLOCK) 

    LEFT JOIN erp.dbo.tcisli200800 B WITH (NOLOCK) ON B.t_brid = A.t_brid 
    LEFT JOIN erp.dbo.ttdsls400800 C WITH (NOLOCK) ON C.t_orno = B.t_slsf 
    LEFT JOIN erp.dbo.tcisli310800 D WITH (NOLOCK) ON D.t_idoc = A.t_idoc AND D.t_tran = A.t_tran 
    LEFT JOIN erp.dbo.ttdsls401800 F WITH (NOLOCK) ON F.t_orno = C.t_orno AND F.t_invn = A.t_idoc AND F.t_pono = D.t_pono 
    LEFT JOIN erp.dbo.ttcibd001800 G WITH (NOLOCK) ON G.t_item =  D.t_item 
    LEFT JOIN Address Ship ON Ship.BP = D.t_stbp
    LEFT JOIN Address CUS ON CUS.BP = D.t_ofbp
	--LEFT join erp.dbo.tcisli220800 H WITH (NOLOCK) on A.t_msid = H.t_msid
	left join    erp.dbo.tcisli225800 J WITH (NOLOCK) on A.t_msid = J.t_msid and J.t_msln = D.t_pono
	left  join erp.dbo.ttcmcs013800 I WITH (NOLOCK) on I.t_cpay = A.t_cpay 
	left join erp.dbo.twhinh430800 DO  WITH (NOLOCK) on DO.t_shpm = B.t_shpf
	--left join erp.dbo.ttttxt010800 tt with (NOLOCK) on tt.t_ctxt = DO.t_text
	left join OneTime on  OneTime.AdCode = A.t_itoa
      WHERE A.t_tran  IN ('S01', 'S02')	and A.t_stat = 6 
      and A.t_idoc not in (25000067)

	) 
  Select * from InvoiceInfo

" ;
  //   $sql .="
  //        WHERE ddd BETWEEN 
  //  DATEADD(HOUR, 16, CAST(CAST(GETDATE() AS DATE) AS DATETIME) - 4) 
  //  AND
  //   DATEADD(HOUR, 16, CAST(CAST(GETDATE() AS DATE) AS DATETIME) - 1) 
      
  //     ";

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
// ) 
// SELECT 
//   * 
// FROM 
//   InvoiceInfo
//     where DATEPART(YEAR, ddd) = 2024 AND DATEPART(MONTH, ddd) = 8
//     order by ddd desc
// ";



$result = sqlsrv_query($conn, $sql);
if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}
// echo $sqlsrv_query;die();
$invoiceIdCounts = array();
$rowNumber = 1;

while ($row = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC)) {
    // Increment row number
    $rowNumber++;
    // Track the occurrence of the InvoiceID
    // If invoice  more than 100 item, then add suffix

    // $invoiceId = $row['InvoiceID'];
    // if (!isset($invoiceIdCounts[$invoiceId])) {
    //     $invoiceIdCounts[$invoiceId] = 0;
    // }
    // $invoiceIdCounts[$invoiceId]++;
    
    // $count = $invoiceIdCounts[$invoiceId];
    // $suffix = '';
    // if ($count > 100) {
    //     // Calculate suffix dynamically
    //     $suffixIndex = ceil($count / 100) - 1;
    //     $suffix = '-' . $suffixIndex;
    // }
    // $modifiedInvoiceId = $invoiceId . $suffix;


    //convert all space or new line to \n  >>>get 1st row(before the 1st \n) > remove custom#: (first 8 character)
    // $customForm1 = $row['CustomForm1'];
    // $normalized = preg_replace("/\r\n|\r|\n/", "\n", $customForm1);
    // $strtokCustomForm1 = strtok($normalized, "\n");
    // // $aggregatedCustomForm1  = str_replace(["\r\n", "\r", "\n"], " ", $customForm1);
    // $aggregatedCustomForm1 = substr($strtokCustomForm1, 8);



    // Set values in specific CSV columns
    $sheet->setCellValue('A' . $rowNumber, $row['DocumentType']);
    $sheet->setCellValue('B' . $rowNumber, $row['InvoiceID']);
    $sheet->setCellValue('C' . $rowNumber, $row['DocumentDate']);
    $sheet->setCellValue('D' . $rowNumber, $row['CusCode']);
    if (empty($row['TIN'])) {
      die("Error: " . $row['CusCode'] . " (" . $row['CusRegName'] . ") has empty TIN");
  } else {
      $sheet->setCellValue('E' . $rowNumber, $row['TIN']);
  }
    $sheet->setCellValue('F' . $rowNumber, $row['BRN']);
    $sheet->setCellValue('K' . $rowNumber, $row['CusRegName']);
    $sheet->setCellValue('M' . $rowNumber, $row['CusAddress1']);
    $sheet->setCellValue('N' . $rowNumber, $row['CusAddress2']);
    $sheet->setCellValue('O' . $rowNumber, $row['CusAddress3']);
    $sheet->setCellValue('S' . $rowNumber, $row['Country']);
    $sheet->setCellValue('W' . $rowNumber, $row['Currency']);
    $sheet->setCellValue('X' . $rowNumber, $row['CurrencyRate']);
    $sheet->setCellValue('Y' . $rowNumber, $row['CusPo']);
    $sheet->setCellValue('Z' . $rowNumber, $row['SoRef']);
    $sheet->setCellValue('AA' . $rowNumber, $row['DoRef']);
    $sheet->setCellValue('AB' . $rowNumber, $row['Terms']);
    $sheet->setCellValue('AC' . $rowNumber, $row['InvoiceTotalAmount']);
    $sheet->setCellValue('AE' . $rowNumber, $row['PartNo']);
    $sheet->setCellValue('AF' . $rowNumber, $row['Classification']);
    $sheet->setCellValue('AG' . $rowNumber, $row['ItemDescription']);
    $sheet->setCellValue('AH' . $rowNumber, $row['InvoiceQty']);
    $sheet->setCellValue('AI' . $rowNumber, $row['OrderUOM']);
    $sheet->setCellValue('AJ' . $rowNumber, $row['UnitPrice']);
    $sheet->setCellValue('AM' . $rowNumber, $row['TaxType']);
    $sheet->setCellValue('AN' . $rowNumber, $row['TaxRate']);
    $sheet->setCellValue('AO' . $rowNumber, $row['TaxAmount']);
    $sheet->setCellValue('AP' . $rowNumber, $row['TaxPrice']);
    $sheet->setCellValue('AQ' . $rowNumber, $row['ItemAmt']);
    $sheet->setCellValue('AR' . $rowNumber, $row['ShipReceiptName']);
    $sheet->setCellValue('AS' . $rowNumber, $row['ShipAddress1']);
    $sheet->setCellValue('AT' . $rowNumber, $row['ShipAddress2']);
    $sheet->setCellValue('AU' . $rowNumber, $row['ShipAddress3']);
    $sheet->setCellValue('BB' . $rowNumber, $row['ShipCountry']);  
    // $sheet->setCellValue('BC' . $rowNumber, $aggregatedCustomForm1);    
    $sheet->setCellValue('BK' . $rowNumber, $row['AmtInMyr']);    

    $city = $row['City'];
    $stateCode = getStateCode($city);
    $sheet->setCellValue('Q' . $rowNumber, $row['City']); 
    $sheet->setCellValue('R' . $rowNumber, $stateCode);

    $tel = $row['Tel'];
    $cleanTel = empty($tel) ? '999' : standardizePhoneNumber($tel);
    $sheet->setCellValue('T' . $rowNumber, $cleanTel);
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


$outputFolder = 'Invoice';

if ($location ==='L'){
// Upload the file to FTP
 uploadToFtp($isLive,$memoryStream, $outputFolder);
}else{
// Generate to Local
generateToLocal($outputFolder,$spreadsheet);
}


?>
