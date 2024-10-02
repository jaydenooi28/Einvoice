<?php
require '../vendor/autoload.php';
include('../../connection/dbcon.php');
include('../api/checkState.php');
include ('../FTP.php');

use PhpOffice\PhpSpreadsheet\IOFactory;


// Define the invoice type
$invoiceType = 'PurchaseInvoice';

$templatePath = "../template/{$invoiceType}_Template.csv";
$spreadsheet = IOFactory::load($templatePath);

$sheet = $spreadsheet->getActiveSheet();

$financialYear = 2024;
$financialMonth = 6;


// $financialYear = '';
// $financialMonth = '';
// $getDate = getdate(); 

// $month = $getDate['mon']; 
// $year = $getDate['year']; 

// if ($month >= 7) {
//     $financialYear = $year + 1;
// } else {
//     $financialYear = $year;
// }
// $financialMonth = date("F", mktime(0, 0, 0, $month, 10));




$sql = "
	WITH Address AS (
    SELECT 
      D.t_bpid as BP,
	  D.t_nama as RegName,A.t_cadr as AdCode,t_dsca as City,t_ccty as Country,D.t_telp as Tel, A.t_nama as OTname,
        t_ln01 + ', ' AS Add1
		, t_ln02  + t_ln03  as Add2, 
		t_ln04+t_ln05 + t_ln06 AS Add3
		
    FROM 
	erp.dbo.ttccom130800 A WITH (NOLOCK) 
	LEFT JOIN erp.dbo.ttccom100800 D WITH (NOLOCK) ON D.t_cadr = A.t_cadr 	where D.t_prst = 2
), 
OneTime AS(
		select 
		t_cadr as AdCode,t_nama as OTname, t_dsca  as City,t_ccty as Country,
		       t_ln01 + ', ' AS Add1
		, t_ln02  + t_ln03  as Add2, 
		t_ln04+t_ln05 + t_ln06 AS Add3
		from  erp.dbo.ttccom130800  WITH (NOLOCK) 
		where t_ccty != 'MYS'
		
),
 BP as (
select distinct t_ifbp as BP ,B.t_nama as BPName  from erp.dbo.ttfacp200800 A WITH (NOLOCK) 
	left join erp.dbo.ttccom100800 B WITH (NOLOCK) on B.t_bpid = A.t_ifbp
	

),InvoiceInfo AS( 
select 
 'PInvoice' AS DocumentType,
 BP.BP+'-'+FORMAT(t_docd, 'MMMM')+'-1' AS InvoiceID

	,FORMAT(DATEADD(HOUR, 8, a.t_docd), 'MM/dd/yyyy HH:mm:ss') AS [DocumentDate],

BP.BP as  [VendorCode],BP.BPName
 , CASE WHEN a.t_ccur = 'MYR' THEN 1 ELSE e.t_rate END AS [CurrencyRate]
 	, a.t_ccur as [Currency]
	,I.t_dsca  as [Terms]
	,a.t_amth_1 as [InvoiceTotalAmount]
	,a.t_amth_1 as UnitPrice
	,1 as InvoiceQty
	,a.t_amth_1 as ItemAmt
,'800' + '/' + CONVERT(VARCHAR(50), a.t_ttyp) + '/' + CONVERT(VARCHAR(50), a.t_ninv) AS ItemDescription
 ,(DATEADD(HOUR, 8, a.t_docd)) as ddd,b.t_tedt as lol
from BP
left join erp.dbo.ttfacp200800 a WITH (NOLOCK) on BP.BP = a.t_ifbp 
left join Address Vendor on Vendor.BP = a.t_ifbp
LEFT JOIN erp.dbo.ttcmcs008800 e WITH (NOLOCK) on e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and    e.t_stdt = (
	select max(e.t_stdt) 

from  erp.dbo.ttcmcs008800 e WITH (NOLOCK) where e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and e.t_stdt<= a.t_docd
    )
	 left  join erp.dbo.ttcmcs013800 I WITH (NOLOCK) on I.t_cpay = a.t_cpay 
	 join erp.dbo.ttfgld100800 b on a.t_btno = b.t_btno and a.t_year =b.t_year
where  t_tdoc = '' and Vendor.Country != 'MYS' and a.t_stap in(4,2) and  LEFT( a.t_orno, 1) != 'F' 
and  a.t_ttyp in ('PIN','PIC')
 and a.t_year = 2025 and t_fprd = 2

)
select * from InvoiceInfo


where  DATEPART(YEAR,(lol)) = 2024 AND DATEPART(month,(lol)) = 8

";



$params1 = array($financialYear, $financialMonth);

$result = sqlsrv_query($conn, $sql,$params1 );
if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}

$invoiceIdCounts = array();
$rowNumber = 1;

while ($row = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC)) {
   
    $rowNumber++;

    // Track the occurrence of the InvoiceID
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

    // // Append suffix to InvoiceID if needed
    // $modifiedInvoiceId = $invoiceId . $suffix;



    $sheet->setCellValue('A' . $rowNumber, $row['DocumentType']);
    $sheet->setCellValue('B' . $rowNumber, $row['InvoiceID']);
    $sheet->setCellValue('C' . $rowNumber, $row['DocumentDate']);
    $sheet->setCellValue('D' . $rowNumber, $row['VendorCode']);
    $sheet->setCellValue('E' . $rowNumber, 'EI00000000030');  
    $sheet->setCellValue('M' . $rowNumber, 'NA');  
    $sheet->setCellValue('N' . $rowNumber, 'NA');  
    $sheet->setCellValue('O' . $rowNumber, 'NA');  
    $sheet->setCellValue('P' . $rowNumber, 'NA');  
    $sheet->setCellValue('Q' . $rowNumber, 'NA');  
    $sheet->setCellValue('R' . $rowNumber, 'NA');  
    $sheet->setCellValue('S' . $rowNumber, 'NA');  
    $sheet->setCellValue('W' . $rowNumber, $row['Currency']);
    $sheet->setCellValue('X' . $rowNumber, $row['CurrencyRate']);
    $sheet->setCellValue('AB' . $rowNumber, $row['Terms']);
    $sheet->setCellValue('AC' . $rowNumber, $row['InvoiceTotalAmount']);
    $sheet->setCellValue('AF' . $rowNumber,  '034');  
    $sheet->setCellValue('AG' . $rowNumber, $row['ItemDescription']);
    $sheet->setCellValue('AH' . $rowNumber,  '1');  
    $sheet->setCellValue('AI' . $rowNumber,  'pcs');  
    $sheet->setCellValue('AJ' . $rowNumber, $row['UnitPrice']);
    $sheet->setCellValue('AM' . $rowNumber, '06');
    $sheet->setCellValue('AN' . $rowNumber, '0');
    $sheet->setCellValue('AO' . $rowNumber, '0');
    $sheet->setCellValue('AP' . $rowNumber, '0');
    $sheet->setCellValue('AQ' . $rowNumber, $row['ItemAmt']);
    $sheet->setCellValue('AR' . $rowNumber,  'NA');  
    $sheet->setCellValue('AS' . $rowNumber,  'NA');  
    $sheet->setCellValue('AT' . $rowNumber, 'NA');  
    $sheet->setCellValue('AU' . $rowNumber,  'NA');  
    $sheet->setCellValue('BB' . $rowNumber, 'NA');  

}


sqlsrv_free_stmt($result);
sqlsrv_close($conn);

// $memoryStream = createMemoryStream($spreadsheet);

// // Generate FTP file path
// $outputFolder = 'PurchaseInvoice';
// $dateTime = date('Y-m-d_H-i-s'); // Format: YYYY-MM-DD_HH-MM-SS
// $ftpFilePath = $outputFolder . '/' . $outputFolder . '_' . $dateTime . '.csv';

// // Upload the file to FTP
// uploadToFtp($memoryStream, $ftpFilePath);


$pathPrefix = "../PurchaseInvoice/PurchaseInvoice_";
$localFilePath = $pathPrefix . date('YmdHis') . ".csv";
$writer = IOFactory::createWriter($spreadsheet, 'Csv');
$writer->setDelimiter('|');
$writer->setEnclosure('');
$writer->save($localFilePath);

echo "CSV file has been saved to: " . $localFilePath;
?>
