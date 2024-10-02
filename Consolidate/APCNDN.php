<?php
require '../vendor/autoload.php';
include('../../connection/dbcon.php');
include('../api/checkState.php');
include ('../FTP.php');

use PhpOffice\PhpSpreadsheet\IOFactory;



$templatePath = "../template/CNDN_Template.csv";
$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();

$year = 2024;
$month = 5;

$formattedDate = (new DateTime("$year-$month-01"))
    ->modify('last day of this month')
    ->setTime(8, 0, 0)
    ->format('m/d/Y H:i:s');
 $escapedDate = addslashes($formattedDate);
// print_r( $escapedDate);
// die();



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
select distinct t_ifbp as BP   from erp.dbo.ttfacp200800 a WITH (NOLOCK) 

),InvoiceInfo AS( 
select 
 'APCNDN' AS DocumentType,
  case when t_ttyp = 'PDN' then 'Debit Note' else 'Credit Note' end as Type,
 BP.BP+'-'+FORMAT(t_docd, 'MMMM')+'-'+ case when t_ttyp = 'PDN' then 'Debit Note' else 'Credit Note'end AS InvoiceID

 ,'$escapedDate'  as DocumentDate,
BP.BP as  [VendorCode]
 , CASE WHEN a.t_ccur = 'MYR' THEN 1 ELSE e.t_rate END AS [CurrencyRate]
 	, 'MYR' as [Currency]
	,'COD'  as [Terms]
	
	 , CASE WHEN a.t_ccur = 'MYR' THEN a.t_amti ELSE e.t_rate*a.t_amti END as [InvoiceTotalAmount]
	 , CASE WHEN a.t_ccur = 'MYR' THEN a.t_amti ELSE e.t_rate*a.t_amti END AS  UnitPrice
	,1 as InvoiceQty
	 , CASE WHEN a.t_ccur = 'MYR' THEN a.t_amti ELSE e.t_rate*a.t_amti END AS  ItemAmt
,'800' + '/' + CONVERT(VARCHAR(50), a.t_ttyp) + '/' + CONVERT(VARCHAR(50), a.t_ninv) AS ItemDescription
 ,(DATEADD(HOUR, 8, a.t_docd)) as ddd
from BP
left join erp.dbo.ttfacp200800 a WITH (NOLOCK) on BP.BP = a.t_ifbp 
left join Address Vendor on Vendor.BP = a.t_ifbp
LEFT JOIN erp.dbo.ttcmcs008800 e WITH (NOLOCK) on e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and    e.t_stdt = (
	select max(e.t_stdt) 

from  erp.dbo.ttcmcs008800 e WITH (NOLOCK) where e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and e.t_stdt<= a.t_docd
    )
	 
where  t_tdoc = '' and Vendor.Country != 'MYS' and a.t_stap in(4,2) and  LEFT( a.t_orno, 1) != 'F' and  a.t_ttyp in ('PCC', 'PDN','PCN') 
)
select * from InvoiceInfo
where  DATEPART(YEAR,(ddd)) = ? AND DATEPART(month,(ddd)) = ?
and LEFT(VendorCode,2) != 'OT'

" ;



$params1 = array($year, $month);

$result = sqlsrv_query($conn, $sql,$params1 );
if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}

$invoiceIdCounts = array();
$rowNumber = 1;

while ($row = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC)) {
    // Increment row number
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

    // Set values in specific CSV columns
    $sheet->setCellValue('A' . $rowNumber, $row['DocumentType']);
    $sheet->setCellValue('B' . $rowNumber, $row['Type']);
    $sheet->setCellValue('c' . $rowNumber, $row['InvoiceID']);
    $sheet->setCellValue('D' . $rowNumber, $row['DocumentDate']);
    $sheet->setCellValue('E' . $rowNumber, $row['VendorCode']);
    $sheet->setCellValue('F' . $rowNumber, 'EI00000000030'); 
    $sheet->setCellValue('L' . $rowNumber,'NA');  
    $sheet->setCellValue('N' . $rowNumber, 'NA');  
    $sheet->setCellValue('O' . $rowNumber, 'NA');  
    $sheet->setCellValue('P' . $rowNumber, 'NA');  
    $sheet->setCellValue('T' . $rowNumber, 'NA');  
    $sheet->setCellValue('X' . $rowNumber, $row['Currency']);
    $sheet->setCellValue('Y' . $rowNumber, $row['CurrencyRate']);
    $sheet->setCellValue('AC' . $rowNumber, $row['Terms']);
    $sheet->setCellValue('AD' . $rowNumber, $row['InvoiceTotalAmount']);
    $sheet->setCellValue('AH' . $rowNumber, $row['ItemDescription']);
    $sheet->setCellValue('AI' . $rowNumber, $row['InvoiceQty']);
    $sheet->setCellValue('AJ' . $rowNumber, 'pcs');  
    $sheet->setCellValue('AK' . $rowNumber, $row['UnitPrice']);
    $sheet->setCellValue('AM' . $rowNumber, '06');
    $sheet->setCellValue('AN' . $rowNumber, '0');
    $sheet->setCellValue('AO' . $rowNumber, '0');
    $sheet->setCellValue('AP' . $rowNumber, '0');
    $sheet->setCellValue('AR' . $rowNumber, $row['ItemAmt']);
    $sheet->setCellValue('AS' . $rowNumber,'NA');  
    $sheet->setCellValue('AT' . $rowNumber, 'NA');  
    $sheet->setCellValue('AU' . $rowNumber,'NA');  
    $sheet->setCellValue('AV' . $rowNumber, 'NA');  
    $sheet->setCellValue('BC' . $rowNumber,'NA');  

        

    }


sqlsrv_free_stmt($result);
sqlsrv_close($conn);


$memoryStream = createMemoryStream($spreadsheet);

// Generate FTP file path
$outputFolder = 'APCNDN';
$dateTime = date('Y-m-d_H-i-s'); // Format: YYYY-MM-DD_HH-MM-SS
$ftpFilePath = $outputFolder . '/' . $outputFolder . '_' . $dateTime . '.csv';

// Upload the file to FTP
uploadToFtp($memoryStream, $ftpFilePath);





// $pathPrefix = "../APCNDN/APCNDN_";
// $localFilePath = $pathPrefix . date('YmdHis') . ".csv";
// $writer = IOFactory::createWriter($spreadsheet, 'Csv');
// $writer->setDelimiter('|');
// $writer->setEnclosure('');
// $writer->save($localFilePath);

// echo "CSV file has been saved to: " . $localFilePath;
?>
