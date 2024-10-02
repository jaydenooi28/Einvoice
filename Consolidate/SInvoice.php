<?php
require '../vendor/autoload.php';
include('../../connection/dbcon.php');
include('../api/checkState.php');
include ('../FTP.php');

use PhpOffice\PhpSpreadsheet\IOFactory;
$invoiceType = 'SaleInvoice';
$templatePath = "../template/{$invoiceType}_Template.csv";
$spreadsheet = IOFactory::load($templatePath);
$sheet = $spreadsheet->getActiveSheet();

$year = 2024;
$month = 8;

$formattedDate = (new DateTime("$year-$month-01"))
    ->modify('last day of this month')
    ->setTime(8, 0, 0)
    ->format('m/d/Y H:i:s');
 $escapedDate = addslashes($formattedDate);


$sql = "

with BP as (
	select  distinct t_itbp AS BP, t_ccty as Country from     erp.dbo.tcisli305800 A WITH (NOLOCK) 
	left join erp.dbo.ttccom100800 B WITH (NOLOCK) on B.t_bpid = A.t_itbp
	left join erp.dbo.ttccom130800 C WITH (NOLOCK) on C.t_cadr = B.t_cadr

  ),InvoiceInfo AS( 
  SELECT 
  'SInvoice' AS DocumentType, 
   BP.BP+'-'+FORMAT(t_idat, 'MMMM') AS InvoiceID
,'$escapedDate'  as DocumentDate
      , A.t_itbp AS CusCode 
     	  , case when BP.Country != 'MYS' then 'EI00000000020' else '' end as TIN
      ,    'MYR' AS Currency
   , A.t_rate_1 AS CurrencyRate
     ,  I.t_dsca AS Terms
      ,A.t_amti*A.t_rate_1 AS InvoiceTotalAmount
	        ,A.t_amti*A.t_rate_1 AS UnitPrice
			,1 as [InvoiceQty]
   ,A.t_amti*A.t_rate_1 AS ItemAmt
       ,'800' + '/' + CONVERT(VARCHAR(50), A.t_tran) + '/' + CONVERT(VARCHAR(50), A.t_idoc) AS ItemDescription
   ,(DATEADD(HOUR, 8, A.t_idat)) as ddd

  
     FROM BP 
  inner join   erp.dbo.tcisli305800 A WITH (NOLOCK)  on A.t_itbp = BP.BP
  
  left  join erp.dbo.ttcmcs013800 I WITH (NOLOCK) on I.t_cpay = A.t_cpay 

  where A.t_tran  IN ('S01', 'S02')	and A.t_stat = 6
  )
  SELECT   *  FROM    InvoiceInfo 
  
    where  DATEPART(YEAR,(ddd)) = ? AND DATEPART(month,(ddd)) = ?
     --and LEFT(CusCode,2) != 'OT'

  ";

$params1 = array($year, $month);

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

    // Set values in specific CSV columns
    $sheet->setCellValue('A' . $rowNumber, $row['DocumentType']);
    $sheet->setCellValue('B' . $rowNumber, $row['InvoiceID']);
    $sheet->setCellValue('C' . $rowNumber, $row['DocumentDate']);
    $sheet->setCellValue('D' . $rowNumber, $row['CusCode']);
    $sheet->setCellValue('E' . $rowNumber, $row['TIN']);
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
    $sheet->setCellValue('AF' . $rowNumber,  '004');  
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

$memoryStream = createMemoryStream($spreadsheet);

// Generate FTP file path
// $outputFolder = 'Invoice';
// $dateTime = date('Y-m-d_H-i-s'); // Format: YYYY-MM-DD_HH-MM-SS
// $ftpFilePath = $outputFolder . '/' . $outputFolder . '_' . $dateTime . '.csv';

// // Upload the file to FTP
// uploadToFtp($memoryStream, $ftpFilePath);




$pathPrefix = "../Soutput/SalesInvoice_";
$localFilePath = $pathPrefix . date('YmdHis') . ".csv";
$writer = IOFactory::createWriter($spreadsheet, 'Csv');
$writer->setDelimiter('|');
$writer->setEnclosure('');
$writer->save($localFilePath);

echo "CSV file has been saved to: " . $localFilePath;
?>
