<?php
require 'vendor/autoload.php';
include('../connection/dbcon.php');
include('api/checkState.php');
include ('FTP.php');

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

// Define the invoice type
$invoiceType = 'PurchaseInvoice';

$templatePath = "template/{$invoiceType}_Template.csv";
$spreadsheet = IOFactory::load($templatePath);

// Get the active sheet
$sheet = $spreadsheet->getActiveSheet();

// Add a value (for example, adding 'Test Value' to cell A1)
// $sheet->setCellValue('A2', 'AAAA');
// $sheet->setCellValue('B2', 'BBB');
// $sheet->setCellValue('AC2', 'EEEE');
$invoiceID = 1; // Example invoice ID



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
		t_cadr as AdCode,t_nama as OTname, t_dsca  as City,t_ccty as Country,t_telp as Tel,
		       t_ln01 + ', ' AS Add1
		, t_ln02  + t_ln03  as Add2, 
		t_ln04+t_ln05 + t_ln06 AS Add3
		from  erp.dbo.ttccom130800  WITH (NOLOCK) 
		where t_ccty != 'MYS'
		
),
InvoiceInfo AS (  

	
		select 
     'PInvoice' AS DocumentType,
		 '800/'+t_ityp +'/'+CONVERT(VARCHAR(50), t_idoc) as [InvoiceID]
		--,FORMAT(DATEADD(HOUR, 8, a.t_docd), 'MM/dd/yyyy HH:mm:ss') AS [DocumentDate]
    ,FORMAT(getdate(), 'MM/dd/yyyy HH:mm:ss') AS DocumentDate
			,a.t_ifbp as [VendorCode]
				,  Vendor.RegName  as 'VendorRegName'
	, Vendor.Add1  as [VendorAddress1]
	, Vendor.Add2  as [VendorAddress2]
	, Vendor.Add3  as [VendorAddress3]
   ,Vendor.[City] as [City]
	, Vendor.Country  as [VendorCountry]
    ,Vendor.Tel as Tel
		, CASE WHEN a.t_ccur = 'MYR' THEN 1 ELSE e.t_rate END AS [CurrencyRate]
		, a.t_ccur as [Currency]

		,a.t_orno as PoRef
		,'' as RcRef
			,a.t_isup as [SupInv]
		,I.t_dsca  as [Terms]
			,a.t_amti as [InvoiceTotalAmount]
				,'' as [PartNo]
	,	'Freight Cost & etc' as [ItemDescription]
		,'036'as Classification
			
		,1 as [InvoiceQty]
		,k.t_amnt as [UnitPrice]
		,'pcs' as OrderUOM
	
	
			,'06' as [TaxType]
	,0 as [TaxRate]
	,0 as [TaxAmount]
	,0 as [TaxPrice]
	,k.t_amnt as [ItemAmt]
	,Ship.RegName as [ShipReceiptName]
	   ,Ship.Add1 AS ShipAddress1, 
    Ship.Add2 AS ShipAddress2, 
    Ship.Add3 AS ShipAddress3, 
    Ship.Country AS ShipCountry
      ,a.t_docd as ddd
from erp.dbo.ttfacp101800 k WITH (NOLOCK) 
inner join erp.dbo.ttfacp200800 a WITH (NOLOCK) on a.t_ninv = k.t_idoc and k.t_ityp = a.t_ttyp and k.t_ityp = 'PIN' and k.t_loco = 0
left join Address Vendor on Vendor.BP = a.t_ifbp
left join Address Ship 	on Ship.BP = 'EG0000001'
 left  join erp.dbo.ttcmcs013800 I WITH (NOLOCK) on I.t_cpay = a.t_cpay 
LEFT JOIN erp.dbo.ttcmcs008800 e WITH (NOLOCK) on e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and    e.t_stdt = (
	select max(e.t_stdt) 

from  erp.dbo.ttcmcs008800 e WITH (NOLOCK) where e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and e.t_stdt<= a.t_docd

    )
	where t_ttyp  ='PIN' and t_tdoc = '' and Vendor.Country != 'MYS' and a.t_stap = 4
and  LEFT( a.t_orno, 1) != 'F'


" ;


// $sql .="
// 	 WHERE ddd BETWEEN DATEADD(HOUR, 16, CAST(CAST(GETDATE() AS DATE) AS DATETIME) - 1) AND GETDATE();
// ";
$sql .="
) 
SELECT 
  * 
FROM 
  InvoiceInfo
  where DATEPART(YEAR, ddd) = 2024 AND DATEPART(MONTH, ddd) = 4 and LEFT(VendorCode,2) != 'OT'
    order by ddd desc
";

// $sql2 = "

	
// 		select 
// 		 '800/'+t_ityp +'/'+CONVERT(VARCHAR(50), t_idoc) as [InvoiceID],
	
// 		'Freight Cost & etc' as [ItemDescription],k.t_amnt as [ItemAmt]
// from erp.dbo.ttfacp101800 k WITH (NOLOCK) 
// inner join erp.dbo.ttfacp200800 a WITH (NOLOCK) on a.t_ninv = k.t_idoc and k.t_ityp = a.t_ttyp and k.t_ityp = 'PIN' and k.t_loco = 0
// 	where DATEPART(YEAR, a.t_docd) = 2024 AND DATEPART(MONTH, a.t_docd) = 6
// ";


$result = sqlsrv_query($conn, $sql);
if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}



// Counter for the row in the CSV
$rowNumber = 1;

while ($row = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC)) {
    // Increment row number
    $rowNumber++;

    // Set values in specific CSV columns
    $sheet->setCellValue('A' . $rowNumber, $row['DocumentType']);
    $sheet->setCellValue('B' . $rowNumber, $row['InvoiceID']);
    $sheet->setCellValue('C' . $rowNumber, $row['DocumentDate']);
    $sheet->setCellValue('D' . $rowNumber, $row['VendorCode']);
    $sheet->setCellValue('K' . $rowNumber, $row['VendorRegName']);
    $sheet->setCellValue('M' . $rowNumber, $row['VendorAddress1']);
    $sheet->setCellValue('N' . $rowNumber, $row['VendorAddress2']);
    $sheet->setCellValue('O' . $rowNumber, $row['VendorAddress3']);
    $sheet->setCellValue('S' . $rowNumber, $row['VendorCountry']);
    $sheet->setCellValue('W' . $rowNumber, $row['Currency']);
    $sheet->setCellValue('Y' . $rowNumber, $row['RcRef']);
    $sheet->setCellValue('X' . $rowNumber, $row['CurrencyRate']);
    $sheet->setCellValue('Z' . $rowNumber, $row['PoRef']);
    $sheet->setCellValue('AA' . $rowNumber, $row['SupInv']);
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

// Generate FTP file path
$outputFolder = 'PurchaseInvoice';
$invoiceType = 'PinAddCost';
$dateTime = date('Y-m-d_H-i-s'); // Format: YYYY-MM-DD_HH-MM-SS
$ftpFilePath = $outputFolder . '/' . $invoiceType . '_' . $dateTime . '.csv';

// Upload the file to FTP
uploadToFtp($memoryStream, $ftpFilePath);
?>
