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
$dateTime = date('Y-m-d H-i-s'); 




// echo "Current Date and Time: " . $dateTime . "<br>";
// echo "Start Time: " . $startTime . "<br>";
// echo "End Time: " . $endTime . "<br>";
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
),OneTime AS(
		select 
		t_cadr as AdCode,t_nama as OTname, t_dsca  as City,t_ccty as Country,t_telp as Tel,
		       t_ln01 + ', ' AS Add1
		, t_ln02  + t_ln03  as Add2, 
		t_ln04+t_ln05 + t_ln06 AS Add3
		from  erp.dbo.ttccom130800  WITH (NOLOCK) 
		
),
		
InvoiceInfo AS (
select 
    'APCNDN' AS DocumentType , case when t_ttyp = 'PDN' then 'Debit Note' else 'Credit Note' end as Type
	,a.t_year as 'Fiscal Year',a.t_prod as 'Fiscal Period'
	,'800/'+t_ttyp +'/'+CONVERT(VARCHAR(50), a.t_ninv) as [InvoiceID]
	,FORMAT(DATEADD(HOUR, 8, a.t_docd), 'MM/dd/yyyy HH:mm:ss') AS [DocumentDate]

	,a.t_ifbp as [VendorCode]

	,  Vendor.RegName  as 'VendorRegName'
	, Vendor.Add1  as [VendorAddress1]
	, Vendor.Add2  as [VendorAddress2]
	,Vendor.Add3  as [VendorAddress3]
  , Vendor.[Country]  as [Country]
   ,Vendor.[City] as [City]
   ,Vendor.Tel as Tel
	, CASE WHEN a.t_ccur = 'MYR' THEN 1 ELSE e.t_rate END AS [CurrencyRate]
	, a.t_ccur as [Currency]
	, bb.t_orno  as [PoRef]
,bb.t_rcno as [RcRef]
  ,a.t_isup as [SupInv]
	,'COD'  as [Terms]
	,a.t_amnt as [InvoiceTotalAmount]
	--,ltrim(d.t_item) as [PartNo]
  ,'' as [PartNo]
	,case when c.t_dsca is null then CAST(a.t_text AS nvarchar(255)) else c.t_dsca end as [ItemDescription]
	--CAST(a.t_text AS nvarchar(255))
	,case  when ltrim(d.t_item) like 'Z%' then '036' else '034' end as Classification
	,case when bb.t_orno is null  then a.t_amnt else  bb.t_iqan end as [InvoiceQty]
	,case when bb.t_orno is null then 1 
		when bb.t_orno is not null and d.t_cupp = '/p2' then d.t_pric/100
	else d.t_pric   end as [UnitPrice]
	,06 as [TaxType]
	,0 as [TaxRate]
	,0 as [TaxAmount]
	,0 as [TaxPrice]
	,case when bb.t_orno is null  then a.t_amnt else bb.t_iamt end as [ItemAmt]

	,Ship.RegName as [ShipReceiptName]
	   ,Ship.Add1 AS ShipAddress1, 
    Ship.Add2 AS ShipAddress2, 
    Ship.Add3 AS ShipAddress3, 
    Ship.Country AS ShipCountry
      ,a.t_docd as ddd
	,case when  bb.t_orno is null then 'pcs' else d.t_cuqp end as OrderUOM
	--,cast(tt.t_text as char(100)) as EInvRefNo 
	,'einv test' as EInvRefNo 


from erp.dbo.ttfacp200800 a WITH (NOLOCK) 
left join erp.dbo.ttttxt010800 tt with (NOLOCK) on tt.t_ctxt = a.t_text
--left JOIN erp.dbo.ttfacp256800 b WITH (NOLOCK) on a.t_ninv = b.t_idoc and a.t_ttyp = b.t_ityp
left join erp.dbo.ttfacp251800 bb  WITH (NOLOCK) on bb.t_idoc =a.t_ninv and a.t_ttyp = bb.t_ityp and   bb.t_rseq =bb.t_rseq 

left JOIN erp.dbo.ttdpur401800 d  WITH (NOLOCK) on  d.t_orno = bb.t_orno and bb.t_pono=d.t_pono and bb.t_sqnb = d.t_sqnb 
LEFT JOIN erp.dbo.ttcibd001800 c WITH (NOLOCK) on c.t_item = d.t_item
LEFT JOIN erp.dbo.ttcmcs008800 e WITH (NOLOCK) on e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and    e.t_stdt = (
	select max(e.t_stdt) 
from  erp.dbo.ttcmcs008800 e WITH (NOLOCK) where e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and e.t_stdt<= a.t_docd
    )
left join Address Vendor on Vendor.BP = a.t_ifbp
left join Address Ship 	on Ship.BP = 'EG0000001'
left join OneTime on  OneTime.AdCode = d.t_sfad

where t_ttyp IN ('PCC', 'PDN','PCN') and t_tdoc = ''  and Vendor.Country != 'MYS' and  LEFT( a.t_orno, 1) != 'F' and a.t_stap in (2,4)


";

$sql .="
	 ) 
SELECT  *FROM  InvoiceInfo
 where [Fiscal Year] = 2024 and [Fiscal Period] = 9
	    order by ddd desc

";

// $sql .="
// ) 
// SELECT 
//   * 
// FROM 

//   InvoiceInfo
//     where DATEPART(YEAR, ddd) = 2024 AND DATEPART(MONTH, ddd) = 5 
//     order by ddd desc
// ";



$result = sqlsrv_query($conn, $sql);
// print_r($sql);

if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}

// Counter for the row in the CSV
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
    $sheet->setCellValue('B' . $rowNumber, $modifiedInvoiceId);
	$sheet->setCellValue('D' . $rowNumber, $row['DocumentDate']);
	$sheet->setCellValue('E' . $rowNumber, $row['VendorCode']);  
	$sheet->setCellValue('L' . $rowNumber, $row['VendorRegName']);
	$sheet->setCellValue('N' . $rowNumber, $row['VendorAddress1']);
	$sheet->setCellValue('O' . $rowNumber, $row['VendorAddress2']);
	$sheet->setCellValue('P' . $rowNumber, $row['VendorAddress3']);
	$sheet->setCellValue('T' . $rowNumber, $row['Country']);
	$sheet->setCellValue('X' . $rowNumber, $row['Currency']);
	$sheet->setCellValue('Z' . $rowNumber, $row['RcRef']);
	$sheet->setCellValue('Y' . $rowNumber, $row['CurrencyRate']);
	$sheet->setCellValue('AA' . $rowNumber, $row['PoRef']);
	$sheet->setCellValue('AB' . $rowNumber, $row['SupInv']);
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
$outputFolder = 'APCNDN';
$location = 'L';


if ($location ==='F'){
	// Upload the file to FTP
	 uploadToFtp($isLive,$memoryStream, $outputFolder);
	}else{
	// Generate to Local
	generateToLocal($outputFolder,$spreadsheet);
	}

?>
