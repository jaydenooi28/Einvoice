<?php
require 'vendor/autoload.php';
include('../connection/dbcon.php');
include('api/checkState.php');
include ('FTP.php');
include ('generateToLocal.php');

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

// Define the invoice type
$invoiceType = 'PurchaseInvoice';

$templatePath = "template/{$invoiceType}_Template.csv";
$spreadsheet = IOFactory::load($templatePath);

// Get the active sheet
$sheet = $spreadsheet->getActiveSheet();



$sql = "
WITH Address AS (
    SELECT 
        D.t_bpid AS BP,
        D.t_nama AS RegName,
        A.t_cadr AS AdCode,
        t_dsca AS City,
        t_ccty AS Country,
        case when D.t_telp = '' then '999' else D.t_telp end AS Tel,
        A.t_nama AS OTname,
        t_ln01 + ', ' AS Add1,
        t_ln02 + t_ln03 AS Add2,
        t_ln04 + t_ln05 + t_ln06 AS Add3
    FROM 
        erp.dbo.ttccom130800 A WITH (NOLOCK)
    LEFT JOIN erp.dbo.ttccom100800 D WITH (NOLOCK) 
        ON D.t_cadr = A.t_cadr
    WHERE 
        D.t_prst = 2
), 

OneTime AS (
    SELECT 
        t_cadr AS AdCode,
        t_nama AS OTname, 
        t_dsca AS City,
        t_ccty AS Country,
        t_telp AS Tel,
        t_ln01 + ', ' AS Add1,
        t_ln02 + t_ln03 AS Add2,
        t_ln04 + t_ln05 + t_ln06 AS Add3
    FROM  
        erp.dbo.ttccom130800 WITH (NOLOCK)
    WHERE 
        t_ccty != 'MYS'
),

InvoiceInfo AS (  
select 
 'PInvoice' AS DocumentType, a.t_year as 'Fiscal Year',a.t_prod as 'Fiscal Period',
	,'800/'+t_ttyp +'/'+CONVERT(VARCHAR(50), a.t_ninv) as [InvoiceID]
	,FORMAT(DATEADD(HOUR, 8, a.t_docd), 'MM/dd/yyyy HH:mm:ss') AS [DocumentDate]
	,a.t_ifbp as [VendorCode]
	,case when a.t_ifbp = 'OT0000002' then OneTime.OTname else  Vendor.RegName end as 'VendorRegName'
	,case when a.t_ifbp = 'OT0000002' then OneTime.Add1 else Vendor.Add1 end as [VendorAddress1]
	,case when a.t_ifbp = 'OT0000002' then OneTime.Add2 else Vendor.Add2 end as [VendorAddress2]
	,case when a.t_ifbp = 'OT0000002' then OneTime.Add3 else Vendor.Add3 end as [VendorAddress3]
	,case when a.t_ifbp = 'OT0000002' then OneTime.Country else Vendor.Country end as [VendorCountry]
	,case when a.t_ifbp = 'OT0000002' then OneTime.City else Vendor.City end as [City]
	,case when a.t_ifbp = 'OT0000002' then OneTime.Tel else Vendor.Tel end as [Tel]

	, CASE WHEN a.t_ccur = 'MYR' THEN 1 ELSE e.t_rate END AS [CurrencyRate]
	, a.t_ccur as [Currency]
	, b.t_orno  as [PoRef]
  ,b.t_rcno as [RcRef]
  ,a.t_isup as [SupInv]
	,I.t_dsca  as [Terms]
	,a.t_amti as [InvoiceTotalAmount]
	,'' as [PartNo]
	,c.t_dsca as [ItemDescription]
	,case  when ltrim(d.t_item) like 'Z%' then '036' else '034' end as Classification
	,b.t_iqan  as [InvoiceQty]
	,case 
		when LEFT(ltrim(d.t_item),1) = 'Z' then b.t_iamt/b.t_iqan 
		when d.t_cupp = c.t_cuni then d.t_pric
		when t_conv<1 then d.t_pric*t_conv
	
		when LEFT(b.t_orno, 1) = 'V' then  b.t_iamt/b.t_iqan 
	
	else d.t_pric/J.t_conv end as [UnitPrice]

	,06 as [TaxType]
	,0 as [TaxRate]
	,0 as [TaxAmount]
	,0 as [TaxPrice]
	,b.t_iamt as [ItemAmt]
	,Ship.RegName as [ShipReceiptName]
	   ,Ship.Add1 AS ShipAddress1, 
    Ship.Add2 AS ShipAddress2, 
    Ship.Add3 AS ShipAddress3, 
    Ship.Country AS ShipCountry
      ,(DATEADD(HOUR, 8, a.t_docd)) as ddd
	  , c.t_cuni as OrderUOM


from erp.dbo.ttfacp200800 a WITH (NOLOCK) 
left JOIN erp.dbo.ttfacp251800 b WITH (NOLOCK) on a.t_ninv = b.t_idoc and a.t_ttyp = b.t_ityp
LEFT JOIN erp.dbo.ttdpur401800 d WITH (NOLOCK)
ON  
  (
    (LEFT(b.t_orno, 1) = 'V' AND d.t_orno = b.t_orno AND b.t_pono = d.t_pono)
    OR
    (LEFT(b.t_orno, 1) != 'V' AND d.t_orno = b.t_orno AND b.t_pono = d.t_pono AND b.t_sqnb = d.t_sqnb --AND b.t_rseq = d.t_rseq 
	)
  )

LEFT JOIN erp.dbo.ttcibd001800 c WITH (NOLOCK) on c.t_item = d.t_item
LEFT JOIN erp.dbo.ttcmcs008800 e WITH (NOLOCK) on e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and    e.t_stdt = (
	select max(e.t_stdt) 
from  erp.dbo.ttcmcs008800 e WITH (NOLOCK) where e.t_bcur = 'MYR' and e.t_ccur = a.t_ccur and e.t_stdt<= a.t_docd
    )
left join Address Vendor on Vendor.BP = a.t_ifbp
left join OneTime on  OneTime.AdCode = d.t_sfad
left join Address Ship 	on Ship.BP = 'EG0000001'
 left  join erp.dbo.ttcmcs013800 I WITH (NOLOCK) on I.t_cpay = a.t_cpay 
LEFT JOIN erp.dbo.ttcibd003800 J WITH (NOLOCK)
ON 
    (J.t_citg = '' AND 
    (
        (c.t_item = J.t_item AND c.t_cuni = J.t_unit AND d.t_cupp = J.t_basu)
        OR 
        (c.t_item = J.t_item AND d.t_cupp = J.t_unit AND c.t_cuni = J.t_basu)
        OR
        (J.t_item = '' AND c.t_cuni = J.t_unit AND d.t_cupp = J.t_basu AND c.t_item <> J.t_item)
        OR 
        (J.t_item = '' AND d.t_cupp = J.t_unit AND c.t_cuni = J.t_basu AND c.t_item <> J.t_item)
    ))
where t_ttyp  ='PIN' and t_tdoc = '' and Vendor.Country != 'MYS' and a.t_stap = 4
and  LEFT( b.t_orno, 1) != 'F'
),

AdditionalCost AS (

		select 
     'PInvoice' AS DocumentType, a.t_year as 'Fiscal Year',a.t_prod as 'Fiscal Period',
		 '800/'+t_ityp +'/'+CONVERT(VARCHAR(50), t_idoc) as [InvoiceID]
		,FORMAT(DATEADD(HOUR, 8, a.t_docd), 'MM/dd/yyyy HH:mm:ss') AS [DocumentDate]
			,a.t_ifbp as [VendorCode]
				,  Vendor.RegName  as 'VendorRegName'
	, Vendor.Add1  as [VendorAddress1]
	, Vendor.Add2  as [VendorAddress2]
	, Vendor.Add3  as [VendorAddress3]
	, Vendor.Country  as [VendorCountry]
		,Vendor.City as City
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
      ,(DATEADD(HOUR, 8, a.t_docd)) as ddd,'pcs' as OrderUOM
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
)
" ;


$sql .="
	 SELECT * 
FROM InvoiceInfo
where [Fiscal Year] = 2024 and [Fiscal Period] = 9
UNION ALL
SELECT * 
FROM AdditionalCost
where [Fiscal Year] = 2024 and [Fiscal Period] = 9
";


// $sql .="
// SELECT * 
// FROM InvoiceInfo
// where DATEPART(YEAR, ddd) = 2024 AND DATEPART(MONTH, ddd) = 8

// UNION ALL
// SELECT * 
// FROM AdditionalCost
// where DATEPART(YEAR, ddd) = 2024 AND DATEPART(MONTH, ddd) = 8
// ";

// $sql2 = "

	


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
// Live = 1
// Test = 0

$isLive = 0;

$location = 'L';


$outputFolder = 'PurchaseInvoice';

if ($location ==='F'){
// Upload the file to FTP
 uploadToFtp($isLive,$memoryStream, $outputFolder);
}else{
// Generate to Local
generateToLocal($outputFolder,$spreadsheet);
}
?>
