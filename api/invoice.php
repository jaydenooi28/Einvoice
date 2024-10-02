<?php

include('../../connection/dbcon.php');
include ('checkState.php');
// include('login.php');

$invoice = 20004457;
$issueTime = date('H:i:s', time() - (9 * 3600)) . 'Z';
$issueDate = date('Y-m-d');


$query = "WITH Address AS (
    SELECT  
      A.t_nama as RegName,A.t_cadr as AdCode,t_dsca as City,t_ccty as Country,D.t_telp as Tel,
        t_ln01 + ', ' + t_ln02  + t_ln03  + t_ln04  + t_ln05 + ', ' + t_ln06 AS 'Address Location' 
    FROM 
	erp.dbo.ttccom130800 A WITH (NOLOCK) 
	LEFT JOIN erp.dbo.ttccom100800 D WITH (NOLOCK) ON D.t_cadr = A.t_cadr
),
InvoiceInfo AS (
    SELECT 
        '800' + ' / ' + CONVERT(VARCHAR(50), A.t_tran) + ' / ' + CONVERT(VARCHAR(50), A.t_idoc) AS InvoiceID,
        FORMAT(GETDATE(), 'yyyy-MM-dd') AS IssueDate,
        CONCAT(FORMAT(DATEADD(HOUR, -9, GETDATE()), 'HH:mm:ss'), 'Z') AS IssueTime,
        A.t_ccur AS Currency,
        C.t_corn AS CusPo,
		Sup.[Address Location] as SupplierAddress,
		Sup.City as SupplierCity,
		'02' as SupState,
		Sup.Country as SupCountry,
		Sup.RegName as SupName, 
		Sup.Tel as SupTel,


		'' as CusTin,
		'' as CusBrn,
		Cus.[Address Location]as CusAddress,
		Cus.City as CusCity,
		'17' as CusState,
		Cus.Country as CusCountry,
		Cus.RegName as CusName,
		Cus.Tel as CusTel,

		0 as TaxAmount,
		E.t_ccur as TaxCurrency,
		'06' as TaxCategoryID,

		D.t_pono as InvoiceLine,
		E.t_ccur as InvCurrency,

		D.t_amti as TaxExclusiveAmount,
		ROUND(F.t_pric, 2) as UnitPrice,
		C.t_ccur as UnitPriceCurrency,
		D.t_tbai as TaxInclusiveAmount,
		E.t_amti as TotalPayableAmount,
		G.t_dsca as ItemDescription	
    FROM 
        erp.dbo.tcisli305800 A WITH (NOLOCK) 
        LEFT JOIN erp.dbo.tcisli200800 B WITH (NOLOCK) ON B.t_brid = A.t_brid
        LEFT JOIN erp.dbo.ttdsls400800 C WITH (NOLOCK) ON C.t_orno = B.t_slsf
		LEFT JOIN Address Sup ON Sup.AdCode = 'LOC001087'
        LEFT JOIN Address Cus ON Cus.AdCode = C.t_stad
		LEFT JOIN erp.dbo.tcisli310800 D WITH (NOLOCK) ON D.t_idoc = A.t_idoc AND D.t_tran= A.t_tran
		LEFT JOIN erp.dbo.tcisli305800 E WITH (NOLOCK) ON E.t_idoc = A.t_idoc AND E.t_tran= A.t_tran
		LEFT JOIN erp.dbo.ttdsls401800 F WITH (NOLOCK) ON F.t_orno = C.t_orno and F.t_invn = A.t_idoc and F.t_pono=D.t_pono 
		LEFT JOIN erp.dbo.ttcibd001800 G WITH (NOLOCK) on G.t_item = F.t_item
		
    WHERE 
        A.t_idoc = ? AND A.t_tran = 'S01'
)
SELECT * FROM InvoiceInfo;

    " ;


$params= array($invoice);
$result = sqlsrv_query($conn, $query,$params);
if ($result === false) {
    die(print_r(sqlsrv_errors(), true));
}

// ECHO $query;

// if (sqlsrv_num_rows($result) === 0) {
// 	$myArray = 'No Data Found!';
// }  else {

// 	$myArray = [];

//     while ($rs = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC)) {
        
//         $myArray[] = [
//             "ID" => ["_" => $rs['InvoiceID']], 
//             "IssueDate" => [["_" => $issueDate]], 
//             "IssueTime" => [["_" => $issueTime]],
//             "InvoiceTypeCode" => [["_" => "01", "listVersionID" => "1.0"]],
//             "DocumentCurrencyCode" => [["_" => $rs['Currency']]],
//             "BillingReference" => [["AdditionalDocumentReference" => ["ID" => [["_" => $rs['CusPo']]]]]],
//             "AccountingSupplierParty" => [
//                 [
//                     "Party" => [
//                         [
//                             "IndustryClassificationCode" => [[
//                                 "_" => "25999",
//                                 "name" => "SMTT Manufacture electronic components and boards"
//                             ]]
//                         ],
//                         [
//                             "PartyIdentification" => [
//                                 [
//                                     "ID" => [["_" => "C5853946080", "schemeID" => "TIN"]]
//                                 ],
//                                 [
//                                     "ID" => [["_" => "199301024828", "schemeID" => "BRN"]]
//                                 ]
//                                 ],
//                                 "PostalAddress" => [
//                                     [
//                                         "CityName" => [
//                                             ["_" =>$rs['SupplierCity']]
//                                         ]
//                                     ],
//                                     [
//                                         "PostalZone" => [
//                                             ["_" => "08000"]
//                                         ]
//                                     ],
//                                     [
//                                         "CountrySubentityCode" => [
//                                             ["_" => $SupStateCode]
//                                         ]
//                                     ]
//                                 ],
                                
//                         ]
//                     ]
//                 ]
//             ]
//         ];
//     }
    
// }



// $jsonData = array(
//     "_D" => "urn:oasis:names:specification:ubl:schema:xsd:Invoice-2",
//     "_A" => "urn:oasis:names:specification:ubl:schema:xsd:CommonAggregateComponents-2",
//     "_B" => "urn:oasis:names:specification:ubl:schema:xsd:CommonBasicComponents-2",
//     "Invoice" => $myArray
	
// );




// echo json_encode($jsonData);

// // // Initialize cURL session
// $ch = curl_init($apiUrl);

// // Set cURL options
// curl_setopt($ch, CURLOPT_CUSTOMREQUEST, "POST");
// curl_setopt($ch, CURLOPT_POSTFIELDS, json_encode($jsonData));
// curl_setopt($ch, CURLOPT_RETURNTRANSFER, true);
// curl_setopt($ch, CURLOPT_HTTPHEADER, array(
//     'Content-Type: application/json',
//     'Content-Length: ' . strlen(json_encode($jsonData))
// ));

// // // Execute cURL session and get the response
// $response = curl_exec($ch);

// // // Check for cURL errors
// if (curl_errno($ch)) {
//     echo 'Curl error: ' . curl_error($ch);
// }

// // Close cURL session
// curl_close($ch);
// sqlsrv_close($conn);


// // // Display the API response
// // echo json_decode($response, true);
// // echo $responseData;
// // echo json_encode($response, true);

// // echo json_encode($response, JSON_UNESCAPED_UNICODE);
// $responseObject = json_decode($response);

// if ($responseObject !== null) {
//     $modifiedResponse = json_encode($responseObject, JSON_UNESCAPED_UNICODE);
//     echo $modifiedResponse;
// } else {
//     echo 'Error decoding JSON response';
// }
$data = sqlsrv_fetch_array($result, SQLSRV_FETCH_ASSOC);

// Convert the data to JSON
$json_data = json_encode($data);

// Output the JSON data
echo $json_data;

// Clean up resources
sqlsrv_free_stmt($result);
sqlsrv_close($conn);
?>