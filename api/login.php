<?php

$curl = curl_init();

curl_setopt_array($curl, array(
  CURLOPT_URL => 'https://preprod-api.myinvois.hasil.gov.my/connect/token',
  CURLOPT_RETURNTRANSFER => true,
  CURLOPT_ENCODING => '',
  CURLOPT_MAXREDIRS => 10,
  CURLOPT_TIMEOUT => 0,
  CURLOPT_FOLLOWLOCATION => true,
  CURLOPT_HTTP_VERSION => CURL_HTTP_VERSION_1_1,
  CURLOPT_CUSTOMREQUEST => 'POST',
  CURLOPT_POSTFIELDS => 'client_id=be355bc0-544b-4360-aa48-9f256d332859&client_secret=6673dddd-a887-4a36-b42e-75df7718b5b9&grant_type=client_credentials&scope=InvoicingAPI',
  CURLOPT_HTTPHEADER => array(
    'Content-Type: application/x-www-form-urlencoded'
  ),
));

$response = curl_exec($curl);

// Close the curl connection
curl_close($curl);

// Decode the JSON response
$decodedResponse = json_decode($response, true);

// Check for errors during decoding
if (json_last_error() === JSON_ERROR_NONE) {
  // Extract the access_token
  $accessToken = $decodedResponse['access_token'];
  echo $decodedResponse;
  return $accessToken; // Or use $accessToken for further processing
} else {
  echo "Error: Failed to decode response: " . json_last_error_msg();
}