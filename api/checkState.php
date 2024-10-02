<?php
// Function to find the code for a given state using wildcard matching
function standardizePhoneNumber($phoneNumber) {
    // Remove any non-numeric characters (including spaces, dashes, and plus signs)
    $cleanedNumber = preg_replace('/[^0-9]/', '', $phoneNumber);
    return $cleanedNumber;
}

function normalizeStateName($stateName) {
    // Remove all spaces from the state name
    return preg_replace('/\s+/', '', $stateName);
};
function getStateCode($stateName) {
    // Include the JSON file
    $jsonData = file_get_contents(__DIR__ . '/json/state.json');
    
    // Decode the JSON data into a PHP array
    $stateCodes = json_decode($jsonData, true);

    
       // Normalize the input state name

       $normalizedStateName = normalizeStateName($stateName);


    switch (true) {
        case stripos($normalizedStateName, normalizeStateName('Sungai Petani')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Alor Setar')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Kulim')) !== false:
            $stateName = 'Kedah';
            break;
        case stripos($normalizedStateName, normalizeStateName('Senai')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Senai  Industrial Estate')) !== false:
        case stripos($normalizedStateName, normalizeStateName('SIMPANG AMPAT')) !== false:
        case stripos($normalizedStateName, normalizeStateName('KULAI')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Muar')) !== false:
            $stateName = 'Johor';
            break;
        case stripos($normalizedStateName, normalizeStateName('BUTTERWORTH')) !== false:
        case stripos($normalizedStateName, normalizeStateName('GELUGOR')) !== false:
        case stripos($normalizedStateName, normalizeStateName('PRAI')) !== false:
        case stripos($normalizedStateName, normalizeStateName('NIBONG TEBAL')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Penang')) !== false:
        case stripos($normalizedStateName, normalizeStateName('SEBERANG PERAI')) !== false:
        case stripos($normalizedStateName, normalizeStateName('SEBERANG PERAI SELATAN')) !== false:
        case stripos($normalizedStateName, normalizeStateName('SEBERANG PERAI Tengah')) !== false:
        case stripos($normalizedStateName, normalizeStateName('TAMAN PERIDUSTRIAN RINGAN JURU')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Mertajam')) !== false:
            
            $stateName = 'Pulau Pinang';
            break;
        case stripos($normalizedStateName, normalizeStateName('DARUL EHSAN')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Kepong')) !== false:
        case stripos($normalizedStateName, normalizeStateName('KELANA JAYA')) !== false:
        case stripos($normalizedStateName, normalizeStateName('PETALING JAYA')) !== false:
        case stripos($normalizedStateName, normalizeStateName('Rawang')) !== false:
        case stripos($normalizedStateName, normalizeStateName('SUBANG JAYA')) !== false:
                $stateName = 'Selangor';
                break;
        case stripos($normalizedStateName, normalizeStateName('DERI MANJUNG')) !== false:
        case stripos($normalizedStateName, normalizeStateName('CHEMOR')) !== false:
            $stateName = 'Perak';
            break;
        // Add more cases here if needed
        default:
            // No special case, continue with the original $stateName
            break;
    }
    // Search for the state code using wildcard matching
    foreach ($stateCodes as $state) {
        if (stripos($state['State'], $stateName) !== false) {
            return $state['Code'];
        }
    }
    return "17"; // Return "17" if the state is not found
}
?>
