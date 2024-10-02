<?php
/**
 * Connects to an FTP server, logs in, and uploads a file from a memory stream.
 *
 * @param resource $memoryStream The memory stream resource containing the file data.
 * @param string $outputFolder The path on the FTP server where the file should be uploaded.
 * @param string $isLive Live or Test env
 * @return void
 */
function uploadToFtp($isLive, $memoryStream, $outputFolder) {
    // FTP server settings
    $ftpServer = "168.168.190.248";
    $ftpUsername = "baan";
    $ftpPassword = "123456";
    $logFile = 'ftp_errors.log'; // Path to the log file

    // Connect to FTP server
    $conn_id = ftp_connect($ftpServer);
    if (!$conn_id) {
        logError("Couldn't connect to $ftpServer", $logFile);
        return;
    }

    // Login to FTP server
    $login_result = ftp_login($conn_id, $ftpUsername, $ftpPassword);
    if (!$login_result) {
        logError("FTP login failed.", $logFile);
        ftp_close($conn_id);
        return;
    }

    // Enable passive mode
    ftp_pasv($conn_id, true);

    $dateTime = date('Y-m-d_H-i-s'); // Format: YYYY-MM-DD_HH-MM-SS
    $ftpFilePath = $outputFolder . '/' . $outputFolder . '_' . $dateTime . '.csv';

    if ($isLive !== 1) {
        $ftpFilePath = 'TestEnv/' . $ftpFilePath;
    }
    // Upload file to FTP server
    $upload = ftp_fput($conn_id, $ftpFilePath, $memoryStream, FTP_BINARY);
    if ($upload) {
        echo "FTP File uploaded successfully to $ftpFilePath\n";
    } else {
        logError("Failed to upload file.", $logFile);
    }

    // Close the FTP connection and memory stream
    fclose($memoryStream);
    ftp_close($conn_id);
}

/**
 * Creates and returns a memory stream containing CSV data.
 *
 * @param PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet The spreadsheet object with the data.
 * @return resource|false The memory stream containing the CSV data, or false on failure.
 */
function createMemoryStream($spreadsheet) {
    $logFile = 'memory_stream_errors.log'; // Path to the log file

    // Create a memory stream for the CSV file
    $memoryStream = fopen('php://temp', 'r+');
    if (!$memoryStream) {
        logMessage("WARNING: Failed to create memory stream.", $logFile);
        return false;
    }

    // Start output buffering
    ob_start();

    // Create CSV writer and save the data to output buffer
    try {
        $writer = new PhpOffice\PhpSpreadsheet\Writer\Csv($spreadsheet);
        $writer->setDelimiter('|');
        $writer->setEnclosure('');
        $writer->save('php://output');
    } catch (Exception $e) {
        logMessage("ERROR: CSV writer error: " . $e->getMessage(), $logFile);
        fclose($memoryStream);
        return false;
    }

    // Get the output buffer content and write it to the memory stream
    $csvContent = ob_get_clean();
    if ($csvContent === false) {
        logMessage("WARNING: Failed to get CSV content from output buffer.", $logFile);
        fclose($memoryStream);
        return false;
    }

    fwrite($memoryStream, $csvContent);
    rewind($memoryStream);

    return $memoryStream;
}

/**
 * Logs a message to a file with a timestamp.
 *
 * @param string $message The message to log.
 * @param string $logFile The path to the log file.
 * @return void
 */
function logMessage($message, $logFile) {
    $timestamp = date('Y-m-d H:i:s');
    $logMessage = "[$timestamp] $message\n";
    file_put_contents($logFile, $logMessage, FILE_APPEND);
}
?>