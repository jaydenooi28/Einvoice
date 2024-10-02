<?php
use PhpOffice\PhpSpreadsheet\IOFactory;
function generateToLocal($outputFolder,$spreadsheet){
    $pathPrefix = $outputFolder . '/' . $outputFolder . '_'; // Concatenate with slashes
    $localFilePath = $pathPrefix . date('YmdHis') . ".csv";
    $writer = IOFactory::createWriter($spreadsheet, 'Csv');
    $writer->setDelimiter('|');
    $writer->setEnclosure('');
    $writer->save($localFilePath);
    echo "CSV file has been saved to: " . $localFilePath;
}

?>