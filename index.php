<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

// Increase memory limit and execution time
ini_set('memory_limit', '1024M');
ini_set('max_execution_time', '600'); // 10 minutes

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    $files = $_FILES['files'];

    if (count($files['name']) === 0) {
        echo "No files were uploaded.";
        exit;
    }

    try {
        $spreadsheet = new Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $currentRow = 1; // Start from the first row

        echo "<p>Processing " . count($files['name']) . " files.</p>";

        for ($i = 0; $i < count($files['name']); $i++) {
            $fileTmpPath = $files['tmp_name'][$i];
            $fileName = $files['name'][$i];
            echo "<p>Processing file: $fileName</p>";

            // Load the current Excel file
            $spreadsheetToMerge = IOFactory::load($fileTmpPath);
            $sheetToMerge = $spreadsheetToMerge->getActiveSheet();

            // Get the highest row and column numbers referenced in the sheet
            $highestRow = $sheetToMerge->getHighestRow();
            $highestColumn = $sheetToMerge->getHighestColumn();

            // Get all rows from the sheet to merge
            $rows = $sheetToMerge->rangeToArray('A1:' . $highestColumn . $highestRow);

            // Write rows to the target spreadsheet, skipping the totals row
            foreach ($rows as $row) {
                if (strtolower(trim($row[0])) !== 'total') {
                    $sheet->fromArray($row, null, 'A' . $currentRow);
                    $currentRow++;
                }
            }

            // Add a 2-row gap between data from different files
            $currentRow += 2;

            // Free memory after processing each file
            $spreadsheetToMerge->disconnectWorksheets();
            unset($spreadsheetToMerge);
            gc_collect_cycles();
        }

        // Create a unique filename with a timestamp
        $timestamp = date('Ymd_His');
        $outputFile = "merged_$timestamp.xlsx";
        $writer = new Xlsx($spreadsheet);
        $writer->save($outputFile);

        echo "Files successfully merged into <a href='$outputFile'>$outputFile</a>.";
    } catch (Exception $e) {
        echo 'Error: ' . $e->getMessage();
    }
} else {
    echo "No files uploaded.";
}
?>

<!DOCTYPE html>
<html>
<head>
    <title>Merge Excel Files</title>
</head>
<body>
    <h1>Upload Excel Files to Merge</h1>
    <form action="index.php" method="post" enctype="multipart/form-data">
        <input type="file" name="files[]" multiple required>
        <button type="submit">Merge Files</button>
    </form>
</body>
</html>
