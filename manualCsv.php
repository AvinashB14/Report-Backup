<?php
if (session_status() === PHP_SESSION_NONE) {
    session_start();
}
$selectedDate = $_SESSION['selectedDate'];
$ftime=$_SESSION['ftime'];
$ttime=$_SESSION['ttime'];
date_default_timezone_set('Asia/Kolkata');
$dataCount = 0;
$serverName = "DESKTOP-OMHOR1A\\CIMPLICITY";
$connectionOptions = [
    "Database" => "master",
    "Uid" => "sa",
    "PWD" => "rechner@123"
];
$conn = sqlsrv_connect($serverName, $connectionOptions);
if ($conn === false) {
    die(print_r(sqlsrv_errors(), true));
}

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$columnWidths = [
    'A'  => 19.50,'N'  => 19.50,'AA' => 19.50,'AN' => 19.50,'BA' => 19.50,'B'=>8.60,'C'=>8.60,'D'=>8.60,'E'=>11.90,'G'=>11.90,'F'=>11.90,'H'=>11.90,'I'=>14.60,'J'=>13.40,'K'=>9.60,'L'=>9.60,
    'O'=>8.90,'P'=>8.90,'Q'=>10.30,'R'=>11.65,'S'=>9.70,'T'=>17.00,'U'=>17.00,'V'=>11.80,'W'=>13.00,'X'=>13.00,'AB'=>13.15,'AC'=>12.10,'AD'=>14.70,'AE'=>14.80,'AF'=>10.80,
    'AG'=>11.00,'AH'=>8.40,'AI'=>8.40,'AJ'=>8.40,'AK'=>13.70,'AO'=>8.30,'AP'=>8.30,'AQ'=>8.50,'AR'=>8.50,'AS'=>8.50,'AT'=>9.60,'AU'=>9.60,'AV'=>9.60,'AW'=>8.60,'AX'=>8.60,
    'AY'=>10.60,'BB'=>9.60,'BC'=>9.60,'BD'=>10.80,'BE'=>10.80,'BF'=>7.60,'BG'=>13.30,'BH'=>13.30,'BI'=>12.80,'BJ'=>13.20,'BK'=>9.00,'BL'=>9.00,'BM'=>6.00,'BN'=>5.50
];
foreach ($columnWidths as $col => $width) {
    $sheet->getColumnDimension($col)->setWidth($width);
}
$sheet->getRowDimension(2)->setRowHeight(60);

$sheet->setBreak('Y1', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_COLUMN);
$sheet->setBreak('AL1', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_COLUMN);
$sheet->setBreak('BM1', \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet::BREAK_COLUMN);
// Function to render a table (reusable for Table 1, 2, etc.)
function renderTable1($sheet, $conn, $sql, $title, $startRow, $date,$table1r,$table2r,$table3r,$table4r) {
    // Query
    $stmt = sqlsrv_query($conn, $sql);
    if ($stmt === false) {
        die(print_r(sqlsrv_errors(), true));
    }

    // ------------------ Detect table width ------------------
    $columns = [];
    $headers = [];
    $rowss=$startRow+1;
    $rowse=$startRow+3;
    // echo "<br>".$rowss.",".$rowse;
    foreach (sqlsrv_field_metadata($stmt) as $field) {
        $colName = $field['Name'];
        $displayName = (substr($colName, -5) === '_VAL0') ? substr($colName, 0, -5) : $colName;
        if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";
            //Span the rows
            $sheet->mergeCells("A{$rowss}:A{$rowse}");
            $sheet->setCellValue("A{$rowss}", $displayName);
            //style it
            $sheet->getStyle("A{$rowss}:A{$rowse}")->getAlignment()
                    ->setVertical(Alignment::VERTICAL_CENTER)
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        //Edit the column name
        if (stripos($displayName, "TET6538") !== false) {$displayName = "TI-T-6538";}
        if (stripos($displayName, "LTT6538") !== false) {$displayName = "LT-T-6538";}              
        if (stripos($displayName, "PTT6538") !== false) {$displayName = "PT-T-6538";}            
        if (stripos($displayName, "VFDP6538AAMP") !== false) {$displayName = "VFD-P-6538A";}       
        if (stripos($displayName, "VFDP6538BAMP") !== false) {$displayName = "VFD-P-6538B";}       
        if (stripos($displayName, "VFDP6538ARPM") !== false) {$displayName = "VFD-P-6538A";}       
        if (stripos($displayName, "VFDP6538BRPM") !== false) {$displayName = "VFD-P-6538B";}
        if (stripos($displayName, "MFMWFE6501") !== false) {$displayName = "MFM-WFE-6501";}
        if (stripos($displayName, "FCVWFE6501") !== false) {$displayName = "FCV-WFE-6501";}
        if (stripos($displayName, "TET6533A") !== false) {$displayName = "TI-T-6533A";}
        if (stripos($displayName, "TET6533B") !== false) {$displayName = "TI-T-6533B";}
        $headers[] = $displayName;
        $columns[] = $colName; 
    }
        
    $colCount = count($headers); 
    $lastCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colCount);

    // ------------------ Header Section ------------------
    $sheet->mergeCells("A{$startRow}:C{$startRow}")->setCellValue("A{$startRow}", "Date: $date");

    $middleEndCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colCount - 3);
    $rightStartCol = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colCount - 2);

    $sheet->mergeCells("D{$startRow}:{$middleEndCol}{$startRow}")
        ->setCellValue("D{$startRow}", "ALKYL AMINES CHEMICAL LIMITED\nKURKUMBH\n$title");

    $secondRow=$startRow+1;
    //Tag Description //Table 1
    
    
    $sheet->mergeCells("B2:B2")->setCellValue("B2", "T-6538\nFeed\nTank\nLevel");//1
    $sheet->mergeCells("C2:C2")->setCellValue("C2", "T-6538\nFeed\nTank\nTemp");//2
    $sheet->mergeCells("D2:D2")->setCellValue("D2", "T-6538\nFeed\nTank\nPressure");//3
    $sheet->mergeCells("E2:H2")->setCellValue("E2", "Feed Pump");//4,5,6,7
    $sheet->mergeCells("I2:I2")->setCellValue("I2", "WFE Feed Rate");//8
    $sheet->mergeCells("J2:J2")->setCellValue("J2", "WFE Feed CV\nOpen");//9
    $sheet->mergeCells("K2:L2")->setCellValue("K2", "Hot Water Tank\nTemp");//10
    //Units
    $sheet->mergeCells("B4:B4")->setCellValue("B4", "%");//1
    $sheet->mergeCells("C4:C4")->setCellValue("C4", "°C");//2
    $sheet->mergeCells("D4:D4")->setCellValue("D4", "kg/cm²g");//3
    $sheet->mergeCells("E4:E4")->setCellValue("E4", "AMP");//4,5,6,7
    $sheet->mergeCells("F4:F4")->setCellValue("F4", "AMP");//4,5,6,7
    $sheet->mergeCells("G4:G4")->setCellValue("G4", "RPM");//8
    $sheet->mergeCells("H4:H4")->setCellValue("H4", "RPM");//8
    $sheet->mergeCells("I4:I4")->setCellValue("I4", "kg/Hr");//9
    $sheet->mergeCells("J4:J4")->setCellValue("J4", "%");//10
    $sheet->mergeCells("K4:K4")->setCellValue("K4", "°C");//11
    $sheet->mergeCells("L4:L4")->setCellValue("L4", "°C");//12
    //Increase the height of tags units row
    //Style tags Description
    $sheet->getStyle("B2:N2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    // Right side (Document info)
    $sheet->mergeCells("{$rightStartCol}{$startRow}:{$lastCol}{$startRow}")
        ->setCellValue("{$rightStartCol}{$startRow}",
            "Document No:\nDocument Name:\nIssued No:\nRevision No:\nPage No: 01 of 05");

    $sheet->getStyle("A{$startRow}:{$lastCol}{$startRow}")->applyFromArray([
     'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("A{$secondRow}:{$lastCol}{$secondRow}")->applyFromArray([
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("D{$startRow}:{$middleEndCol}{$startRow}")->getFont()->setBold(true);

    // Align header
    $sheet->getStyle("A{$startRow}:C{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->getStyle("D{$startRow}:{$middleEndCol}{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER)
          ->setWrapText(true);

    $sheet->getStyle(($colCount > 3 ? $middleEndCol : "D") . "{$startRow}:{$lastCol}{$startRow}")
          ->getAlignment()->setVertical(Alignment::VERTICAL_TOP)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT)
          ->setWrapText(true);

    $sheet->getRowDimension($startRow)->setRowHeight(75);
    // ------------------ Column Headers ------------------
    $headerRow = $secondRow + 1;
    $sheet->fromArray($headers, NULL, "A$headerRow");
    $sheet->getStyle("A$headerRow:{$lastCol}$headerRow")->getFont()->setBold(true);

    // ------------------ Data rows ------------------ 
    // $t2r = 0;$t3r = 0;$t4r = 0;
    $rowNum = $headerRow + 2;
    $rowCount=0;
    $rarray=[];
    while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
        $excelRow = [];
        foreach ($columns as $colIndex => $col) {
            $value = $row[$col];
            if ($colIndex === 0 && $value instanceof DateTime) {
                $value = $value->format('d-m-Y h:i A');
            }
            if ($value === null || $value === '') {
                $value = 0;
            }
            if ($col === "VP6503ONOFF_VAL0") {
                if ($value == 0) {
                    $value = "OFF";
                } elseif ($value == 1) {
                    $value = "ON";
                }
            }
            $excelRow[] = $value;
        }
        $sheet->fromArray($excelRow, NULL, "A$rowNum");
        $sheet->setBreak('M1', Worksheet::BREAK_COLUMN);
        $rowNum++;
    }
    $lastRow = $rowNum - 1;
    // echo $lastRow;
    $sheet->getStyle("A$headerRow:{$lastCol}$lastRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ]);
    // ------------------ Footer ------------------
    $footerRow = $lastRow + 1;
    $sheet->mergeCells("A$footerRow:{$lastCol}$footerRow")
          ->setCellValue("A$footerRow", "CONTROLLER (Name and Sign):");
    $sheet->getStyle("A$footerRow:{$lastCol}$footerRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    $footerRow2 = $footerRow + 1;
    $sheet->mergeCells("A$footerRow2:{$lastCol}$footerRow2")
          ->setCellValue("A$footerRow2", "SIC (Name and Sign):");
    $sheet->getStyle("A$footerRow2:{$lastCol}$footerRow2")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    return $footerRow2+2;
}

// Table 1
$sql1 = "WITH ABC AS (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY DATEDIFF(MINUTE, '$selectedDate $ftime', [timestamp]) / 120
                    ORDER BY [timestamp] ASC
                ) AS rn
            FROM [master].[dbo].[TRENDS]
            WHERE 
                [timestamp] >= '$selectedDate $ftime'
                AND [timestamp] <= '$selectedDate $ttime' 
        )
        SELECT  
            [timestamp],
            [LTT6538_VAL0],
            [TET6538_VAL0],
            [PTT6538_VAL0],
            [VFDP6538AAMP_VAL0],
            [VFDP6538BAMP_VAL0],
            [VFDP6538ARPM_VAL0],
            [VFDP6538BRPM_VAL0],
            [MFMWFE6501_VAL0],
            [FCVWFE6501OPENING_VAL0],
            [TET6533A_VAL0],
            [TET6533B_VAL0]
        FROM ABC
        WHERE rn = 1
        ORDER BY [timestamp] ASC;";

$nextStartRow = renderTable1($sheet, $conn, $sql1, "DES WFE PLANT CONTROLLER LOGSHEET \n", 1, $selectedDate, 1 ,0,0,0);

//TABLE 2
function renderTable2($sheet, $conn, $sql, $title, $startRow, $date) {
    // Query
    $stmt = sqlsrv_query($conn, $sql);
    if ($stmt === false) {
        die(print_r(sqlsrv_errors(), true));
    }

    // ------------------ Detect table width ------------------
    $columns = [];
    $headers = [];
    $rowss=$startRow+1;
    $rowse=$startRow+3;
    // echo "<br>".$rowss.",".$rowse;
    foreach (sqlsrv_field_metadata($stmt) as $field) {
        $colName = $field['Name'];
        $displayName = (substr($colName, -5) === '_VAL0') ? substr($colName, 0, -5) : $colName;
        if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";
            //Span the rows
            $sheet->mergeCells("N{$rowss}:N{$rowse}");
            $sheet->setCellValue("N{$rowss}", $displayName);
            //style it
            $sheet->getStyle("N{$rowss}:N{$rowse}")->getAlignment()
                    ->setVertical(Alignment::VERTICAL_CENTER)
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        //Edit the column name
        if (stripos($displayName, "PTT65331") !== false) {$displayName = "PT-T-6533";}
        if (stripos($displayName, "LTT6533") !== false) {$displayName = "LT-T-6533";}
        if (stripos($displayName, "TET6533A") !== false) {$displayName = "TCV-T-6533";}
        if (stripos($displayName, "FTHWSWFE6501") !== false) {$displayName = "FI-HWS-6501";}
        if (stripos($displayName, "FCVHWRWFE6501OPENING") !== false) {$displayName = "FCV-HWR-";}
        if (stripos($displayName, "TEHWSWFE6501") !== false) {$displayName = "TE-HWS-WFE-6501";}
        if (stripos($displayName, "TEHWRWFE6501") !== false) {$displayName = "TE-HWR-WFE-6501";}
        if (stripos($displayName, "PTWFE6501") !== false) {$displayName = "PT-WFE-6501";}
        if (stripos($displayName, "TEWFE65011") !== false) {$displayName = "TI-WFE-6501-1";}
        if (stripos($displayName, "TEWFE65012") !== false) {$displayName = "TI-WFE-6501-2";}
        if (stripos($displayName, "TEWFE65013") !== false) {$displayName = "TI-WFE-6501-3";}
        $headers[] = $displayName;
        $columns[] = $colName; 
    }
    $colCount = count($headers); 
    $lastCol="X";
    // ------------------ Header Section ------------------
    $sheet->mergeCells("N{$startRow}:P{$startRow}")->setCellValue("N{$startRow}", "Date: $date");
    $middleEndCol="U";
    $rightStartCol="V";
    $sheet->mergeCells("Q{$startRow}:{$middleEndCol}{$startRow}")
        ->setCellValue("Q{$startRow}", "ALKYL AMINES CHEMICAL LIMITED\nKURKUMBH\n$title");

    $secondRow=$startRow+1;
    //Description row height
    // $sheet->getRowDimension($startRow+1)->setRowHeight(60);
    //Tag Description //Table 1
    $sheet->mergeCells("O2:O2")->setCellValue("O2", "Hot\nWater\nTank\nPressure");//11
    $sheet->mergeCells("P2:P2")->setCellValue("P2", "Hot\nWater\nTank\nLevel");//11
    $sheet->mergeCells("Q2:Q2")->setCellValue("Q2", "Hot Water\ntank steam\nCV Open");//1
    $sheet->mergeCells("R2:R2")->setCellValue("R2", "Flow rate of\nHot water to\nWFE");//2
    $sheet->mergeCells("S2:S2")->setCellValue("S2", "Hot Water\nflow CV\nOpen");//3
    $sheet->mergeCells("T2:T2")->setCellValue("T2", "WFE Hot Water\nSupply temp");//4
    $sheet->mergeCells("U2:U2")->setCellValue("U2", "WFE Hot Water\nreturn temp");//5
    $sheet->mergeCells("V2:V2")->setCellValue("V2", "WFE\nPressure");//6
    $sheet->mergeCells("W2:W2")->setCellValue("W2", "Cyclone\nBottom Line");//7
    $sheet->mergeCells("X2:X2")->setCellValue("X2", "WFE Bottom\ntemp");//8

    //Units
    $sheet->mergeCells("O4:O4")->setCellValue("O4", "kg/cm²g");//13
    $sheet->mergeCells("P4:P4")->setCellValue("P4", "%");//14
    $sheet->mergeCells("Q4:Q4")->setCellValue("Q4", "%");//3
    $sheet->mergeCells("R4:R4")->setCellValue("R4", "Kg/Hr");//4,5,6,7
    $sheet->mergeCells("S4:S4")->setCellValue("S4", "%");//4,5,6,7
    $sheet->mergeCells("T4:T4")->setCellValue("T4", "°C");//8
    $sheet->mergeCells("U4:U4")->setCellValue("U4", "°C");//8
    $sheet->mergeCells("V4:V4")->setCellValue("V4", "mmHg");//9
    $sheet->mergeCells("W4:W4")->setCellValue("W4", "°C");//10
    $sheet->mergeCells("X4:X4")->setCellValue("X4", "°C");//11
    //Style tags Description
    $sheet->getStyle("N2:Y2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    // Right side (Document info)
    $sheet->mergeCells("{$rightStartCol}{$startRow}:{$lastCol}{$startRow}")
        ->setCellValue("{$rightStartCol}{$startRow}",
            "Document No:\nDocument Name:\nIssued No:\nRevision No:\nPage No: 02 of 05");

    $sheet->getStyle("N{$startRow}:{$lastCol}{$startRow}")->applyFromArray([
     'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("N{$secondRow}:{$lastCol}{$secondRow}")->applyFromArray([
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("Q{$startRow}:{$middleEndCol}{$startRow}")->getFont()->setBold(true);

    // Align header
    $sheet->getStyle("N{$startRow}:P{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->getStyle("Q{$startRow}:{$middleEndCol}{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER)
          ->setWrapText(true);

    $sheet->getStyle(($colCount > 3 ? $middleEndCol : "Q") . "{$startRow}:{$lastCol}{$startRow}")
          ->getAlignment()->setVertical(Alignment::VERTICAL_TOP)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT)
          ->setWrapText(true);

    $sheet->getRowDimension($startRow)->setRowHeight(75);
    // ------------------ Column Headers ------------------
    $headerRow = $secondRow + 1;
    $sheet->fromArray($headers, NULL, "N$headerRow");
    $sheet->getStyle("N$headerRow:{$lastCol}$headerRow")->getFont()->setBold(true);

    // ------------------ Data rows ------------------
    $rowNum = $headerRow + 2;
    $rowCount=0;
    $rarray=[];
    while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
        $excelRow = [];
        foreach ($columns as $colIndex => $col) {
            $value = $row[$col];
            if ($colIndex === 0 && $value instanceof DateTime) {
                $value = $value->format('d-m-Y h:i A');
            }
            if ($value === null || $value === '') {
                $value = 0;
            }
            if ($col === "VP6503ONOFF_VAL0") {
                if ($value == 0) {
                    $value = "OFF";
                } elseif ($value == 1) {
                    $value = "ON";
                }
            }
            $excelRow[] = $value;
        }
        $sheet->fromArray($excelRow, NULL, "N$rowNum");
        $rowNum++;
    }
    $lastRow = $rowNum - 1;
    // echo $lastRow;
    $sheet->getStyle("N$headerRow:{$lastCol}$lastRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ]);
    // ------------------ Footer ------------------
    $footerRow = $lastRow + 1;
    $sheet->mergeCells("N$footerRow:{$lastCol}$footerRow")
          ->setCellValue("N$footerRow", "CONTROLLER (Name and Sign):");
    $sheet->getStyle("N$footerRow:{$lastCol}$footerRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    $footerRow2 = $footerRow + 1;
    $sheet->mergeCells("N$footerRow2:{$lastCol}$footerRow2")
          ->setCellValue("N$footerRow2", "SIC (Name and Sign):");
    $sheet->getStyle("N$footerRow2:{$lastCol}$footerRow2")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    return $footerRow2 + 2;
}
// Table 2
$sql2 = "WITH ABC AS (
                    SELECT *,
                        ROW_NUMBER() OVER (
                            PARTITION BY DATEDIFF(MINUTE, '$selectedDate $ftime', [timestamp]) / 120
                            ORDER BY [timestamp] ASC
                        ) AS rn
                    FROM [master].[dbo].[TRENDS]
                    WHERE 
                        [timestamp] >= '$selectedDate $ftime'
                        AND [timestamp] <= '$selectedDate $ttime'
                )
                SELECT  
                    [timestamp],
                    [PTT65331_VAL0],
                    [LTT6533_VAL0],
                    [TET6533A_VAL0],
                    [FTHWSWFE6501_VAL0],
                    [FCVHWRWFE6501OPENING_VAL0],
                    [TEHWSWFE6501_VAL0],
                    [TEHWRWFE6501_VAL0],
                    [PTWFE6501_VAL0],
                    [TEWFE65011_VAL0],
                    [TEWFE65012_VAL0]
                FROM ABC
                WHERE rn = 1
                ORDER BY [timestamp] ASC;";

$nextStartRow = renderTable2($sheet, $conn, $sql2, "DES WFE PLANT CONTROLLER LOGSHEET \n", 1, $selectedDate);
//TABLE 3
function renderTable3($sheet, $conn, $sql, $title, $startRow, $date) {
    // Query
    $stmt = sqlsrv_query($conn, $sql);
    if ($stmt === false) {
        die(print_r(sqlsrv_errors(), true));
    }

    // ------------------ Detect table width ------------------
    $columns = [];
    $headers = [];
    $rowss=$startRow+1;
    $rowse=$startRow+3;
    // echo "<br>".$rowss.",".$rowse;
    foreach (sqlsrv_field_metadata($stmt) as $field) {
        $colName = $field['Name'];
        $displayName = (substr($colName, -5) === '_VAL0') ? substr($colName, 0, -5) : $colName;
        if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";
            //Span the rows
            $sheet->mergeCells("AA{$rowss}:AA{$rowse}");
            $sheet->setCellValue("AA{$rowss}", $displayName);
            //style it
            $sheet->getStyle("AA{$rowss}:AA{$rowse}")->getAlignment()
                    ->setVertical(Alignment::VERTICAL_CENTER)
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        //Edit the column name
        if (stripos($displayName, "TEWFE65013") !== false) {$displayName = "TI-WFE-6501-3";}
        if (stripos($displayName, "TEHE6503") !== false) {$displayName = "TI-DPHE-6503";}
        if (stripos($displayName, "FTCHBSHE6503") !== false) {$displayName = "FT-CHBS-HE6503";}
        if (stripos($displayName, "FTCHBSHC6501") !== false) {$displayName = "FT-CHBS-HC6501";}
        if (stripos($displayName, "TEHC6501B") !== false) {$displayName = "TI-HC-6501B";}
        if (stripos($displayName, "TEHC6501A") !== false) {$displayName = "TI-HC-6501A";}
        if (stripos($displayName, "TET6535") !== false) {$displayName = "TI-T-6535";}
        if (stripos($displayName, "LTT6535") !== false) {$displayName = "LT-T-6535";}
        if (stripos($displayName, "PTT6535") !== false) {$displayName = "PT-T-6535";}
        if (stripos($displayName, "FTCWST6535") !== false) {$displayName = "FT-CWS-T-6535";}
        $headers[] = $displayName;
        $columns[] = $colName; 
    }
    $colCount = count($headers); 
    $lastCol="AK";
    // ------------------ Header Section ------------------
    $sheet->mergeCells("AA{$startRow}:AC{$startRow}")->setCellValue("AA{$startRow}", "Date: $date");
    $middleEndCol="AH";
    $rightStartCol="AI";
    $sheet->mergeCells("AD{$startRow}:{$middleEndCol}{$startRow}")
        ->setCellValue("AD{$startRow}", "ALKYL AMINES CHEMICAL LIMITED\nKURKUMBH\n$title");

    $secondRow=$startRow+1;
    //Description row height
    // $sheet->getRowDimension($startRow+1)->setRowHeight(60);
    //Tag Description
    $sheet->mergeCells("AB2:AB2")->setCellValue("AB2", "Vapour Line\ntemp");//11
    $sheet->mergeCells("AC2:AC2")->setCellValue("AC2", "Double Pipe\nOutlet");//11
    $sheet->mergeCells("AD2:AD2")->setCellValue("AD2", "Double Pipe HE");//1
    $sheet->mergeCells("AE2:AE2")->setCellValue("AE2", "Chilling Flow to\nCondenser");//2
    $sheet->mergeCells("AF2:AF2")->setCellValue("AF2", "DES\nCollection\nLine temp");//3
    $sheet->mergeCells("AG2:AG2")->setCellValue("AG2", "Chilling\nCondenser\nVapour\nOutlet");//4
    $sheet->mergeCells("AH2:Ak2")->setCellValue("AH2", "Concentrate Collection Tank");//5
    //Units
    $sheet->mergeCells("AB4:AB4")->setCellValue("AB4", "°C");//13
    $sheet->mergeCells("AC4:AC4")->setCellValue("AC4", "°C");//14
    $sheet->mergeCells("AD4:AD4")->setCellValue("AD4", "Kg/Hr");//3
    $sheet->mergeCells("AE4:AE4")->setCellValue("AE4", "Kg/Hr");//4,5,6,7
    $sheet->mergeCells("AF4:AF4")->setCellValue("AF4", "°C");//4,5,6,7
    $sheet->mergeCells("AG4:AG4")->setCellValue("AG4", "°C");//8
    $sheet->mergeCells("AH4:AH4")->setCellValue("AH4", "°C");//8
    $sheet->mergeCells("AI4:AI4")->setCellValue("AI4", "%");//9
    $sheet->mergeCells("AJ4:AJ4")->setCellValue("AJ4", "mmHg");//10
    $sheet->mergeCells("AK4:AK4")->setCellValue("AK4", "Kg/Hr");//11
    //Increase the height of tags units row
    //Style tags Description
    $sheet->getStyle("AA2:AK2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle("AH2:AK2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    // Right side (Document info)
    $sheet->mergeCells("{$rightStartCol}{$startRow}:{$lastCol}{$startRow}")
        ->setCellValue("{$rightStartCol}{$startRow}",
            "Document No:\nDocument Name:\nIssued No:\nRevision No:\nPage No: 03 of 05");

    $sheet->getStyle("AA{$startRow}:{$lastCol}{$startRow}")->applyFromArray([
     'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("AA{$secondRow}:{$lastCol}{$secondRow}")->applyFromArray([
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("AD{$startRow}:{$middleEndCol}{$startRow}")->getFont()->setBold(true);

    // Align header
    $sheet->getStyle("AA{$startRow}:AC{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->getStyle("AD{$startRow}:{$middleEndCol}{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER)
          ->setWrapText(true);

    $sheet->getStyle(($colCount > 3 ? $middleEndCol : "AD") . "{$startRow}:{$lastCol}{$startRow}")
          ->getAlignment()->setVertical(Alignment::VERTICAL_TOP)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT)
          ->setWrapText(true);

    $sheet->getRowDimension($startRow)->setRowHeight(75);
    // ------------------ Column Headers ------------------
    $headerRow = $secondRow + 1;
    $sheet->fromArray($headers, NULL, "AA$headerRow");
    $sheet->getStyle("AA$headerRow:{$lastCol}$headerRow")->getFont()->setBold(true);

    // ------------------ Data rows ------------------
    $rowNum = $headerRow + 2;
    $rowCount=0;
    $rarray=[];
    while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
        $excelRow = [];
        foreach ($columns as $colIndex => $col) {
            $value = $row[$col];
            if ($colIndex === 0 && $value instanceof DateTime) {
                $value = $value->format('d-m-Y h:i A');
            }
            if ($value === null || $value === '') {
                $value = 0;
            }
            if ($col === "VP6503ONOFF_VAL0") {
                if ($value == 0) {
                    $value = "OFF";
                } elseif ($value == 1) {
                    $value = "ON";
                }
            }
            $excelRow[] = $value;
        }
        $sheet->fromArray($excelRow, NULL, "AA$rowNum");
        $rowNum++;
    }
    $lastRow = $rowNum - 1;
    // echo $lastRow;
    $sheet->getStyle("AA$headerRow:{$lastCol}$lastRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ]);
    // ------------------ Footer ------------------
    $footerRow = $lastRow + 1;
    $sheet->mergeCells("AA$footerRow:{$lastCol}$footerRow")
          ->setCellValue("AA$footerRow", "CONTROLLER (Name and Sign):");
    $sheet->getStyle("AA$footerRow:{$lastCol}$footerRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    $footerRow2 = $footerRow + 1;
    $sheet->mergeCells("AA$footerRow2:{$lastCol}$footerRow2")
          ->setCellValue("AA$footerRow2", "SIC (Name and Sign):");
    $sheet->getStyle("AA$footerRow2:{$lastCol}$footerRow2")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    return $footerRow2 + 2;
}
//Table 3
$sql3 = "WITH ABC AS (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY DATEDIFF(MINUTE, '$selectedDate $ftime', [timestamp]) / 120
                    ORDER BY [timestamp] ASC
                ) AS rn
            FROM [master].[dbo].[TRENDS]
            WHERE 
                [timestamp] >= '$selectedDate $ftime'
                AND [timestamp] <= '$selectedDate $ttime'
        )
        SELECT  
            [timestamp],
            [TEWFE65013_VAL0],
            [TEHE6503_VAL0],
            [FTCHBSHE6503_VAL0],
            [FTCHBSHC6501_VAL0],
            [TEHC6501B_VAL0],
            [TEHC6501A_VAL0],
            [TET6535_VAL0],
            [LTT6535_VAL0],
            [PTT6535_VAL0],
            [FTCWST6535_VAL0]
        FROM ABC
        WHERE rn = 1
        ORDER BY [timestamp] ASC;";
$nextStartRow = renderTable3($sheet, $conn, $sql3, "DES WFE PLANT CONTROLLER LOGSHEET \n", 1, $selectedDate);
//TABLE 4
function renderTable4($sheet, $conn, $sql, $title, $startRow, $date) {
    // Query
    $stmt = sqlsrv_query($conn, $sql);
    if ($stmt === false) {
        die(print_r(sqlsrv_errors(), true));
    }

    // ------------------ Detect table width ------------------
    $columns = [];
    $headers = [];
    $rowss=$startRow+1;
    $rowse=$startRow+3;
    // echo "<br>".$rowss.",".$rowse;
    foreach (sqlsrv_field_metadata($stmt) as $field) {
        $colName = $field['Name'];
        $displayName = (substr($colName, -5) === '_VAL0') ? substr($colName, 0, -5) : $colName;
        if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";
            //Span the rows
            $sheet->mergeCells("AN{$rowss}:AN{$rowse}");
            $sheet->setCellValue("AN{$rowss}", $displayName);
            //style it
            $sheet->getStyle("AN{$rowss}:AN{$rowse}")->getAlignment()
                    ->setVertical(Alignment::VERTICAL_CENTER)
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        //Edit the column name
        if (stripos($displayName, "FTCWST6535") !== false) {$displayName = "FT-CWS-T-6535";}
        if (stripos($displayName, "TECHBS") !== false) {$displayName = "TI-CHBS";}
        if (stripos($displayName, "TECHBR") !== false) {$displayName = "TI-CHBR";}
        if (stripos($displayName, "LTT6536") !== false) {$displayName = "LT-T-6536";}
        if (stripos($displayName, "TET6536") !== false) {$displayName = "TI-T-6536";}
        if (stripos($displayName, "PTT6536") !== false) {$displayName = "PT-T-6536";}
        if (stripos($displayName, "LTST6502") !== false) {$displayName = "LT-ST-6502";}
        if (stripos($displayName, "TIST6502") !== false) {$displayName = "TI-ST-6502";}
        if (stripos($displayName, "LTT6504") !== false) {$displayName = "LT-ST-6504";}
        if (stripos($displayName, "LTT6539") !== false) {$displayName = "LT-T-6539";}
        if (stripos($displayName, "PTT6539") !== false) {$displayName = "PT-T-6539";}
        if (stripos($displayName, "PTT65332") !== false) {$displayName = "PT-T-6533-2";}
        if (stripos($displayName, "LTT6505") !== false) {$displayName = "LT-ST-6505";}
        $headers[] = $displayName;
        $columns[] = $colName; 
    }
    $colCount = count($headers); 
    $lastCol="AY";
    // ------------------ Header Section ------------------
    $sheet->mergeCells("AN{$startRow}:AP{$startRow}")->setCellValue("AN{$startRow}", "Date: $date");
    $middleEndCol="AU";
    $rightStartCol="AV";
    $sheet->mergeCells("AQ{$startRow}:{$middleEndCol}{$startRow}")
        ->setCellValue("AQ{$startRow}", "ALKYL AMINES CHEMICAL LIMITED\nKURKUMBH\n$title");

    $secondRow=$startRow+1;
    //Description row height
    // $sheet->getRowDimension($startRow+1)->setRowHeight(60);
    //Tag Description //Table 1
    $sheet->mergeCells("AO2:AP2")->setCellValue("AO2", "Chilling\nTempratures");//11
    // $sheet->mergeCells("AP2:AP2")->setCellValue("AP2", "Hot\nWater\nTank\nLevel");//11
    $sheet->mergeCells("AQ2:AS2")->setCellValue("AQ2", "DES Collection Tank (T-6536)");//1
    $sheet->mergeCells("AT2:AU2")->setCellValue("AT2", "Purging Tank (T-6502)");//2
    $sheet->mergeCells("AV2:AV2")->setCellValue("AV2", "ST 6504");//3
    $sheet->mergeCells("AW2:AX2")->setCellValue("AW2", "Off Spec Storage\nTank (T-6539)");//4
    $sheet->mergeCells("AY2:AY2")->setCellValue("AY2", "Hot water tank\nparameters");//5

    //Units
    $sheet->mergeCells("AO4:AO4")->setCellValue("AO4", "°C");//13
    $sheet->mergeCells("AP4:AP4")->setCellValue("AP4", "°C");//14
    $sheet->mergeCells("AQ4:AQ4")->setCellValue("AQ4", "%");//3
    $sheet->mergeCells("AR4:AR4")->setCellValue("AR4", "°C");//4,5,6,7
    $sheet->mergeCells("AS4:AS4")->setCellValue("AS4", "mmHg");//4,5,6,7
    $sheet->mergeCells("AT4:AT4")->setCellValue("AT4", "%");//8
    $sheet->mergeCells("AU4:AU4")->setCellValue("AU4", "°C");//8
    $sheet->mergeCells("AV4:AV4")->setCellValue("AV4", "%");//9
    $sheet->mergeCells("AW4:AW4")->setCellValue("AW4", "%");//10
    $sheet->mergeCells("AX4:AX4")->setCellValue("AX4", "Kg/cm²g");//11
    $sheet->mergeCells("AY4:AY4")->setCellValue("AY4", "Kg/cm²g");//11
    //Style tags Description
    $sheet->getStyle("AN2:AY2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    $sheet->getStyle("AN{$startRow}:AY{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->getStyle("AN2:AY2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    // Right side (Document info)
    $sheet->mergeCells("{$rightStartCol}{$startRow}:{$lastCol}{$startRow}")
        ->setCellValue("{$rightStartCol}{$startRow}",
            "Document No:\nDocument Name:\nIssued No:\nRevision No:\nPage No: 04 of 05");

    $sheet->getStyle("AN{$startRow}:{$lastCol}{$startRow}")->applyFromArray([
     'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("AN{$secondRow}:{$lastCol}{$secondRow}")->applyFromArray([
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("AQ{$startRow}:{$middleEndCol}{$startRow}")->getFont()->setBold(true);

    // Align header
    $sheet->getStyle("AN{$startRow}:P{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->getStyle("AQ{$startRow}:{$middleEndCol}{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER)
          ->setWrapText(true);

    $sheet->getStyle(($colCount > 3 ? $middleEndCol : "AQ") . "{$startRow}:{$lastCol}{$startRow}")
          ->getAlignment()->setVertical(Alignment::VERTICAL_TOP)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT)
          ->setWrapText(true);

    $sheet->getRowDimension($startRow)->setRowHeight(75);
    // ------------------ Column Headers ------------------
    $headerRow = $secondRow + 1;
    $sheet->fromArray($headers, NULL, "AN$headerRow");
    $sheet->getStyle("AN$headerRow:{$lastCol}$headerRow")->getFont()->setBold(true);

    // ------------------ Data rows ------------------
    $rowNum = $headerRow + 2;
    $rowCount=0;
    $rarray=[];
    while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
        $excelRow = [];
        foreach ($columns as $colIndex => $col) {
            $value = $row[$col];
            if ($colIndex === 0 && $value instanceof DateTime) {
                $value = $value->format('d-m-Y h:i A');
            }
            if ($value === null || $value === '') {
                $value = 0;
            }
            if ($col === "VP6503ONOFF_VAL0") {
                if ($value == 0) {
                    $value = "OFF";
                } elseif ($value == 1) {
                    $value = "ON";
                }
            }
            $excelRow[] = $value;
        }
        $sheet->fromArray($excelRow, NULL, "AN$rowNum");
        $rowNum++;
    }
    $lastRow = $rowNum - 1;
    // echo $lastRow;
    $sheet->getStyle("AN$headerRow:{$lastCol}$lastRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ]);
    // ------------------ Footer ------------------
    $footerRow = $lastRow + 1;
    $sheet->mergeCells("AN$footerRow:{$lastCol}$footerRow")
          ->setCellValue("AN$footerRow", "CONTROLLER (Name and Sign):");
    $sheet->getStyle("AN$footerRow:{$lastCol}$footerRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    $footerRow2 = $footerRow + 1;
    $sheet->mergeCells("AN$footerRow2:{$lastCol}$footerRow2")
          ->setCellValue("AN$footerRow2", "SIC (Name and Sign):");
    $sheet->getStyle("AN$footerRow2:{$lastCol}$footerRow2")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    return $footerRow2 + 2;
}
//Table 4
$sql4 = "WITH ABC AS (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY DATEDIFF(MINUTE, '$selectedDate $ftime', [timestamp]) / 120
                    ORDER BY [timestamp] ASC
                ) AS rn
            FROM [master].[dbo].[TRENDS]
            WHERE 
                [timestamp] >= '$selectedDate $ftime'
                AND [timestamp] <= '$selectedDate $ttime'
        )
        SELECT  
            [timestamp],
            [TECHBS_VAL0],
            [TECHBR_VAL0],
            [LTT6536_VAL0],
            [TET6536_VAL0],
            [PTT6536_VAL0],
            [LTST6502_VAL0],
            [TIST6502_VAL0],
            [LTT6504_VAL0],
            [LTT6539_VAL0],
            [PTT6539_VAL0],
            [PTT65332_VAL0]
        FROM ABC
        WHERE rn = 1
        ORDER BY [timestamp] ASC;";
$nextStartRow = renderTable4($sheet, $conn, $sql4, "DES WFE PLANT CONTROLLER LOGSHEET \n", 1, $selectedDate);
// /TABLE 5
function renderTable5($sheet, $conn, $sql, $title, $startRow, $date) {
    // Query
    $stmt = sqlsrv_query($conn, $sql);
    if ($stmt === false) {
        die(print_r(sqlsrv_errors(), true));
    }

    // ------------------ Detect table width ------------------
    $columns = [];
    $headers = [];
    $rowss=$startRow+1;
    $rowse=$startRow+3;
    // echo "<br>".$rowss.",".$rowse;
    foreach (sqlsrv_field_metadata($stmt) as $field) {
        $colName = $field['Name'];
        $displayName = (substr($colName, -5) === '_VAL0') ? substr($colName, 0, -5) : $colName;
        if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";
            //Span the rows
            $sheet->mergeCells("BA{$rowss}:BA{$rowse}");
            $sheet->setCellValue("BA{$rowss}", $displayName);
            //style it
            $sheet->getStyle("BA{$rowss}:BA{$rowse}")->getAlignment()
                    ->setVertical(Alignment::VERTICAL_CENTER)
                    ->setHorizontal(Alignment::HORIZONTAL_CENTER);
        }
        //Edit the column name
        if (stripos($displayName, "timestamp") !== false) {$displayName = "Date&Time";}
        if (stripos($displayName, "LTT6505") !== false) {$displayName = "LT-ST-6505";}
        if (stripos($displayName, "LTT6506") !== false) {$displayName = "LT-ST-6506";}
        if (stripos($displayName, "LTST6702B") !== false) {$displayName = "LT-ST-6702B";}
        if (stripos($displayName, "TIST6702B") !== false) {$displayName = "TI-ST-6702B";}
        if (stripos($displayName, "VP6503ONOFF") !== false) {$displayName = "VP-6503";}
        if (stripos($displayName, "VP6502AAMP") !== false) {$displayName = "VFD-VP-6502A";}
        if (stripos($displayName, "VP6502ARPM") !== false) {$displayName = "VFD-VP-6502A";}
        if (stripos($displayName, "VP6502BAMP") !== false) {$displayName = "VFD-VP-6502B";}
        if (stripos($displayName, "VP6502BRPM") !== false) {$displayName = "VFD-VP-6502B";}
        if (stripos($displayName, "PTT6537") !== false) {$displayName = "PT-T-6537";}
        if (stripos($displayName, "TET6537") !== false) {$displayName = "TE-T-6537";}
        if (stripos($displayName, "VT01") !== false) {$displayName = "VT-01";}
        if (stripos($displayName, "TT011") !== false) {$displayName = "TT-01";}
        $headers[] = $displayName;
        $columns[] = $colName; 
    }
    $colCount = count($headers); 
    $lastCol="BN";
    // ------------------ Header Section ------------------
    $sheet->mergeCells("BA{$startRow}:BC{$startRow}")->setCellValue("BA{$startRow}", "Date: $date");
    $middleEndCol="BH";
    $rightStartCol="BI";
    $sheet->mergeCells("BD{$startRow}:{$middleEndCol}{$startRow}")
        ->setCellValue("BD{$startRow}", "ALKYL AMINES CHEMICAL LIMITED\nKURKUMBH\n$title");

    $secondRow=$startRow+1;
    //Description row height
    // $sheet->getRowDimension($startRow+1)->setRowHeight(60);
    //Tag Description //Table 1

    $sheet->mergeCells("BB2:BC2")->setCellValue("BB2", "Crude DES Tank");//10
    $sheet->mergeCells("BD2:BE2")->setCellValue("BD2", "ST-6702B");//11
    $sheet->mergeCells("BF2:BN2")->setCellValue("BF2", "DES Vaccum pump parameters");//12
    //Units
    $sheet->mergeCells("BB4:BB4")->setCellValue("BB4", "%");//13
    $sheet->mergeCells("BC4:BC4")->setCellValue("BC4", "%");//14
    $sheet->mergeCells("BD4:BD4")->setCellValue("BD4", "%");//3
    $sheet->mergeCells("BE4:BE4")->setCellValue("BE4", "°C");//4,5,6,7
    $sheet->mergeCells("BF4:BF4")->setCellValue("BF4", "On/Off");//4,5,6,7
    $sheet->mergeCells("BG4:BG4")->setCellValue("BG4", "AMP");//8
    $sheet->mergeCells("BH4:BH4")->setCellValue("BH4", "RPM");//8
    $sheet->mergeCells("BI4:BI4")->setCellValue("BI4", "AMP");//9
    $sheet->mergeCells("BJ4:BJ4")->setCellValue("BJ4", "RPM");//10
    $sheet->mergeCells("BK4:BL4")->setCellValue("BK4", "Vaccum Trap\nT-6537");//11
    $sheet->mergeCells("BM4:BM4")->setCellValue("BM4", "mmHg");//10
    $sheet->mergeCells("BN4:BN4")->setCellValue("BN4", "°C");//10


    //Style tags Description
    $sheet->getStyle("BA2:BK2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle("BH2:BK2")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER);

    // Right side (Document info)
    $sheet->mergeCells("{$rightStartCol}{$startRow}:{$lastCol}{$startRow}")
        ->setCellValue("{$rightStartCol}{$startRow}",
            "Document No:\nDocument Name:\nIssued No:\nRevision No:\nPage No: 05 of 05");

    $sheet->getStyle("BA{$startRow}:{$lastCol}{$startRow}")->applyFromArray([
     'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("BA{$secondRow}:{$lastCol}{$secondRow}")->applyFromArray([
    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);
    $sheet->getStyle("BD{$startRow}:{$middleEndCol}{$startRow}")->getFont()->setBold(true);

    // Align header
    $sheet->getStyle("BA{$startRow}:BC{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT);

    $sheet->getStyle("BD{$startRow}:{$middleEndCol}{$startRow}")->getAlignment()
          ->setVertical(Alignment::VERTICAL_CENTER)
          ->setHorizontal(Alignment::HORIZONTAL_CENTER)
          ->setWrapText(true);

    $sheet->getStyle(($colCount > 3 ? $middleEndCol : "BD") . "{$startRow}:{$lastCol}{$startRow}")
          ->getAlignment()->setVertical(Alignment::VERTICAL_TOP)
          ->setHorizontal(Alignment::HORIZONTAL_LEFT)
          ->setWrapText(true);

    $sheet->getRowDimension($startRow)->setRowHeight(75);
    // ------------------ Column Headers ------------------
    $headerRow = $secondRow + 1;
    $sheet->fromArray($headers, NULL, "BA$headerRow");
    $sheet->getStyle("BA$headerRow:{$lastCol}$headerRow")->getFont()->setBold(true);

    // ------------------ Data rows ------------------
    $rowNum = $headerRow + 2;
    $rowCount=0;
    $rarray=[];
    while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
        $excelRow = [];
        foreach ($columns as $colIndex => $col) {
            $value = $row[$col];
            if ($colIndex === 0 && $value instanceof DateTime) {
                $value = $value->format('d-m-Y h:i A');
            }
            if ($value === null || $value === '') {
                $value = 0;
            }
            if ($col === "VP6503ONOFF_VAL0") {
                if ($value == 0) {
                    $value = "OFF";
                } elseif ($value == 1) {
                    $value = "ON";
                }
            }
            $excelRow[] = $value;
        }
        $sheet->fromArray($excelRow, NULL, "BA$rowNum");
        $rowNum++;
    }
    $lastRow = $rowNum - 1;
    // echo $lastRow;
    $sheet->getStyle("BA$headerRow:{$lastCol}$lastRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
        'alignment' => [
            'horizontal' => Alignment::HORIZONTAL_CENTER,
            'vertical'   => Alignment::VERTICAL_CENTER,
        ],
    ]);
    // ------------------ Footer ------------------
    $footerRow = $lastRow + 1;
    $sheet->mergeCells("BA$footerRow:{$lastCol}$footerRow")
          ->setCellValue("BA$footerRow", "CONTROLLER (Name and Sign):");
    $sheet->getStyle("BA$footerRow:{$lastCol}$footerRow")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    $footerRow2 = $footerRow + 1;
    $sheet->mergeCells("BA$footerRow2:{$lastCol}$footerRow2")
          ->setCellValue("BA$footerRow2", "SIC (Name and Sign):");
    $sheet->getStyle("BA$footerRow2:{$lastCol}$footerRow2")->applyFromArray([
        'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
    ]);

    return $footerRow2 + 2;
}
//Table 5
$sql5 = "WITH ABC AS (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY DATEDIFF(MINUTE, '$selectedDate $ftime', [timestamp]) / 120
                    ORDER BY [timestamp] ASC
                ) AS rn
            FROM [master].[dbo].[TRENDS]
            WHERE 
                [timestamp] >= '$selectedDate $ftime'
                AND [timestamp] <= '$selectedDate $ttime'
        )
        SELECT  
                [timestamp],
                [LTT6505_VAL0],
                [LTT6506_VAL0],
                [LTST6702B_VAL0],
                [TIST6702B_VAL0],
                [VP6503ONOFF_VAL0],
                [VP6502AAMP_VAL0],
                [VP6502ARPM_VAL0],
                [VP6502BAMP_VAL0],
                [VP6502BRPM_VAL0],
                [PTT6537_VAL0],
                [TET6537_VAL0],
                [VT01_VAL0],
                [TT011_VAL0]
        FROM ABC
        WHERE rn = 1
        ORDER BY [timestamp] ASC;";
$nextStartRow = renderTable5($sheet, $conn, $sql5, "DES WFE PLANT CONTROLLER LOGSHEET \n", 1, $selectedDate);
// Save File
$filename = date('d-m-Y_H-i') . ".xlsx";
$dateFolder = date('d-m-Y');
$basePath = "D:/Avinash/REPORT/";
$fullPath = $basePath . $dateFolder . '/';
if (!is_dir($fullPath)) {
    mkdir($fullPath, 0777, true);
}
$excelFilePath = $fullPath . $filename;

$writer = new Xlsx($spreadsheet);
$writer->save($excelFilePath);


?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Excel Created</title>
    <link rel="stylesheet" href="generatePdf.css">
</head>
<body>
    <div class="container">
        <?php if (file_exists($excelFilePath)) { ?>
            <p class="success-msg">Excel created successfully!</p>
            <?php echo "<br><br>Path:".$excelFilePath; ?></div>
            <?php } ?>
    </div>
</body>
</html>