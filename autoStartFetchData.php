<?php
session_start();
$selectedDate = date('Y-m-d');
$serverName = "DESKTOP-7L13FDU\CIMPLICITY";
$connectionOptions = [
    "Database" => "master", 
    // "Uid" => "sa",
    // "PWD" => "rechner@123"
];
$conn = sqlsrv_connect($serverName, $connectionOptions);
if ($conn === false) {
    die(print_r(sqlsrv_errors(), true));
}
    $sql = "WITH ABC AS
        (
            SELECT *,
                ROW_NUMBER() OVER (
                    PARTITION BY DATEDIFF(MINUTE, CAST(CAST(getdate() AS DATE) AS DATETIME), [timestamp])/120
                    ORDER BY [timestamp] ASC
                ) AS rn
            FROM [master].[dbo].[TRENDS]
            WHERE [timestamp] >= DATEADD(hour, -8, GETDATE())
            AND CAST([timestamp] AS DATE) = CAST(GETDATE() AS DATE)
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
            [FCV_WFE_6501_OPENING_VAL0],
            [TET6533A_VAL0],
            [TET6533B_VAL0]
        FROM ABC
        WHERE rn = 1
        ORDER BY [timestamp] ASC;";
        $params = [$selectedDate];
        $stmt = sqlsrv_query($conn, $sql, $params);
        if ($stmt === false) {
            die(print_r(sqlsrv_errors(), true));
        }
?>
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Project php</title>
  <link rel="stylesheet" href="style.css">
</head>
<body>
<!-- Data Table 1 -->
<table id="table1" border="1" style="width: 100%; border-collapse: collapse; text-align: center; height: auto;">
       <thead>
            <tr>
                <th style="width: 9%; height: 20px; text-align: left; vertical-align: center;">
                    Date: <?php echo "<br>".$selectedDate?>
                </th>

                <th colspan="8" style="position: relative; text-align: center;">
                    <div style="font-weight: bold;">ALKYL AMINES CHEMICAL LIMITED</div>
                    <div><br>KURKUMBH<br><br></div>
                    <p style="border: 1px solid black; width: 50%; text-align: center; margin: 0 auto;">
                        DES WFE PLANT CONTROLLER LOGSHEET
                    </p>
                    <!-- <div style="position:relative; width:25%; margin-top:-5em;"><img src="eci.png" alt="ECI"></div> -->
                </th>
                <th colspan="3" style="text-align: left; vertical-align: top;">
                    Document No: <br> <hr style="margin: 2px 0;">
                    Document Name: <br> <hr style="margin: 2px 0;">
                    Issued No: <br> <hr style="margin: 2px 0;">
                    Revision No: <br> <hr style="margin: 2px 0;">
                    Page No: 01 of 06
                </th>
            </tr>
            <tr>
                <th colspan="12">DES WFE Distillation Logsheet</th>
            </tr>
            <tr>
                <td></td>
                <td>T-6538 Feed Tank Level</td>
                <td>T-6538 Feed Tank Temp</td>
                <td>T-6538 Feed Tank Pressure</td>
                <td colspan="4">Feed Pump</td>
                <td>WFE Feed Rate</td>
                <td>WFE Feed CV Open</td>
                <td colspan="2">Hot Water Tank Temprature</td>
            </tr>
        </thead>
    <tbody>
        <?php

            $columns = [];
            $headers = [];
            foreach (sqlsrv_field_metadata($stmt) as $field) {
                $colName = $field['Name'];
                if (substr($colName, -5) === '_VAL0') {
                    $displayName = substr($colName, 0, -5);
                } else {
                    $displayName = $colName;
                }
                if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";}
                if (stripos($displayName, "TET6538") !== false) {$displayName = "TI-T-6538";}
                if (stripos($displayName, "LTT6538") !== false) {$displayName = "LT-T-6538";}              
                if (stripos($displayName, "PTT6538") !== false) {$displayName = "PT-T-6538";}            
                if (stripos($displayName, "VFDP6538AAMP") !== false) {$displayName = "VFD-P-6538A";}       
                if (stripos($displayName, "VFDP6538BAMP") !== false) {$displayName = "VFD-P-6538B";}       
                if (stripos($displayName, "VFDP6538ARPM") !== false) {$displayName = "VFD-P-6538A";}       
                if (stripos($displayName, "VFDP6538BRPM") !== false) {$displayName = "VFD-P-6538B";}
                if (stripos($displayName, "MFMWFE6501") !== false) {$displayName = "MFM-WFE-6501";}
                if (stripos($displayName, "FCV_WFE_6501_OPENING") !== false) {$displayName = "FCV-WFE-6501";}
                if (stripos($displayName, "TET6533A") !== false) {$displayName = "TI-T-6533A";}
                if (stripos($displayName, "TET6533B") !== false) {$displayName = "TI-T-6533B";}
                $headers[] = $displayName;
                $columns[] = $colName;
            }
            echo "<thead><tr>";
            foreach ($headers as $header) {
                    if ($header==="Date & Time") {
                        echo "<td rowspan='2'>" . htmlspecialchars($header) . "</td>";
                    } else {
                        echo "<td>" . htmlspecialchars($header) . "</td>";
                    }
                }
                echo "</tr>";
                echo "<tr>";
                echo "<th>%</th>";echo "<th>°C</th>";echo "<th>kg/cm²g</th>"; echo "<th>AMP</th>";echo "<th>AMP</th>";echo "<th>RPM</th>";echo "<th>RPM</th>"; echo "<th>Kg/Hr</th>"; echo "<th>%</th>";
                echo "<th>°C</th>"; echo "<th>°C</th>";
                //echo "<th>kg/cm²g</th>";echo "<th>%</th>";
                echo "</tr>";
            echo "</tr></thead>";
            echo "<tbody>";
            $rowIndex = 0;
            $startTime = null;

            while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
                echo "<tr>";
                foreach ($columns as $colIndex => $col) {
                    $value = $row[$col];
                    if ($colIndex === 0 && $value instanceof DateTime) {
                        if ($value instanceof DateTime) {                    
                            $value = $value->format('h:i A');
                            }
                        } else {
                        if ($value instanceof DateTime) {
                            if (stripos($col, "time") !== false) {
                                $value = $value->format('H:i');  
                            } else {
                                $value = $value->format('H:i'); 
                            }
                        } elseif (is_string($value) && stripos($col, "time") !== false) {
                            $dateTime = DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('h:i A', $value);
                            if ($dateTime) {
                                $value = $dateTime->format('H:i');  
                            }
                        }
                    }
                    //Show only 2 digits after the decimal point
                    if ($col==="timestamp") {echo "<td>" . htmlspecialchars((string)$value) . "</td>";}
                    else{echo "<td>" . htmlspecialchars(sprintf("%.2f",(float)$value)) . "</td>";}
                    if ($value === null || $value === '') {
                        $value = 0;
                    }

                    // echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                }
                echo "</tr>";
                $rowIndex++;
            }
            echo "</tbody>";
            sqlsrv_free_stmt($stmt);
        ?>
        <!-- Signature Rows -->
        <tr>
            <td colspan="3" style="height:20px;">CONTROLLER(Name and Sign):</td>
            <td colspan="9"></td>
        </tr>
        <tr>
            <td colspan="3" style="height:20px;">SIC(Name and Sign):</td>
        </tr>
    </tbody>
</table> <br><br><br>
<!-- Data table 2 -->
<?php
        $sql = "WITH ABC AS
            (
                SELECT *,
                    ROW_NUMBER() OVER (
                        PARTITION BY DATEDIFF(MINUTE, CAST(CAST(getdate() AS DATE) AS DATETIME), [timestamp])/120
                        ORDER BY [timestamp] ASC
                    ) AS rn
                FROM [master].[dbo].[TRENDS]
                WHERE [timestamp] >= DATEADD(hour, -8, GETDATE())
                AND CAST([timestamp] AS DATE) = CAST(GETDATE() AS DATE)
            )
                SELECT  
                    [timestamp],
                    [PTT65331_VAL0],
                    [LTT6533_VAL0],
                    [TET6533A_VAL0],
                    [FTHWSWFE6501_VAL0],
                    [FCV_HWR_WFE_6501_OPENING_VAL0],
                    [TEHWSWFE6501_VAL0],
                    [TEHWRWFE6501_VAL0],
                    [PTWFE6501_VAL0],
                    [TEWFE65011_VAL0],
                    [TEWFE65012_VAL0]
                    -- [TEWFE65013_VAL0]
                    -- [TEHE6503_VAL0],
                    -- [FTCHBSHE6503_VAL0],
                    -- [FTCHBSHC6501_VAL0],
                    -- [TEHC6501B_VAL0],
                    -- [TEHC6501A_VAL0]
                FROM ABC
                WHERE rn = 1
                ORDER BY [timestamp] ASC;";
        $params = [$selectedDate];
        $stmt = sqlsrv_query($conn, $sql, $params);
        if ($stmt === false) {
            die(print_r(sqlsrv_errors(), true));
        }
?>
<table  border="1" style="width: 100%; border-collapse: collapse; text-align: center; height: auto;">
       <thead>
            <tr>
                <th style="width: 9%; height: 20px; text-align: left; vertical-align: center;">
                    Date: <?php echo "<br>".$selectedDate?>
                </th>

                <th colspan="7" style="position: relative; text-align: center;">
                    <div style="font-weight: bold;">ALKYL AMINES CHEMICAL LIMITED</div>
                    <div><br>KURKUMBH<br><br></div>
                    <p style="border: 1px solid black; width: 50%; text-align: center; margin: 0 auto;">
                        DES WFE PLANT CONTROLLER LOGSHEET
                    </p>
                </th>
                <th colspan="3" style="text-align: left; vertical-align: top;">
                    Document No: <br> <hr style="margin: 2px 0;">
                    Document Name: <br> <hr style="margin: 2px 0;">
                    Issued No: <br> <hr style="margin: 2px 0;">
                    Revision No: <br> <hr style="margin: 2px 0;">
                    Page No: 02 of 06
                </th>
                <tr>
                    <th colspan="11">DES WFE Distillation Logsheet</th>
                </tr>
            <tr>
                <td></td>
                <!-- <td>Hot Water Tank Temprature</td> -->
                <td>Hot Water Tank Pressure</td>
                <td>Hot Water Tank Level</td>
                <td>Hot Water tank steam CV open</td>
                <td>Flow rate of Hot water to WFE</td>
                <td>Hot Water flow CV Open</td>
                <td>WFE Hot Water Supply temp</td>
                <td>WFE Hot Water return temp</td>
                <td>WFE Pressure</td>
                <td>Cyclone Bottom Line</td>
                <td>WFE Bottom temp</td>
                <!-- <td>Vapour Line temp</td> -->
                <!-- <td>Double Pipe Outlet</td> -->
                <!-- <td>Double Pipe HE</td>
                <td>Chilling Flow to Condenser</td>
                <td>DES Collection Line temp</td>
                <td>Chilling Condenser Vapour Outlet</td>       -->
            </tr>
            </tr>
        </thead>
    <tbody>
        <?php

            $columns = [];
            $headers = [];
            foreach (sqlsrv_field_metadata($stmt) as $field) {
                $colName = $field['Name'];
                if (substr($colName, -5) === '_VAL0') {
                    $displayName = substr($colName, 0, -5);
                } else {
                    $displayName = $colName;
                }
                if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";}
                // if (stripos($displayName, "TET6533B") !== false) {$displayName = "TI-T-6533B";}
                if (stripos($displayName, "PTT65331") !== false) {$displayName = "PT-T-6533";}
                if (stripos($displayName, "LTT6533") !== false) {$displayName = "LT-T-6533";}
                if (stripos($displayName, "TET6533A") !== false) {$displayName = "TCV-T-6533";}
                if (stripos($displayName, "FTHWSWFE6501") !== false) {$displayName = "FI-HWS-6501";}
                if (stripos($displayName, "FCV_HWR_WFE_6501_OPENING") !== false) {$displayName = "FCV-HWR-OPEN";}
                if (stripos($displayName, "TEHWSWFE6501") !== false) {$displayName = "TE-HWS-WFE-6501";}
                if (stripos($displayName, "TEHWRWFE6501") !== false) {$displayName = "TE-HWR-WFE-6501";}
                if (stripos($displayName, "PTWFE6501") !== false) {$displayName = "PT-WFE-6501";}
                if (stripos($displayName, "TEWFE65011") !== false) {$displayName = "TI-WFE-6501-1";}
                if (stripos($displayName, "TEWFE65012") !== false) {$displayName = "TI-WFE-6501-2";}
                if (stripos($displayName, "TEWFE65013") !== false) {$displayName = "TI-WFE-6501-3";}
                $headers[] = $displayName;
                $columns[] = $colName;
            }
            echo "<thead><tr>";
            foreach ($headers as $header) {
                if($header === "Date & Time"){
                    echo "<td rowspan='2'>" . htmlspecialchars($header) . "</td>";
                }
                else{
                    echo "<td>" . htmlspecialchars($header) . "</td>";
                }
            }
            echo "<tr>";
            echo "<th>kg/cm²g</th>";echo "<th>%</th>";
            echo "<th>%</th>";echo "<th>Kg/Hr</th>"; echo "<th>%</th>";echo "<th>°C</th>";echo "<th>°C</th>"; echo "<th>mmHg</th>";echo "<th>°C</th>";echo "<th>°C</th>";    
            // echo "<th>°C</th>";
            // echo "<th>°C</th>";echo "<th>kg/Hr</th>";
            //  echo "<th>Kg/Hr</th>"; echo "<th>°C</th>"; echo "<th>°C</th>"; 
            echo "</tr>";
            echo "</tr></thead>";
            echo "<tbody>";
            $rowIndex = 0;
            $startTime = null;

            while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
                echo "<tr>";
                foreach ($columns as $colIndex => $col) {
                    $value = $row[$col];
                    if ($colIndex === 0 && $value instanceof DateTime) {
                        if ($value instanceof DateTime) {
                        // $value = $value->format('d-m-Y:- h:i');
                            // if ((int)$value->format('i') >= 59) {
                            //     $value->modify('+'.(60 - (int)$value->format('i')).' minutes');
                            // }
                            $value = $value->format('h:i A');
                            }
                        } else {
                        if ($value instanceof DateTime) {
                            if (stripos($col, "time") !== false) {
                                $value = $value->format('H:i');  
                            } else {
                                $value = $value->format('H:i');
                            }
                        } elseif (is_string($value) && stripos($col, "time") !== false) {
                            $dateTime = DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('h:i A', $value);
                            if ($dateTime) {
                                $value = $dateTime->format('H:i');  
                            }
                        }
                    }
                    //Show only 2 digits after the decimal point
                    if ($col==="timestamp") {echo "<td>" . htmlspecialchars((string)$value) . "</td>";}
                    else{echo "<td>" . htmlspecialchars(sprintf("%.2f",(float)$value)) . "</td>";}

                    if ($value === null || $value === '') {
                        $value = 0;
                    }

                    //echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                }
                echo "</tr>";
                $rowIndex++;
            }
            echo "</tbody>";
            sqlsrv_free_stmt($stmt);
        ?>
        <!-- Signature Rows -->
        <tr>
            <td colspan="3" style="height:20px;">CONTROLLER(Name and Sign):</td>
            <td colspan="8"></td>
        </tr>
        <tr>
            <td colspan="3" style="height:20px;">SIC(Name and Sign):</td>
        </tr>
    </tbody>
</table> <br> <br>
<!-- Data table 3 -->
<?php
$sql = "WITH ABC AS
(
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY DATEDIFF(MINUTE, CAST(CAST(getdate() AS DATE) AS DATETIME), [timestamp])/120
               ORDER BY [timestamp] ASC
           ) AS rn
    FROM [master].[dbo].[TRENDS]
    WHERE [timestamp] >= DATEADD(hour, -8, GETDATE())
      AND CAST([timestamp] AS DATE) = CAST(GETDATE() AS DATE)
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

        $params = [$selectedDate];
        $stmt = sqlsrv_query($conn, $sql, $params);
        if ($stmt === false) {
            die(print_r(sqlsrv_errors(), true));
        }
?>
<table border="1" style="width: 100%; border-collapse: collapse; text-align: center; height: auto;">
       <thead>
            <tr>
                <th style="width: 9%; height: 20px; text-align: left; vertical-align: center;">
                    Date: <?php echo "<br>".$selectedDate?>
                </th>

                <th colspan="7" style="position: relative; text-align: center;">
                    <div style="font-weight: bold;">ALKYL AMINES CHEMICAL LIMITED</div>
                    <div><br>KURKUMBH<br><br></div>
                    <p style="border: 1px solid black; width: 50%; text-align: center; margin: 0 auto;">
                        DES WFE PLANT CONTROLLER LOGSHEET
                    </p>
                </th>
                <th colspan="3" style="text-align: left; vertical-align: top;">
                    Document No: <br> <hr style="margin: 2px 0;">
                    Document Name: <br> <hr style="margin: 2px 0;">
                    Issued No: <br> <hr style="margin: 2px 0;">
                    Revision No: <br> <hr style="margin: 2px 0;">
                    Page No: 03 of 06
                </th>
                <tr>
                    <th colspan="11">DES WFE Distillation Logsheet</th>
                </tr>
                <tr>
                    <td></td>
                <!-- <td>WFE Bottom temp</td> -->
                <!-- <td>WFE Bottom temp</td>-->
                <td>Vapour Line temp</td> 
                <td>Double Pipe Outlet</td>
                <td>Double Pipe HE</td>
                <td>Chilling Flow to Condenser</td>
                <td>DES Collection Line temp</td>
                <td>Chilling Condenser Vapour Outlet</td> 
                <td colspan="4">Concentrate Collection Tank</td>
                </tr>
                
            </tr>
        </thead>
    <tbody>
        <?php

            $columns = [];
            $headers = [];
            foreach (sqlsrv_field_metadata($stmt) as $field) {
                $colName = $field['Name'];
                if (substr($colName, -5) === '_VAL0') {
                    $displayName = substr($colName, 0, -5);
                } else {
                    $displayName = $colName;
                }
                if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";}
                if (stripos($displayName, "TEWFE65012") !== false) {$displayName = "TI-WFE-6501-2";}
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
                if (stripos($displayName, "TECHBS") !== false) {$displayName = "TI-CHBS";}

                $headers[] = $displayName;
                $columns[] = $colName;
            }
            echo "<thead><tr>";
            foreach ($headers as $header) {
                if($header === "Date & Time"){
                    echo "<td rowspan='2'>" . htmlspecialchars($header) . "</td>";
                }
                else{
                    echo "<td>" . htmlspecialchars($header) . "</td>";
                }
                    
            }
            echo "<tr>";
            echo "<th>°C</th>";echo "<th>°C</th>";echo "<th>kg/Hr</th>";
             echo "<th>Kg/Hr</th>"; echo "<th>°C</th>"; echo "<th>°C</th>"; 
            echo "<th>°C</th>";echo "<th>%</th>";echo "<th>mmHg</th>";echo "<th>Kg/Hr</th>";
            echo "</tr>";
            echo "</tr></thead>";
            echo "<tbody>";
            $rowIndex = 0;
            $startTime = null;

            while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
                echo "<tr>";
                foreach ($columns as $colIndex => $col) {
                    $value = $row[$col];
                    if ($colIndex === 0 && $value instanceof DateTime) {
                        if ($value instanceof DateTime) {
                        // $value = $value->format('d-m-Y:- h:i');
                            // if ((int)$value->format('i') >= 59) {
                            //     $value->modify('+'.(60 - (int)$value->format('i')).' minutes');
                            // }
                            $value = $value->format('h:i A');
                            }
                        } else {
                        if ($value instanceof DateTime) {
                            if (stripos($col, "time") !== false) {
                                $value = $value->format('H:i');  
                            } else {
                                $value = $value->format('H:i');
                            }
                        } elseif (is_string($value) && stripos($col, "time") !== false) {
                            $dateTime = DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('h:i A', $value);
                            if ($dateTime) {
                                $value = $dateTime->format('H:i');  
                            }
                        }
                    }
                    //Show only 2 digits after the decimal point
                    if ($col==="timestamp") {echo "<td>" . htmlspecialchars((string)$value) . "</td>";}
                    else{echo "<td>" . htmlspecialchars(sprintf("%.2f",(float)$value)) . "</td>";}

                    if ($value === null || $value === '') {
                        $value = 0;
                    }

                    //echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                }
                echo "</tr>";
                $rowIndex++;
            }
            echo "</tbody>";
            sqlsrv_free_stmt($stmt);
        ?>
        <!-- Signature Rows -->
        <tr>
            <td colspan="3" style="height:20px;">CONTROLLER(Name and Sign):</td>
            <td colspan="8"></td>
        </tr>
        <tr>
            <td colspan="3" style="height:20px;">SIC(Name and Sign):</td>
        </tr>
    </tbody>
</table> <br> <br>
<!-- Data table 4 -->
<?php
$sql = "WITH ABC AS
(
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY DATEDIFF(MINUTE, CAST(CAST(getdate() AS DATE) AS DATETIME), [timestamp])/120
               ORDER BY [timestamp] ASC
           ) AS rn
    FROM [master].[dbo].[TRENDS]
    WHERE [timestamp] >= DATEADD(hour, -8, GETDATE())
      AND CAST([timestamp] AS DATE) = CAST(GETDATE() AS DATE)
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

        $params = [$selectedDate];
        $stmt = sqlsrv_query($conn, $sql, $params);
        if ($stmt === false) {
            die(print_r(sqlsrv_errors(), true));
        }
?>
<table border="1" style="width: 100%; border-collapse: collapse; text-align: center; height: auto;">
       <thead>
            <tr>
                <th style="width: 9%; height: 20px; text-align: left; vertical-align: center;">
                    Date: <?php echo "<br>".$selectedDate?>
                </th>

                <th colspan="8" style="position: relative; text-align: center;">
                    <div style="font-weight: bold;">ALKYL AMINES CHEMICAL LIMITED</div>
                    <div><br>KURKUMBH<br><br></div>
                    <p style="border: 1px solid black; width: 50%; text-align: center; margin: 0 auto;">
                        DES WFE PLANT CONTROLLER LOGSHEET
                    </p>
                </th>
                <th colspan="3" style="text-align: left; vertical-align: top;">
                    Document No: <br> <hr style="margin: 2px 0;">
                    Document Name: <br> <hr style="margin: 2px 0;">
                    Issued No: <br> <hr style="margin: 2px 0;">
                    Revision No: <br> <hr style="margin: 2px 0;">
                    Page No: 04 of 06
                </th>
                <tr>
                    <th colspan="12">DES WFE Distillation Logsheet</th>
                </tr>
            </tr>
            <tr>
                <td></td>
                <td colspan="2">Chilling Tempratures</td>
                <td colspan="3">DES Collection Tank (T-6536)</td>
                <td colspan="2">Purging Tank (T-6502)</td>
                <td>ST 6504</td>
                <td colspan="2">Off Spec Storage Tank (T-6539)</td>
                <td>Hot water tank parameters</td>
            </tr>
        </thead>
    <tbody>
        <?php

            $columns = [];
            $headers = [];
            foreach (sqlsrv_field_metadata($stmt) as $field) {
                $colName = $field['Name'];
                if (substr($colName, -5) === '_VAL0') {
                    $displayName = substr($colName, 0, -5);
                } else {
                    $displayName = $colName;
                }
                if (stripos($displayName, "timestamp") !== false) {$displayName = "Date & Time";}
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
            echo "<thead><tr>";
            foreach ($headers as $header) {
                if($header === "Date & Time"){
                    echo "<td rowspan='2'>" . htmlspecialchars($header) . "</td>";
                }else{
                    echo "<td>" . htmlspecialchars($header) . "</td>";
                }

            }
            echo "</tr><tr>";
            echo "<th>°C</th>";echo "<th>°C</th>";echo "<th>%</th>";echo "<th>°C</th>";
            echo "<th>mmHg</th>";echo "<th>%</th>";echo "<th>°C</th>";echo "<th>%</th>";echo "<th>%</th>";echo "<th>Kg/cm²g</th>";echo "<th>Kg/cm²g</th>";
            echo "</tr></thead>";
            echo "<tbody>";
            $rowIndex = 0;
            $startTime = null;

            while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
                echo "<tr>";
                foreach ($columns as $colIndex => $col) {
                    $value = $row[$col];
                    if ($colIndex === 0 && $value instanceof DateTime) {
                        if ($value instanceof DateTime) {
                            $value = $value->format('h:i A');
                            }
                        } else {
                        if ($value instanceof DateTime) {
                            if (stripos($col, "time") !== false) {
                                $value = $value->format('H:i');  
                            } else {
                                $value = $value->format('H:i');
                            }
                        } elseif (is_string($value) && stripos($col, "time") !== false) {
                            $dateTime = DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('h:i A', $value);
                            if ($dateTime) {
                                $value = $dateTime->format('H:i');  
                            }
                        }
                    }
                    if ($col === "VP6503ONOFF_VAL0") {
                        if ($value == 0) {
                            $value = "OFF";
                        } elseif ($value == 1) {
                            $value = "ON";
                        }
                    }
                    //Show only 2 digits after the decimal point
                    if ($col==="timestamp") {echo "<td>" . htmlspecialchars((string)$value) . "</td>";}
                    else{echo "<td>" . htmlspecialchars(sprintf("%.2f",(float)$value)) . "</td>";}

                    if ($value === null || $value === '') {
                        $value = 0;
                    }

                    //echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                }
                echo "</tr>";
                $rowIndex++;
            }
            echo "</tbody>";
            sqlsrv_free_stmt($stmt);
        ?>
        <!-- Signature Rows -->
        <tr>
            <td colspan="3" style="height:20px;">CONTROLLER(Name and Sign):</td>
            <td colspan="9"></td>
        </tr>
        <tr>
            <td colspan="3" style="height:20px;">SIC(Name and Sign):</td>
        </tr>
    </tbody>
</table><br><br> 
<!-- Data table 5 -->
<?php
$sql = "WITH ABC AS
(
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY DATEDIFF(MINUTE, CAST(CAST(getdate() AS DATE) AS DATETIME), [timestamp])/120
               ORDER BY [timestamp] ASC
           ) AS rn
    FROM [master].[dbo].[TRENDS]
    WHERE [timestamp] >= DATEADD(hour, -8, GETDATE())
      AND CAST([timestamp] AS DATE) = CAST(GETDATE() AS DATE)
)
        SELECT  
                [timestamp],
                [LTT6505_VAL0],
                [LTT6506_VAL0],
                [LTST6702B_VAL0],
                [TIST6702B_VAL0],
                [VP_6503_ON_OFF_VAL0],
                [VP6502AAMP_VAL0]
        FROM ABC
        WHERE rn = 1
        ORDER BY [timestamp] ASC;";

        $params = [$selectedDate];
        $stmt = sqlsrv_query($conn, $sql, $params);
        if ($stmt === false) {
            die(print_r(sqlsrv_errors(), true));
        }
?>
<table border="1" style="width: 100%; border-collapse: collapse; text-align: center; height: auto;">
       <thead>
            <tr>
                <th style="width: 9%; height: 20px; text-align: left; vertical-align: center;">
                    Date: <?php echo "<br>".$selectedDate?>
                </th>

                <th colspan="4" style="position: relative; text-align: center;">
                    <div style="font-weight: bold;">ALKYL AMINES CHEMICAL LIMITED</div>
                    <div><br>KURKUMBH<br><br></div>
                    <p style="border: 1px solid black; width: 50%; text-align: center; margin: 0 auto;">
                        DES WFE PLANT CONTROLLER LOGSHEET
                    </p>
                </th>
                <th colspan="2" style="text-align: left; vertical-align: top;">
                    Document No: <br> <hr style="margin: 2px 0;">
                    Document Name: <br> <hr style="margin: 2px 0;">
                    Issued No: <br> <hr style="margin: 2px 0;">
                    Revision No: <br> <hr style="margin: 2px 0;">
                    Page No: 05 of 06
                </th>
                <tr>
                    <th colspan="7">DES WFE Distillation Logsheet</th>
                </tr>
            </tr>
            <tr>
                <td></td>
                <td colspan="2">Crude DES Tank</td>
                <td colspan="2">ST-6702B</td>
                <td colspan="2">DES Vaccum pump parameters</td>
            </tr>
        </thead>
    <tbody>
        <?php

            $columns = [];
            $headers = [];
            foreach (sqlsrv_field_metadata($stmt) as $field) {
                $colName = $field['Name'];
                if (substr($colName, -5) === '_VAL0') {
                    $displayName = substr($colName, 0, -5);
                } else {
                    $displayName = $colName;
                }
                if (stripos($displayName, "timestamp") !== false) {$displayName = "Date&Time";}
                if (stripos($displayName, "LTT6505") !== false) {$displayName = "LT-ST-6505";}
                if (stripos($displayName, "LTT6506") !== false) {$displayName = "LT-ST-6506";}
                if (stripos($displayName, "LTST6702B") !== false) {$displayName = "LT-ST-6702B";}
                if (stripos($displayName, "TIST6702B") !== false) {$displayName = "TI-ST-6702B";}
                if (stripos($displayName, "VP_6503_ON_OFF") !== false) {$displayName = "VP-6503";}
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
            echo "<thead><tr>";
            foreach ($headers as $header) {
                if($header === "Date&Time"){
                    echo "<td rowspan='2'>" . htmlspecialchars($header) . "</td>";
                }else{
                    echo "<td>" . htmlspecialchars($header) . "</td>";
                }

            }
            echo "</tr><tr>";echo "<th>%</th>";echo "<th>%</th>";
            echo "<th>%</th>";echo "<th>°C</th>";echo "<th>On/Off</th>";echo "<th>AMP</th>";
            echo "</tr></thead>";
            echo "<tbody>";
            $rowIndex = 0;
            $startTime = null;

            while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
                echo "<tr>";
                foreach ($columns as $colIndex => $col) {
                    $value = $row[$col];
                    if ($colIndex === 0 && $value instanceof DateTime) {
                        if ($value instanceof DateTime) {
                            $value = $value->format('h:i A');
                            }
                        } else {
                        if ($value instanceof DateTime) {
                            if (stripos($col, "time") !== false) {
                                $value = $value->format('H:i');  
                            } else {
                                $value = $value->format('H:i');
                            }
                        } elseif (is_string($value) && stripos($col, "time") !== false) {
                            $dateTime = DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('h:i A', $value);
                            if ($dateTime) {
                                $value = $dateTime->format('H:i');  
                            }
                        }
                    }
                    if ($col === "VP_6503_ON_OFF_VAL0") {
                        if ($value == 0) {
                            $value = "OFF";
                        } elseif ($value == 1) {
                            $value = "ON";
                        }
                    }
                    //Shown only 2 digits after decimal point
                    if ($col==="timestamp") {
                        echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                    }elseif ($col === "VP_6503_ON_OFF_VAL0") {
                            if ($value == 0) {
                                $value = "OFF";
                            } elseif ($value == 1) {
                                $value = "ON";
                            }
                        echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                        }
                    else{
                        echo "<td>" . htmlspecialchars(sprintf("%.2f",(float)$value)) . "</td>";
                    }
                    if ($value === null || $value === '') {
                        $value = 0;
                    }

                    //echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                }
                echo "</tr>";
                $rowIndex++;
            }
            echo "</tbody>";
            sqlsrv_free_stmt($stmt);
        ?>
        <!-- Signature Rows -->
        <tr>
            <td colspan="3" style="height:20px;">CONTROLLER(Name and Sign):</td>
            <td colspan="4"></td>
        </tr>
        <tr>
            <td colspan="3" style="height:20px;">SIC(Name and Sign):</td>
        </tr>
    </tbody>
</table><br><br>
<!-- Data table 6 -->
<?php
$sql = "WITH ABC AS
(
    SELECT *,
           ROW_NUMBER() OVER (
               PARTITION BY DATEDIFF(MINUTE, CAST(CAST(getdate() AS DATE) AS DATETIME), [timestamp])/120
               ORDER BY [timestamp] ASC
           ) AS rn
    FROM [master].[dbo].[TRENDS]
    WHERE [timestamp] >= DATEADD(hour, -8, GETDATE())
      AND CAST([timestamp] AS DATE) = CAST(GETDATE() AS DATE)
)
        SELECT  
                [timestamp],
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

        $params = [$selectedDate];
        $stmt = sqlsrv_query($conn, $sql, $params);
        if ($stmt === false) {
            die(print_r(sqlsrv_errors(), true));
        }
?>
<table border="1" style="width: 100%; border-collapse: collapse; text-align: center; height: auto;">
       <thead>
            <tr>
                <th style="width: 9%; height: 20px; text-align: left; vertical-align: center;">
                    Date: <?php echo "<br>".$selectedDate?>
                </th>

                <th colspan="5" style="position: relative; text-align: center;">
                    <div style="font-weight: bold;">ALKYL AMINES CHEMICAL LIMITED</div>
                    <div><br>KURKUMBH<br><br></div>
                    <p style="border: 1px solid black; width: 50%; text-align: center; margin: 0 auto;">
                        DES WFE PLANT CONTROLLER LOGSHEET
                    </p>
                </th>
                <th colspan="2" style="text-align: left; vertical-align: top;">
                    Document No: <br> <hr style="margin: 2px 0;">
                    Document Name: <br> <hr style="margin: 2px 0;">
                    Issued No: <br> <hr style="margin: 2px 0;">
                    Revision No: <br> <hr style="margin: 2px 0;">
                    Page No: 06 of 06
                </th>
                <tr>
                    <th colspan="8">DES WFE Distillation Logsheet</th>
                </tr>
            </tr>
            <tr>
                <td></td>
                <td colspan="7">DES Vaccum Pump Parameters</td>
            </tr>
        </thead>
    <tbody>
        <?php

            $columns = [];
            $headers = [];
            foreach (sqlsrv_field_metadata($stmt) as $field) { 
                $colName = $field['Name'];
                if (substr($colName, -5) === '_VAL0') {
                    $displayName = substr($colName, 0, -5);
                } else {
                    $displayName = $colName;
                }
                if (stripos($displayName, "timestamp") !== false) {$displayName = "Date&Time";}
                if (stripos($displayName, "LTT6505") !== false) {$displayName = "LTST6505";}
                if (stripos($displayName, "LTT6506") !== false) {$displayName = "LTST6506";}
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
            echo "<thead><tr>";
            foreach ($headers as $header) {
                if($header === "Date&Time"){
                    echo "<td rowspan='2'>" . htmlspecialchars($header) . "</td>";
                }else{
                    echo "<td>" . htmlspecialchars($header) . "</td>";
                }

            }
            echo "</tr><tr>";echo "<th>RPM</th>";echo "<th>AMP</th>";
            echo "<th>RPM</th>";echo "<th colspan='2'>Vaccum Trap T-6537</th>";echo "<th>mmHg</th>";echo "<th>°C</th>";
            echo "</tr></thead>";
            echo "<tbody>";
            $rowIndex = 0;
            $startTime = null;

            while ($row = sqlsrv_fetch_array($stmt, SQLSRV_FETCH_ASSOC)) {
                echo "<tr>";
                foreach ($columns as $colIndex => $col) {
                    $value = $row[$col];
                    if ($colIndex === 0 && $value instanceof DateTime) {
                        if ($value instanceof DateTime) {
                            $value = $value->format('h:i A');
                            }
                        } else {
                        if ($value instanceof DateTime) {
                            if (stripos($col, "time") !== false) {
                                $value = $value->format('H:i');  
                            } else {
                                $value = $value->format('H:i');
                            }
                        } elseif (is_string($value) && stripos($col, "time") !== false) {
                            $dateTime = DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('H:i', $value)
                                    ?: DateTime::createFromFormat('h:i A', $value);
                            if ($dateTime) {
                                $value = $dateTime->format('H:i');  
                            }
                        }
                    }
                    if ($col === "VP6503ONOFF_VAL0") {
                        if ($value == 0) {
                            $value = "OFF";
                        } elseif ($value == 1) {
                            $value = "ON";
                        }
                    }if ($col==="timestamp") {echo "<td>" . htmlspecialchars((string)$value) . "</td>";}
                    else{echo "<td>" . htmlspecialchars(sprintf("%.2f",(float)$value)) . "</td>";}
                    if ($value === null || $value === '') {
                        $value = 0;
                    }

                    //echo "<td>" . htmlspecialchars((string)$value) . "</td>";
                }
                echo "</tr>";
                $rowIndex++;
            }
            echo "</tbody>";
            sqlsrv_free_stmt($stmt);
        ?>
        <!-- Signature Rows -->
        <tr>
            <td colspan="3" style="height:20px;">CONTROLLER(Name and Sign):</td>
            <td colspan="5"></td>
        </tr>
        <tr>
            <td colspan="3" style="height:20px;">SIC(Name and Sign):</td>
        </tr>
    </tbody>
</table>
</body>
</html>
