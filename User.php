<?php
session_start();
?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Generate Report</title>
    <link rel="stylesheet" href="User.css">
</head>
<body>

    <!-- Header Section -->
    <div class="header-container">
        <!-- Left Logo -->
        <img src="alkyl.png" alt="ALKYL Logo" class="left-logo">

        <!-- Right Logo -->
        <img src="eci.png" alt="ECI Logo" class="right-logo">

        <!-- Company Title -->
        <p class="company-title">ALKYL AMINES CHEMICAL LIMITED</p>
        <p class="sub-title">KURKUMBH</p> 

        <!-- Logsheet Title -->
        <p class="logsheet-title">DES WFE PLANT CONTROLLER LOGSHEET</p>
    </div>
    <!-- Form Section -->
    <div class="form-container">
        <form action="process.php" method="post">
            <label for="date">Select Date:</label>
            <input type="date" name="date" id="date" required> <br><br>
            
            <label for="from_time">Select From Time:</label>
            <input type="time" name="ftime" id="ftime" required> <br><br>
            
            <label for="to_time">Select To Time:</label>
            <input type="time" name="ttime" id="ttime" required> <br><br>

            <label for="itime">Select Time Interval:</label>
            <input type="time" name="itime" id="itime" required> <br><br>
            
            <button type="submit">Generate PDF</button>
        </form>
    </div>
</body>
</html>
<script> 
    // Get today's date
    const today = new Date();
    const dd = String(today.getDate()).padStart(2, '0');
    const mm = String(today.getMonth() + 1).padStart(2, '0');
    const yyyy = today.getFullYear();
    const todayDate = `${yyyy}-${mm}-${dd}`;

    // Calculate last 30 days
    const last30 = new Date();
    last30.setDate(today.getDate() - 30);
    const dd30 = String(last30.getDate()).padStart(2, '0');
    const mm30 = String(last30.getMonth() + 1).padStart(2, '0');
    const yyyy30 = last30.getFullYear();
    const minDate = `${yyyy30}-${mm30}-${dd30}`;

    // Set min & max attributes dynamically
    const dateInput = document.getElementById("date");
    dateInput.max = todayDate;
    dateInput.min = minDate;
</script>
