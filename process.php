<?php
session_start();
if (isset($_POST['date']) && isset($_POST['ftime']) && isset($_POST['ttime']) && isset($_POST['itime'])) {
    $_SESSION['selectedDate'] = $_POST['date'];
    $_SESSION['ftime'] = $_POST['ftime'];
    $_SESSION['ttime'] = $_POST['ttime'];
    $itime=$_POST['itime'];
    list($hr,$min)=explode(":",$itime);
    $totalmin=($hr*60)+$min;
    $_SESSION['itime']=$totalmin;
    ob_start();
    include 'fetchData.php';
    ob_end_clean();

    header("Location: generatePdf.php");
    exit();
} else {
    echo "Invalid Request!";
}
?>