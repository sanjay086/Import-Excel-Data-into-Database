<?php
error_reporting(E_ALL);
ini_set('display_errors', 1);

session_start();
include('dbconfig.php');
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

if (isset($_POST['save_excel_data'])) {
    $fileName = $_FILES['import_file']['name'];
    $file_ext = pathinfo($fileName, PATHINFO_EXTENSION);

    $allowed_ext = ['xls', 'csv', 'xlsx'];

    if (in_array($file_ext, $allowed_ext)) {
        $inputFileNamePath = $_FILES['import_file']['tmp_name'];
        $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileNamePath);
        $data = $spreadsheet->getActiveSheet()->toArray();
       // echo "Hiiiiiiiiiii";
        foreach ($data as $row) {
            $SLNO = $row['0'];
            $AI_RANK = $row['1'];
            $rollno = $row['2'];
            $appno = $row['3'];
            $CNAME = $row['4'];
            $FNAME = $row['5'];
            $MNAME = $row['6'];
            $DOB = $row['7'];
            $GENDER = $row['8'];
            $CAT = $row['9'];
            $PWD = $row['10'];
            $NATION = $row['11'];
            $STATE_RES = $row['12'];
            $TOT_MRK = $row['13'];
            $PSCORE = $row['14'];
            $MRK_FIG = $row['15'];
            $PS_FIG = $row['16'];
            $RESULT = $row['17'];
            $CAT_QUAL = $row['18'];
            $YEAR = $row['19'];

           $studentQuery = "INSERT INTO GPAT (SLNO, AI_RANK, rollno, appno, CNAME, FNAME, MNAME, DOB, GENDER, CAT, PWD, NATION, STATE_RES, TOT_MRK, PSCORE, MRK_FIG, PS_FIG, RESULT, CAT_QUAL, YEAR) 
                            VALUES ('$SLNO', '$AI_RANK', '$rollno', '$appno', '$CNAME', '$FNAME', '$MNAME', '$DOB', '$GENDER', '$CAT', '$PWD', '$NATION', '$STATE_RES', '$TOT_MRK', '$PSCORE', '$MRK_FIG', '$PS_FIG', '$RESULT', '$CAT_QUAL', '$YEAR')";
            
            $result = mysqli_query($con, $studentQuery);

            if (!$result) {
                $_SESSION['message'] = "Error in query: " . mysqli_error($con);
                header('Location: index.php');
                exit(0);
            }
        }

        $_SESSION['message'] = "Successfully Imported";
        header('Location: index.php');
        exit(0);
    } else {
        $_SESSION['message'] = "Invalid File";
        header('Location: index.php');
        exit(0);
    }
}
?>
