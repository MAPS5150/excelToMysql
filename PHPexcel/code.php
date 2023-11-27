<?php
    session_start();
    
    $con = mysqli_connect('localhost', 'root', 'ra1zc0mpleja', 'escuela');

    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Spreadsheet;
    use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

    if(isset($_POST['save_excel_data'])){
        $fileName = $_FILES['import_file']['name'];
        $file_ext = pathinfo($fileName, PATHINFO_EXTENSION);

        $allowed_ext = ['xls', 'csv', 'xlsx'];

        if(in_array($file_ext, $allowed_ext)){
            $inputFileNamePath = $_FILES['import_file']['tmp_name'];
            $spreadsheet = \PhpOffice\PhpSpreadsheet\IOFactory::load($inputFileNamePath);
            $data = $spreadsheet->getActiveSheet()->toArray();

            $counter = "0";
            foreach($data as $row){
                if($counter > 0){
                    $fullname = $row['0'];
                    $email = $row['1'];
                    $phone = $row['2'];
                    $course = $row['3'];

                    $Query = "SELECT * FROM `students` WHERE email='$email'";
                    $output = mysqli_query($con, $Query);

                    if(mysqli_num_rows($output) > 0){

                    }else{
                        $studentQuery = "INSERT INTO `students` (fullname, email, phone, course) VALUES ('$fullname', '$email', '$phone', '$course')";
                        $result = mysqli_query($con, $studentQuery);
                        $msg = true;
                    }
                }else{
                    $counter = "1";
                }
            }
            if(isset($msg)){
                $_SESSION['message'] = "El archivo se importo exitosamente";
                header('Location: index.php');
                exit(0);
            }else{
                $_SESSION['message'] = "Archivo no importado";
                header('Location: index.php');
                exit(0);
            }

        }else{
            $_SESSION['message'] = "Archivo invalido";
            header('Location: index.php');
            exit(0);
        }

    }

?>