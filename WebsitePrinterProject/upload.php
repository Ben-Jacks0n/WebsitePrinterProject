<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/4.7.0/css/font-awesome.min.css">
    <title>3RPrint</title>
    <link rel="stylesheet" href="style.css">
</head>
<body>
    <div class="wrapper">
        <div class="box">
            <div class="input-bx">
                <h2 class="upload-area-title">3R Printer</h2>
                <img src="printer.jpeg" alt="image" class="printer">
                    <label>
                        <h2 class="phpDone">Your file is being printed. Check printer</h2>
                    </label>
                  
            </div>
        </div>


    </div>
    <?php
        $targetPath = "uploads/" . basename($_FILES["document"]["name"]);
        move_uploaded_file($_FILES["document"]["tmp_name"], $targetPath);

        // Run the Python script with the uploaded file path
        $pythonScript = "print.py";
        $command = "python $pythonScript \"$targetPath\"";
        exec($command);
    ?>

    <script src="script.js"></script>
</body>
</html>
