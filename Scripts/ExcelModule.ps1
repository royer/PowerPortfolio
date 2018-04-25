# import Excel Module
# read a worksheet from excel file.
# this function used Excel ComObject. it needs Excel installed in your system.
function Import-Excel {
    param ([string]$filename
        , [string]$sheetname 
        , [bool]$FirstRowIsHeader = $true
    )


    if ($filename -eq "") {
        throw "Please provide path of Excel File."
        exit
    }
    if (-not (Test-Path $filename)) {
        throw "File '$filename' does not exist."
        exit
    }

    $filename = Resolve-Path $filename;

    $ExcelApp = New-Object -ComObject excel.application;
    $ExcelApp.Visible = $false;
    $workbook = $ExcelApp.workbooks.open($filename,$null,$true);

    if (-not $sheetname) {
        Write-Warning "Not provide sheetname. use actived sheet.";
        $sheet = $workbook.ActiveSheet
    } else {
        $sheet = $workbook.sheets.Item($sheetname);
    }
    if (-not $sheet) {
        throw "Unable to open worksheet $sheetname .";
        exit;
    }

    $cols =$sheet.UsedRange.Columns.Count;
    $rows = $sheet.UsedRange.Rows.Count;

    $Headers = @();

    for ($col = 1; $col -le $cols; $col++) {
        if ($FirstRowIsHeader) {
            $column = $sheet.Cells.Item(1,$col).Value().trim()
        } else {
            $column = "Column"+$col;
        }
        $Headers += $column;
    }

    $initRow = 1;
    if ($FirstRowIsHeader -eq $true) {$initRow = 2;}

    $DataTable = @();
    for ($row = $initRow; $row -le $rows; $row++) {
        $objRow = New-Object psobject;
        for ($col = 1; $col -le $cols; $col++) {
            $val  = $sheet.Cells.Item($row, $col).Value();
            
            $objRow | Add-Member -MemberType NoteProperty -Name $Headers[$col-1] -Value $val; 
        }
        $DataTable += $objRow;
    } 

    $s= $workbook.Close();
    $s = $ExcelApp.Quit();
    $s = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ExcelApp);


    return ,$DataTable;
}

function Import-Excel-ListObject {
    param (
        [string]$filename,
        [string]$sheetname,
        [string]$listname
    )

    if ($filename -eq "") {
        throw "Please provide path of Excel File."
        exit
    }
    if (-not (Test-Path $filename)) {
        throw "File '$filename' does not exist."
        exit
    }

    if (-not $listname) { return $null}


    
    $filename = Resolve-Path $filename;

    $app = New-Object -ComObject excel.application;
    $workbook = $app.workbooks.open($filename, $null, $true)

    $DataTable = @()
    $sheet = $null
    $listObject = $null
    if ($sheetname) {
        $sheet = $workbook.sheets.item($sheetname);
        $listObject = $sheet.ListObjects.item($listname);
    } else {
        foreach ($s in $workbook.sheets) {
            foreach ($listo in $s.ListObjects) {
                if ($listo.name -eq $listname) {
                    $sheet = $s;
                    $listObject = $listo;
                    break;
                }
            }
            if ($listObject) {
                break;
            }
        }
    }

    if (-not $listObject) {
        throw "Unable to find ListObject $listname .";
        exit;
    }

    $Header = @();
    foreach($c in $listObject.HeaderRowRange.Columns) {
        $Header += $c.Value()
    }
    foreach ($row in $listObject.DataBodyRange.Rows) {
        $objrow = New-Object psobject;
        $col = 0
        foreach ($c in $row.cells) {
            $objrow | Add-Member -MemberType NoteProperty -Name $Header[$col] -Value $c.Value();
            $col += 1
        }
        $DataTable += $objrow;
    }

    $r = $workbook.Close()
    $r = $app.Quit()
    $r = [System.Runtime.Interopservices.Marshal]::ReleaseComObject($app);

    return ,$DataTable
}
