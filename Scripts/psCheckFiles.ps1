# 2017-Sep-10. Created by Maxim T. 
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: CheckFiles".PadRight(28);

$errFileDupl = $psDataFolder + "\Error.txt"; if (Test-Path $errFileDupl) { Remove-Item $errFileDupl;} # If error found, then copy of log file is put into extract folder
$verbose=$false; # Overriding config. We do not want details of checked files on screen, unless testing.
$allFilesOK = $true;

# ##################################################################
# ########################## Checking file Dates.csv
# ##################################################################
$fn = $psDataFolder + "\Dates.csv"; $fc = @(Get-Content $fn); $fileError = $false;
$str = (Get-Date).ToString("HH:mm:ss") + " File Dates.csv. Record count: " + $fc.Count; if($verbose) {$str}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Dates.csv. Row count - must be at least 2 (header and at least one date)".PadRight(82); $testOK=$true; if (!($fc.Count -ge 2)) {$testOK=$false}
if ($testOK) {$str += "- OK";} else {$str += "- ERROR"; $fileError = $true;} if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Dates.csv. First row (file header) should be 'Date'".PadRight(82); $testOK=$true; $h=$fc[0];
if ($h -ne "Date") {$testOK = $false}; if($testOK) {$str += "- OK";} else {$str += "- ERROR. Actual value: '$h'"; $fileError = $true;}
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Dates.csv. All but first row (header) should be in format YYYY-MM-DD".PadRight(82); $testOK=$true;
$br = ""; For($i=1; $i -lt $fc.count; $i++) {try {$dd=[DateTime]::ParseExact($fc[$i], "yyyy-MM-dd", $null)} catch{$dd=$null;} if (![bool]$dd) {$br+=$fc[$i]+"`r`n"; $testOK=$false; }}
if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;}
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$fcsv = @(import-csv $fn | Sort-Object -Property "Date"); $minDateInFile = $fcsv[0].Date; $maxDateInFile = $fcsv[$fcsv.Count-1].Date; 

$str = " Dates.csv. Minimum date in file should be configured MinDate".PadRight(82); $testOK=$true; if ($minDateInFile -ne $minDate) {$testOK = $false;}
if ($testOK) {$str += "- OK ($minDateInFile)";} else {$str += "- ERROR. Actual minimum date in file $minDateInFile"; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Dates.csv. Maximum date in file should be today".PadRight(82); $testOK=$true; if ($maxDateInFile -ne (Get-Date).ToString("yyyy-MM-dd")) {$testOK=$false;}
if ($testOK) {$str += "- OK ($maxDateInFile)";} else {$str += "- ERROR. Actual maximum date in file $maxDateInFile"; $fileError = $true;}
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Dates.csv. Dates should be unique".PadRight(82); $testOK=$true; 
$br = ""; For($i=1; $i -lt $fcsv.count; $i++) {if($fcsv[$i].Date -eq $fcsv[$i-1].Date) {$testOK=$false; $br+=$fcsv[$i].Date+"`r`n"}}
if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Duplicate records: `r`n" + $br; $fileError = $true;}
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = (Get-Date).ToString("HH:mm:ss") + " File Dates.csv check completed."; if ($fileError) {$str+=" Errors found - please review"; $allFilesOK=$false;} else {$str+=" No issues found."};
if($verbose -or (!$testOK)) {$str+"`r`n"}; $str | Out-File $logFile -Encoding OEM -Append;
# ##################################################################

# ##################################################################
# ########################## Checking file Quotes.csv
# ##################################################################
$fn = $psDataFolder + "\Quotes.csv"; $fc = @(Get-Content $fn); $fileError = $false;
$str = (Get-Date).ToString("HH:mm:ss") + " File Quotes.csv. Record count: " + $fc.Count; if($verbose) {$str}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Quotes.csv. Row count - must be at least 2 (header and at least one record)".PadRight(82); $testOK=$true; 
if ($fc.Count -ge 2) {$str += "- OK";} else {$str += "- ERROR"; $testOK = $false; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Quotes.csv. First row (file header) should be 'Date,Symbol,Close'".PadRight(82).Replace(",",$colSep); $testOK=$true; $h=$fc[0];
if ($h -ne "Date,Symbol,Close".Replace(",",$colSep)) {$testOK = $fase;} if($testOK) {$str += "- OK";} else {$str += "- ERROR. Actual value: '$h'"; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " Quotes.csv. Each row should have exactly two column separators (',')".PadRight(82).Replace(",",$colSep); $testOK=$true; $br="";
ForEach($row in ($fc | Select-Object -skip 1)) {
    if (([regex]::Matches($row, $colSep)).count -ne 2) {$testOK = $fase; $br+=$row+"`r`n";}
    if ($testOK) {
        $parten = "^[12]\d{3}-\d{2}-\d{2}[,]{1}[^,]*[,]{1}\d+\.?\d*\s*$";
        if ($colSep -eq '`t') { $regsep = "\t"} else {$regsep = $colSep; }
        $parten = $parten.replace(',',$regsep);
        if (![regex]::IsMatch($row,$parten)) { $testOK = $fase; $br+=$row+"`r`n";}
    }
}
if ($br -eq "") {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;


$fcsv = @(import-csv $fn -Delimiter $colSep | Sort-Object -Property "Date","Symbol");
if ($fcsv.Count -gt 0) { # If there is at least 1 record
    $str = " Quotes.csv. All rows should have 'Date' column in format YYYY-MM-DD".PadRight(82); $testOK=$true;
    $br = ""; For($i=1; $i -lt $fcsv.count; $i++) {try {$dd=[DateTime]::ParseExact($fcsv[$i].Date, "yyyy-MM-dd", $null)} catch{$dd=$null;} if (![bool]$dd) {$testOK = $fase; $br+=$fcsv[$i].Date+"," +","+$fcsv[$i].Symbol+$fcsv[$i].Close+"`r`n"}}
    if ($br -eq "") {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    #$str = " Quotes.csv. All rows should have 'Close' column in North America number format".PadRight(82); $testOK=$true;
    #$br = ""; For($i=1; $i -lt $fcsv.count; $i++) {if (!($fcsv[$i].Close -match "^[+-]?([0-9]*\.?[0-9]+|[0-9]+\.?[0-9]*)([eE][+-]?[0-9]+)?$")) {$testOK = $fase; $br+=$fcsv[$i].Date+","+$fcsv[$i].Close +","+$fcsv[$i].Symbol+"`r`n"}}
    #if ($br -eq "") {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;} 
    #if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    $str = " Quotes.csv. Minimum date in file should be after configured MinDate".PadRight(82); $testOK=$true; $minDateInFile = $fcsv[0].Date;
    if ($minDateInFile -lt $minDate) {$testOK = $fase;} if ($testOK) {$str += "- OK ($minDateInFile)";} else {$str += "- ERROR. Actual minimum date in file $minDateInFile"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    $str = " Quotes.csv. Maximum date in file should be today or before".PadRight(82); $maxDateInFile = $fcsv[$fcsv.Count-1].Date; 
    if ($maxDateInFile -gt (Get-Date).ToString("yyyy-MM-dd")) {$testOK = $fase;} if ($testOK) {$str += "- OK ($maxDateInFile)";} else {$str += "- ERROR. Actual maximum date in file $maxDateInFile"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    # Check if dates+symbol records are unique
    $str = " Quotes.csv. Date+Symbol should be unique".PadRight(82); $testOK=$true;
    $br = ""; For($i=1; $i -lt $fcsv.count; $i++) {if($fcsv[$i].Date+$fcsv[$i].Symbol -eq $fcsv[$i-1].Date+$fcsv[$i-1].Symbol) {$testOK = $fase; $br+=$fcsv[$i].Date+","+$fcsv[$i].Close +","+$fcsv[$i].Symbol+"`r`n"}}
    if ($br -eq "") {$str += "- OK";} else {$str += "- ERROR. Duplicate records: `r`n" + $br; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;
}

$str = (Get-Date).ToString("HH:mm:ss") + " File Quotes.csv check completed."; if ($fileError) {$str+=" Errors found - please review"; $allFilesOK=$false;} else {$str+=" No issues found."};
if($verbose -or (!$testOK)) {$str+"`r`n"}; $str | Out-File $logFile -Encoding OEM -Append;
# ##################################################################


# ##################################################################
# ########################## Checking file fxrates.csv
# ##################################################################
$fn = $psDataFolder + "\fxrates.csv"; $fc = @(Get-Content $fn); $fileError = $false;
$str = (Get-Date).ToString("HH:mm:ss") + " File fxrates.csv. Record count: " + $fc.Count; if($verbose) {$str}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " fxrates.csv. Row count - must be at least 1 (header)".PadRight(82); $testOK=$true; if ($fc.Count -ge 1) {$str += "- OK";} else {$str += "- ERROR"; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " fxrates.csv. First row should be 'Date,Code,Rate'".PadRight(82).Replace(",",$colSep); $testOK=$true; $h=$fc[0];
if ($h -ne "Date,Code,Rate".Replace(",",$colSep)) {$testOK = $fase;} if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Actual value: '$h'"; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$str = " fxrates.csv. Each row should have exactly two column separators (',')".PadRight(82).Replace(",",$colSep); $testOK = $true; $br="";
ForEach($row in ($fc | Select-Object -skip 1)) {
    if (([regex]::Matches($row, $colSep )).count -ne 2) {$testOK = $fase; $br+=$row+"`r`n";}
    $parten = "^[12]\d{3}-\d{2}-\d{2}[,]{1}[^,]*[,]{1}\d+\.?\d*\s*$";
    if ($testOK) {
        if ($colSep -eq '`t') { $regSep = "\t"} else { $regSep = $colSep}
        $parten = $parten.Replace(',',$regSep);
        if (![regex]::IsMatch($row, $parten)) {
            $testOK = $fase; $br+=$row+"`r`n";
        }
    }
}
if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;} 
if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

$fcsv = @(import-csv $fn -Delimiter $colSep | Sort-Object -Property "Date","Code");
if ($fcsv.Count -gt 0) {
    $str = " fxrates.csv. All rows should have 'Date' column in format YYYY-MM-DD".PadRight(82); $testOK = $true;
    $br = ""; For($i=1; $i -lt $fcsv.count; $i++) {try {$dd=[DateTime]::ParseExact($fcsv[$i].Date, "yyyy-MM-dd", $null)} catch{$dd=$null;} if (![bool]$dd) {$testOK = $fase; $br+=$fcsv[$i].Date+","+$fcsv[$i].Code +","+$fcsv[$i].Rate+"`r`n"}}
    if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    #$str = " CurrencyConv.csv. All rows should have 'ExchRate' in North America number format".PadRight(82); $testOK = $true;
    #$br = ""; For($i=1; $i -lt $fcsv.count; $i++) {if (!($fcsv[$i].ExchRate -match "^[+-]?([0-9]*\.?[0-9]+|[0-9]+\.?[0-9]*)([eE][+-]?[0-9]+)?$")) {$testOK = $fase; $br+=$fcsv[$i].Date+","+$fcsv[$i].ExchRate +","+$fcsv[$i].CurrencyFrom+","+$fcsv[$i].CurrencyTo+"`r`n"}}
    #if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Bad records bellow:`r`n$br"; $fileError = $true;} 
    #if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    $str = " fxrates.csv. Minimum date in file should be configured MinDate".PadRight(82); $testOK  =$true; $minDateInFile = $fcsv[0].Date;
    if ($minDateInFile -lt $minDate) {$testOK = $fase;} if ($testOK) {$str += "- OK ($minDateInFile)";} else {$str += "- ERROR. Actual minimum date in file $minDateInFile"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    $str = " fxrates.csv. Maximum date in file should be today or before".PadRight(82); $maxDateInFile = $fcsv[$fcsv.Count-1].Date; 
    if ($maxDateInFile -gt (Get-Date).ToString("yyyy-MM-dd")) {$testOK = $fase;} if ($testOK) {$str += "- OK ($maxDateInFile)";} else {$str += "- ERROR. Actual maximum date in file $maxDateInFile"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;

    $str = " fxrates.csv. Date+Code should be unique".PadRight(82); $testOK = $true;
    $br = ""; For($i=1; $i -lt $fcsv.count; $i++) {if($fcsv[$i].Date+$fcsv[$i].Code -eq $fcsv[$i-1].Date+$fcsv[$i-1].Code) {$testOK = $fase; $br+=$fcsv[$i].Date+","+$fcsv[$i].Code +","+$fcsv[$i].Rate+"`r`n"}}
    if ($testOK) {$str += "- OK";} else {$str += "- ERROR. Duplicate records: `r`n$br"; $fileError = $true;} 
    if($verbose -or (!$testOK)) {Write-Host $str -ForegroundColor Red}; $str | Out-File $logFile -Encoding OEM -Append;
}

$str = (Get-Date).ToString("HH:mm:ss") + " File fxrates.csv check completed."; if ($fileError) {$str+=" Errors found - please review"; $allFilesOK=$false;} else {$str+=" No issues found."};
if($verbose -or (!$testOK)) {$str+"`r`n"}; $str | Out-File $logFile -Encoding OEM -Append;
# ##################################################################


$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
$str =(Get-Date).ToString("HH:mm:ss") + " --- Finished. Duration: $duration`r`n";
$str+="=======================================================================================`r`n";
$logSummary + ". Duration: $duration";

if ($allFilesOK) {Write-Host  "     ***** Generated PPS files were checked - no issues found *****" -ForegroundColor Green} 
else {Write-Host "     ***** Generated PPS files were checked - errors were found, please review *****" -ForegroundColor Red}; 
$str | Out-File $logFile -Encoding OEM -append;

if (!$allFilesOK) {Get-Content $logFile | out-File $errFileDupl -Encoding OEM;} # if there is error, copy log file into extract folder. Name file Error.txt


# #######################################################################################
# ######################## Checking if Portfolio Slicer is up to date
# #######################################################################################
# $url = "http://portfolioslicer.com/PSMsg/v2.0.html";
# $wc = new-object system.net.WebClient; 
# try {
#     $webpage = $wc.DownloadData($url); $mhtml = [System.Text.Encoding]::ASCII.GetString($webpage);
#     if ($mhtml.IndexOf("<message>") -ge 0) {$m = $mhtml.Substring($mhtml.IndexOf("<message>")+9); if ($m.IndexOf("</message>") -ge 0) {$m = $m.Substring(0, $m.IndexOf("</message>"))}}
# }
# catch {$m="N/A";};  

# $str="`r`n========== Message from PortfolioSlicer.com: '$m'`r`n"

# $str | Out-File $logFile -Encoding OEM -Append; 
# if ($str.Contains("upgrade")) {Write-Host $str -ForegroundColor Yellow} #else {Write-Host $str};
# # #######################################################################################



