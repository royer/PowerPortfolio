####################################################################################################
#
# Make All Data Files in exdata folder.
#

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\ExcelModule.ps1");     # Adding script for import excel file module.


$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: MakeDataFiles".PadRight(28);

$includeArchiveFlag = $false; 
$includeArchiveValue = ($config | 
    Select-Object -Index(($config.IndexOf("<IncludeQuoteArchiveFolder>"))+1)).Replace("</IncludeQuoteArchiveFolder>",""); # Getting IncludeArchive value
if ($includeArchiveValue -ne $null) {
    if ($includeArchiveValue.Trim().ToLower() -eq "yes") {
        $includeArchiveFlag = $true;
    }
} 

# #######################################################################################
# ######################## Creating Currency Exchange file.
# ######################## exdata\fxrates.csv
# #######################################################################################
$str=(Get-Date).ToString("HH:mm:ss") + " Starting creating file fxrates.csv."; $str | Out-File $logFile -Encoding OEM -Append; if ($verbose) {$str};
$outFile = $psDataFolder + "\fxrates.csv"; $fileCount = 0; $recCount = 0;
"Date,Code,Rate".Replace(",", $colSep).Replace(".",$decSep) | Out-File $outFile -Encoding OEM;
(Get-ChildItem $currExchFolder | Where-Object {$_.extension -eq ".csv"}) | 
Where-Object{
    $fc=@(Get-Content -Path $_.FullName); 
    $fileCount+=1; $recCount+=$fc.Count; 
    $fc.Replace(",", $colSep).Replace(".",$decSep) | Out-file $outFile -Encoding OEM -Append 
};

# #######################################################################################
# ######################## Creating Quotes file
# #######################################################################################
$outFile = $psDataFolder + "\Quotes.csv"; $fileCount = 0; $recCount = 0;
"Date,Symbol,Close".Replace(",", $colSep) | Out-File $outFile -Encoding OEM;

#Add All currency cash fake price. get cash currency symbol from portfoliox.xlsx
$securities = (Import-Excel-ListObject $ExcelFilePath "securities" "Symbols");
#get all cash symbol
$allcash = $securities | select-object -Property "symbol", "secutype" | Where-Object secutype -EQ "cash"
$allcash | ForEach-Object {
    $str = $minDate + "," + $_.symbol + ",1.0000000001" ;
    $str.replace(",", $colSep) | Out-File $outFile -Encoding OEM -Append;
}
$fl = @(Get-ChildItem $quotesFolder -Recurse | Where-Object {$_.extension -eq ".csv" -and (!($_.BaseName -like "*_Archive") -or $includeArchiveFlag)}); 
ForEach($f in $fl) {
    $fc=@(Get-Content -Path $f.FullName | Where-Object {$_.Trim() -ne ""}); 
    if ($fc.Count -le 0) {continue;} 
    $fileCount+=1; $recCount+=$fc.Count; 
    $fc.Replace(",", $colSep) | Out-file $outFile -Encoding OEM -Append;
}
$str=(Get-Date).ToString("HH:mm:ss") + " Finished creating file       Quotes.csv. Source file count: " + $fileCount.ToString().PadLeft(3) + ". Record count: $recCount.";
$str | Out-File $logFile -Encoding OEM -Append; if ($verbose) {$str};

$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished creating all data files. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + "Finished creating all data files. Duration: $duration";