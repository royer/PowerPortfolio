$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: Create Dates".PadRight(28);
"PSData folder: $psDataFolder" | Out-File $logFile -Encoding OEM -Append;

# #######################################################################################
# ######################## Create Dates.csv file
# #######################################################################################
$outFile = $psDataFolder + "\Dates.csv"; 
write-Host "The Dates.csv file locate: " + $outFile;
$dToday = Get-Date; 
$dDate = [datetime]::ParseExact($minDate,"yyyy-MM-dd",$null); 
$strD = "Date";
while ($dDate -le $dToday) { 
    $strD += "`r`n" + $dDate.ToString("yyyy-MM-dd") ; 
    $dDate = $dDate.AddDays(1); 
    $reqRowsT++;
}; 
$strD | Out-File $outFile -Encoding OEM; 


$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. Dates: $reqRowsT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". Dates: $reqRowsT. Duration: $duration";
# #######################################################################################

