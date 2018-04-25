# ######################################################
# 2017-Sep-10. Created by Maxim T. 
# URL: https://stooq.com/q/d/l/?s=cadusd&d1=20160102&d2=20170913&i=d
#
# ***** NOTICE & LIMIT *****************
# Has daily request limit. will return "Exceeded the daily hits limit"
# **************************************
#
# ######################################################

# ######################################################
# 2018-Apl-14 Modify by Royer
# ######################################################

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\excelModule.ps1");     # Adding load excel table module

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("yyyy-mm-dd HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
if ($verbose) { "" + $MyInvocation.MyCommand.Name + " starting..."}

$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: ExchRate Stooq".PadRight(28);
$currExchIDFile = $currExchIDFolder + "GoogleCurrExchIntraday.txt"; if (Test-Path $currExchIDFile) { Remove-Item $currExchIDFile;} # Removing intraday file before each load
$currExchIDFile = $currExchIDFolder + "YahooCurrExchIntraday.txt"; if (Test-Path $currExchIDFile) { Remove-Item $currExchIDFile;} # Removing intraday file before each load

# Get Currency List from powerportfolio excel file
if ($verbose) {"Get currency list from " + $ExcelFilePath + " ..."}
$currList = (Import-Excel-ListObject $ExcelFilePath "currencies" "Currencys") | Select-Object -Property "Code" | Where-Object { ($_.code -ne "" -and $_.code -ne "USD") };
$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
$logmsg = "Load currency list from $ExcelFilePath used $duration";
$logmsg | Out-File $logFile -Encoding OEM -Append;
if ($verbose) {$logmsg; }

if ($currList.count -eq 0) {
    $errmsg = "Currency list in " + $ExcelFilePath | Split-Path -Leaf + " is Empty(except USD). script exit."
    $errmsg | Out-File $logFile -Encoding OEM -Append;
    if ($verbose) {$errmsg | Write-Warning}
    exit(1);
}


$urlBase = "https://stooq.com/q/d/l/?s=USD@@CurrTo@@&d1=@@DateFrom@@&d2=@@DateTo@@&i=d".Replace("@@DateTo@@", $todayYMD.Replace("-",""));

foreach ($c in $currList) {
    $ret = GetFxRateInfo $c.code $currExchFolder $minDate;
    $filepath = $ret[0];  $startdate = $ret[1];
    if ($startdate -notmatch "\d{4}-\d{2}-\d{2}") {
        $errmsg = "get " + $filepath + " last date error! please check this file. skiped this currency.";
        $errmsg | Out-File $logFile -Encoding OEM -Append;
        $errmsg | write-Host -ForegroundColor Red;
        continue;
    }
    if ($startdate -gt $todayYMD) {
        if ($verbose) {
            $c.code + " is updated newest. skip it."
        }
        continue;
    }
    $appendstring = ""

    $url = $urlBase.Replace("@@CurrTo@@", $c.code).Replace("@@DateFrom@@", $startdate.Replace('-',''));
    if ($verbose) { "".PadRight(4) + "Requesting: "+$url;}
    $reqCount++; 
    $wc = new-object System.Net.WebClient;
    try {
        $webdata = $wc.DownloadData($url);
        $reqSucceed++;
    } catch {
        $reqFailed++;
        $errmsg = "  " + $c.code + " request failed. " + $_;
        $errmsg | Out-File $logFile -Encoding OEM -Append;
        $errmsg | Write-Warning;
    }
    $data = [System.Text.Encoding]::ASCII.getString($webdata);
    if ($data.contains("<html>")) { 
        $reqFailed++; "  No new data for " + $c.code + " (returned html)" | Out-File $logFile -Encoding OEM -Append; 
        if($verbose) {$log + " No new data received."}; 
        continue; 
    } # Result is html file, something went wrong, ignore result that was just returned
    if ($data.contains("No data")) { 
        "  No new data for " + $c.code + " (returned No data)" | Out-File $logFile -Encoding OEM -Append; 
        if($verbose) {$log + " No new data received."}; 
        continue; 
    } # Result is "No data", ignore result that was just returne

    $data = $data | ConvertFrom-Csv;
    foreach ($row in $data) {
        $line = ($row.date, $c.code, $row.close) -join ',';
        if ($appendstring -ne "") { $appendstring += "`r`n"}
        $appendstring += $line;
    }
    $reqRowsT += $data.count;
    $logmsg = "  get rate of " + $c.code + " " + $data.count + " rows";
    $logmsg | Out-File $logFile -Encoding OEM -Append;
    if ($verbose) { $logmsg; }
    if ($appendstring -ne "") {
        $appendstring | Out-File $filepath -Encoding OEM -Append;
    }
}


$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration";
