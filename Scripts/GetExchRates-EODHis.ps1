# Apl 20 2018 created by Royer
#
# Fetch Exchange rate from https://eodhistoricaldata.com



$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\ExcelModule.ps1");     # Adding Excel COM Object read function

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
# starting logging to file.
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; 
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: ExchRate EODHistorical".PadRight(28);


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

$urlBase = "https://eodhistoricaldata.com/api/eod/@symbol@.FOREX?api_token=@apikey@&order=a&fmt=json&from=@FromDate@&to=@ToDate@";

# get apikey
$EOD_APIKEY=""; $EOD_APIKEY = ($apikeys | Select-Object -Index(($apikeys.IndexOf("<EODHistorical>"))+1)).Replace("</EODHistorical>","");
if ($EOD_APIKEY.Length -lt 1 )  {
    $str="*** Error. EODHistorical APIKey in apikey.txt file length is zero. exit script!"; 
    $str | Out-File $logFile -Encoding OEM -Append; $str | write-Host -ForegroundColor Red; exit(1);
}


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

    $url = $urlBase.Replace("@symbol@", $c.code).Replace("@apikey@",$EOD_APIKEY).Replace("@FromDate@",$startdate).Replace("@ToDate@", $todayYMD);

    try {
        $logmsg = "Get " + $c.code + " from $startdate to $todayYMD ...";
        $logmsg | Out-File $logFile -Encoding OEM -Append;
        if ($verbose) { $logmsg; }
        $reqCount++;
        $response = Invoke-WebRequest -Uri $url;

        $reqSucceed++;

        $data = $response.content | ConvertFrom-Json;

        $logmsg = "Get " + $c.code + " from $startdate to $todayYMD successful. total " + $data.count + " records."
        $logmsg | Out-File $logFile -Encoding OEM -Append;
        if ($verbose) { $logmsg; }
        $wc = 0;
        foreach ($row in $data) {
            if ( ($row.date -ge $startdate ) -and ($row.date -le $todayYMD) -and $row.close -ne "" ) {
                if ($appendstring -ne "") { $appendstring += "`r`n";}
                $line = ($row.date, $c.code, $row.close) -join ',';
                $appendstring += $line;
                $wc++;
            } else {
                $errmsg = "Invalid data in " + $c.code + " return: " + $row | Out-String;
                $errmsg | Out-File $logFile -Encoding OEM -Append;
                $errmsg | Write-Warning ;
            }
        }

        $reqRowsT += $wc;

        if ($appendstring -ne "") {
            $appendstring | Out-File $filepath -Encoding OEM -Append;
        }

    } catch {
        $reqFailed++;
        $errmsg = "Catch Web Error: " + $_;
        $errmsg | Out-File $logFile -Encoding OEM -Append;
        if ($verbose) {$errmsg | write-Host -ForegroundColor Red;}
    }
}

$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration";
