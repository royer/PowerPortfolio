# Download fxrate from Bank of Canada
# http://www.bankofcanada.ca/valet/observations/FXUSDCAD/csv?start_date=2017-04-28&end_date=2017-05-03
# 
# 2017-Sep-10. Created by Maxim T. 
# 2018-Apl-12. Modify by Royer
#
# ***********************!!! WARNNING !!! **************************************
# Bank of Canada provide the earlest daily rate date is last year. if you query at 2018, it only got data from
# 2017-01-01 to query date. 

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\ExcelModule.ps1");     # Adding Excel COM Object read function


$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: ExchRate BoC".PadRight(28);

# Remove Intraday fxrate file. Why?
$currExchIDFile = $currExchIDFolder + "GoogleCurrExchIntraday.txt"; if (Test-Path $currExchIDFile) { Remove-Item $currExchIDFile;} # Removing intraday file before each load
$currExchIDFile = $currExchIDFolder + "YahooCurrExchIntraday.txt"; if (Test-Path $currExchIDFile) { Remove-Item $currExchIDFile;} # Removing intraday file before each load

# Get Currency List from powerportfolio excel file
$currList = (Import-Excel-ListObject $ExcelFilePath "currencies" "Currencys") | Select-Object -Property "Code" | Where-Object { ($_.code -ne "" -and $_.code -ne "USD") };

$currList = $currList | Where-Object { $_.code -ne "USD"}
if ($currList.count -eq 0) {
    "Currencries List in " + $excelfilepath | split-path -Leaf + " is empty. Exiting script." | Out-File $logFile -Encoding OEM -Append;
    exit(1);
}
"Currency count: " + $currList.Count + ". MinDate: $minDate" | Out-File $logFile -Encoding OEM -Append;

$urlUSDCAD = "http://www.bankofcanada.ca/valet/observations/FXUSDCAD/json?start_date=@@DateFrom@@&end_date=@@DateTo@@"
$urlBase = "http://www.bankofcanada.ca/valet/observations/FX@@currency@@CAD/json?start_date=@@DateFrom@@&end_date=@@DateTo@@"
# Bank of Canada always provide currency <=> CAD rate, we need use USDCAD / HKDCAD to connvert USDCNY for match our record.

$currMap = @{}
$currdatesort = @();

foreach ($c in $currList) {
    $ret = GetFxRateInfo $c.code $currExchFolder $minDate;
    $currMap.Add($c.code,@{'filepath'=$ret[0];'startdate'=$ret[1];'appendstring'='';'rows'=0});
    $currdatesort += New-Object -TypeName PSObject -Prop @{'code'=$c.code;'date'=$ret[1]};
}
$currdatesort = $currdatesort | Sort-Object -Property 'date';
$MinDownloadDate = $currdatesort[0].date;

# firstable download USDCAD from $MinDownloadDate
$urlUSDCAD = $urlUSDCAD.Replace("@@DateFrom@@", $MinDownloadDate).Replace("@@DateTo@@", $todayYMD)
$wc = New-Object System.Net.WebClient;
try {
    $webdata = $wc.DownloadData($urlUSDCAD);
    $webdata = [System.Text.Encoding]::ASCII.GetString($webdata);
    $j = $webdata | ConvertFrom-Json;
    $USDCAD = $j.observations

} catch {
    $errmsg = "Download USDCAD from BankOfCanada.com failed. " + $_;
    $errmsg | Out-File $logFile -Encoding OEM -Append;
    $errmsg | Write-Warning;
    exit(1)
}

foreach ($kvp in $currMap.GetEnumerator()) {
    $startdate = $kvp.value.startdate
    if ($kvp.key -eq 'CAD') {
        $CADJson = $USDCAD | Where-Object d -ge $startdate
        $reqCount++; $reqSucceed++;
        foreach ($row in $CADJson) {
            $date = $row.d; $code = 'CAD'; $rate = $row.FXUSDCAD.v;
            $line = ($date, $code, $rate) -join ','
            
            if ($kvp.value.appendstring -ne "") {$kvp.value.appendstring += "`r`n";}
            $kvp.value.appendstring += $line;
            $kvp.value.rows++; $reqRowsT++;
        }
        if ($kvp.value.appendstring -ne "" ) {
            $kvp.value.appendstring | Out-File $kvp.value.filepath -Encoding OEM -Append;
            "get " + $kvp.value.rows + " rows of " + $kvp.key + " exrate." | Out-File $logFile -Encoding OEM -Append;
        }
    } else {
        $usdcad_match = $USDCAD | where-object d -ge $startdate;

        # download rate 
        $url = $urlBase.Replace("@@currency@@", $kvp.key).Replace("@@DateFrom@@", $startdate).Replace("@@DateTo@@", $todayYMD)
        try {
            $reqCount++;
            $webdata = Invoke-WebRequest -Uri $url;
            $j = $webdata.content | ConvertFrom-Json;
            if ($usdcad_match.count -ne $j.observations.count) {
                $errmsg = "Get " + $kvp.key + " " + $j.observations.count + " rows not match USDCAD " + $usdcad_match.count + ". skip this currency data";
                $errmsg | Out-File $logFile -Encoding OEM -Append;
                $errmsg | Write-Warning;
                $reqFailed++;
                continue;
            }else {
                $alldatematch = $true
                for ($i = 0; $i -lt $usdcad_match.count; $i++) {
                    if ($usdcad_match[$i].d -eq $j.observations[$i].d) {
                        $date = $usdcad_match[$i].d;
                        $code = $kvp.key;
                        $fieldname = $j.seriesDetail.PSObject.Properties.Name;
                        $rate =  [math]::Round($usdcad_match[$i].FXUSDCAD.v / $j.observations[$i].$fieldname.v,7);
                        $line = ($date, $code, $rate) -join ','
                        if ($kvp.value.appendstring -ne "") {$kvp.value.appendstring += "`r`n";}
                        $kvp.value.appendstring += $line;
                        $kvp.value.rows++; $reqRowsT++;
                    } else {
                        $errmsg = "found usdcad date not match " + $kvp.key + " date. skip this currency.";
                        $errmsg | Out-File $logFile -Encoding OEM -Append;
                        $errmsg | Write-Warning; 
                        $alldatematch = $false;
                        break;
                    }
                }
                if ($alldatematch) { 
                    $reqSucceed++;
                    if ($kvp.value.appendstring -ne "" ) {
                        $kvp.value.appendstring | Out-File $kvp.value.filepath -Encoding OEM -Append;
                        "get " + $kvp.value.rows + " rows of " + $kvp.key + " exrate." | Out-File $logFile -Encoding OEM -Append;
                    }
                } else { $reqFailed++ ;}
               
            }
        } catch {
            $errmsg = "Download " + $kvp.key +" failed. " + $_;
            $errmsg | Out-File $logFile -Encoding OEM -Append;
            $errmsg | Write-Warning
            $reqFailed++;
            exit(1)
        }
    }
}

exit(1)


$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration";
