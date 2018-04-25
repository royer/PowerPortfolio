# ######################################################
# Sep 10 2017. Created by Royer. 
# 
# Fetch currency exchange rate from currencylayer
# url: https://currencylayer.com
# ######################################################

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\ExcelModule.ps1");     # Adding Excel COM Object read function

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
# starting logging to file.
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; 
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: ExchRate CurrencyLayer".PadRight(28);



<#
$listStart = $config.IndexOf("<Currency>"); 
$listEnd = $config.IndexOf("</Currency>"); 
if ($listStart -eq -1 -or $listEnd -eq -1 -or $listStart+1 -ge $listEnd) {
    "<currency> Symbol list is empty. Exiting script." | Out-File $logFile -Encoding OEM -Append; 
    exit(1);
}
#list of symbols we will work on
$currList = @($config | Select-Object -Index(($listStart+1)..($listEnd-1))); 
$currCount= $currList.Count; 
"Currency count: $currCount. MinDate: $minDate" | Out-File $logFile -Encoding OEM -Append;
#>

#$excelfilepath = $scriptPathParent + "\Powerportfolio.xlsx";
$currList = (Import-Excel-ListObject $ExcelFilePath "currencies" "Currencys") | Select-Object -Property "Code" | Where-Object { $_.code -ne "" };

# DO NOT need USD. remove it.
$currList = $currList | Where-Object { $_.code -ne "USD"}
if ($currList.count -eq 0) {
    "Currencries List in " + $excelfilepath | split-path -Leaf + " is empty. Exiting script." | Out-File $logFile -Encoding OEM -Append;
    exit(1);
}
"Currency count: " + $currList.Count + ". MinDate: $minDate" | Out-File $logFile -Encoding OEM -Append;



$CURRENCYLAYER_APIKEY=""; $CURRENCYLAYER_APIKEY = ($apikeys | Select-Object -Index(($apikeys.IndexOf("<CurrencyLayerAPIKey>"))+1)).Replace("</CurrencyLayerAPIKey>","");
if ($CURRENCYLAYER_APIKEY.Length -lt 1 )  {$str="*** Error. CurrencyLayer APIKey in apikey.txt file length is zero. Terminating script"; $str | Out-File $logFile -Encoding OEM -Append; $str | write-Host -ForegroundColor Red; exit(1);}
$urlBase = "https://apilayer.net/api/timeframe?access_key=@@AccessAPIKEY@@&start_date=@@start_date@@&end_date=@@end_date@@&currencies=@@currencies@@".Replace("@@AccessAPIKEY@@",$CURRENCYLAYER_APIKEY);


$currencies = "";
$currMap = @{}
$currdatesort = @();

foreach ($c in $currList) {
    $ret = GetFxRateInfo $c.code $currExchFolder $minDate;
    $currMap.Add($c.code,@{'filepath'=$ret[0];'startdate'=$ret[1];'appendstring'='';'rows'=0});
    $currdatesort += New-Object -TypeName PSObject -Prop @{'code'=$c.code;'date'=$ret[1]};
}
$currdatesort = $currdatesort | Sort-Object -Property 'date';
$request_infos = @();

$currencies = $currdatesort[0].code;
for ($i = 1; $i -lt $currdatesort.count ; $i++) {
    if ($currdatesort[$i].date -ne $currdatesort[$i-1].date) {
        $endd = ([datetime]::ParseExact($currdatesort[$i].date,"yyyy-MM-dd",$null)).AddDays(-1).ToString("yyyy-MM-dd");
        $request_infos += New-Object -TypeName PSObject -Property @{'start_date'=$currdatesort[$i-1].date; 'end_date'= $endd; 'currencies'=$currencies};
    }
    $currencies += "," + $currdatesort[$i].code;
}
$request_infos += New-Object -TypeName PSObject -Property @{'start_date'=$currdatesort[$currdatesort.count-1].date;'end_date'=(Get-Date).ToString("yyyy-MM-dd"); 'currencies'=$currencies};

$reqRowsT = 0;
for ($i = 0; $i -lt $request_infos.count; $i++) {
    $startYMD = [datetime]::ParseExact($request_infos[$i].start_date,'yyyy-MM-dd',$null);
    $lastday = [datetime]::ParseExact($request_infos[$i].end_date,'yyyy-MM-dd',$null);
    $endYMD = (($startYMD).AddDays(365), $lastday | Measure-Object -Min).Minimum;

    while (($startYMD -le $lastday) -and ($endYMD -le $lastday)) {
        $url = $urlBase.replace("@@start_date@@",$startYMD.toString("yyyy-MM-dd")).replace("@@end_date@@",$endYMD.ToString("yyyy-MM-dd")).replace("@@currencies@@",$request_infos[$i].currencies);
        $answerObj = ""; $reqCount++; 
        try { $answerObj = Invoke-RestMethod -Uri $url;}
        catch { $reqFailed++; "  " + $request_infos[$i].currencies + " - Not Found.(web err) `r`n" |Out-File $logFile -Encoding OEM -Append; }
        if ($answerObj.success -eq $false) {
            $reqFailed++;
            "Get fxrate failed. err info: " + $answerObj.error.info | Out-File $logFile -Encoding OEM -Append;
            if ($verbose) { $log + " No new fxrates data received."}
        }

        $reqSucceed++;
        
        foreach ($pdate in $answerObj.quotes.PSObject.Properties) {
            $sdate = $pdate.name;
            foreach ($quote in $pdate.value.PSObject.Properties) {
                $currencycode =  $quote.name.Substring(3);
                $rate = $quote.value;
                $line = ($sdate, $currencycode, $rate) -join ','
                if ($currMap[$currencycode].appendstring -ne "") { $currMap[$currencycode].appendstring += "`r`n";}
                $currMap[$currencycode].appendstring += $line;
                $currMap[$currencycode].rows ++;
                $reqRowsT++;
            }
        }

        $startYMD = $endYMD.AddDays(1);
        $endYMD = (($startYMD).AddDays(365), $lastday | Measure-Object -Min).Minimum;
    }
}

foreach ($code in $currMap.Keys) {
    if ($currMap[$code].appendstring -ne "") {
        $currMap[$code].appendstring | Out-File $currMap[$code].filepath -Encoding OEM -Append ;
    }
    "get " + $currMap[$code].rows + " rows of " + $code + " exrate." | Out-File $logFile -Encoding OEM -Append;
}


$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". CurrExch Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration";
