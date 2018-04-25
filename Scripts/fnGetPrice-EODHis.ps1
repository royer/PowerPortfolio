# Fetch quotes from https://eodhistoricaldata.com/
# 2018-Apl-15 Created by Royer
#
# eodhistoricaldata.com is not free data provider
# They provide:
#   * FUNDAMENTAL data
#   * REAL-TIME and DAILY historical stock prices for stocks 
#   * ETFs and Mutual Funds all around the world. 
#   * 120+ CRYPTO currencies
#   * 150+ FOREX pairs. 
# Market coverage is more than 40+ stock exchanges and more than 120.000 symbols in total.

function GetPriceHistory_EODHis($symbol, $sdate, $edate, $lastprice, $apikey,  $logfile) {

    $retvalue = New-Object -TypeName psobject -Property @{'result'=$false; "prices"=@()};

    $urlBase = "https://eodhistoricaldata.com/api/eod/@Symbol@?from=@DateFrom@&to=@DateTo@&api_token=@apikey@&fmt=json&period=d"

    $url = $urlBase.Replace("@Symbol@", $symbol).Replace("@apikey@", $apikey).Replace("@DateFrom@",$sdate).Replace("@DateTo@", $edate);
    try {
        $webresponse = Invoke-WebRequest -Uri $url;
    } catch {
        $errmsg = "EODHistorical: Get " + $symbol + " quotes failed. Web Error: " + $_;
        if ($logfile) { $errmsg | Out-File $logfile -Encoding OEM -Append;}
        $errmsg | Write-Host -ForegroundColor Red;
        return $retvalue;
    }

    $data = $webresponse.content | ConvertFrom-Json;

    foreach ($row in $data) {
        if ($row.date -ge $sdate -and $row.date -le $edate ) {
            if ($row.Close -ne "" -and ($row.Close -match "^\d*\.?\d*$") -and $row.Close -gt 0) {
                $lastprice = $row.Close;
            }
            $prop = @{'Date'=$row.date; 'Close'=$lastprice;}
            $obj = New-Object -TypeName psobject -Property $prop;

            $retvalue.prices += $obj;
        }
    }

    if ($logfile) {
        if ($retvalue.prices.Count -ne 0) {
            $d1 = $retvalue.prices[0].Date;
            $d2 = $retvalue.prices[$retvalue.prices.Count-1].Date
        } else {
            $d1 = $sdate; $d2 = $edate;
        }
        "EODHis: get $symbol $d1 - $d2 successful. 0 rows added."  | Out-File $logfile -Encoding OEM -Append;
    }

    $retvalue.result = $true;
    return $retvalue;
}

