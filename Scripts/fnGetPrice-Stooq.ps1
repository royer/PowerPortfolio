# Get Quotes from www.stooq.com
# Apl 15, 2018 by Royer
# https://stooq.com/q/d/l/?s=aapl.us&d1=20160907&d2=20170912&i=d&o=1111111
#
# ***** NOTICE & LIMIT *****************
# Has daily request limit. will return "Exceeded the daily hits limit"
# **************************************
#
# param:
#   s:  equity symbol
#   d1: start date
#   d2: end date
#   i:  interval
#       d:  daily
#       w:  weekly
#       m:  monthly
#       q:  Quarterly
#       y:  Yearly
#   o:  optional: bitmask option setting
#       splits:     1000000  set sting[0] = 1, mean do not adjust for split
#       dividends:  0100000  set string[1] = 1, mean do not adjust for divided
#       preemptive rights .....
#       prepurchase rights ....
#       preaccession rights ....
#       denominations ....
#       others ....
#
# return in response: csv format data
# Date,Open,High,Low,Close,Volume
# 2018-04-02,167.88,168.94,164.47,166.68,37586791
# 2018-04-03,167.64,168.746,164.88,168.39,30278046
#
# https://stooq.com/q/d/l/?s=aapl.us&i=d&o=1111111
# if do not provide d1 and d2 param, it will download all quotes of this symbol since it IPO
#
# stooq.com is a polski website. it provide stock , currency, futures, Cryptocurrency quotes.
# stock include US, UK, Hongkong, Japan,German, Polish,Hungarian
# Futures include stock indices Futures
#

function GetPriceHistory_Stooq($symbol, $sdate, $edate, $lastprice, $logfile) {
    $retvalue = New-Object -TypeName psobject -Property @{'result'=$false; 'prices'=@()};

    $urlBase = "https://stooq.com/q/d/l/?s=@Symbol@&d1=@DateFrom@&d2=@DateTo@&i=d&o=1111111";

    $url = $urlBase.Replace("@Symbol@", $symbol).Replace("@DateFrom@", $sdate.Replace('-','')).Replace("@DateTo@", $edate.Replace('-',''));

    try {
        $webResponse = Invoke-WebRequest -Uri $url;
    }catch {
        $errmsg = "request " + $url + " Failed. " + $_;
        if ($logfile) { $errmsg | Out-File $logfile -Encoding OEM -Append};
        $errmsg | Write-Warning;
        return $retvalue; 
    }

    if ($webResponse.content.Contains("No data")) {
        $errmsg = $symbol + " not found. (return 'No data')";
        if ($logfile) { $errmsg | Out-File $logfile -Encoding OEM -Append;}
        $errmsg | Write-Warning;
        return $retvalue;
    }

    if ($webResponse.content.Contains("<html>")) {
        $errmsg = $symbol + " not found. (return html)";
        if ($logfile) { $errmsg | Out-File $logfile -Encoding OEM -Append;}
        $errmsg | Write-Warning;
        return $retvalue;
    }

    if ($webResponse.content.length -le 31) {
        # response less then header line, must error
        $errmsg = "response total lenth less than header line charaters. response: " + $webResponse ; 
        if ($logfile) { $errmsg | Out-File $logfile -Encoding OEM -Append;}
        $errmsg | Write-Warning;
        return $retvalue;
    }

    $csvdata = $webResponse.Content | ConvertFrom-Csv;
    foreach ($row in $csvdata) {
        if (($row.Date -le $edate) -and $row.Date -ge $sdate) {
            if ($row.Close -ne "" -and ($row.Close -match "^\d*\.?\d*$") -and $row.Close -gt 0) {
                $lastprice = $row.Close;
            }
            $prop = @{'Date'=$row.Date; 'Close'=$lastprice}
            $obj = New-Object -TypeName psobject -Property $prop;
            $retvalue.prices += $obj;
        }
    }

    $retvalue.result = $true;

    return $retvalue;
}
