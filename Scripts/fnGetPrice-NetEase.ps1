# function of get price from netease
#
# base Url: http://quotes.money.163.com/service/chddata.html?code=<code>&start=YYYYMMDD&end=YYYYMMDD&fields=TCLOSE;HIGH;...
#
# return: csv file. encoding by GB2312
#
# Note: 网易的数据是不复权的数据，正是本应用需要的数据
#
# Param:
#   code:       stock code.
#       for shanghai exchange stock. add 0 before original code. for example: 600036 -> 0600036.
#       for shengzhen exchange stock. add 1 before original code. for example: 000059 -> 1000059
#
#   start:      begin date for query. format: YYYYMMDD
#
#   end:        end date for query.   format: YYYYMMDD             
#
#   fields:     the query fields list, seperated by ;  , fields name must be uppercase.
#       TCLOSE:     close price(收盘价)
#       HIGH:       high price(最高价)
#       LOW:        low price(最低价)
#       TOPEN:      open price(开盘价)
#       LCLOSE:     yestoday close price(前收盘)
#       CHG:        change (涨跌额)
#       PCHG:       % change(涨跌幅)
#       TURNOVER:   turnover(换手率)
#       VOTURNOVER: volumn(成交量)
#       VATURNOVER: (成交金额)
#       TCAP:       Cap(总市值)
#       MCAP:       流通市值

function getPriceHistory_NetEase($neteasesymbol, $sdate, $edate, $lastprice, $logfile) {
    
    $urlBase = " http://quotes.money.163.com/service/chddata.html?code=@code@&start=@sdate@&end=@edate@&fields=TCLOSE";

    $lastQuote = (Get-Date $sdate).AddDays(-1).toString('yyyy-MM-dd');

    $fromdate = $sdate.replace('-','')
    $todate = $edate.replace('-', '')

    $url = $urlBase.Replace('@code@',$neteasesymbol).replace('@sdate@', $fromdate).replace('@edate@',$todate);

    $retvalue = New-Object -TypeName PSObject -Property @{'result'=$false; 'prices'=@();}

    $wr = '';

    try {
        $wr = Invoke-WebRequest -Uri $url;

        if ($wr.StatusCode -ne 200) {
            $errmessage = "ERROR: get $neteasesymbol from NetEase failed. web request return status code: "+$wr.StatusCode + "`r`n";
            if ($logfile) { $errmessage | Out-File $logfile -Encoding OEM -Append;}
            $errmessage | Write-Warning
            
            return $retvalue;
        }

        $gbk = [System.Text.Encoding]::GetEncoding("GB2312");
        # $utf8 = [System.Text.Encoding]::UTF8;
        # $qtemp = [System.Text.Encoding]::Convert($gbk, $utf8, $wr.content)
        # $quoteTxt = $utf8.getString($qtemp)
        $quoteTxt = $gbk.getString($wr.Content);
        if ($quoteTxt.Length -le 14) {
            $errmessage = "NetEase get $neteasesymbol failed. return data less than 14 (head line character count)";
            if ($logfile) { $errmessage | Out-File -Encoding OEM -Append; }
            $errmessage | Write-Warning
            return $retvalue;
        }

        if ($quoteTxt.Contains("<html>")) {
            $errmessage = "NetEase get $neteasesymbol failed. return a html page not quote data.";
            if ($logfile) { $errmessage | Out-File -Encoding OEM -Append; }
            $errmessage | Write-Warning
            return $retvalue;            
        }
        
        #$quoteTxt = $quoteTxt.replace("`u{日期}", "Date").replace("`u{股票代码}", "Symbol").Replace("`u{名称}","SymbolName").Replace("`u{收盘价}","Close")
        # replace header line
        $quoteTxt = $quoteTxt.Replace($quoteTxt.substring(0,$quoteTxt.IndexOf("`n")),"Date,Symbol,Symbolname,Close");
        $quotecsv = ($quoteTxt | ConvertFrom-Csv) | Sort-Object -Property "Date";

        $quotecsv | ForEach-Object {
            if ($_.Date -ge $sdate -and $_.Date -ge $lastQuote -and $_.Date -le $edate) {
                $lastQuote = $_.Date;
                if ($_.Close -ne "" -and ($_.Close -match "^\d*\.?\d*$") -and $_.Close -gt 0) {
                    $lastprice = $_.Close;
                }
                $prop = @{'Date'=$_.Date; 'Close'=$lastprice};
                $obj = New-Object -TypeName psobject -Property $prop;
                $retvalue.prices += $obj;
            }
            
        }
        if ($retvalue.prices.count -eq 0) {
            $retvalue.result = $false;
        } else {
            $retvalue.result = $true;
        }

    }
    catch {
        $errmessage = "Warning: Get $neteasesymbol from NetEase failed. Catch web error:" + $_;
        if ($logfile) { $errmessage | Out-File $logfile -Encoding OEM -Append; }
        $errmessage | Write-Warning
        return $retvalue;
    }

    if ($logfile) { 
        if ($retvalue.result) { $sf = "Successful"} else {$sf = "Failed"}
        "NetEase: get $neteasesymbol price. $sdate - $edate $sf. " + $retvalue.prices.count + " rows added." | Out-File $logFile -Encoding OEM -Append; 
    }

    return $retvalue;
}