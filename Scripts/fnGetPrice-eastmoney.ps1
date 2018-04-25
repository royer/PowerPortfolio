# Get Quotes of Chinese Mutual Fund from eastmoeny.com
# http://fund.eastmoney.com/f10/F10DataApi.aspx?type=lsjz&code=159915&page=1&per=50&sdate=2014-03-01&edate=2017-05-30&rt=0.273778463698463
#
# Apl 15, 2018 create by royer

function GetPriceHistory_EastMoney ($symbol, $sdate, $edate, $lastprice, $logfile) {

    $urlBase = "http://fund.eastmoney.com/f10/F10DataApi.aspx?type=lsjz&code=@symbol@&page=@page@&per=50&sdate=@DateFrom@&edate=@DateTo@&rt=0.273778463698463"

    $retvalue = New-Object -TypeName PSObject -Property @{'result'=$false; 'prices'=@();}

    $currPage = 1
    $pages = 2
    $regexpartten = "(?<=apidata=)(.*)(?=;)"
    while ($currpage -le $pages) {

        $url = $urlBase.Replace("@symbol@", $symbol).Replace("@page@", $currpage).Replace("@DateFrom@", $sdate).Replace("@DateTo@", $edate)
        try {
            $webresult = Invoke-WebRequest -Uri $url;
        } catch {
            $errmsg = "Request Web failed when get $symbol from eastmoney.";
            if ($logfile) { $errmsg | Out-File -Encoding OEM -Append; }
            $errmsg | Write-Warning;
            return $retvalue;
        }
        $regex_result = [regex]::Match($webresult.Content, $regexpartten)
        
        if ($regex_result.Success) {
            $dirtyJson = $regex_result.Captures.groups[0];
            # fix lazy json string
            # eastmoney return json dict proerty does not use ""
            # https://stackoverflow.com/a/40794738/1036923

            $strclean = [regex]::Replace($dirtyJson, "{\s*'?(\w)", '{"$1');
            $strclean = [regex]::Replace($strclean, ",\s*'?(\w)",',"$1');
            $strclean = [regex]::Replace($strclean, "(\w)'?\s*:(?!/)",'$1":');
            $strclean = [regex]::Replace($strclean, ":\s*'(\w+)'\s*([,}])",':"$1"$2');
            $strclean = [regex]::Replace($strclean, ",\s*]", ']');


            $jsondata = $strclean | ConvertFrom-Json;
            $pages = $jsondata.pages;
            if ($jsondata.records -eq 0) {
                
                break;
            }

            [System.Xml.XmlDocument] $xmld = New-Object System.Xml.XmlDocument;
            $xmld.loadXml($jsondata.content);
            $headnode = $xmld.SelectNodes("/table/thead/tr/th");
            $col_date = -1; $col_close = -1; $col = 0;
            foreach ($th in $headnode) {
                if ($th.get_InnerXml() -eq "净值日期") {$col_date = $col}
                if ($th.get_InnerXml() -eq "单位净值") { $col_close = $col}
                $col++;
            }

            $tbodyrows = $xmld.SelectNodes("/table/tbody/tr");
            foreach ($row in $tbodyrows) {
                $tds = $row.SelectNodes("td");
                [string] $date = $tds[$col_date].get_InnerXml();
                $close = $tds[$col_close].get_InnerXml();
                if ($date.StartsWith("*")) { $date = $date.Substring(1)}
                if ($date.StartsWith("暂无数据")) {
                    $retvalue.result = $true;
                    
                    return $retvalue;
                }
                if ($date -notmatch "\d{4}-\d{2}-\d{2}") {
                    $errmsg = " " + $date + " is not a date at page: " + $currPage +" when get " + $symbol + " quotes.";
                    if ($logfile) { $errmsg | Out-File $logfile -Encoding OEM -Append; }
                    $errmsg | Write-Warning;
                    return $retvalue;
                }

                if ($date -le $edate) {
                    if ($close -ne "" -and ($close -match "^\d*\.?\d*$") -and $close -gt 0) {
                        $lastprice = $close;
                    }
                    $prop = @{'Date'=$date; 'Close'=$lastprice};
                    $obj = New-Object -TypeName psobject -Property $prop;
                    $retvalue.prices += $obj;
                }
            }
        }    
        else {
            "Match apidata failed." | Write-Warning;
            exit(1);
        }
        $currPage++;

    }

    $retvalue.result = $true;
    if ($retvalue.prices.Count -gt 0) {
        $retvalue.prices = $retvalue.prices | Sort-Object -Property "Date";
    }
    return $retvalue;
}


# $ret = GetPriceHistory_EastMoney "159915" "2018-01-01" "2018-04-15" $null

# "return " + $ret.result | Write-Host;
# foreach ($quote in $ret.prices) {
#     "Date: " + $quote.Date + "`tClose: " + $quote.close | Write-Output;
# } 