# function of get price from yahoo
#
# ************** NOTICE & LIMIT ***********************
# From 2017-05 Yahoo requires session with cookie and crumb value at the end of URL to return data.
# Example: https://query1.finance.yahoo.com/v7/finance/download/VYM?period1=1500137940&period2=1502816340&interval=1d&events=history&crumb=nF2PUWr9OBA
#
# ************** BUGS *****************************
# some country stock history data is missed.
# For example: 600309.SS has no history data
#
function GetYahooCrumbAndWebSession($logFile=$null) {

    $wrCookie = ""; $wr = ""; $crumbStr = "`"CrumbStore`":{`"crumb`":`""; $crumb = "";

    #This URL will be used to get establish session and get crumb value
    $urlCookie = "https://finance.yahoo.com/quote/AAPL/history?p=AAPL"; 
    for ($i=0;$i -le 2 -and $crumb -eq ""; $i++) { # will attempt to get crumb 3 times

        $wrCookie = Invoke-WebRequest -Uri $urlCookie -SessionVariable websession; 
        # Session with cookie now is in $websession variable. 

        $webCookie = $wrCookie.Content;   # Need to parse content and look for string $crumbStr value.

        #Identifing crumb location in file. Get start location of the value;
        $crumbStart = $webCookie.IndexOf($crumbStr); 
        if ($crumbStart -eq 0) {
            if ($logFile -ne $null) {
                "Crumb start not found for Yahoo" | Out-File $logFile -Encoding OEM -Append; 
            }
            continue;
        } 
        
        # Get 100 characters from the crumb start
        $crumb = $webCookie.Substring($crumbStart+$crumbStr.Length,100); 

        # Find crumb string end.
        $crumbEnd = $crumb.IndexOf("`"}"); 
        if ($crumbEnd -eq 0) {
            if ($logFile -ne $null) {
                "Crumb end not found for yahoo" | Out-File $logFile -Encoding OEM -Append;
            }
            continue;
        } 
        # ================ At this point we have crumb value in $crumb and web session with cookie established in $websession;
        $crumb = $crumb.Substring(0, $crumbEnd);  
        if (-not($crumb -match "^[a-zA-Z0-9\s]+$") -and $i -ne 2) {$crumb = "";}
    }
    
    $crumb;          #0
    $websession;     #1
}
function GetPriceHistroy_Yahoo($crumb, $websession, $yahoosymbol, $fromdate, $todate, $lastprice, $logFile) {

    $urlBase = "https://query1.finance.yahoo.com/v7/finance/download/@Symbol@?period1=@FromDay@&period2=@ToDay@&interval=1d&events=history&crumb=@CRUMB@"
    $toDay   = [string] [math]::Floor((get-date $todate -UFormat %s)); #today in unix timestamp
    $urlBase = $urlBase.Replace("@CRUMB@", $crumb).Replace("@ToDay@", $toDay);    

    $fromDay = [string] [math]::Floor((get-date $fromdate -UFormat %s));

    $lastQuote = (get-date $fromdate).AddDays(-1).toString("yyyy-MM-dd");
    $url = $urlBase.replace("@Symbol@",$yahoosymbol).replace("@FromDay@",$fromDay);


    $retvalue = New-Object -TypeName PSObject -Property @{'result'=$false; 'prices'=@();}

    $wr = "";
    try {
        $wr = Invoke-WebRequest -Uri $url -WebSession $websession;

        if ($wr.StatusCode -ne 200) {
            $errmessage = "Error: YAHOO get $yahoosymbol Failed. http return StatusCode: " + $wr.StatusCode + "`r`n";
            if ($logFile) { $errmessage | Out-File $logFile -Encoding OEM -Append;}
            $errmessage | Write-Warning;

            return $retvalue;
        }

        $QuoteTxt = $wr.Content;
        # This variable now contains downloaded quotes in text format:
        #    Date,Open,High,Low,Close,Adj Close,Volume
        #    2017-07-17,78.830002,78.930000,78.760002,78.839996,78.839996,472700
        if ($QuoteTxt.Length -le 45) {
            # Check if data received makes sense. If request size is less than 45 bytes (header length is 42), then there is something wrong with data received. 
            $errmessage = "Error: YAHOO get $yhaoosymbol Failed. return data length less than 45 bytes. ";
            if ($logFile) { $errmessage | Out-File $logFile -Encoding OEM -Append;}
            $errmessage | Write-Warning ;
            return $retvalue;
        }
        if ($QuoteTxt.Contains("<html>")) {
            $errmessage = "Error: YAHOO get $yahoosymbol Failed. return a html page not quote data.";
            if ($logFile) {$errmessage | Out-File $logFile -Encoding OEM -Append;}
            $errmessage | Write-Warning ;
            return $retvalue;
        }

        # Sometimes duplicate records could come back, need to sort records and not load duplicate values
        $ql = $QuoteTxt.Split("`n") | Where-Object {$_.trim() -ne ""  -and ($_.StartsWith("1") -or $_.StartsWith("2"))} | Sort-Object; 
        $ql | ForEach-Object {
            $a=$_.Split(","); 
            if($a[0] -ge $fromdate -and $a[0] -gt $lastQuote `
                -and $a[0] -le $todate ) {
                if ($a[4] -ne "" -and ($a[4] -match "^\d*\.?\d*$") -and $a[4] -gt 0) {
                    $lastprice = $a[4];
                }
                $prop = @{'Date'=$a[0]; 'Close'=$lastprice};
                $lastQuote=$a[0]; 
                $obj = new-Object -TypeName PSObject -Property $prop;
                $retvalue.prices += $obj;
            } 
            $retvalue.result = $true;
        }

    }
    catch {
        $errmessage = "ERROR: YAHOO get $yahoosymbol Failed. web error by catched." + $_;
        if ($logFile) { $errmessage | Out-File $logFile -Encoding OEM -Append; }
        $errmessage + " " + $_ | Write-Warning ;
        return $retvalue ;
    }
    if ($logFile) {"YAHOO get $yahoosymbol price. $fromdate - $todate Successful. " + $retvalue.prices.count + " rows added." | Out-File $logFile -Encoding OEM -Append;}
    
    #return following values:
    $retvalue;
    

}