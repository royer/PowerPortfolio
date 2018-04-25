# Money fund has not price, it always is 1$
# make a fake quotes data file

function MakeMoneyFundQuotes($symbol, $fromdate, $todate, $logfile) {

    $retvalue = New-Object -TypeName psobject -Property @{'result'=$true; 'prices'=@();};

    $dateFrom = get-Date $fromdate;

    $dateTo = get-Date $todate;

    $d = $dateFrom
    while ($d -le $dateTo) {
        if ($d.DayOfWeek -ne 'Saturday' -and $d.DayOfWeek -ne 'Sunday') {
            $strD = $d.ToString("yyyy-MM-dd");
            $prop = @{'Date'=$strD; 'Close'="1.0"} ;
            $obj = New-Object -TypeName psobject -Prop $prop;
            $retvalue.prices += $obj; 
        }
        
        $d = $d.AddDays(1)
    }

    if ($logfile) {"Make MoneyFund Quotes $fromdate - $todate successful." + $retvalue.prices.count + " row added." | Out-File $logfile -Encoding OEM -Append;}

    $retvalue;
}