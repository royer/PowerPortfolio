# 2017-Sep-10. Created by Maxim T. 


function GetSymbolLastInfo($symbol,$quotesFolder, $minDate) {
    $symbolQuoteFile = $quotesFolder + $symbol.replace(' ',"_").replace('^','_').replace('&','_')+".csv";
    $nextDate = $minDate;
    $lastprice = 1.0;
    if (Test-Path -path $symbolQuoteFile) {
        #quote file exist. need to get last available quote date.
        $prices = @(import-Csv $symbolQuoteFile -Header "Date","Symbol", "Close" | Sort-Object -Prop "Date");
        $lastDateInFile = "";
        if ($prices.count -gt 0) { $lastDateInFile = $prices[$prices.count-1].Date; }
        if ($lastDateInFile -eq $null) { $lastDateInFile = "";}
        if ($lastDateInFile -ne "") {
            # Adding one day to max date found in quote file - this will be next request start date
            $nextDate = ([datetime]::ParseExact($lastDateInFile,"yyyy-MM-dd",$null)).AddDays(1).ToString("yyyy-MM-dd"); 
            $lastprice = $prices[$prices.count-1].Close;
        }
    }
    
    $lastQuoteDate = ([datetime]::ParseExact($nextDate,"yyyy-MM-dd",$null)).AddDays(-1).ToString("yyyy-MM-dd");

    #return following values:
    $symbolQuoteFile;   #0
    $nextDate;          #1
    $lastQuoteDate;     #2
    $lastprice;         #3
}


function GetFxRateInfo($currencycode, $currExchFolder, $minDate) {

    $currExchFile = $currExchFolder + $currencycode + ".csv";
    $nextDate = $minDate;

    # Getting date of last quote for currency in quote file we already have.
    # if there is no quote file, use minDate from parameter file.
    if (Test-Path -Path $currExchFile) {
        # File exists, need to find last available quote date
        $fc = @(Import-Csv $currExchFile -Header "Date", "code", "rate" | Sort-Object -Property "Date");
        $lastDateInFile = "";
        if ($fc.count -gt 0) { $lastDateInFile = $fc[$fc.count-1].Date;}
        # Adding one day to max date found in quote file - this will be next request start date
        if ($lastDateInFile -ne "") { $nextDate = ([datetime]::ParseExact($lastDateInFile,"yyyy-MM-dd",$null)).AddDays(1).ToString("yyyy-MM-dd");}
    }
    $currExchFile;      # 0
    $nextDate;          # 1
}

