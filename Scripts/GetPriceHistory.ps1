$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\ExcelModule.ps1");     # Adding script for import excel file module.

. ($scriptPath + "\fnGetPrice-Yahoo.ps1");  # Add get Price history from yahoo module.
. ($scriptPath + "\fnMakeMoneyFundQuotes.ps1"); # Add make money fund quotes module.
. ($scriptPath + "\fnGetPrice-NetEase.ps1");    # Add gete price history from NetEase.
. ($scriptPath + "\fnGetPrice-eastmoney.ps1");  # get China Mutual Fund quotes from EastMmoney.com
. ($scriptpath + "\fnGetPrice-stooq.ps1")       # get quotes from www.stooq.com. suggest data source.
. ($scriptPath + "\fnGetPrice-EODHis.ps1")      # get quotes from https://eodhistoricaldata.com


$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("yyyy-MM-dd HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: Get Price History".PadRight(28);
"Script: " + $MyInvocation.MyCommand.Name + " starting..."
#$excelfilepath = $scriptPathParent + "\powerportfolio.xlsx";

$logmsg = "Get symbol list from  $excelfilepath";
$logmsg | Out-File $logFile -Encoding OEM -Append; 
if ($verbose) { $logmsg}
$securities = Import-Excel-ListObject $ExcelFilePath "securities" "Symbols";
$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
$logmsg = "Load symbol list from $ExcelFilePath used $duration";
$logmsg | Out-File $logFile -Encoding OEM -Append;
if ($verbose) {$logmsg; }



# do not include *Cash
$securities = $securities | Select-Object -Property "symbol","Exchange","PriceProvider","PPSymbol", "secuType","Activited" |
    Where-Object {$_.Activited -ne "0" } | Where-Object {$_.secuType -ne "Cash"} | Where-Object { $_.ppsymbol -ne $null} ;
$logmsg = "Total " + $securities.Length + " symbols get from excel file" ;
$logmsg | Out-File $logFile -Encoding OEM -Append;
if ($verbose) { $logmsg;}

$providers = $securities | select-Object -Property "PriceProvider"  -Unique;


foreach ($provider in $providers) {
    switch ($provider.PriceProvider) {
        "EODHIS" {
            $eodhiss = $securities | Where-Object PriceProvider -EQ "EODHIS";
            if ($verbose) {"Total " + $eodhiss.count + " symbols will get from https://eodhistoricaldata.com ..."}
            # get apikey
            $EOD_APIKEY=""; $EOD_APIKEY = ($apikeys | Select-Object -Index(($apikeys.IndexOf("<EODHistorical>"))+1)).Replace("</EODHistorical>","");
            if ($EOD_APIKEY.Length -lt 1 )  {
                $str="*** Error. EODHistorical APIKey in apikey.txt file length is zero. skip fetch all EODHIS provider symbol!"; 
                $str | Out-File $logFile -Encoding OEM -Append; $str | write-Host -ForegroundColor Red; exit(1);
            }
            foreach ($symbol in $eodhiss) {
                $ret = GetSymbolLastInfo $symbol.symbol $quotesFolder $minDate;
                $symbolQuoteFile = $ret[0]; $nextDate = $ret[1]; $lastQuoteDate = $ret[2]; $lastprice = $ret[3];
                if ($nextDate -le $todayYMD) {
                    $reqCount++;
                    if ($verbose) {"Start get " + $symbol.symbol + " quotes $nextDate - $todayYMD from eodhistoricaldata.com ..."}
                    $result = GetPriceHistory_EODHis $symbol.PPSymbol $nextDate $todayYMD $lastprice $EOD_APIKEY $logFile;
                    if ($result.result) {
                        $reqSucceed++;
                        $reqRowsT += $result.prices.count;
                        $result.prices | ForEach-Object {
                            $_.Date + "," + $symbol.symbol + "," + $_.Close | Out-File $symbolQuoteFile -Encoding OEM -Append;
                        }
                        if ($verbose) {"Get "+$symbol.symbol + " from EODHistorical successful."}
                    } else {
                        $reqFailed++;
                        $errmsg = "Get " + $symbol.symbol + " Failed."
                        $errmsg | Out-File $logFile -Encoding OEM -Append;
                        $errmsg | Write-Warning; 
                    }
                }
            }
        }
        "STOOQ" {
            $stooqs = $securities | Where-Object PriceProvider -EQ "STOOQ";
            if ($verbose) {"Total " + $stooqs.count + " symbols will get from www.stooq.com ..."}
            foreach ($symbol in $stooqs) {
                $ret = GetSymbolLastInfo $symbol.symbol $quotesFolder $minDate;
                $symbolQuoteFile = $ret[0]; $nextDate = $ret[1]; $lastQuoteDate = $ret[2]; $lastprice = $ret[3];
                if ($nextDate -le $todayYMD) {
                    $reqCount++;
                    if ($verbose) {"Start get " + $symbol.symbol + " quotes $nextDate - $todayYMD from stooq.com ..."}
                    $result = getPriceHistory_Stooq $symbol.PPSymbol $nextDate $todayYMD $lastprice $logFile;
                    if ($result.result -eq $true) {
                        $reqSucceed++;
                        $reqRowsT += $result.prices.count;
                        $result.prices | ForEach-Object {
                            $_.Date + "," + $symbol.symbol + "," + $_.Close | Out-File $symbolQuoteFile -Encoding OEM -Append;
                        }
                        if ($verbose) { "Get " + $symbol.symbol + " from stooq Successful."}
                    }else {
                        $reqFailed++;
                        $errmsg = "Get " + $symbol.symbol + " Failed.";
                        $errmsg | Out-File $logFile -Encoding OEM -Append;
                        $errmsg | Write-Warning;
                    }
                }
            }
        }
        "YAHOO" {
            $yahoos = $securities | Where-Object PriceProvider -EQ "YAHOO";
            if ($verbose) {"Total " + $yahoos.count + " symbols will get from finance.yahoo.com ..."}
            $ret = GetYahooCrumbAndWebSession $logFile; $crumb = $ret[0]; $websession = $ret[1];
            foreach ($symbol in $yahoos) {
                $ret = GetSymbolLastInfo $symbol.symbol $quotesFolder $minDate;
                $symbolQuoteFile = $ret[0]; $nextDate = $ret[1]; $lastQuoteDate = $ret[2]; $lastprice = $ret[3];
                if ($nextDate -le $todayYMD) {
                    $reqCount++;
                    if ($verbose) {"Start get " + $symbol.symbol + " quotes $nextDate - $todayYMD from yahoo..."}
                    $result = GetPriceHistroy_Yahoo $crumb $websession $symbol.PPSymbol $nextDate $todayYMD $lastprice $logFile;
                    if ($result.result -eq $true) {
                        $reqSucceed++;
                        $reqRowsT += $result.prices.count;
                        $prices = $result.prices;
                        $prices | ForEach-Object {
                            $_.Date + "," + $Symbol.symbol + "," + $_.Close | Out-File $symbolQuoteFile -Encoding OEM -Append;
                        };
                        if ($verbose) {"Get "+$symbol.symbol +" from yahoo successful."}
                    } else {
                        $reqFailed++;
                        "Get "+$symbol.symbol+" Failed." | Write-Warning ;
                    }
                }
            }
        }
        "NETEASE" {
            $neteases = $securities | Where-Object PriceProvider -EQ "NETEASE";
            if ($verbose) {"Total " + $neteases.count + " symbols will get from www.163.com ..."}
            foreach ($symbol in $neteases) {
                $ret = GetSymbolLastInfo $symbol.symbol $quotesFolder $minDate;
                $symbolQuoteFile = $ret[0]; $nextDate = $ret[1]; $lastQuoteDate = $ret[2]; $lastprice=$ret[3];
                if ($nextDate -le $todayYMD) {
                    $reqCount++; 
                    if ($verbose) {"Start get " + $symbol.symbol + " quotes $nextDate - $todayYMD from 163.com..."}
                    $result = getPriceHistory_NetEase $symbol.PPSymbol.Substring(1) $nextDate $todayYMD $lastprice $logFile;
                    
                    
                    if ($result.result -eq $true) {
                        $reqSucceed++;
                        $reqRowsT += $result.prices.count;
                        $prices = $result.prices;
                        $prices | ForEach-Object {
                            $_.Date + "," + $symbol.symbol + "," + $_.Close | Out-File $symbolQuoteFile -Encoding OEM -Append;
                        }
                        if ($verbose) {"Get " + $symbol.symbol + " from 163.com successful."}
                    } else {
                        $reqFailed++;
                        "Get " + $symbol.symbol + " Failed." | Write-Warning;
                    }
                    # sleep a while for next request, or 163.com will block request.
                    $delayms = Get-Random -Maximum 1500 -Minimum 500;
                    if ($verbose) {" delay $demayms"+"ms for next request from 163.com..."}
                    Start-Sleep -Milliseconds $delayms;
                }
            }
        }
        "EASTMONEY" {
            $eastmoneys = $securities | Where-Object PriceProvider -EQ "EASTMONEY";
            if ($verbose) {"Total " + $eastmoneys.count + " symbols will get from www.eastmoney.com ..."}
            foreach ($symbol in $eastmoneys) {
                $ret = GetSymbolLastInfo $symbol.symbol $quotesFolder $minDate;
                $symbolQuoteFile = $ret[0]; $nextDate = $ret[1]; $lastQuoteDate = $ret[2]; $lastprice = $ret[3];
                if ($nextDate -le $todayYMD) {
                    $reqCount++;
                    if ($verbose) {"Start get " + $symbol.symbol + " quotes $nextDate - $todayYMD from eastmoney.com..."}
                    $result = GetPriceHistory_EastMoney $symbol.PPSymbol $nextDate $todayYMD $lastprice $logFile;
                    if ($result.result -eq $true) {
                        if ($result.prices.count -ne 0) {
                            $reqSucceed++;
                            $reqRowsT += $result.prices.count;
                            $result.prices | ForEach-Object {
                                $_.Date + "," + $Symbol.symbol + ',' + $_.Close | Out-File $symbolQuoteFile -Encoding OEM -Append;
                            }
                        }
                        if ($verbose) {"Get "+$symbol.symbol + " from eastmoney.com successful."}
                    } else {
                        $reqFailed++;
                        "Get " + $symbol.symbol + " Failed." | Write-Warning;
                    }
                }
            } 
        }
        "MONEYFUND" {
            $moneyfunds = $securities | Where-Object PriceProvider -EQ "MONEYFUND";
            if ($verbose) {"get " + $moneyfunds.count + " symbols from Fake MoneyFound ..."}
            foreach ($symbol in $moneyfunds) {
                $ret = GetSymbolLastInfo $symbol.symbol $quotesFolder $minDate;
                $symbolQuoteFile = $ret[0]; $nextDate = $ret[1]; $lastQuoteDate = $ret[2];
                if ($nextDate -le $todayYMD) {
                    $reqCount++;
                    if ($verbose) {"Start make fake quotes for "+$symbol.symbol + " ..."}
                    $result = MakeMoneyFundQuotes $symbol.PPSymbol $nextDate $todayYMD $logFile;
                    if ($result.result -eq $true) {
                        $reqSucceed++;
                        $reqRowsT += $result.prices.count;
                        $prices = $result.prices;
                        $prices | ForEach-Object {
                            $_.Date + "," + $Symbol.symbol + "," + $_.Close | Out-File $symbolQuoteFile -Encoding OEM -Append;
                        };
                        if ($verbose) {"Make " + $symbol.symbol + " fake quotes successful."}
                    } else {
                        $reqFailed++;
                        "Get "+$symbol.symbol+" Failed." | Write-Warning ;
                    }
                }                
            }
        }
        default {
            "Unknow price Provider: " + $provider.PriceProvider | Out-File $logFile -Encoding OEM -Append;
            "Unknow price Provider: "+ $provider.PriceProvider | Write-Warning;
        }
    }
}


$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. Quotes Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". Quotes Requested/Succeed/Failed/Rows: $reqCount/$reqSucceed/$reqFailed/$reqRowsT. Duration: $duration";