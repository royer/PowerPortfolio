# 2017-Sep-10. Created by Maxim T. 
# 2018-Apl-15. Modified by Royer.

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables
. ($scriptPath + "\ExcelModule.ps1");      # Adding Read Excel Module

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: Append GenQ".PadRight(28);
" PSData folder: $psDataFolder" | Out-File $logFile -Encoding OEM -Append;

$includeArchiveFlag = $false; $includeArchiveValue = ($config | Select-Object -Index(($config.IndexOf("<IncludeQuoteArchiveFolder>"))+1)).Replace("</IncludeQuoteArchiveFolder>",""); # Getting IncludeArchive value
if ($includeArchiveValue -ne $null) {if ($includeArchiveValue.Trim().ToLower() -eq "yes") {$includeArchiveFlag = $true;}} 

$outFile = $psDataFolder + "\Quotes.csv"; 

$st_loadexcel = get-date;

$genQuotes = import-Excel-ListObject $ExcelFilePath "GenQuotes" "GenerateQuotes"
$duration = (NEW-TIMESPAN -Start $st_loadexcel -End (Get-Date)).TotalMilliseconds.ToString("#,##0") + " ms.";
" Load all generated quotes from excel file used $duration" | Out-File $logFile -Encoding OEM -Append;
$symbolCount = 0;
$newRCT = 0;


if ($genQuotes.Count -gt 0) {

    $dMinDate = [datetime]::ParseExact($minDate, "yyyy-MM-dd", $null);
    $dToday = Get-Date;

    $symbols = @( $genQuotes | Select-Object -Prop "Symbol" -Unique );
    $symbolCount = $symbols.Count;

    foreach ($symbol in $symbols) {
        $rows = @( $genQuotes | Where-Object Symbol -eq $symbol.Symbol | Sort-Object -Prop "Date")
        $str = ""
        $dLastQuote = $dMinDate;
        $dLastXDate = $dToday.addDays(-50);     # 50 calendar days is about 30 business days
        $LastPrice = 0;
        $newRC = 0;
        for ($i = 0; $i -lt $rows.Count; $i++) {
            $LastPrice = $rows[$i].Close
            $currDate = $rows[$i].Date
            if ($currDate -lt $dMinDate) {continue; } # skip date early than mindate
            if ($i+1 -eq $rows.Count ) { $nextDate = $currDate} else {$nextDate = $rows[$i+1].Date}
            if ($i -eq 0 -or $i+1 -eq $rows.Count -or $currDate.Month -ne $nextDate.Month -or $currDate -ge $dLastXDate) {
                # first/last record, next reocrd month is diffirent or date to today less than 50 calendar days.
                # add current row to 
                $dLastQuote = $rows[$i].Date
                
                $str += $rows[$i].Date.ToString("yyyy-MM-dd") + $colSep + $rows[$i].Symbol + $colSep + $rows[$i].Close + "`r`n";
                $newRC++;

                if (($nextDate - $currDate).days -gt 31) {
                    # insert new record
                    $ddate = $currDate;
                    do {
                        $ddate = $ddate.AddMonths(1);
                        $ddate = $ddate.addDays((-1)*$ddate.Day + 1);  # make it as first day of month
                        if (($ddate.ToString("ddd"), $culture) -eq "Sat") { $ddate = $ddate.AddDays(2); }
                        if (($ddate.ToString("ddd", $culture)) -eq "Sun") { $ddate = $ddate.AddDays(1); }
                        if ($ddate -lt $nextDate -and $ddate -le $dToday) {
                            $str += $ddate.toString("yyyy-MM-dd") + $colSep + $rows[$i].Symbol + $colSep + $rows[$i].Close + "`r`n";
                            $dLastQuote = $ddate;
                            $newRC++;
                        }
                    }while($ddate -lt $nextDate -and $ddate -le $dToday)

                }
            }
        }
        if (($dToday - $dLastQuote).Days -gt 31) {
            $ddate = $dLastQuote;
            do {
                $ddate = $ddate.AddMonths(1);
                $ddate = $ddate.addDays((-1)*$ddate.Day + 1);
                if (($ddate.ToString("ddd"), $culture) -eq "Sat") { $ddate = $ddate.AddDays(2); }
                if (($ddate.ToString("ddd", $culture)) -eq "Sun") { $ddate = $ddate.AddDays(1); }
                if ($ddate -le $dToday) {
                    $str += $ddate.toString("yyyy-MM-dd") + $colSep + $symbol.Symbol + $colSep + $LastPrice + "`r`n";
                    $newRC++; 
                }
            }while ($ddate -le $dToday)
        }
        if ($str.Length -ge 2) {$str = $str.Substring(0, $str.Length-2);}
        if ($str.Length -gt 0) {
            $str | Out-File $outFile -Encoding OEM -Append;
        }
        " Symbol: " + $symbol.Symbol +", old record count: " + $rows.Count + ". Generated record count: $newRC" | Out-File $logFile -Encoding OEM -append;
        $newRCT += $newRC;
    }
    
}
else { "** No symbols specified for generated quotes" | Out-File $logFile -Encoding OEM -append;}
# #######################################################################################
$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
(Get-Date).ToString("HH:mm:ss") + " --- Finished. GenQuotes SymbolCount/RecCount: $symbolCount/$newRCT. Duration: $duration`r`n" | Out-File $logFile -Encoding OEM -append;
$logSummary + ". GenQuotes SymbolCount/RecCount: $symbolCount/$newRCT. Duration: $duration";
