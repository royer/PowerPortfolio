# 2017-Sep-10. 
$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; 
. ($scriptPath + "\psFunctions.ps1");     # Adding script with reusable functions
. ($scriptPath + "\psSetVariables.ps1");  # Adding script to assign values to common variables

$logFile        = $scriptPath + "\Log\" + $MyInvocation.MyCommand.Name.Replace(".ps1",".txt"); 
(Get-Date).ToString("HH:mm:ss") + " --- Starting script " + $MyInvocation.MyCommand.Name | Out-File $logFile -Encoding OEM; # starting logging to file.
$logSummary = (Get-Date).ToString("HH:mm:ss") + " Script: Archive Quotes".PadRight(28);

if ($config.IndexOf("<ArchiveQuotes>") -eq -1) {$ArchiveQuotes = "";}
else {$ArchiveQuotes = ($config | Select-Object -Index(($config.IndexOf("<ArchiveQuotes>"))+1)).Replace("</ArchiveQuotes>","")}; 
if ($ArchiveQuotes -eq $null) {exit(1)}; if ($ArchiveQuotes.ToLower() -ne "yes") {exit(1);} # checking if need to archive quotes, if no, exit script.
$lastXdays = (Get-Date).AddDays(-50).ToString("yyyy-MM-dd"); # 50 calendar days is about 30 business days

$fl =  @(Get-ChildItem $quotesFolder -Recurse | Where-Object {$_.extension -eq ".csv" -and (!($_.BaseName -like "*_Archive"))});
ForEach($f in $fl) {
    $fc = @(import-csv $f.FullName -Header "Date","Symbol","Close" | Sort-Object -Property "Date");
    $fileRecCount = $fc.count;
    if ($fc.count -gt 0) { # File has at least one record
        $str = ""; $strArch = ""; $newRC=0; $arcRC=0;
        For ($i = 0; $i -lt $fileRecCount; $i++) {
            $currDate = $fc[$i].Date; 
            if ($i+1 -eq $fileRecCount) {$nextDate = $currDate} else {$nextDate = $fc[$i+1].Date;}
            if ($i -eq 0 -or $i+1 -eq $fileRecCount -or $currDate.Substring(5,2) -ne  $nextDate.Substring(5,2) -or $currDate -ge $lastXdays) { 
                # if first or last record or if month changed, then keep record
                $str+= $fc[$i].Date + "," + $fc[$i].Symbol +"," + $fc[$i].Close + "`r`n"; $newRC++;
            }
            else { # otherwise move record to archive folder
                $strArch+=$fc[$i].Date + "," + $fc[$i].Symbol  +"," + $fc[$i].Close + "`r`n"; $arcRC++;
            }
        }
        if ($str.Length -gt 1) {$str = $str.Substring(0, $str.Length-2);}
        if ($strArch.Length -gt 1) {$strArch = $strArch.Substring(0, $strArch.Length-2);}
        $fNew = $f.FullName.Replace(".csv", "_Archive.csv"); # Getting archive file name

        if ($strArch.Length -ne 0) {
            $strArch | out-file $fNew -Encoding oem -Append; # Appending to archive file
            if ($str.Length -ne 0) {$str | out-file $f.FullName -Encoding oem;} # replacing existing file, but just when there was something to archive
        }

        (Get-Date).ToString("HH:mm:ss") + " " + $f.Name.PadRight(22) + ". Was rec: " + $fileRecCount.ToString().PadLeft(5) + ". New rec: " + $newRC.ToString().PadLeft(5) + ". Archived rec: " + $arcRC.ToString().PadLeft(5) + "."  | Out-File $logFile -Encoding OEM -Append;
        $recRowsT+=$arcRC;
    }
}; 

$duration = (NEW-TIMESPAN -Start $startTime -End (Get-Date)).TotalSeconds.ToString("#,##0") + " sec.";
"Finished. Archived rec count: $recRowsT.  Duration: " + $duration | Out-File $logFile -Encoding OEM -Append;
$logSummary + ". Archived rec count: $recRowsT. Duration: $duration";
