# standalone Powershell script to tetch stock price from EODHistoricaldata.com
param(
    # Parameter help description
    [Parameter(mandatory=$true)][string]$Symbol,
    [string]$FromDate="",
    [string]$ToDate = (Get-Date).ToString("yyyy-MM-dd"),
    [string]$Interval="d",
    [string]$Format = "csv"
)

$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition;

$apikeysFile = $scriptPath + "\apikey.txt";
if ( Test-Path $apikeysFile ) {
    $apikeys = Get-Content $apikeysFile | Where-Object {$_.trim() -ne "" -and !$_.StartsWith("#")};
    $APIKEY = ($apikeys | Select-Object -Index(($apikeys.IndexOf("<EODHistorical>"))+1)).Replace("</EODHistorical>","");

    if (($APIKEY -eq "demo" ) -or ($APIKEY -eq "")) {
        "The APIkey in $apikeysFile semms incorrect." | write-Host -ForegroundColor Red;
        return; 
    }
} else {
    "Cannnot find apikey.txt file in $scriptPath. script exit." | Write-Host -ForegroundColor Red;
    return;
}

$url = "https://eodhistoricaldata.com/api/eod/" + $Symbol + "?api_token="+$APIKEY;
if ($FromDate -ne "") {
    $url += ("&from=" + $FromDate);
}
$url += ("&to="+$ToDate);
$url += ("&fmt=$Format&period=$Interval");

try {

    $webResponse = Invoke-WebRequest -Uri $url;
    $result = $webResponse.Content;
    if ($Format -eq "csv") {
        # remove the last line. which is total characters
        $result = $webResponse.Content.SubString(0,$webResponse.Content.LastIndexOf("`n"))
    }
    return $result;
   
}catch {
    "Error: " + $_ | Write-Host -ForegroundColor Red;
}




