# ####################################################################################################################################################
# Maxim T.
# This script should be included at the top of Quotes/Currency Exchange scripts. 
# This script reads config file into $config variable, sets $minDate variable and all folder related variables. 
# This script should be re-used in all main scripts.
#
$startTime      = GET-DATE; $reqCount = 0; $reqFailed = 0; $reqSucceed = 0;
$scriptPath     = Split-Path -parent $MyInvocation.MyCommand.Definition; 
$scriptPathParent = Split-Path -path $scriptPath -Parent; 
if (!(Test-Path -Path  ($scriptPath + "\Log\"))) {
    New-Item ($scriptPath + "\Log\") -type directory
};

$configFile     = $scriptPath + "\psConfig.txt"; 
if (!(Test-Path -Path $configFile)) {
    Write-Host "Config file not found: $configFile" -ForegroundColor Red; 
    exit(1)
};

$dataRootFolder = $scriptPathParent; 

# Getting config file without empty lines and comments
$config = Get-Content $configFile | Where-Object {$_.trim() -ne "" -and !$_.StartsWith("#") }; 

# get all apikeys
$apikeys = "";
$apikeysFile = $scriptPath + "\apikey.txt";
if (Test-Path -Path $apikeysFile) {
    $apikeys = Get-Content $apikeysFile | Where-Object {$_.trim() -ne "" -and !$_.StartsWith("#")};
} 
$verbose = $false; 
# Check if configured to do detail output
if (($config | Select-Object -Index(($config.IndexOf("<DetailOutput>"))+1)).Replace("</DetailOutput>","").ToLower() -eq "yes") {
    $verbose=$true;
};

# get Main Excel file path
$ExcelFile = ($config | Select-Object -Index(($config.indexOf("<ExcelFile>"))+1)).Replace("</ExcelFile>","");
$ExcelFilePath = $dataRootFolder + "\" + $ExcelFile;
if (!(Test-Path -Path $ExcelFilePath)) {
    Write-Host "*** Error. Main Excel File must exist before run Any PowerShell script to get data." -ForegroundColor Red;
    exit(1);
}

$minDate = ($config | Select-Object -Index(($config.IndexOf("<MinDate>"))+1)).Replace("</MinDate>","");
if ($minDate.Length -lt 10 -or $minDate -le "1960-01-01" -or $minDate -ge "2050-01-01") {
    Write-Host "*** Error. Min date in psconfig file is $minDateFile but should be between 1960-01-01 and 2050-01-01" -ForegroundColor Red; 
    exit(1);
}
$dataRootFolderCfg = ($config | Select-Object -Index(($config.IndexOf("<DataRootFolder>"))+1)).Replace("</DataRootFolder>","");
if ($dataRootFolderCfg -ne $null -and $dataRootFolderCfg -ne "") {
    $dataRootFolder = $dataRootFolderCfg
}; 
if (!(Test-Path -Path $dataRootFolder)) {New-Item $dataRootFolder -type directory};

$quotesFolder    = $dataRootFolder + "\Quotes\";           if (!(Test-Path $quotesFolder))     {New-Item $quotesFolder -type directory | Out-Null; if ($verbose) {"Make directory: $quotesFolder`r`n"}}; 
$currExchFolder  = $dataRootFolder + "\CurrExch\";         if (!(Test-Path $currExchFolder))   {New-Item $currExchFolder -type directory | Out-Null; if ($verbose) {"Make directory: $currExchFolder`r`n"}};
$todayYMD = (Get-Date).ToString("yyyy-MM-dd"); 
$reqCount = 0; $reqFailed = 0; $reqSucceed = 0; $reqRows = 0; $reqRowsT = 0; $roundTo = 6;

# Getting culture for proper formatting of number
$culture = New-Object System.Globalization.CultureInfo("en-US"); 

$psDataFolder   = $scriptPathParent + "\PSData"; 
$psDataFolderCfg = ($config | Select-Object -Index(($config.IndexOf("<PSDataFolder>"))+1)).Replace("</PSDataFolder>","");
if ($psDataFolderCfg -ne $null -and $psDataFolderCfg -ne "") {
    $psDataFolder = $psDataFolderCfg;
} 
if (!(Test-Path $psDataFolder)) {New-Item $psDataFolder -type directory}; 

#$colSepTxt = ($config | Select-Object -Index(($config.IndexOf("<ColumnSeparator>"))+1)).Replace("</ColumnSeparator>",""); $colSep = "`t";
#if ($colSepTxt.ToLower() -eq "comma") {$colSep = ","}; if ($colSepTxt.ToLower() -eq "tab") {$colSep = "`t"}; if ($colSepTxt.ToLower() -eq "verticalbar") {$colSep = "|"};
$colSep = "`t"; # Ignoring setting, we need to use this value!!!!
$decSep = "."; $decSep = ($config | Select-Object -Index(($config.IndexOf("<DecimalSeparator>"))+1)).Replace("</DecimalSeparator>",""); $decSep = $decSep.Substring(0, 1);
# ####################################################################################################################################################

# 3 lines below should be included into all "main" scripts, so that this inserted script will pre-define all required variables
#$scriptPath = Split-Path -parent $MyInvocation.MyCommand.Definition; $psSetVariables = $scriptPath + "\psSetVariables.ps1";
#. $psSetVariables;
