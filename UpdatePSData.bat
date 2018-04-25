echo off
REM powershell -ExecutionPolicy Bypass .\Scripts\GetQuotes-YahooIntraday.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetQuotes-Google.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetQuotes-GoogleWeb.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetQuotes-AlphaVantage.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetQuotes-Stooq.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetQuotes-GoogleIntraday.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-ECB.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-YahooIntraday.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-GoogleIntraday.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\psMakeAllDataFiles.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-BoC.ps1
REM powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-CurrencyLayer.ps1
rem powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-Stooq.ps1
powershell -ExecutionPolicy Bypass .\Scripts\GetExchRates-EODHis.ps1
powershell -ExecutionPolicy Bypass .\Scripts\GetPriceHistory.ps1
powershell -ExecutionPolicy Bypass .\Scripts\psArchiveOldQuotes.ps1
powershell -ExecutionPolicy Bypass .\Scripts\psMakeDataFiles.ps1
powershell -ExecutionPolicy Bypass .\Scripts\psCreateDatesFile.ps1
powershell -ExecutionPolicy Bypass .\Scripts\psAppendGeneratedQuotes.ps1
powershell -ExecutionPolicy Bypass .\Scripts\psCheckFiles.ps1
choice /C Y /T 10 /D Y /M "Waiting 10sec before closing"

