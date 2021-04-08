$ExcelWorkbookPath = Read-Host -Prompt "Input a full path name of Schedule_UC_FYXX.xlsx without filename"
$ExcelWorksheetName = Read-Host -Prompt "Input a worksheet name"
$WebHookUrl = Read-Host -Prompt "Input a url of Teams Incoming Webhoook"
$jsonobj = ConvertTo-Json @{
    excelworkbookpath = $ExcelWorkbookPath
    excelworksheet = $ExcelWorksheetName
    url = $WebHookUrl
    Debugflag = $false
}
Set-Content -Path .\config.json -Value $jsonobj
