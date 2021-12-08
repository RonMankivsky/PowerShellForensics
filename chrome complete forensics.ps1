try {
    Import-Module ImportExcel
}
catch{
Write-Host "Open as admin and install the following module:            Install-Module -Name ImportExcel"
}





$path = "C:\Users\"+$env:UserName+"\AppData\Local\Google\Chrome\User Data\Default\Preferences"
$Content = Get-Content $path | ConvertFrom-Json  
$popups = $Content.profile.content_settings.exceptions.popups

#$popups
$Account = $Content.account_info
#$Account
$priner = $Content.printing.print_preview_sticky_settings.appState | ConvertFrom-Json 
$priner = $priner.recentDestinations
#$priner
$crashes = $Content.sessions.event_log
#$crashes
$installsig = $Content.extensions.install_signature.ids
$uninstall = $Content.extensions.external_uninstalls
$DesktopPath = [Environment]::GetFolderPath("Desktop").ToString() + "\output.xlsx"




$Account | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes1 -WorksheetName Account
$popups | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes2 -WorksheetName UserAllowedPopups
$priner | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes3 -WorksheetName Printers
$installsig | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes4 -WorksheetName Extentions_Install
$uninstall | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes5 -WorksheetName Extentions_Uninstall
$crashes | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes6 -WorksheetName ChromeCrashes

GetChromeHistory | Export-Excel -Path $DesktopPath -AutoSize -TableName Processes7 -WorksheetName ChromeHistory

Write-Host("File saved at: " + $DesktopPath )



function GetChromeHistory
{
$Path = 'C:\Users\'+$env:UserName+'\AppData\Local\Google\Chrome\User Data\Default\History'
    if (-not (Test-Path -Path $Path)) {
        Write-Verbose [!] Could not find Chrome History for username: $UserName
    }
    $Regex = '(http|https)://([\w-]+\.)+[\w-]+(/[\w- ./?%&=]*)*?'
    $Value = Get-Content -Path $path | Select-String -AllMatches $regex |% {($_.Matches).Value} |Sort -Unique
    $Value | ForEach-Object {
        $Key = $_
        if ($Key -match $Search){
            New-Object -TypeName PSObject -Property @{
                User = $env:UserName
                Browser = 'Chrome'
                DataType = 'History'
                Data = $_
            }
        }
    } 
}