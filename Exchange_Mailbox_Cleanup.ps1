<#
.SYNOPSIS
Collects Mailbox details for exchange Cleanup

.DESCRIPTION
Collects last sent email, WhenChanged, IsMailboxEnabled etc. to help identify wtich mailboxes can be removed. 

.INPUTS
None.

.NOTES
* Only runs on server with 2013 or higher installed
** if LastSent is blank then emails have not been sent for 30 days (or whatever exchange logging has been set to)

Author: Theo Bird
#>

Function MailboxlastUsed {
    
    try {
    
        Add-PSSnapin "Microsoft.Exchange.Management.PowerShell.SnapIn" -ErrorAction stop
        
    }
    catch {  
        Write-Error "Exchange 2013 or higher not installed locally... "
        Write-Warning "Run Script on server with Exchange 2013 or higher installed"
        Start-Sleep 30
        Stop-Process $PID
    }
    
    $Path = "$env:SystemRoot\temp\MailboxlastUsed.csv"
    $WarningPreference = 'SilentlyContinue'
    $ErrorActionPreference = 'SilentlyContinue'
    $UserList = Get-mailbox -Resultsize unlimited
    $MasterList = @()
    
    if (Test-Path $Path) { Remove-Item $Path }
    
    foreach ($User in $UserList) {
    
        $MyObject = New-Object PSObject -Property @{
            MailboxSize      = $null
            LastSent         = $null
            WhenChanged      = $null
            WhenCreated      = $null
            EmailAddress     = $null
            IsMailboxEnabled = $null
            MailBoxType      = $null

        }
    
        Write-Host "Collecting data for $user.PrimarySmtpAddress..." -ForegroundColor Green
    
        $MyObject.LastSent = ((Get-TransportService | Get-MessageTrackingLog -ResultSize unlimited -Sender $user.PrimarySmtpAddress | Sort-Object timestamp)[-1]).timestamp
        $MyObject.EmailAddress = ($User).PrimarySmtpAddress
        $MyObject.MailboxSize = (Get-MailboxStatistics $User).TotalItemSize
        $MyObject.WhenCreated = ($User).WhenCreated
        $MyObject.WhenChanged = ($User).WhenChanged
        $MyObject.IsMailboxEnabled = ($User).IsMailboxEnabled
        $MyObject.MailBoxType = ($User).RecipientTypeDetails        
    
        $MasterList += $MyObject
    }
    
    $ErrorActionPreference = 'Continue'
    $WarningPreference = 'Continue'
    
    $MasterList | Select-Object EmailAddress, IsMailboxEnabled, MailBoxType, WhenCreated, WhenChanged, LastSent, MailboxSize | Export-Csv $Path
    Write-Verbose "List exported to $Path" -Verbose
    
    Start-Sleep 20
}   
 

$Admin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")

if ($Admin -eq $true) {
    MailboxlastUsed
}
else {
    Write-Warning "Run as administrator..." 
    Start-Sleep 30
    Stop-Process $PID
}
