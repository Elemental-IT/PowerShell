   
<#
.SYNOPSIS
Bulk mailbox enabled 
.DESCRIPTION
Bulk mail box enable via CSV 
.NOTES
  $CSV_Column: Is the column that is CSV has the user names for the enabled mailboxes. 
  $ConnectionUri: This is the fully qualified domain name for the Exchange Server 
   
Author Theo bird (Bedlem55)
  
#>

# 
$CSV_Path = ''
$CSV_Column = ''
$DataBase = ''
$ConnectionUri = ''

# connected to server
$UserCredential = Get-Credential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri "http://$ConnectionUri/PowerShell/" -Authentication Kerberos -Credential $UserCredential
Import-PSSession $Session -DisableNameChecking

# Imports csv
$CSV = Import-Csv -Path $CSV_Path
$Users =  $csv.($CSV_Column)

# for each account enable mailbox and output to shell
Foreach($User in $Users) {

    Try {
    Enable-Mailbox -Identity $User -Database $DataBase -ErrorAction Stop 
    $Mailbox = (Get-mailbox $User).UserPrincipalName
    # Write-Host "$Mailbox mailbox enabled"
    } Catch { $Error[0] }

}