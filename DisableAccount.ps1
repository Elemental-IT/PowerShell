<#

INSTRUCTIONS: Enter user username and confirm 
FUNCTION: AD: Disable AD account, moves to disabled OU then removes *ALL* groups. 
FUNCTION: Exchange: Disables ActivesyncEnabled and OWA then gives the option to hide mailbox.
 
#>


#Import Sessions/Modules  
$session=New-PSSession -ConfigurationName microsoft.exchange -connectionuri http://mail01.doamin.local/powershell -Authentication Kerberos
Import-PSSession $session -AllowClobber
Import-Module ActiveDirectory
cls

#disclaimers 
Write-Warning "Don't use without training from a Senior Systems Engineer."
Write-Host "Script created By Theo.Bird." -ForegroundColor Cyan

#Set Variables
do{
$user = Read-Host "Enter username of account you want to disable [e.g. John.Doe]"
  try{
    $SamAccountName = $null
    $SamAccountName = get-aduser $user -ErrorAction Stop
    write-verbose "Username '$user' is valid" -Verbose    
  }catch{
    write-warning "Please enter a valid name"
  }
}until($SamAccountName)

$confirmation = Read-Host "CONFIRM: Disable $user [y/n]"
while($confirmation -ne "y")
{if ($confirmation -eq 'n') {exit}
$confirmation = Read-Host "CONFIRM: Disable $user [y/n]"}

<#Get date format#>
$date = get-date -Format y
$month,$year = $date.Split(' ')
$month = $month.Remove(3)
$year = $year.substring(2)
$OUdate = ($month + '-' + $year)

$DisDate = Get-Date -DisplayHint date
$SetDescription = ("Disabled on " + $DisDate)
$OUdate = get-Date -Format y
$OUdatePath = "OU=$OUdate,OU=Disable Accounts,OU=Solotel Users,DC=solotel,DC=local"

#Create new OU and catch error if it exists 
try {
New-ADOrganizationalUnit -Name $OUdate -Path "OU=Disable Accounts,OU=domain Users,DC=domain,DC=local" -Description "disabled users for $DisDate"
} catch {}

#Disable account, move to Disabled OU and remove all groups
Disable-ADAccount -Identity $User
Set-ADUser -Identity $User -Description $SetDescription
Get-ADUser -Identity $User | Move-ADObject -TargetPath $OUdatePath
Get-ADUser -Identity $User -Properties MemberOf | ForEach-Object {$_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false}
#Get-ADUser -Identity $User -Properties MemberOf | Out-File "filetxt"
# create log with 

#Exchange
Set-CasMailbox -Identity $User -OWAEnabled $False
Set-CasMailbox -Identity $User -ActivesyncEnabled $False 
Set-CASMailbox -Identity $User -OWAforDevicesEnabled $false 

$confirmation2 = Read-Host "CONFIRM: Hide $user from the Globle Address List [y/n]"
while($confirmation2 -ne "y")
{
    if ($confirmation2 -eq 'n') {exit}
    $confirmation2 = Read-Host "CONFIRM: Hide $user from the Globle Address List [y/n]"
}

Set-Mailbox -Identity $User -HiddenFromAddressListsEnabled $true
Remove-PSSession $session

write-verbose "$User account disabled successfully" -Verbose
write-host "Press any key to exit..."
[void][System.Console]::ReadKey($true)