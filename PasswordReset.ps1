<#
Resets users Password passsword to MonthYear Ie. May2019, unlocks account, sets account require new password at login 
#>

#Modules
Import-Module ActiveDirectory
cls

#disclaimers
"Enter Accountname - This resets the passsword to MonthYear Ie. May2019"

#Set Variables
$date = Get-Date -Format "Y" 
$date = $date -replace (' ', '')

$date = get-date -Format y
$month,$year = $date.Split(' ')
$month = $month.Remove(3)
$year
$OUdate = ($month + $year + "!")

do{
  $User = Read-Host "Enter username [e.g. John.Doe]"
  try{
    $SamAccountName = $null
    $SamAccountName = get-aduser $User -ErrorAction Stop
    write-verbose "Username '$User' is valid" -Verbose    
  }catch{
    write-warning "Please enter a valid name"
  }
}until($SamAccountName)        
 
$confirmation = Read-Host "CONFIRM: Reset password for $User [y/n]"
while($confirmation -ne "y")
{
    if ($confirmation -eq 'n') {exit}
    $confirmation = Read-Host "CONFIRM: Reset password for $Use [y/n]"
}

#unlock and reset password
Set-ADAccountPassword -Identity $User -Reset -NewPassword (ConvertTo-SecureString -AsPlainText $date -Force)
Set-ADuser -Identity $User -ChangePasswordAtLogon $True
Unlock-ADAccount -Identity $User

Write-Verbose "$User Password has been reset to $date"
write-host "Press any key to exit..."
[void][System.Console]::ReadKey($true)
