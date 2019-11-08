<#
This Script is for copying user accounts
1. Then username for new account." 
2. Then user account details, ie Phone number etc..." 
3. Passsword format is MonthYear Ie. May2019"
4. This will copy the groups to the new account and enable the mailbox" 
#>

#Import Modules
$session=New-PSSession -ConfigurationName microsoft.exchange -connectionuri http://mail01.Domain.local/powershell -Authentication Kerberos
Import-PSSession $session -AllowClobber
Import-Module ActiveDirectory
cls

#disclaimers
Write-Warning "Don't use without training from a Senior Systems Engineer."
Write-Host "Script created By Theo.Bird." -ForegroundColor Cyan

#Set Variables
do{
  $UserCopy = Read-Host "Enter username Of account you want to copy [e.g. John.Doe]"
  try{
    $SamAccountName = $null
    $SamAccountName = get-aduser $UserCopy -ErrorAction Stop
    write-verbose "Username '$UserCopy' is valid" -Verbose    
  }catch{
    write-warning "Please enter a valid name"
  }
}until($SamAccountName)          

$UserNew = Read-Host "Enter username of new user [e.g. John.Doe]"
if (Get-ADUser -F {SamAccountName -eq $UserNew})
       {
   Write-Warning "A user account $UserNew has already exist in Active Directory."
   $UserNew = Read-Host  "Enter username of new user [e.g. John.Doe]"
       }

do{
  $Manger= Read-Host "Enter new user's Manage's user name [e.g. John.Doe]"
  try{
    $SamAccountNameManager = $null
    $SamAccountNameManager = get-aduser $Manger -ErrorAction Stop
    write-verbose "Username '$Manger' is valid" -Verbose    
  }catch{
    write-warning "Please enter a valid name"
  }
}until($SamAccountNameManager)  

<#Get OU#>
$Department = Read-Host "Office user? [y/n]"
if ($Department -eq 'y') 
    {$OU = "OU=Head Office,OU=Doamin Users,DC=Domain,DC=local"}
elseif ($Department -eq 'n')
    {$OU = "OU=Hotels,OU=Doamin Users,DC=Domain,DC=local"}

<#Set other AD Variables#>
    $FirstName = Read-Host "Enter Frist name"
    $LastName = Read-Host "Enter Last name"
    $JobTitle = Read-Host "Enter Job title"
    $Mobile = Read-Host "Enter Mobile Number"
    $Useremail = ($UserNew + "@Domain.com.au")
    $Fullname = ($FirstName + " " + $LastName)
    $userInstance = Get-ADUser -Identity $UserCopy

<#Set password#> #could just make this $olotel*123
    $date = get-date -Format y
    $month,$year = $date.Split(' ')
    $month = $month.Remove(3)
    $sym = Get-Random -InputObject '!','@','#','$','%','^','&','*','(',')'
    $PWFormat = ($month + $year + $sym)
    $Password = (ConvertTo-SecureString -AsPlainText $PWFormat -Force)

<#Get mail database#>
$letter = $UserNew.Remove(1)

switch -Regex ($letter) {
    
   [ab]         { $DBs = "MAIL01 Firstname A-B" }
   [cd]         { $DBs = "MAIL01 Firstname c-d" }
   [efhi]       { $DBs = "MAIL01 Firstname E-I" }
   [jkl]        { $DBs = "MAIL01 Firstname J-L" } 
   [mno]        { $DBs = "MAIL01 Firstname M-O" } 
   [pqs]        { $DBs = "MAIL01 Firstname P-R" }
   [stuvwxyz]   { $DBs = "MAIL01 Firstname S-Z" }
}
 
#on-prem Account creation
$aduserparam = @{
   'Instance' = $userInstance
    'Name' = $Fullname 
    'GivenName' = $FirstName
    'Surname' = $LastName
    'DisplayName' = $Fullname
    'SamAccountName' = $UserNew 
    'UserPrincipalName' = $Useremail
    'Path' = $OU
    'Enabled' = $true
    'EmailAddress' = $Useremail
    'Title' = $JobTitle 
    'Manager' = $Manger
    'MobilePhone'  =$Mobile
    'Department' = $Department
    'AccountPassword' = $Password
    'ChangePasswordAtLogon' = $True
}

#Create ADaccount 
New-ADUser @aduserparam | out-null

#Adduser memberships
$CopyFromUser = Get-ADUser $UserCopy -Properties MemberOf
$CopyToUser = Get-ADUser $UserNew -Properties MemberOf
$CopyFromUser.MemberOf | Where{$CopyToUser.MemberOf -notcontains $_} |  Add-ADGroupMember -Member $CopyToUser | out-null

#Create mailbox
Enable-Mailbox -Identity $Useremail -Database $DBs | out-null
Remove-PSSession $session

write-verbose "$UserNew account created successfully the passowrd is $PWFormat"
write-host "Press any key to exit..."
[void][System.Console]::ReadKey($true) 
