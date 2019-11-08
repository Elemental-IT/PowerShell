<#
This Script is for creating accounts with MS froms, MS flow using a gatesway and Schedlue tasks
1. Fill in custom form 
2. Flow creates Newuser.csv via gateway
3. Schedlue tasks Tiggers this script and creates AD account
4. Emails service desk to with user details and requirements
#>


Import-Module ActiveDirectory

Import-Csv -Delimiter "," -path "C:\temp\Newuser.csv" | ForEach-Object {
 $manager = $_.manager
 $Startdate = $_.Startdate
 $Enddate = $_.Enddate
 $givenname = $_.Firstname
 $surname = $_.lastname
 $department = $_.department
 $Software = $_.Softwarerequired
 $Hardware = $_.Hardwarerequired
 $Special = $_.Special
 }

$name = ("$givenname" + " " + "$surname" )
$SAM = ("$givenname" + "." + "$surname" )
$UPN = ("$SAM" + "@test.local")

switch ($Department) 
{
"Sales"{
$template = "Sales.Template" 
$OU = "OU=Users,OU=Sales,OU=Test Domain,DC=test,DC=local"
$ADDepartment = "Sales"
}
"Marketing"{
$template = "Markrting.Template"
$OU = "OU=Users,OU=Marketing,OU=Test Domain,DC=test,DC=local"
$ADDepartment = "Marketing"
}
"Human Resourecs"{
$template = "HR.template"
$OU = "OU=Users,OU=HR,OU=Test Domain,DC=test,DC=local"
$ADDepartment = "Human Resourecs" 
}
"Information Technology"{
$template = "IT.template"
$OU = "OU=Users,OU=IT,OU=Test Domain,DC=test,DC=local"
$ADDepartment = "Information Technology"
}
}

<#Setpassword#>
$date = get-date -Format y
$month,$year = $date.Split(' ')
$month = $month.Remove(3)
$PWFormat = ($month + $year)
$Password = (ConvertTo-SecureString -AsPlainText $PWFormat -Force)

$User = @{ 
    'Instance' = $template
    'Name' = $name 
    'GivenName' = $givenname
    'Surname' = $surname
    'DisplayName' = $Name
    'SamAccountName' = $SAM
    'UserPrincipalName' = $UPN
    'Path' = $OU
    'Enabled' = $false
    'EmailAddress' = $UPN
    'Title' = $JobTitle 
    'Manager' = $Manger
    'MobilePhone'  = $Mobile
    'Department' = $ADDepartment
    'AccountPassword' = $Password
    'ChangePasswordAtLogon' = $True
    'AccountExpirationDate' = $Enddate
    }

start-sleep 5

New-ADUser @User 

<#Adduser memberships#>
$CopyFromUser = Get-ADUser $template -Properties MemberOf
$CopyToUser = Get-ADUser $SAM -Properties MemberOf
$CopyFromUser.MemberOf | Where{$CopyToUser.MemberOf -notcontains $_} |  Add-ADGroupMember -Members $CopyToUser

start-sleep 5

$cred = New-Object system.management.automation.pscredential -ArgumentList "Emailaccount@domain.com", (Get-Content "password" | ConvertTo-SecureString)
$toaddress = "Emailaccount@domain.com"
$Fromaddress = "Emailaccount@domain.com"
$Smtpserver = "smtp.office365.com"
$port = "587"
$Subject = "New account $SAM has been created"

<#Email Message#> 
$body = "<HTML><HEAD><META http-equiv=""Content-Type"" content=""text/html; charset=iso-8859-1"" /><TITLE></TITLE></HEAD>"
$body += "<BODY bgcolor=""#FFFFFF"" style=""font-size: Small; font-family: TAHOMA; color: #000000""><P>"
$body += "Dear Servicedesk <br> <br>"
$body += "Account $SAM has been created with passowrd $Password <br>"
$body += "Please Provision $SAM with the Following and enabled account on $Startdate <br>"
$body += "Software: $Software <br>"
$body += "Hardware: $Hardware <br>" 
$body += "Access: $Special <br>"


$Email = @{
 'Credential' = $cred
 'body' = $body 
 'Subject' = $Subject
 'from' = $Fromaddress
 'to' = $toaddress
 'port' = $port
}

Send-MailMessage @email -BodyAsHtml 
