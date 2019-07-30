<#
This creates a CSV with username and SecureString Password used by other FTP Scripts for uploading  
#>

$Path = "C:\temp\Key.csv"
$User = "UserName"
$Pswd = ConvertTo-SecureString -String 'P@$$w0rD!' -AsPlainText -Force
$Cred = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $User, $Pswd
$User = $Cred.UserName | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString 
$Pswd = $Cred.Password | ConvertFrom-SecureString 
New-Item -Path "C:\temp\OK\" -ItemType directory -ErrorAction SilentlyContinue -Force
New-Item -Path $Path -ItemType file -ErrorAction SilentlyContinue -Force
Add-Content -Path $Path -Value '"User","Password"' -Force
$NewLine = "{0},{1}" -f $User,$Pswd
$NewLine | add-content -Path $Path -Force