<#

.SYNOPSIS
office 2016 installer

.DESCRIPTION
Silent unintaller for office 2016(MSI verion) if WMI Method uninstallation throws a 1603 error 

Code
$computername =''
$appname = 'Microsoft Office Professional Plus 2016'

(Get-WmiObject -Class Win32_Product -ComputerName $computername | Where-Object{​​​​​​$_.Name -eq $appname}​​​​).Uninstall()

.NOTES
Script Based of these steps:
https://support.microsoft.com/en-us/office/manually-uninstall-office-4e2904ea-25c8-4544-99ee-17696bb3027b
 
Author Theo bird (Bedlem55)
  
#>


#Start logs
Start-Transcript "$env:SystemDrive\Remove office.log"

# create PSdrive to HKEY_CLASSES_ROOT
New-PSDrive -PSProvider registry -Root HKEY_CLASSES_ROOT -Name HKCR

# Back up Reg 
New-Item "$env:SystemDrive\RegBack" -ItemType Directory -Force

$Export_paths = @(

    "$env:SystemDrive\RegBack\HKCR.Reg"
    "$env:SystemDrive\RegBack\HKCU.Reg"
    "$env:SystemDrive\RegBack\HKLM.Reg"
    "$env:SystemDrive\RegBack\HKU.Reg"
    "$env:SystemDrive\RegBack\HKCC.Reg"

)

Foreach ($Export in $Export_paths) { reg export HKLM $Export /y }

# remove folders and shortcuts 
$Paths = @(

    "$env:SystemDrive\users\public\desktop\Excel 2016.lnk"
    "$env:SystemDrive\users\public\desktop\Outlook 2016.lnk"
    "$env:SystemDrive\users\public\desktop\Word 2016.lnk"
    "$Env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Excel 2016.lnk"
    "$Env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Outlook 2016.lnk"
    "$Env:ALLUSERSPROFILE\Microsoft\Windows\Start Menu\Programs\Word 2016.lnk"
    "$env:SystemDrive\Program Files (x86)\Microsoft Office\Office16"
    "$env:SystemDrive\Program Files\Microsoft Office\Office16"

)

Foreach($Path in $Paths) { $test = Test-Path $Path

    $test = Test-Path $Path
    If ($test -eq $true) {Remove-Item $Path -Force -Recurse -Verbose}
 
 }

# remove Reg Keys
$Reg = @( 

    # 32 Bit 
    "HKCR:\Software\Microsoft\Office\16.0"
    "HKLM:\SOFTWARE\Microsoft\Office\16.0"
    "HKLM:\SOFTWARE\Microsoft\Office\Delivery\SourceEngine\Downloads\*0FF1CE}-*"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*0FF1CE*"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\Upgrade Codes\*F01FEC"
    "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Installer\UserData\S-1-5-18\Products\*F01FEC"
    "HKLM:\CurrentControlSet\Services\ose"
    "HKCR:\Installer\Features\*F01FEC"
    "HKCR:\Installer\Products\*F01FEC"
    "HKCR:\Installer\UpgradeCodes\*F01FEC"
    "HKCR:\Installer\Win32Assemblies\*Office16*"

    #64 Bit 
    "HKCR:\Software\Microsoft\Office\16.0"
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Office\16.0"
    "HKLM:\Wow6432Node\Microsoft\Office\Delivery\SourceEngine\Downloads\*0FF1CE}-*"
    "HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*0FF1CE*"
    "HKLM:\SYSTEM\CurrentControlSet\Services\ose"
    "HKCR:\Installer\Features\*F01FEC"
    "HKCR:\Installer\Products\*F01FEC"
    "HKCR:\Installer\UpgradeCodes\*F01FEC"
    "HKCR:\Installer\Win32Asemblies\*Office16*"

)

# test path for each Reg and remove if path is true
Foreach($Reg_Path in $Reg) {
 
  $test = Test-Path $Reg_Path
  If ($test -eq $true) {Remove-Item $Reg_Path -Recurse -Verbose -Force}
  
}
