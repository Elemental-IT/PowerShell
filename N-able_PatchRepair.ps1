<#
.synopsis
This script is to help resolve updates for patch V2 alerts in Nable on local device.

.Description
Resets Nable agent patching config, clears SoftwareDistribution folder and invokes windows updates and starts installing if any are found.

.Author Theo Bird

#>

Function Ultimate_patch_repair_script {

    $date = get-date -Format "MM/dd/yyyy"
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Windows Agent Maintenance Service"
    $patchCachePath = "C:\Program Files (x86)\N-able Technologies\PatchManagement"
    $patchConfig = "C:\Program Files (x86)\N-able Technologies\Windows Agent\config\PatchConfig.xml"
    $Key = "DelayedAutostart"
    $Value = "1"
    
    $services = @(
        "wuauserv"
        "Windows Agent Maintenance Service"
        "Windows Agent Service"
        "bits'"
        "cryptsvc"
    )

    if (test-path $env:SystemDrive\temp) { Start-Transcript "$env:SystemDrive\temp\Ultimate_patch_repair_script-$date" }
    else {New-Item -Path $env:SystemDrive\temp -ItemType Directory }

    # Set Services to Manual
    foreach ($service in $services) { Set-Service $service -StartupType Manual -Verbose }
    
    # Stop Services
    foreach ($service in $services) { Stop-Service $service -Force -Verbose }
        
    If (!(Test-Path $registryPath)) {
        Write-Warning "Cannot find N-Able Registry entry of $registryPath"
    }
    
    If (!(Get-ItemProperty -Path $registryPath -Name $Key).DelayedAutoStart -eq "0") {
        New-ItemProperty -Path $registryPath -Name $Key -Value $Value -PropertyType DWORD -Force
    }
        
    $Value = "2"
    
    If ((Get-ItemProperty -Path $registryPath -Name $Key).Start -ne "2") {
        New-ItemProperty -Path $registryPath -Name $Key -Value $Value -PropertyType DWORD -Force
    }
    
    $registryPath = "HKLM:\SYSTEM\CurrentControlSet\Services\Windows Agent Service"
    
    If (!(Test-Path $registryPath)) {
        Write-Warning "Cannot find N-Able Registry entry of $registryPath"
    }
    
    If ((Get-ItemProperty -Path $registryPath -Name $Key).Start -ne "2") {
        New-ItemProperty -Path $registryPath -Name $Key -Value $Value -PropertyType DWORD -Force
    }
    
    # Delete N-Able Patch Cache
    If (Test-Path $patchCachePath) {
        Write-Host "Removing N-Able Patch Cache folder and config file.."
        Remove-Item -Path $patchCachePath -Recurse -Force
        Remove-Item -Path $patchConfig
    }
    Else {
        Write-Warning "Cannot find N-Able Patch Cache Folder.."
    }
    
    #Remove Windows update cache & windowsupdate.log
    Remove-Item -Path "C:\Windows\SoftwareDistribution"-Recurse -Force -Verbose
    Remove-Item -Path "C:\windows\windowsupdate.log" -Force -Verbose
    Remove-Item -Path "C:\windows\System32\Catroot2" -Force -Verbose -Recurse

    #Registering DLL’s pertaining to Windows Update
    $RegDll = @(        
        "c:\windows\system32\vbscript.dll"
        "c:\windows\system32\mshtml.dll"
        "c:\windows\system32\msjava.dll"
        "c:\windows\system32\jscript.dll"
        "c:\windows\system32\msxml.dll"
        "c:\windows\system32\actxprxy.dll"
        "c:\windows\system32\shdocvw.dll"
        "wuapi.dll"
        "wuaueng1.dll"
        "wuaueng.dll"
        "wucltui.dll"
        "wups2.dll"
        "wups.dll"
        "wuweb.dll"
        "Softpub.dll"
        "Mssip32.dll"
        "Initpki.dll"
        "softpub.dll"
        "wintrust.dll"
        "initpki.dll"
        "dssenh.dll"
        "rsaenh.dll"
        "gpkcsp.dll"
        "sccbase.dll"
        "slbcsp.dll"
        "cryptdlg.dll"
        "Urlmon.dll"
        "Shdocvw.dll"
        "Msjava.dll"
        "Actxprxy.dll"
        "Oleaut32.dll"
        "Mshtml.dll"
        "msxml.dll"
        "msxml2.dll"
        "msxml3.dll"
        "Browseui.dll"
        "shell32.dll"
        "wuapi.dll"
        "wuaueng.dll"
        "wuaueng1.dll"
        "wucltui.dll"
        "wups.dll"
        "wuweb.dll"
        "jscript.dll"
        "Mssip32.dll"
    )

    foreach ($DLL in $RegDll) {
        $regsvrp = Start-Process regsvr32.exe -ArgumentList "/s $DLL" -PassThru
        $regsvrp.WaitForExit(5000) # Wait (up to) 5 seconds
        if($regsvrp.ExitCode -ne 0)
        {
            Write-Warning "regsvr32 exited with error $($regsvrp.ExitCode)"
        }
    }

    #Set Windows update services to Automatic
    foreach ($service in $services) { Set-Service $service -StartupType Automatic -Verbose }
        
    #restart Services 
    foreach ($service in $services) { restart-Service $service -Force -Verbose }

    #Wait for Services 
    Start-Sleep 5
    
    #Start updates

    If (($PSVersionTable).PSVersion.Major -ge '4') { 
    
        try {
            Get-InstalledModule  -Name NuGet -MinimumVersion 2.8.5.201 -ErrorAction Stop
        }
        catch {
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force | Out-Null 
            Install-Module PSWindowsUpdate -AllowClobber -Force | Out-null
        }
            
        Get-WindowsUpdate -RootCategories 'Critical Updates', 'Security Updates' -Install -IgnoreReboot -AcceptAll -Verbose   
        
        } else {
            Write-Warning "this Script Requires Powershell version 4 or greator to in still updates"
    }

    stop-Transcript
    
} Ultimate_patch_repair_script
    
