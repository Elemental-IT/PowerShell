#Requires -RunAsAdministrator

<#
.Synopsis
System Clean Up Temp, System and User files

.Author Theo bird
#>


Function Cleanup {
        
    $LogDate = Get-Date -Format dd-mm-yyyy
    Start-Transcript -Path "$env:SystemDrive\Temp\SystemClean_$LogDate.log"
    
    $ErrorActionPreference = "SilentlyContinue"

    Clear-Host
                    
    # Stops the windows update service. 
    Get-Service -Name wuauserv | Stop-Service -Force -Verbose 

    # Deletes the contents of windows software distribution.
    Get-ChildItem "$env:SystemDrive\Windows\SoftwareDistribution\*" -Recurse -Force | remove-item -force -Verbose -recurse

    # Deletes the contents of the Windows Temp folder. 
    Get-ChildItem "$env:SystemDrive\Windows\Temp\*" -Recurse -Force | remove-item -force -Verbose -recurse 

    # Removes items in the downloads folder not accessed for more than 3 months
    Get-ChildItem -Path "$env:SystemDrive\users\*\downloads\*" -Recurse | Where-Object { ($_.LastAccessTime -le $(Get-Date).AddDays(-90)) } | remove-item -force -Verbose -recurse 

    # Remove all files and folders in user's Temporary Internet Files. 
    Get-ChildItem "$env:SystemDrive\users\*\AppData\Local\Microsoft\Windows\Temporary Internet Files\*" -Recurse -Force | remove-item -force -recurse -Verbose
                    
    # Deletes the contents of the recycling Bin.
    Clear-RecycleBin -force 

    # Starts the Windows Update Service
    Get-Service -Name wuauserv | restart-Service -Verbose

    $ErrorActionPreference = "Continue"

    Stop-Transcript
    Start-Sleep 15
    Stop-Process $PID

} Cleanup

$PSV = ($PSVersionTable).PSVersion.Major 

switch ( $PSV ) {
    5 { Cleanup } 
    Default { 
        Write-Error "Powershell version 5 required"
        Start-Sleep 15
        Stop-Process $PID
    }
}
