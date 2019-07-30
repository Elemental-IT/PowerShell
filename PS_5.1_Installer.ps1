<#
PS verions 5.1 updater for server 2008 r2 and 2012 r2
#>

function InstallWMF5.1Server2008 {
    
    #7zip Variables 
    $URI_7Zip = "https://www.7-zip.org/a/7z1900.exe"
    $INS_7Zip = "C:\Temp\7z1900.exe"

    #WMF 5.1 Variables
    $URI_WMF = "https://go.microsoft.com/fwlink/?linkid=839523"
    $INS_WMFzip = "C:\Temp\WMF 5.1.zip"
    $INS_WMFexe = "C:\Temp\Win7AndW2K8R2-KB3191566-x64.msu" 
    $INS_DIR = "C:\Temp"
    $Arguments = "/quiet /norestart"

    #Create Path C:\Temp 
    If (-not(Get-Item $INS_DIR)) { 
        New-Item -Path "C:\" -Name "Temp" -ItemType directory -ErrorAction SilentlyContinue -Force 
    }

    Set-Location $INS_DIR
 
    #Check for KB2506143 and uninstall
    if (get-hotfix -id KB2506143) {
        Invoke-Expression "wusa /uninstall /kb:2506143 /norestart /quiet" 
        Start-Sleep 30
    } 
 
    #Download and install 7zip
    (New-Object Net.WebClient).DownloadFile($URI_7Zip, $INS_7Zip)
    Start-Process -FilePath $INS_7Zip -ArgumentList "/S" -Wait
    Start-Sleep 5

    #Download, UnZIP and install WMF 5.1
    (New-Object Net.WebClient).DownloadFile($URI_WMF, $INS_WMFzip)
    set-alias sz "$env:ProgramFiles (x86)\7-Zip\7z.exe"  
    sz x $INS_WMFzip
    Start-Sleep 5 
    Start-Process -FilePath $INS_WMFexe -ArgumentList $Arguments -Wait

}

function InstallWMF5.1Server2012 { 
    
    #WMF 5.1 Variables
    $URI_WMF = "https://download.microsoft.com/download/6/F/5/6F5FF66C-6775-42B0-86C4-47D41F2DA187/Win8.1AndW2K12R2-KB3191564-x64.msu"
    $INS_WMFexe = "C:\Temp\Win8.1AndW2K12R2-KB3191564-x64.msu" 
    $INS_DIR = "C:\Temp"
    $Arguments = "/quiet /norestart"

    If (-not(Get-Item $INS_DIR)) { 
        New-Item -Path "C:\" -Name "Temp" -ItemType directory -ErrorAction SilentlyContinue -Force 
    }

    #Download and install WMF 5.1
    (New-Object Net.WebClient).DownloadFile($URI_WMF, $INS_WMFexe)
    Start-Sleep 5 
    Start-Process -FilePath $INS_WMFexe -ArgumentList $Arguments -Wait
    
}

$OS = (Get-WmiObject -Class Win32_OperatingSystem).name -split ' '

switch ($OS[3]) {
    2008 { InstallWMF5.1Server2008 }
    2012 { InstallWMF5.1Server2012 }
}