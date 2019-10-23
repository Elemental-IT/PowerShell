<#
.Synopsis
Installs cleanmgr.exe without installing desktop experience server role and adds desktop shortcut to cleanmgr.exe

.Author Theo bird 
#>

if ((get-windowsfeature -name "desktop-experience").installed -ne $true) {
    function set-shortcut {
        param ( [string]$SourceLnk, [string]$DestinationPath )
        $WshShell = New-Object -comObject WScript.Shell
        $Shortcut = $WshShell.CreateShortcut($SourceLnk)
        $Shortcut.TargetPath = $DestinationPath
        $Shortcut.Save()
    }
    
    if ((test-path "$env:SystemDrive\Users\Public\Desktop") -ne $true) {
        New-Item -path "$env:SystemDrive\Users\Public\Desktop" -ItemType Directory -Force
    }
    
    Try { 
        Copy-Item -path "C:\Windows\winsxs\amd64_microsoft-windows-cleanmgr_31bf3856ad364e35_6.0.6001.18000_none_c962d1e515e94269\cleanmgr.exe" -Destination "C:\Windows\System32\" -ErrorAction Stop
        Copy-Item -path "C:\Windows\winsxs\amd64_microsoft-windows-cleanmgr.resources_31bf3856ad364e35_6.0.6001.18000_en-us_b9f50b71510436f2\cleanmgr.exe.mui" -Destination  "C:\Windows\System32\en-US\" -ErrorAction Stop
        set-shortcut "$env:SystemDrive\Users\Public\Desktop\cleanmgr.exe.lnk" "C:\Windows\System32\cleanmgr.exe"
    }
    Catch {
        Copy-Item -path "C:\Windows\winsxs\amd64_microsoft-windows-cleanmgr_31bf3856ad364e35_6.1.7600.16385_none_c9392808773cd7da\cleanmgr.exe" -Destination "C:\Windows\System32\"
        Copy-Item -path "C:\Windows\winsxs\amd64_microsoft-windows-cleanmgr.resources_31bf3856ad364e35_6.1.7600.16385_en-us_b9cb6194b257cc63\cleanmgr.exe.mui" -Destination "C:\Windows\System32\en-US\"
        set-shortcut "$env:SystemDrive\Users\Public\Desktop\cleanmgr.exe.lnk" "C:\Windows\System32\cleanmgr.exe"
    }
  } else {
    Write-Warning "desktop-experience enabled"
    Start-Sleep 10
    Stop-Process $PID
}    


