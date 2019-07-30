<#
Author: Theo Bird
Purpose: Connect each Chkdsk result and update to FTP
Date: Tuesday, 7 May 2019
#>

# Kill script on VM (CHKDSK Is not required for VM)
if ((Get-WmiObject -Class:win32_ComputerSystem).model -eq "VMware Virtual Machine" ) { exit }
# need to added kill switch for SSD and Huper-V VMs

function FTP_UPLOAD {
    
    # Decrypt-FTP Key    
    $CSV = Import-Csv "C:\temp\Key.csv"
    foreach ($Object in $CSV) {
        $User = $Object.User | ConvertTo-SecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($User)
        $User = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $Password = $Object.Password | ConvertTo-SecureString
    }

    # Uplaod CSV to FTP Server
    $ftp = [System.Net.FtpWebRequest]::Create($FTP_link)
    $ftp = [System.Net.FtpWebRequest]$ftp
    $ftp.Method = [System.Net.WebRequestMethods+Ftp]::UploadFile
    $ftp.Credentials = new-object System.Net.NetworkCredential($User, $Password)
    $ftp.UseBinary = $true
    $ftp.UsePassive = $true
    $content = [System.IO.File]::ReadAllBytes($File_Path)
    $ftp.ContentLength = $content.Length
    $rs = $ftp.GetRequestStream()
    $rs.Write($content, 0, $content.Length)
    $rs.Close()
    $rs.Dispose()
    
}

# Check for each fixed drive - recerate CHKDSK folder
$Drive_Letter = (Get-CimInstance -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -match '3' }).DeviceID
Remove-Item "c:\temp\CHKDSK" -ErrorAction SilentlyContinue -Force -Recurse 
New-Item -Path "c:\temp" -Name "CHKDSK" -ItemType Directory -ErrorAction SilentlyContinue -Force

# Run Chkdsk for each fixed drive and outpot file log to c:\temp\CHKDSK
foreach ($Drive in $Drive_Letter) {
    $Out_Drive_Letter = $Drive -replace ':', ''
    $File_Path = "c:\temp\CHKDSK\$env:COMPUTERNAME`_$Out_Drive_Letter`_Scandisk.txt"
    Chkdsk $Drive > $File_Path 
    $Out_file = (Get-Content $File_Path) | Sort-Object -Descending
    Clear-Content $File_Path -Force 
    Add-Content -Path $File_Path -Value $Out_file -Force 
    
    # Set Status 
    $Status = Get-Content $File_Path -First 1 ; $Status = $Status -replace ' ', ''
    Switch ($Status) {
        Windowshasscannedthefilesystemandfoundnoproblems. { $Status = "OK" }
        Windowshascheckedthefilesystemandfoundnoproblems. { $Status = "OK" }
        ThetypeofthefilesystemisREFS. { $Status = "OK" }
        Default { $Status = "Fail" }
    }
    
    # Rename file with Status
    Rename-Item $File_Path -NewName "c:\temp\CHKDSK\$env:COMPUTERNAME`_$Out_Drive_Letter`_Scandisk_$Status.txt" -Force
    $File_Path = "c:\temp\CHKDSK\$env:COMPUTERNAME`_$Out_Drive_Letter`_Scandisk_$Status.txt"
    
    # Set FTP link
    $Internal_FTP = ($FTP_link = "ftp:/127.0.0.1/FTPSITE/$env:USERDOMAIN/$env:COMPUTERNAME`_$Out_Drive_Letter`_Scandisk_$Status.txt")
    $External_FTP = ($FTP_link = "ftp://FTPSITE.com/$env:USERDOMAIN/$env:COMPUTERNAME`_$Out_Drive_Letter`_Scandisk_$Status.txt")
    switch ($env:USERDOMAIN) {           
        LA  { $Internal_FTP }
        AUSITSOLUTIONS { $Internal_FTP }
        OKHOSTED { $Internal_FTP }
        Default { $External_FTP }
  
    }

    FTP_UPLOAD
   
} 