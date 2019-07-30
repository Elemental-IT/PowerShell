<#
Author: Theo Bird
Purpose: Check DNS scavenging results and update to FTP
Date: Tuesday, 9 May 2019
#>

function FTP_UPLOAD {
    
    # Decrypt-FTP Key
    $CSV = Import-Csv "C:\temp\Key.csv"
    foreach ($Object in $CSV) {
        $User = $Object.User | ConvertTo-SecureString
        $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($User)
        $User = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
        $Password = $Object.Password | ConvertTo-SecureString
    }

    # Uplaod File to FTP Server
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

# Get DNS server role or exit 
$dns = get-windowsfeature -name "DNS*"
if ($dns.Installed) {

    # recerate DNS folder
    Remove-Item "c:\temp\DNS" -ErrorAction SilentlyContinue -Force -Recurse 
    New-Item -Path "c:\temp" -Name "DNS" -ItemType Directory -ErrorAction SilentlyContinue -Force

    # Check if scavenging is enabled
    $sysinfo = Get-WmiObject -Class Win32_ComputerSystem
    $Domain = $sysinfo.Domain
    $File_Path = "C:\Temp\DNS\$env:COMPUTERNAME`_DNS_Scav.txt"
    $ZONE = Get-WmiObject -Namespace root\MicrosoftDNS -Class Microsoftdns_zone | Where-Object { $_.name -match "^$Domain" } 
    $ZONE | Out-File $File_Path
    $Aging = ($ZONE).aging
            
    switch ($Aging) {
        True { $Status = "Enabled" }
        Default { $Status = "Disabled" } 
    }
            
    # Rename file with Status
    Rename-Item $File_Path -NewName "C:\Temp\DNS\$env:COMPUTERNAME`_DNS_Scav_$Status.txt" -Force
    $File_Path = "C:\Temp\DNS\$env:COMPUTERNAME`_DNS_Scav_$Status.txt"
     
    # set FTP link 
    $Internal_FTP = ($FTP_link = "ftp://127.0.0.1/FTPSITE/$env:USERDOMAIN/$env:COMPUTERNAME`_DNS_Scav_$Status.txt")
    $External_FTP = ($FTP_link = "FTPSITE.com/$env:USERDOMAIN/$env:COMPUTERNAME`_DNS_Scav_$Status.txt")
    switch ($env:USERDOMAIN) {   
        LA { $Internal_FTP  }
        AUSITSOLUTIONS { $Internal_FTP  }
        OKHOSTED { $Internal_FTP }
        Default { $External_FTP  }
    }

    FTP_UPLOAD

    $File_Path = "C:\Temp\DNS\$Domain`_DNS_ZONE_Aging.csv"
    get-WmiObject -Namespace root\MicrosoftDNS -Class Microsoftdns_zone | Select-Object Name,Aging | Export-Csv $File_Path
    
    # set FTP link
    $Internal_FTP = ($FTP_link = "ftp://127.0.0.1/FTPSITE/$env:USERDOMAIN/$Domain`_DNS_ZONE_Aging.csv")
    $External_FTP = ($FTP_link = "ftp://FTPSITE.com/$env:USERDOMAIN/$Domain`_DNS_ZONE_Aging.csv")
    switch ($env:USERDOMAIN) {   
        LA  {$Internal_FTP }
        AUSITSOLUTIONS { $Internal_FTP }
        OKHOSTED { $Internal_FTP }
        Default { $External_FTP }
    }

    FTP_UPLOAD

}

else { exit }