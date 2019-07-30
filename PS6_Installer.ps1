$Link = "https://github.com/PowerShell/PowerShell/releases/download/v6.2.0/PowerShell-6.2.0-win-x64.msi"
$Path = "C:\temp\PS6\PS6.msi"
$Arguments = "/quiet /norestart"

New-Item -Path "C:\temp\PS6" -ItemType Directory -Force -ErrorAction SilentlyContinue
(New-Object Net.WebClient).DownloadFile($Link, $path)

Start-Process -FilePath $Path -ArgumentList $Arguments -Wait