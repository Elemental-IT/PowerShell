$Link = 'https://download.microsoft.com/download/2/7/A/27AF1BE6-DD20-4CB4-B154-EBAB8A7D4A7E/officedeploymenttool_11509-33604.exe'
$Path = "C:\temp\Office_365\officedeploymenttool.exe"
$Arguments = "/quiet /extract:C:\temp\Office_365"

New-Item -Path "C:\temp\Office_365" -ItemType Directory -Force -ErrorAction SilentlyContinue
(New-Object Net.WebClient).DownloadFile($Link, $path)

Start-Process -FilePath $Path -ArgumentList $Arguments -Wait

Start-Sleep 3

$Config_XML = "C:\temp\Office_365\Configuration.xml" 
New-Item -Path $Config_XML -ItemType File -Force

$Config_Body = @'
 <Configuration>
 
 <Add OfficeClientEdition="32" Channel="Monthly">
   <Product ID="O365ProPlusRetail">
     <Language ID="en-us" />
   </Product>
   <Product ID="VisioProRetail">
     <Language ID="en-us" />
  </Product>
 </Add>
 <Updates Enabled="TRUE" Channel="Monthly" />
 <Property Name="AUTOACTIVATE" Value="1" />

 </Configuration>
'@

Add-Content -Path $Config_XML -Value $Config_Body

$Path = "C:\temp\Office_365\setup.exe"
$Arguments = "/configure C:\temp\Office_365\Configuration.xml"

Start-Process -FilePath $Path -ArgumentList $Arguments