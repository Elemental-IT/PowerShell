$Exportlist = @()

Set-Location $env:SystemDrive\

Get-ChildItem -Recurse -ErrorAction SilentlyContinue | Where-Object { $_.PSIsContainer } | ForEach-Object {

$obj = New-Object PSObject -Property @{

   Fullname = $_.FullName 
   Size = "{0} MB" -f [math]::round(((Get-ChildItem  $_.FullName -Recurse -ErrorAction SilentlyContinue |  Measure-Object -Property Length -Sum -ErrorAction SilentlyContinue).Sum / 1MB))
   CreationTime = $_.CreationTime
   LastAccessTime = $_.LastAccessTime
      
   } 
   $Exportlist += $obj
} 

$Exportlist | Select-Object Fullname,Size,CreationTime,LastAccessTime | Export-Csv "$env:SystemDrive\$env:COMPUTERNAME`_foldrsize.csv" -NoTypeInformation
