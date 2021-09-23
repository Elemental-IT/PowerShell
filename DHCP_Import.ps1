$CSV = ""
$Scopes = Import-Csv $CSV
$Server = ""

Foreach ($Scope in $Scopes) {

Add-DhcpServerv4Scope -ComputerName $Server -Name $Scope.name -Description $Scope.Description -SubnetMask $Scope.SubnetMask -StartRange $Scope.StartRange -EndRange $Scope.EndRange

}