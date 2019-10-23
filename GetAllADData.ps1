<#
  .Author
   Theo Bird

  .SYNOPSIS
    Collect data for AD cleanup 

  .DESCRIPTION
    Exports AD user, computers and groups accounts to CSV
#>

# Export all AD users and Memberships
$domain = (Get-ADDomain).Name
$FolderPath = "$env:SystemDrive\temp\$domain"
New-Item $FolderPath -ItemType Directory -ErrorAction SilentlyContinue -Force
Import-Module activedirectory

function AdUserExport {
    Get-ADUser -filter * -Properties * | ForEach-Object {
        New-Object PSObject -Property @{
            UserName      = $_.DisplayName
            Name          = $_.name
            Email         = $_.mail
            Groups        = ($_.memberof | Get-ADGroup | Select-Object -ExpandProperty Name) -join ", "
            Enabled       = $_.Enabled
            Created       = $_.whenCreated
            LastLogonDate = $_.LastLogonDate

        }
    } | Select-Object UserName, Name, Email, Groups, Enabled, Created, LastLogonDate | Export-Csv $FolderPath\ADUsersList.csv -NoTypeInformation
}

# Get all groups and members for each Group
function GroupsExport {
    Get-ADGroup -Filter * -Properties * | ForEach-Object {
        New-Object PSObject -Property @{
            GroupName = $_.Name
            Type      = $_.groupcategory
            Members   = ($_.Name | Get-ADGroupMember | Select-Object -ExpandProperty Name) -join ", "
    
        }
    } | Select-Object GroupName, Type, Members | Export-Csv $FolderPath\ADGroupsList.csv -NoTypeInformation
}

# Computers objects list
function ComputerExport { 
   $Computers = Get-ADComputer -Filter * -Properties *
   $Computers | Select-Object Name, Created, Enabled, OperatingSystem, IPv4Address, LastLogonDate, logonCount | export-csv $FolderPath\Computers.csv
}

# Run functions
AdUserExport
GroupsExport
ComputerExport