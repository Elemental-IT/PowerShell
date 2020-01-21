<#
.SYNOPSIS
Folder info - gets selected folder and subfolders ACL permissions 

.NOTES
Authored By Theo bird (Bedlem55)
#>

## Assemblys 
#===========================================================
Add-Type -AssemblyName system.windows.forms

## functions 
#===========================================================
function FolderImport {
    $GetFolderTB.Clear()
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.RootFolder = "MyComputer"
    $FolderBrowser.ShowDialog()
    $GetFolderTB.AppendText($FolderBrowser.SelectedPath.tostring())
}

function CSVExport {
    $ExportCSVTB.Clear()
    $SaveFile = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $SaveFile.Title = "Export ACL Permissions"
    $SaveFile.FileName = "ACL Export"
    $SaveFile.Filter = "CSV Files (*.csv)|*.csv"
    $SaveFile.ShowDialog()
    $ExportCSVTB.AppendText($SaveFile.FileName.tostring())
}

function Run {
Get-ChildItem -Path $GetFolderTB.Text.ToString() -Recurse | Where-Object{$_.psiscontainer}|
Get-Acl | foreach {
    $path = $_.Path
    $_.Access | % {
        New-Object PSObject -Property @{
            Folder = $path.Replace("Microsoft.PowerShell.Core\FileSystem::","")
            Access = $_.FileSystemRights
            Control = $_.AccessControlType
            User = $_.IdentityReference
            Inheritance = $_.IsInherited
            }
        }
    } | select-object -Property Folder,User,Access,Control,Inheritance | export-csv $ExportCSVTB.Text.ToString() -NoTypeInformation -force
          
}


## From
#===========================================================

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '460,110'
$Form.text = " Get ACL Folder Info"
$Form.TopMost = $false
$Form.ShowIcon = $false
$Form.FormBorderStyle = 1

$RunButtion =  New-Object System.Windows.Forms.Button
$RunButtion.location = "400, 10"
$RunButtion.Size = "50,85"
$RunButtion.Text = "Run"
$RunButtion.add_Click({ Run })


## Get Folder GroupBox
#===========================================================

$GetFolderGB = New-Object System.Windows.Forms.GroupBox
$GetFolderGB.location = "10, 5"
$GetFolderGB.Size =  "380,45"
$GetFolderGB.Text = "Select Folder"

$GetFolderTB = New-Object System.Windows.Forms.TextBox
$GetFolderTB.location = "10, 15"
$GetFolderTB.Size =  "330,30"

$GetFolderButtion =  New-Object System.Windows.Forms.Button
$GetFolderButtion.location = "345, 14"
$GetFolderButtion.Size = "25, 22"
$GetFolderButtion.Text = "..."
$GetFolderButtion.add_Click({ FolderImport })

$GetFolderGB.controls.AddRange(@($GetFolderTB,$GetFolderButtion))

## Export CSV GroupBox
#===========================================================

$ExportCSVGB = New-Object System.Windows.Forms.GroupBox
$ExportCSVGB.location = "10, 50"
$ExportCSVGB.size = "380,45"
$ExportCSVGB.Text = "Select Export Location"

$ExportCSVTB = New-Object System.Windows.Forms.TextBox
$ExportCSVTB.location = "10, 15"
$ExportCSVTB.Size =  "330,30"

$ExportCSVButtion =  New-Object System.Windows.Forms.Button
$ExportCSVButtion.location = "345, 14"
$ExportCSVButtion.Size = "25, 22"
$ExportCSVButtion.Text = "..."
$ExportCSVButtion.add_Click({ CSVExport })

$ExportCSVGB.controls.AddRange(@($ExportCSVTB,$ExportCSVButtion))

# Controls
#===========================================================

$Form.controls.AddRange(@($GetFolderGB,$ExportCSVGB,$RunButtion))
[void]$Form.ShowDialog()