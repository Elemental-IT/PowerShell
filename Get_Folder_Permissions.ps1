<#
.SYNOPSIS
Folder info - gets selected folder and subfolders ACL permissions 

.NOTES
Authored By Theo bird (Bedlem55)
#>

## Assemblys 
#===========================================================
Add-Type -AssemblyName system.windows.forms

# help message
$About = @'

Folder info - gets selected folder and subfolders ACL permissions and exports them to a CSV

Authored By Theo bird

'@


## functions 
#===========================================================
function FolderImport {
    $GetFolderTB.Clear()
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.RootFolder = "MyComputer"
    $FolderBrowser.ShowDialog()
    $GetFolderTB.AppendText($FolderBrowser.SelectedPath.tostring())
}


function Run {

    if ( $GetFolderTB.Text -eq '' ) {

    [System.Windows.Forms.MessageBox]::Show("No folder selected", "Warning:",0,48) 

    }
    Else {

        $SaveFile = New-Object -TypeName System.Windows.Forms.SaveFileDialog
        $SaveFile.Title = "Export ACL Permissions"
        $SaveFile.FileName = "Folder Permissions Export"
        $SaveFile.Filter = "CSV Files (*.csv)|*.csv"
        $SaveFile.ShowDialog()

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
        } | select-object -Property Folder,User,Access,Control,Inheritance | export-csv $SaveFile.FileName.tostring() -NoTypeInformation -force
    }
}


## From
#===========================================================

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '460,90'
$Form.text = " Get ACL Folder Info"
$Form.TopMost = $false
$Form.ShowIcon = $false
$Form.FormBorderStyle = 1
$Form.MaximizeBox = $false

$Menu = New-Object System.Windows.Forms.MenuStrip

$MenuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuFile.Text = "&File"
[void]$Menu.Items.Add($MenuFile)

$MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuExit.Text = "&Exit"
$menuExit.Add_Click({ $Form.close() })
[void]$MenuFile.DropDownItems.Add($MenuExit)

$MenuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuHelp.Text = "&Help"
[void]$Menu.Items.Add($MenuHelp)

$MenuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAbout.Text = "&About"
$MenuAbout.Add_Click({ [System.Windows.Forms.MessageBox]::Show("$About", "    About") })
[void]$MenuHelp.DropDownItems.Add($MenuAbout)

$RunButtion =  New-Object System.Windows.Forms.Button
$RunButtion.location = "400, 35"
$RunButtion.Size = "50,40"
$RunButtion.Text = "Export"
$RunButtion.add_Click({ Run })


## Get Folder GroupBox
#===========================================================

$GetFolderGB = New-Object System.Windows.Forms.GroupBox
$GetFolderGB.location = "10, 30"
$GetFolderGB.Size =  "380,45"
$GetFolderGB.Text = "Select Folder"

$GetFolderTB = New-Object System.Windows.Forms.TextBox
$GetFolderTB.location = "10, 15"
$GetFolderTB.Size =  "330,30"
$GetFolderTB.Enabled = $false

$GetFolderButtion =  New-Object System.Windows.Forms.Button
$GetFolderButtion.location = "345, 14"
$GetFolderButtion.Size = "25, 22"
$GetFolderButtion.Text = "..."
$GetFolderButtion.add_Click({ FolderImport })

$GetFolderGB.controls.AddRange(@($GetFolderTB,$GetFolderButtion))

# Controls
#===========================================================

$Form.controls.AddRange(@($Menu,$GetFolderGB,$RunButtion))
[void]$Form.ShowDialog()
