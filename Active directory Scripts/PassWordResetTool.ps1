<#
.Synopsis
Simple GUI for resetting passwords quickly

.Author Theo bird
#>

Function StartApp {
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()
 
Import-Module activedirectory 
$ADusers = get-aduser -Filter * | select -ExpandProperty SamAccountName

# Base From
#===========================================================

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '473,150'
$Form.text = "    Password reset tool"
$Form.TopMost = $false
$Form.MaximizeBox = $false
$Form.MinimizeBox = $false
$Form.ForeColor = "white"
$Form.BackColor = "CadetBlue"
$Form.FormBorderStyle = "FixedSingle"
$Form.ShowIcon = $false

# Select GroupBox 
#===========================================================

$SelectGroupBox = New-Object System.Windows.Forms.GroupBox
$SelectGroupBox.location = New-Object System.Drawing.Point(12, 5)
$SelectGroupBox.Text = "Select User Account"
$SelectGroupBox.ForeColor = "white"
$SelectGroupBox.width = 450
$SelectGroupBox.height = 70

$SelectCombobox = New-Object System.Windows.Forms.ComboBox 
$SelectCombobox.location = New-Object System.Drawing.Point(7, 14)
$SelectCombobox.Width = 305
$SelectCombobox.FlatStyle = 'flat'
$SelectCombobox.AutoCompleteSource = "CustomSource" 
$SelectCombobox.AutoCompleteMode = "SuggestAppend"
$ADusers | % {[void]$SelectCombobox.AutoCompleteCustomSource.Add($_)}
ForEach ($user in $ADusers) {[void]$SelectCombobox.Items.Add($user)}

Function ResetPass {
Try {
    $OutputTextbox.Clear()
    $Seasons = Get-Random @('Spring','Summer','Autumn','Winter') 
    $Num = Get-Random @(10..99)
    if ($SelectCheckBox.CheckState -eq "checked") {Set-ADuser -Identity $SelectCombobox.Text -ChangePasswordAtLogon $True }  
    Set-ADAccountPassword -Identity $SelectCombobox.Text -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Seasons$Num" -Force)  -ErrorAction Stop 
    $OutputTextbox.AppendText("New password is $Seasons$Num")
} catch {
    $OutputTextbox.Clear()
    $OutputTextbox.AppendText("$Error")}
}

$SelectButton = New-Object System.Windows.Forms.Button
$SelectButton.location = New-Object System.Drawing.Point(320, 13)
$SelectButton.Text = "Reset Password"
$SelectButton.width = 122
$SelectButton.height = 48
$SelectButton.FlatStyle = 'flat'
$SelectButton.ForeColor = "black"
$SelectButton.BackColor = "Azure"
$SelectButton.FlatAppearance.BorderSize = '0'
$SelectButton.add_Click({ ResetPass })

$SelectCheckBox = New-Object System.Windows.Forms.CheckBox
$SelectCheckBox.location = New-Object System.Drawing.Point(142, 40)
$SelectCheckBox.Text = "Reset password at next login"
$SelectCheckBox.Width = 170
$SelectCheckBox.CheckAlign = "MiddleRight"
$SelectCheckBox.Checked = "Checked"


# Output GroupBox
#===========================================================

$OutputGroupBox = New-Object System.Windows.Forms.GroupBox
$OutputGroupBox.location = New-Object System.Drawing.Point(12, 80)
$OutputGroupBox.Text = "Password Output"
$OutputGroupBox.ForeColor = "white"
$OutputGroupBox.width = 450
$OutputGroupBox.height = 60

$OutputTextbox = New-Object System.Windows.Forms.TextBox
$OutputTextbox.location = New-Object System.Drawing.Point(7, 14)
$OutputTextbox.Width = 435
$OutputTextbox.height = 37
$OutputTextbox.Enabled = $false
$OutputTextbox.Multiline = $true
$OutputTextBox.BackColor = "White"
$OutputTextbox.BorderStyle = 'none'

$SelectGroupBox.Controls.AddRange(@($SelectCombobox,$SelectButton, $SelectCheckBox))
$OutputGroupBox.Controls.AddRange(@( $OutputTextbox  ))
$Form.controls.AddRange(@($SelectGroupBox,$OutputGroupBox ))
[void]$Form.ShowDialog()
} StartApp
