<#
.Synopsis
Helpdesk tool for making active directory administration easier.

Author - Theo bird (Bedlem55)
#>

Function StartApp {

# Assembly and Modules
#===========================================================
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()
     
Import-Module activedirectory 
    
# VariableS
#===========================================================
$FormBackColor = "DodgerBlue"
$FormForeColor = "white"
$ButtionForeColor = "Black"
$ButtionBackColor = "Aqua"
$domain = (Get-ADDomain).Name
$ADusers = Get-ADUser -Filter * | select -ExpandProperty SamAccountName
$Groups = Get-ADGroup -Filter * | select -ExpandProperty SamAccountName
$Computers = Get-ADComputer -Filter * | select -ExpandProperty SamAccountName
    
#$SavePath = $SaveFile.FileName.ToString()
#$userAccount = $AccountsCombobox.Text.ToString()


# Base Functions
#===========================================================

Function OutError {
    $OutputTextbox.Clear()
    $Err = $Error[0]
    $OutputTextbox.AppendText("$Err")
}
    
# Account Functions
#===========================================================

Function ResetPass {
        


    if ($AccountsCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No User Account Selected") 
    }
    Else {
        Try {
            $userAccount = $AccountsCombobox.Text.ToString()
            $UserPrompt = new-object -comobject wscript.shell
            $Answer = $UserPrompt.popup("         Reset $userAccount Password?", 0, "Reset Password Prompt", 4)
            
            If ($Answer -eq 6) {
                $OutputTextbox.Clear()
                $Seasons = Get-Random @('Spring', 'Summer', 'Autumn', 'Winter') 
                $Num = Get-Random @(10..99)
                Set-ADAccountPassword -Identity $AccountsCombobox.SelectedItem -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Seasons$Num" -Force)  -ErrorAction Stop
                if ($NextLoginCheckBox.CheckState -eq $true) { Set-ADuser -Identity $AccountsCombobox.Text -ChangePasswordAtLogon $True } 
                $OutputTextbox.AppendText("$userAccount's password has been reset to $Seasons$Num")}
            Else { 
                $OutputTextbox.Clear()
                $OutputTextbox.AppendText("Reset Password action canceled")
                }
        }
        catch { OutError }
    }
}
    
Function Unlock {
        
    if ($AccountsCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No User Account Selected") 
    }
    Else {
        Try {
            $userAccount = $AccountsCombobox.Text.ToString()                
            $OutputTextbox.Clear()
            Unlock-ADAccount -Identity $AccountsCombobox.Text -ErrorAction Stop
            $OutputTextbox.AppendText("$userAccount's account is now unlocked")
        }
        catch { OutError }
    }
}
                             
                                                                                                                                                                
Function Disable {
        
    if ($AccountsCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No User Account Selected") 
    }
    
    Else {
        Try {

            $userAccount = $AccountsCombobox.Text.ToString()
            $UserPrompt = new-object -comobject wscript.shell
            $Answer = $UserPrompt.popup("        Disable $userAccount`?", 0, "Disable Account Prompt", 4)
            
            If ($Answer -eq 6) {
                $OutputTextbox.Clear()
                Disable-ADAccount -Identity $AccountsCombobox.SelectedItem -ErrorAction Stop
                $OutputTextbox.AppendText("$userAccount account is now disabled")
            }
    
            Else {
                $OutputTextbox.Clear()
                $OutputTextbox.AppendText("Account disabled action canceled") 
            }
    
        }
        catch { OutError }
    }
}
    
# Groups functions
#===========================================================
    
function AddMember {
        
    if ($GroupCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No Group Selected") 
    }
    Else {
        Try {
            $GroupOBJ = $GroupCombobox.Text.ToString()
            $OutputTextbox.Clear()
            Get-ADUser -Filter * | select  Name, SamAccountName, mail | Out-GridView  -Title "Select Member(s) to add to $GroupOBJ" -PassThru | Tee-Object -Variable "Members"
            Add-ADGroupMember -Identity $GroupCombobox.text -Members $Members.SamAccountName -ErrorAction Stop 
            $OutputTextbox.AppendText("Members added to $GroupOBJ")
        }
        Catch { OutError }
    }
}
    
function RemoveMember {
 if ($GroupCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No Group Selected") 
    }
    Else {
        Try {$OutputTextbox.Clear()
            $GroupOBJ = $GroupCombobox.Text.ToString()
            Get-ADGroup $GroupCombobox.SelectedItem | Get-ADGroupMember | Select-Object -ExpandProperty Name | Out-GridView -Title "Select Member(s) to remove from $GroupOBJ" -PassThru | Tee-Object -Variable "Members"
            Get-ADGroup $GroupCombobox.SelectedItem | remove-adgroupmember -Members $members -Confirm:$false -ErrorAction Stop
            $OutputTextbox.AppendText("Removed members from $GroupOBJ")
        }
        Catch { OutError }
    }
}
    
# AD Export functions
#===========================================================
    
function AdUserExport {
    Get-ADUser -filter * -Properties * | ForEach-Object {
    New-Object PSObject -Property @{
       
    UserName      = $_.DisplayName
    Name          = $_.name
    Email         = $_.mail
    Groups        = ($_.memberof | Get-ADGroup | Select-Object -ExpandProperty Name) -join ", "
    Enabled       = $_.Enabled
    Created       = $_.whenCreated
    LastLogonDate = $_.LastLogonDate }
       
    } | Tee-Object -Variable "AdUserExport"
    
    $SaveFile = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $SaveFile.Title = "Export User Accounts"
    $SaveFile.FileName = "$domain User Export"
    $SaveFile.Filter = "CSV Files (*.csv)|*.csv"
    $SaveFile.ShowDialog()
    
    $AdUserExport | Select-Object UserName, Name, Email, Groups, Enabled, Created, LastLogonDate | Export-Csv $SaveFile.FileName -NoTypeInformation
    $SavePath = $SaveFile.FileName.ToString()
    $OutputTextbox.Clear()
    $OutputTextbox.AppendText("Exported user acconunts to $SavePath")
}
    
function GroupsExport {
    Get-ADGroup -Filter * -Properties * | ForEach-Object {
    New-Object PSObject -Property @{
        
    GroupName = $_.Name
    Type      = $_.groupcategory
    Members   = ($_.Name | Get-ADGroupMember | Select-Object -ExpandProperty Name) -join ", " }
        
    } | Tee-Object -Variable "GroupsExport"
        
    $SaveFile = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $SaveFile.Title = "Export Groups"
    $SaveFile.FileName = "$domain Groups Export"
    $SaveFile.Filter = "CSV Files (*.csv)|*.csv"
    $SaveFile.ShowDialog()
       
    $GroupsExport | Select-Object GroupName, Type, Members | Export-Csv $SaveFile.FileName -NoTypeInformation
    $SavePath = $SaveFile.FileName.ToString()
    $OutputTextbox.Clear()
    $OutputTextbox.AppendText("Exported Groups acconunts to $SavePath")
}
    
function ComputerExport { 
    Get-ADComputer -Filter * -Properties * | Tee-Object -Variable "ComputersExport"
       
    $SaveFile = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $SaveFile.Title = "Export Computers"
    $SaveFile.FileName = "$domain Computers Export"
    $SaveFile.Filter = "CSV Files (*.csv)|*.csv"
    $SaveFile.ShowDialog()
       
    $ComputersExport | Select-Object Name, Created, Enabled, OperatingSystem, IPv4Address, LastLogonDate, logonCount | export-csv $SaveFile.FileName -NoTypeInformation
    $SavePath = $SaveFile.FileName.ToString()
    $OutputTextbox.Clear()
    $OutputTextbox.AppendText("Exported Computer acconunts to $SavePath")
}
    
# Base From
#===========================================================
    
$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '445,390'
$Form.text = "    TooEazyAD"
$Form.TopMost = $false
$Form.MaximizeBox = $false
$Form.MinimizeBox = $false
$Form.ForeColor = $FormForeColor 
$Form.BackColor = $FormBackColor
$Form.FormBorderStyle = "FixedSingle"
$Form.ShowIcon = $false
    
# User accounts GroupBox 
#===========================================================
    
$AccountsGroupBox = New-Object System.Windows.Forms.GroupBox
$AccountsGroupBox.location = New-Object System.Drawing.Point(20, 20)
$AccountsGroupBox.Size = New-Object System.Drawing.Size(409, 107)
$AccountsGroupBox.Text = "Select User Account"
$AccountsGroupBox.ForeColor = "white"
    
$AccountsCombobox = New-Object System.Windows.Forms.ComboBox 
$AccountsCombobox.location = New-Object System.Drawing.Point(7, 14)
$AccountsCombobox.Width = 394
$AccountsCombobox.FlatStyle = 'flat'
ForEach ($user in $ADusers) {[void]$AccountsCombobox.Items.Add($user)}
$AccountsCombobox.AutoCompleteSource = "CustomSource" 
$AccountsCombobox.AutoCompleteMode = "SuggestAppend"
$ADusers | % {[void]$AccountsCombobox.AutoCompleteCustomSource.Add($_)}
    
$ButtonX = 35
$ButtonY = 128
    
$ResetPasswordButton = New-Object System.Windows.Forms.Button
$ResetPasswordButton.location = New-Object System.Drawing.Point(7, 42)
$ResetPasswordButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$ResetPasswordButton.Text = "Reset Password"
$ResetPasswordButton.FlatStyle = 'flat'
$ResetPasswordButton.ForeColor = $ButtionForeColor 
$ResetPasswordButton.BackColor = $ButtionBackColor
$ResetPasswordButton.FlatAppearance.BorderSize = 0
$ResetPasswordButton.add_Click({ ResetPass })
    
$UnlockButton = New-Object System.Windows.Forms.Button
$UnlockButton.location = New-Object System.Drawing.Point(140, 42)
$UnlockButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$UnlockButton.Text = "Unlock Account"
$UnlockButton.FlatStyle = 'flat'
$UnlockButton.ForeColor = $ButtionForeColor 
$UnlockButton.BackColor = $ButtionBackColor
$UnlockButton.FlatAppearance.BorderSize = 0
$UnlockButton.add_Click({ Unlock })
    
$DisableButton = New-Object System.Windows.Forms.Button
$DisableButton.location = New-Object System.Drawing.Point(273, 42)
$DisableButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$DisableButton.Text = "Disable Account"
$DisableButton.FlatStyle = 'flat'
$DisableButton.ForeColor = $ButtionForeColor 
$DisableButton.BackColor = $ButtionBackColor
$DisableButton.FlatAppearance.BorderSize = 0
$DisableButton.add_Click({ Disable })
    
$NextLoginCheckBox = New-Object System.Windows.Forms.CheckBox
$NextLoginCheckBox.location = New-Object System.Drawing.Point(7, 80)
$NextLoginCheckBox.Text = "Reset password at next login"
$NextLoginCheckBox.Width = 170
$NextLoginCheckBox.Checked = "Checked"
    
# Group GroupBox
#===========================================================
    
$GroupGroupBox = New-Object System.Windows.Forms.GroupBox
$GroupGroupBox.location = New-Object System.Drawing.Point(20, 135)
$GroupGroupBox.Size = New-Object System.Drawing.Size(409, 88)
$GroupGroupBox.Text = "Select Group"
$GroupGroupBox.ForeColor = "white"
    
$GroupCombobox = New-Object System.Windows.Forms.ComboBox 
$GroupCombobox.location = New-Object System.Drawing.Point(7, 14)
$GroupCombobox.Width = 394
$GroupCombobox.FlatStyle = 'flat'
ForEach ($Group in $Groups) {[void]$GroupCombobox.Items.Add($Group)}
$GroupCombobox.AutoCompleteSource = "CustomSource" 
$GroupCombobox.AutoCompleteMode = "SuggestAppend"
$Groups | % {[void]$GroupCombobox.AutoCompleteCustomSource.Add($_)}
    
$GroupButtonX = 35
$GroupButtonY = 193
    
$AddButton = New-Object System.Windows.Forms.Button
$AddButton.location = New-Object System.Drawing.Point(7, 42)
$AddButton.Size = New-Object System.Drawing.Size($GroupButtonY, $GroupButtonX)
$AddButton.Text = "Add Members to Group"
$AddButton.FlatStyle = 'flat'
$AddButton.ForeColor = $ButtionForeColor 
$AddButton.BackColor = $ButtionBackColor
$AddButton.FlatAppearance.BorderSize = 0
$AddButton.add_Click({ AddMember })
    
$RemoveButton = New-Object System.Windows.Forms.Button
$RemoveButton.location = New-Object System.Drawing.Point(207, 42)
$RemoveButton.Size = New-Object System.Drawing.Size($GroupButtonY, $GroupButtonX)
$RemoveButton.Text = "Remove Members from Group"
$RemoveButton.FlatStyle = 'flat'
$RemoveButton.ForeColor = $ButtionForeColor 
$RemoveButton.BackColor = $ButtionBackColor
$RemoveButton.FlatAppearance.BorderSize = 0
$RemoveButton.add_Click({ RemoveMember })
    
# Export GroupBox
#===========================================================
    
$ExportGroupBox = New-Object System.Windows.Forms.GroupBox
$ExportGroupBox.location = New-Object System.Drawing.Point(20,235)
$ExportGroupBox.Size = New-Object System.Drawing.Size(409, 57)
$ExportGroupBox.Text = "Export to CSV"
$ExportGroupBox.ForeColor = "white"
    
$ExportButtonX = 35
$ExportButtonY = 128
    
$ExportAccountButton = New-Object System.Windows.Forms.Button
$ExportAccountButton.location = New-Object System.Drawing.Point(7, 13)
$ExportAccountButton.Size = New-Object System.Drawing.Size($ExportButtonY, $ExportButtonX)
$ExportAccountButton.Text = "User Accounts"
$ExportAccountButton.FlatStyle = 'flat'
$ExportAccountButton.ForeColor = $ButtionForeColor 
$ExportAccountButton.BackColor = $ButtionBackColor
$ExportAccountButton.FlatAppearance.BorderSize = 0
$ExportAccountButton.add_Click({ AdUserExport })
    
$ExportGroupsButton = New-Object System.Windows.Forms.Button
$ExportGroupsButton.location = New-Object System.Drawing.Point(140, 13)
$ExportGroupsButton.Size = New-Object System.Drawing.Size($ExportButtonY, $ExportButtonX)
$ExportGroupsButton.Text = "Groups and Members"
$ExportGroupsButton.FlatStyle = 'flat'
$ExportGroupsButton.ForeColor = $ButtionForeColor 
$ExportGroupsButton.BackColor = $ButtionBackColor
$ExportGroupsButton.FlatAppearance.BorderSize = 0
$ExportGroupsButton.add_Click({ GroupsExport })
    
$ExportComptuersButton = New-Object System.Windows.Forms.Button
$ExportComptuersButton.location = New-Object System.Drawing.Point(273, 13)
$ExportComptuersButton.Size = New-Object System.Drawing.Size($ExportButtonY, $ExportButtonX)
$ExportComptuersButton.Text = "Comptuers Accounts"
$ExportComptuersButton.FlatStyle = 'flat'
$ExportComptuersButton.ForeColor = $ButtionForeColor 
$ExportComptuersButton.BackColor = $ButtionBackColor
$ExportComptuersButton.FlatAppearance.BorderSize = 0
$ExportComptuersButton.add_Click({ ComputerExport })
    
# Output GroupBox
#===========================================================
    
$OutputGroupBox = New-Object System.Windows.Forms.GroupBox
$OutputGroupBox.location = New-Object System.Drawing.Point(20,300)
$OutputGroupBox.Size = New-Object System.Drawing.Size(409, 70)
$OutputGroupBox.Text = "Output"
$OutputGroupBox.ForeColor = "white"
    
$OutputTextbox = New-Object System.Windows.Forms.TextBox
$OutputTextbox.location = New-Object System.Drawing.Point(7, 14)
$OutputTextBox.Size = New-Object System.Drawing.Size(395, 48)
$OutputTextbox.Enabled = $false
$OutputTextbox.Multiline = $true
$OutputTextBox.BackColor = "White"
$OutputTextbox.BorderStyle = 'none'

    
# Controls
#===========================================================
    
$AccountsGroupBox.Controls.AddRange(@( $AccountsCombobox, $ResetPasswordButton ,$UnlockButton, $DisableButton, $NextLoginCheckBox ))
$GroupGroupBox.Controls.AddRange(@( $GroupCombobox, $AddButton, $RemoveButton ))
$ExportGroupBox.Controls.AddRange(@( $ExportAccountButton, $ExportGroupsButton, $ExportComptuersButton, $ExportComptuersButton ))
$OutputGroupBox.Controls.AddRange(@( $OutputTextbox ))
$Form.controls.AddRange(@( $AccountsGroupBox ,$GroupGroupBox ,$ComputersGroupBox  ,$ExportGroupBox, $OutputGroupBox ))
[void]$Form.ShowDialog()
} StartApp
