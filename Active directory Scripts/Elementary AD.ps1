<#
  .SYNOPSIS
   Active Directory Administration Tool - A tool that works with any Active Directory
   
  .DESCRIPTION
   This tool is to make getting Active Directory Administration simple and easy
       
  .NOTES
   Authored By Theo bird (Bedlem55)

#>


Function StartApp {


# Assembly and Modules
#===========================================================
Add-Type -AssemblyName system.windows.forms
[System.Windows.Forms.Application]::EnableVisualStyles()
     
Import-Module activedirectory -ErrorAction Stop
    
# VariableS
#===========================================================
$FormBackColor = "DodgerBlue"
$FormForeColor = "white"
$ButtionForeColor = "Black"
$ButtionBackColor = "Aqua"
$GroupTextColor = "White"
$Style = 0
$domain = (Get-ADDomain).Name
$ADusers = (Get-ADUser -Filter *).SamAccountName
$Groups = (Get-ADGroup -Filter *).SamAccountName
$Computers = (Get-ADComputer -Filter *).SamAccountName
$OUs = (Get-ADOrganizationalUnit -Filter *).name    

$About = @'
  Elementary AD - Active Directory Administration Tool 
  
  * Designed to wors with any on premises Active Directory 
  * This tool is to make Active Directory administration quick and easy

  Authored By Theo bird (Bedlem55)
    
'@

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
        } catch { OutError }
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
        } catch { OutError }
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
    
        } catch { OutError }
    }
}
    

Function Accountinfo {
        
    if ($AccountsCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No User Account Selected") 
    }
    
    Else {
        Try {
            $OutputTextbox.Clear()
            $userAccount = $AccountsCombobox.Text.ToString()
            $OutputTextbox.text = get-aduser $userAccount -Properties * | Format-List |  Out-String -Width 2147483647
                
        } catch { OutError }
    }
}



Function Allusersinfo {
        
            Try {
            $OutputTextbox.Clear()
            $OutputTextbox.text = Get-ADUser -filter * -Properties * | ForEach-Object {
                New-Object PSObject -Property @{
       
                    UserName      = $_.DisplayName
                    Name          = $_.name
                    Email         = $_.mail
                    Groups        = ($_.memberof | Get-ADGroup | Select-Object -ExpandProperty Name) -join ", "
                    Enabled       = $_.Enabled
                    Created       = $_.whenCreated
                    LastLogonDate = $_.LastLogonDate }
       
        } | Select-Object UserName, Name, Email, Groups, Enabled, Created, LastLogonDate | Out-String -Width 2147483647
                
    } catch { OutError }
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
            $OutputTextbox.AppendText("Members added to $GroupOBJ Group")
        } Catch { OutError }
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
            $OutputTextbox.AppendText("Removed members from $GroupOBJ Group")
        } Catch { OutError }
    }
}

function GroupInfo {
    if ($GroupCombobox.SelectedItem -eq $null) {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("No Group Selected") 
    }
    Else {
        Try {$OutputTextbox.Clear()
            $OutputTextbox.text = Get-ADGroup $GroupCombobox.SelectedItem -Properties * | FL| Out-String -Width 2147483647
        } Catch { OutError }
    }
}


function ShowallGroups {
        
    Try {$OutputTextbox.text =  Get-ADGroup -Filter * -Properties * | ForEach-Object {
            New-Object PSObject -Property @{
        
                GroupName = $_.Name
                Type      = $_.groupcategory
                Members   = ($_.Name | Get-ADGroupMember | Select-Object -ExpandProperty Name) -join ", " }
    
         } |  Select-Object GroupName, Type, Members | Out-String -Width 2147483647
    } Catch { OutError }
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
    $Answer = $SaveFile.ShowDialog(); $Answer
    if ( $Answer -eq "OK") {
    
        $AdUserExport | Select-Object UserName, Name, Email, Groups, Enabled, Created, LastLogonDate | Export-Csv $SaveFile.FileName -NoTypeInformation
        $SavePath = $SaveFile.FileName.ToString()
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("Exported user acconunts to $SavePath")
    } Else  {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText(" Exported canceled")
    }
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
    $Answer = $SaveFile.ShowDialog(); $Answer
    if ( $Answer -eq "OK") {
       
        $GroupsExport | Select-Object GroupName, Type, Members | Export-Csv $SaveFile.FileName -NoTypeInformation
        $SavePath = $SaveFile.FileName.ToString()
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("Exported Groups acconunts to $SavePath")
    } Else {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("Exported canceled")
    }
}
    
function ComputerExport { 
    Get-ADComputer -Filter * -Properties * | Tee-Object -Variable "ComputersExport"
       
    $SaveFile = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $SaveFile.Title = "Export Computers"
    $SaveFile.FileName = "$domain Computers Export"
    $SaveFile.Filter = "CSV Files (*.csv)|*.csv"
    $Answer = $SaveFile.ShowDialog(); $Answer
    if ( $Answer -eq "OK") {
       
        $ComputersExport | Select-Object Name, Created, Enabled, OperatingSystem, OperatingSystemVersion, IPv4Address, LastLogonDate, logonCount | export-csv $SaveFile.FileName -NoTypeInformation
        $SavePath = $SaveFile.FileName.ToString()
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("Exported Computer acconunts to $SavePath")
    } Else {
        $OutputTextbox.Clear()
        $OutputTextbox.AppendText("Exported canceled")
    }
}
    
# Base From & menus
#===========================================================
    
$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '1080,590'
$Form.text = "     Elementary AD - Active Directory Administration Tool"
$Form.TopMost = $false
$Form.ForeColor = $FormForeColor 
$Form.BackColor = $FormBackColor
$Form.ShowIcon = $false


$Menu  = New-Object System.Windows.Forms.MenuStrip

$MenuFile  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuFile.Text = "&File"
    [void]$Menu.Items.Add($MenuFile)

        $MenuExit  = New-Object System.Windows.Forms.ToolStripMenuItem
        $MenuExit.Text = "&Exit"
        $menuExit.Add_Click({ $Form.close() })
[void]$MenuFile.DropDownItems.Add($MenuExit)


$MenuExport  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuExport.Text = "&Export"
    [void]$Menu.Items.Add($MenuExport)

        $ExportUsers = New-Object System.Windows.Forms.ToolStripMenuItem
        $ExportUsers.Text = "&Export Users to CSV"
        $ExportUsers.Add_Click({ AdUserExport })

        $ExportGroups = New-Object System.Windows.Forms.ToolStripMenuItem
        $ExportGroupS.Text = "&Export Group to CSV"
        $ExportGroupS.Add_Click({ GroupsExport })

        $ExportComputers = New-Object System.Windows.Forms.ToolStripMenuItem
        $ExportComputers.Text = "&Export Computers to CSV"
        $ExportComputers.Add_Click({ ComputerExport })

[void]$MenuExport.DropDownItems.AddRange(@($ExportUsers,$ExportGroups,$ExportComputers ))


$MenuHelp  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuHelp.Text = "&Help"
    [void]$Menu.Items.Add($MenuHelp)

        $MenuAbout  = New-Object System.Windows.Forms.ToolStripMenuItem
        $MenuAbout.Text = "&About"
        $MenuAbout.Add_Click({ [System.Windows.Forms.MessageBox]::Show("$About","    About") })
[void]$MenuHelp.DropDownItems.Add($MenuAbout)
    

# User accounts GroupBox 
#===========================================================
    
$AccountsGroupBox = New-Object System.Windows.Forms.GroupBox
$AccountsGroupBox.location = New-Object System.Drawing.Point(20,40)
$AccountsGroupBox.Size = New-Object System.Drawing.Size(409, 167)
$AccountsGroupBox.Text = "User Account Actions"
$AccountsGroupBox.ForeColor = $GroupTextColor

$AccountsLabel = New-Object System.Windows.Forms.Label
$AccountsLabel.location = New-Object System.Drawing.Point(7, 17)
$AccountsLabel.Text = "Select User Account:"
$AccountsLabel.Width = 110
    
$AccountsCombobox = New-Object System.Windows.Forms.ComboBox 
$AccountsCombobox.location = New-Object System.Drawing.Point(120, 14)
$AccountsCombobox.DropDownStyle = "DropDown"
$AccountsCombobox.Width = 280
$AccountsCombobox.FlatStyle = $Style
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
$ResetPasswordButton.FlatStyle = $Style 
$ResetPasswordButton.ForeColor = $ButtionForeColor 
$ResetPasswordButton.BackColor = $ButtionBackColor
$ResetPasswordButton.FlatAppearance.BorderSize = 0
$ResetPasswordButton.add_Click({ ResetPass })
    
$UnlockButton = New-Object System.Windows.Forms.Button
$UnlockButton.location = New-Object System.Drawing.Point(140, 42)
$UnlockButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$UnlockButton.Text = "Unlock Account"
$UnlockButton.FlatStyle = $Style 
$UnlockButton.ForeColor = $ButtionForeColor 
$UnlockButton.BackColor = $ButtionBackColor
$UnlockButton.FlatAppearance.BorderSize = 0
$UnlockButton.add_Click({ Unlock })
    
$DisableButton = New-Object System.Windows.Forms.Button
$DisableButton.location = New-Object System.Drawing.Point(273, 42)
$DisableButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$DisableButton.Text = "Disable/Enable Account"
$DisableButton.FlatStyle = $Style 
$DisableButton.ForeColor = $ButtionForeColor 
$DisableButton.BackColor = $ButtionBackColor
$DisableButton.FlatAppearance.BorderSize = 0
$DisableButton.add_Click({ Disable })
 
<# Move this to reset password output
$NextLoginCheckBox = New-Object System.Windows.Forms.CheckBox
$NextLoginCheckBox.location = New-Object System.Drawing.Point(7, 78)
$NextLoginCheckBox.Text = "Reset password at next login"
$NextLoginCheckBox.Width = 170
$NextLoginCheckBox.Checked = "Checked"
#>
    
$NewButton = New-Object System.Windows.Forms.Button
$NewButton.location = New-Object System.Drawing.Point(7, 82)
$NewButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$NewButton.Text = "New User Account"
$NewButton.FlatStyle = $Style 
$NewButton.ForeColor = $ButtionForeColor 
$NewButton.BackColor = $ButtionBackColor
$NewButton.FlatAppearance.BorderSize = 0
$NewButton.add_Click({  })

$PasNvrExpBut = New-Object System.Windows.Forms.Button
$PasNvrExpBut.location = New-Object System.Drawing.Point(140, 82)
$PasNvrExpBut.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$PasNvrExpBut.Text = "Set Password to never Expires"
$PasNvrExpBut.FlatStyle = $Style 
$PasNvrExpBut.ForeColor = $ButtionForeColor 
$PasNvrExpBut.BackColor = $ButtionBackColor
$PasNvrExpBut.FlatAppearance.BorderSize = 0
$PasNvrExpBut.add_Click({ SetPasneverExpires })

$PasNoChangeBut = New-Object System.Windows.Forms.Button
$PasNoChangeBut.location = New-Object System.Drawing.Point(273, 82)
$PasNoChangeBut.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$PasNoChangeBut.Text = "Set user cannot change Password"
$PasNoChangeBut.FlatStyle = $Style 
$PasNoChangeBut.ForeColor = $ButtionForeColor 
$PasNoChangeBut.BackColor = $ButtionBackColor
$PasNoChangeBut.FlatAppearance.BorderSize = 0
$PasNoChangeBut.add_Click({  })

$MoveToOUBut = New-Object System.Windows.Forms.Button
$MoveToOUBut.location = New-Object System.Drawing.Point(7, 122)
$MoveToOUBut.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$MoveToOUBut.Text = "Move to OU"
$MoveToOUBut.FlatStyle = $Style 
$MoveToOUBut.ForeColor = $ButtionForeColor 
$MoveToOUBut.BackColor = $ButtionBackColor
$MoveToOUBut.FlatAppearance.BorderSize = 0
$MoveToOUBut.add_Click({  })

$AllUserinfo = New-Object System.Windows.Forms.Button
$AllUserinfo.location = New-Object System.Drawing.Point(140, 122)
$AllUserinfo.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$AllUserinfo.Text = "Get all user info"
$AllUserinfo.FlatStyle = $Style 
$AllUserinfo.ForeColor = $ButtionForeColor 
$AllUserinfo.BackColor = $ButtionBackColor
$AllUserinfo.FlatAppearance.BorderSize = 0
$AllUserinfo.add_Click({ Allusersinfo })
    
$AccountinfoButton = New-Object System.Windows.Forms.Button
$AccountinfoButton.location = New-Object System.Drawing.Point(273, 122)
$AccountinfoButton.Size = New-Object System.Drawing.Size($ButtonY, $ButtonX)
$AccountinfoButton.Text = "Account info"
$AccountinfoButton.FlatStyle = $Style 
$AccountinfoButton.ForeColor = $ButtionForeColor 
$AccountinfoButton.BackColor = $ButtionBackColor
$AccountinfoButton.FlatAppearance.BorderSize = 0
$AccountinfoButton.add_Click({ Accountinfo })


<# Move this to $NewButton output
$CopyCheckBox = New-Object System.Windows.Forms.CheckBox
$CopyCheckBox.location = New-Object System.Drawing.Point(7, 138)
$CopyCheckBox.Text = "Copy selected user Account"
$CopyCheckBox.Width = 400
$CopyCheckBox.Checked = "Checked"
#>
    
# Group GroupBox
#===========================================================
    
$GroupGroupBox = New-Object System.Windows.Forms.GroupBox
$GroupGroupBox.location = New-Object System.Drawing.Point(20, 220)
$GroupGroupBox.Size = New-Object System.Drawing.Size(409, 127)
$GroupGroupBox.Text = "Select Group"
$GroupGroupBox.ForeColor = $GroupTextColor
    
$GroupCombobox = New-Object System.Windows.Forms.ComboBox 
$GroupCombobox.location = New-Object System.Drawing.Point(7, 14)
$GroupCombobox.Width = 394
$GroupCombobox.DropDownStyle = "DropDown"
$GroupCombobox.FlatStyle = $Style 
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
$AddButton.FlatStyle = $Style 
$AddButton.ForeColor = $ButtionForeColor 
$AddButton.BackColor = $ButtionBackColor
$AddButton.FlatAppearance.BorderSize = 0
$AddButton.add_Click({ AddMember })
    
$RemoveButton = New-Object System.Windows.Forms.Button
$RemoveButton.location = New-Object System.Drawing.Point(207, 42)
$RemoveButton.Size = New-Object System.Drawing.Size($GroupButtonY, $GroupButtonX)
$RemoveButton.Text = "Remove Members from Group"
$RemoveButton.FlatStyle = $Style 
$RemoveButton.ForeColor = $ButtionForeColor 
$RemoveButton.BackColor = $ButtionBackColor
$RemoveButton.FlatAppearance.BorderSize = 0
$RemoveButton.add_Click({ RemoveMember })

$InfoGroup = New-Object System.Windows.Forms.Button
$InfoGroup.location = New-Object System.Drawing.Point(7, 82)
$InfoGroup.Size = New-Object System.Drawing.Size($GroupButtonY, $GroupButtonX)
$InfoGroup.Text = "Show Group info"
$InfoGroup.FlatStyle = $Style 
$InfoGroup.ForeColor = $ButtionForeColor 
$InfoGroup.BackColor = $ButtionBackColor
$InfoGroup.FlatAppearance.BorderSize = 0
$InfoGroup.add_Click({ GroupInfo })
    
$AllGroups = New-Object System.Windows.Forms.Button
$AllGroups.location = New-Object System.Drawing.Point(207, 82)
$AllGroups.Size = New-Object System.Drawing.Size($GroupButtonY, $GroupButtonX)
$AllGroups.Text = "Show all Groups and Members"
$AllGroups.FlatStyle = $Style 
$AllGroups.ForeColor = $ButtionForeColor 
$AllGroups.BackColor = $ButtionBackColor
$AllGroups.FlatAppearance.BorderSize = 0
$AllGroups.add_Click({ ShowallGroups })
    
# Output GroupBox
#===========================================================
    
$OutputGroupBox = New-Object System.Windows.Forms.GroupBox
$OutputGroupBox.location = New-Object System.Drawing.Point(450,40)
$OutputGroupBox.Size = New-Object System.Drawing.Size(609, 530)
$OutputGroupBox.Text = "Output"
$OutputGroupBox.ForeColor = $GroupTextColor
$OutputGroupBox.Anchor = "Top, Bottom, Left, Right"
    
$OutputTextbox = New-Object System.Windows.Forms.RichTextBox
$OutputTextbox.location = New-Object System.Drawing.Point(7, 14)
$OutputTextBox.Size = New-Object System.Drawing.Size(595, 507)
$OutputTextbox.ScrollBars = "both"
$OutputTextbox.Multiline = $true
$OutputTextBox.BackColor = "White"
$OutputTextbox.BorderStyle = $Style 
$OutputTextbox.WordWrap = $false
$OutputTextbox.Anchor = "Top, Bottom, Left, Right"
$OutputTextbox.Font = "lucida console"
$OutputTextbox.RightToLeft = "No"
$OutputTextbox.Cursor = "IBeam"
    
# Controls
#===========================================================
    
$AccountsGroupBox.Controls.AddRange(@($AccountsLabel, $AccountsCombobox, $ResetPasswordButton,
$UnlockButton, $DisableButton, $PasNoChangeBut, $PasNvrExpBut, $NewButton, $MoveToOUBut,$AllUserinfo, $AccountinfoButton ))

$GroupGroupBox.Controls.AddRange(@( $GroupCombobox, $AddButton, $RemoveButton, $InfoGroup, $AllGroups  ))
$OutputGroupBox.Controls.AddRange(@( $OutputTextbox ))
$Form.controls.AddRange(@($Menu, $AccountsGroupBox ,$GroupGroupBox ,$ComputersGroupBox  ,$ExportGroupBox, $OutputGroupBox ))
[void]$Form.ShowDialog()
} 

StartApp 


<#
Try { 
    Get-Module activedirectory -ErrorAction Stop | Out-Null
    StartApp 
    } Catch {
    Add-Type -AssemblyName system.windows.forms
    [System.Windows.Forms.Application]::EnableVisualStyles()
    [System.Windows.Forms.MessageBox]::Show("$error","Error")
    Stop-Process $PID
    }


    #>


<#
$OULabel = New-Object System.Windows.Forms.Label
$OULabel.location = New-Object System.Drawing.Point(7, 120)
$OULabel.Text = "SeOrganizational Unit:"
$OULabel.Width = 110

$OUCombobox = New-Object System.Windows.Forms.ComboBox 
$OUCombobox.location = New-Object System.Drawing.Point(120, 120)
$OUCombobox.Width = 280
$OUCombobox.FlatStyle = $Style
ForEach ($OU in $OUs) {[void]$OUCombobox.Items.Add($OU)}
$OUCombobox.AutoCompleteSource = "CustomSource" 
$OUCombobox.AutoCompleteMode = "SuggestAppend"
$OUs | % {[void]$OUCombobox.AutoCompleteCustomSource.Add($_)}
#>