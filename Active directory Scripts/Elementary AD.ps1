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

# Styles and Colors 
$FormBackColor = "DodgerBlue"
$FormForeColor = "white"
$ButtionForeColor = "Black"
$ButtionBackColor = "Aqua"
$GroupTextColor = "White"
$Style = 0
$Font = 'Microsoft Sans Serif,10'

# Short Button size
$ShortBX = 35
$ShortBY = 128

# Long Button size
$LongBX = 35
$LongBY = 193

# AD data
$domain = (Get-ADDomain).Name
$ADusers = (Get-ADUser -Filter *).SamAccountName
$Groups = (Get-ADGroup -Filter *).SamAccountName
$Computers = (Get-ADComputer -Filter *).SamAccountName
$OUs = (Get-ADOrganizationalUnit -Filter *).name    

# help message
$About = @'
  Elementary AD - Active Directory Administration Tool 
  
  * Designed to wors with any on premises Active Directory 
  * This tool is to make Active Directory administration quick and easy

  Authored By Theo bird (Bedlem55)
    
'@

# Base Functions
#===========================================================

Function OutError {
    $OutputTB.Clear()
    $Err = $Error[0]
    $OutputTB.AppendText("$Err")
}
    
# Account Functions
#===========================================================

Function ResetPass {        
        
    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }
    Else {
        Try {
            $userAccount = $AccCB.Text.ToString()
            $UserPrompt = new-object -comobject wscript.shell
            $Answer = $UserPrompt.popup("         Reset $userAccount Password?", 0, "Reset Password Prompt", 4)
            
            If ($Answer -eq 6) {
                $OutputTB.Clear()
                $Seasons = Get-Random @('Spring', 'Summer', 'Autumn', 'Winter') 
                $Num = Get-Random @(10..99)
                Set-ADAccountPassword -Identity $AccCB.SelectedItem -Reset -NewPassword (ConvertTo-SecureString -AsPlainText "$Seasons$Num" -Force)  -ErrorAction Stop
                if ($NextLoginCheckBox.CheckState -eq $true) { Set-ADuser -Identity $AccCB.Text -ChangePasswordAtLogon $True } 
                $OutputTB.AppendText("$userAccount's password has been reset to $Seasons$Num")}
            Else { 
                $OutputTB.Clear()
                $OutputTB.AppendText("Reset Password action canceled")
                }
        } catch { OutError }
    }
}
    
Function Unlock {
        
    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }
    Else {
        Try {
            $userAccount = $AccCB.Text.ToString()                
            $OutputTB.Clear()
            Unlock-ADAccount -Identity $AccCB.Text -ErrorAction Stop
            $OutputTB.AppendText("$userAccount's account is now unlocked")
        } catch { OutError }
    }
}
                             
                                                                                                                                                                
Function Disable-Enable {
        
    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }
    
    Else {
        if ((Get-ADUser -Identity $AccCB.SelectedItem).Enabled -eq $true) { 
            Try {
            
                $userAccount = $AccCB.Text.ToString()
                $UserPrompt = new-object -comobject wscript.shell
                $Answer = $UserPrompt.popup("        Disable $userAccount`?", 0, "Disable Account Prompt", 4)
            
                If ($Answer -eq 6) {
                    $OutputTB.Clear()
                    Disable-ADAccount -Identity $AccCB.SelectedItem -ErrorAction Stop
                    $OutputTB.AppendText("$userAccount account is now disabled")
                }
    
                Else {
                    $OutputTB.Clear()
                    $OutputTB.AppendText("Account disabled action canceled") 
                }
    
            } catch { OutError }
        } else { 
            Try {
            
                $userAccount = $AccCB.Text.ToString()
                $UserPrompt = new-object -comobject wscript.shell
                $Answer = $UserPrompt.popup("        $userAccount is disabled, Enable this account`?", 0, "Enable Account Prompt", 4)
            
                If ($Answer -eq 6) {
                    $OutputTB.Clear()
                    Enable-ADAccount -Identity $AccCB.SelectedItem -ErrorAction Stop
                    $OutputTB.AppendText("$userAccount account is now Enabled")
                }
    
                Else {
                    $OutputTB.Clear()
                    $OutputTB.AppendText("Account Enable action canceled") 
                }
    
            } catch { OutError }
        }
    }
}


Function Passneverexp {
        
    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }
    
    Else {
        if((Get-ADUser $AccCB.SelectedItem -Properties *).PasswordNeverExpires -eq $false ){

            Try {
                $OutputTB.Clear()
                $userAccount = $AccCB.Text.ToString()
                set-aduser $AccCB.SelectedItem -PasswordNeverExpires:$true 
                $OutputTB.AppendText("$userAccount``s account is set to 'Password Never Expires'")
                
            } catch { OutError }
        }else{
            Try {
                $OutputTB.Clear()
                $userAccount = $AccCB.Text.ToString()
                set-aduser $AccCB.SelectedItem -PasswordNeverExpires:$false 
                $OutputTB.AppendText("$userAccount's Password is now expired and must be changed")
            } catch { OutError }
        }
    }
}

Function PasscantChange {
        
    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }
    
    Else {
        if((Get-ADUser $AccCB.SelectedItem -Properties *).CannotChangePassword -eq $false ){

            Try {
                $OutputTB.Clear()
                $userAccount = $AccCB.Text.ToString()
                set-aduser $AccCB.SelectedItem -CannotChangePassword:$true
                $OutputTB.AppendText("$userAccount's account is set to 'Cannot Change Password'")
                
            } catch { OutError }
        }else{
            Try {
                $OutputTB.Clear()
                $userAccount = $AccCB.Text.ToString()
                set-aduser $AccCB.SelectedItem -CannotChangePassword:$false 
                $OutputTB.AppendText("$userAccount's Password can now be changed by user")
            } catch { OutError }
        }
    }
}


# Create new user 
Function NewUser {

$NewForm = New-Object Windows.forms.form
$NewForm.text = "New User creation tool"
$NewForm.size = "425, 360"
$NewForm.TopMost = $false
$NewForm.ForeColor = "white"
$NewForm.BackColor = "DodgerBlue"
$NewForm.ShowIcon = $false
$NewForm.ShowInTaskbar = $False
$NewForm.MinimizeBox = $False
$NewForm.MaximizeBox = $False
$NewForm.FormBorderStyle = 3

$1stLab = New-Object System.Windows.Forms.Label
$1stLab.location = "15, 15"
$1stLab.Width = 80
$1stLab.Text = "First name:"
$1stLab.ForeColor = "White"
$1stLab.Font = $Font

$1stPanl = New-Object System.Windows.Forms.Panel
$1stPanl.Location = "93, 15"
$1stPanl.Size = "310,20"
$1stPanl.BorderStyle = "None"
$1stPanl.BackColor = "White"

$1stNameTB = New-Object System.Windows.Forms.TextBox
$1stNameTB.Location = "3, 2"
$1stNameTB.Width = 300
$1stNameTB.BorderStyle = "None"
$1stNameTB.Font = $Font
$1stPanl.Controls.Add( $1stNameTB )

$LasLab = New-Object System.Windows.Forms.Label
$LasLab.location = "15, 45"
$LasLab.Width = 80
$LasLab.Text = "Last name:"
$LasLab.ForeColor = "White"
$LasLab.Font = $Font

$LasPanl= New-Object System.Windows.Forms.Panel
$LasPanl.Location =   "93, 45"
$LasPanl.Size = "310,20"
$LasPanl.BorderStyle = "None"
$LasPanl.BackColor = "White"

$LasTB = New-Object System.Windows.Forms.TextBox
$LasTB.Location =  "3, 2"
$LasTB.Width = 300
$LasTB.BorderStyle = "None"
$LasTB.Font = $Font
$LasPanl.Controls.Add( $LasNameTB )

$UPNGB = New-Object System.Windows.Forms.GroupBox
$UPNGB.location = "15, 70"
$UPNGB.Size = "387, 45"
$UPNGB.Text = "User Name"
$UPNGB.ForeColor = "White"

$UPNPanl= New-Object System.Windows.Forms.Panel
$UPNPanl.Location =   "8, 15"
$UPNPanl.Size = "250,21"
$UPNPanl.BorderStyle = "None"
$UPNPanl.BackColor = "White"

$UPNTB = New-Object System.Windows.Forms.TextBox
$UPNTB.Location =  "3, 2"
$UPNTB.Width = 250
$UPNTB.BorderStyle = "None"
$UPNTB.Font = $Font
$UPNPanl.Controls.Add( $UPNTB )

$UPNCB = New-Object System.Windows.Forms.ComboBox 
$UPNCB.location = "260, 15"
$UPNCB.Width = 120
$UPNCB.FlatStyle = 'flat'
$UPNCB.AutoCompleteSource = "CustomSource" 
$UPNCB.AutoCompleteMode = "SuggestAppend"

$UPNGB.Controls.AddRange(@( $UPNPanl , $UPNCB ))

$OUGB = New-Object System.Windows.Forms.GroupBox
$OUGB.location = "15, 120"
$OUGB.Size = "387, 45"
$OUGB.Text = "Select OU"
$OUGB.ForeColor = "White"

$OUCB = New-Object System.Windows.Forms.ComboBox 
$OUCB.location = "8, 14"
$OUCB.Width = 372
$OUCB.FlatStyle = 'flat'  
ForEach ($OU in $OUs) {[void]$OUCB.Items.Add($OU)}
$OUCB.AutoCompleteSource = "CustomSource" 
$OUCB.AutoCompleteMode = "SuggestAppend"
$OUs | % {[void]$OUCB.AutoCompleteCustomSource.Add($_)}
$OUGB.Controls.Add($OUCB)

$PasLabel = New-Object System.Windows.Forms.Label
$PasLabel.location = "15, 175"
$PasLabel.Width = 80
$PasLabel.Text = "Password:"
$PasLabel.ForeColor = "White"
$PasLabel.Font = $Font

$PasPanl = New-Object System.Windows.Forms.Panel
$PasPanl.Location = "93, 175"
$PasPanl.Size = "310,20"
$PasPanl.BorderStyle = "None"
$PasPanl.BackColor = "White"

$PasTB = New-Object System.Windows.Forms.TextBox
$PasTB.Location = "3, 2"
$PasTB.Width = 300
$PasTB.BorderStyle = "None"
$PasTB.Font = $Font
$PasPanl.Controls.Add( $PasTB )

$CopyTic = New-Object System.Windows.Forms.CheckBox
$CopyTic.location = "15, 200"
$CopyTic.Text = "Copy selected user Account"
$CopyTic.Width = 400

$UserGB = New-Object System.Windows.Forms.GroupBox
$UserGB.location = "15,225"
$UserGB.Size = "387, 45"
$UserGB.Text = "User account to copy"
$UserGB.ForeColor = "White"

$UserCB = New-Object System.Windows.Forms.ComboBox 
$UserCB.location = "7, 14"
$UserCB.DropDownStyle = "DropDown"
$UserCB.Width = 372
$UserCB.FlatStyle = "Flat"
ForEach ($user in $ADusers) {[void]$UserCB.Items.Add($user)}
$UserCB.AutoCompleteSource = "CustomSource" 
$UserCB.AutoCompleteMode = "SuggestAppend"
$ADusers | % {[void]$UserCB.AutoCompleteCustomSource.Add($_)}
$UserGB.Controls.Add($UserCB)


$NewCan = New-Object System.Windows.Forms.Button
$NewCan.location = "140, 280"
$NewCan.Size = "$ShortBY, $ShortBX"
$NewCan.Text = "Cancel"
$NewCan.FlatStyle = $Style 
$NewCan.ForeColor = $ButtionForeColor 
$NewCan.BackColor = $ButtionBackColor
$NewCan.FlatAppearance.BorderSize = 0
$NewCan.add_Click({ $NewForm.Close() })

$NewOk = New-Object System.Windows.Forms.Button
$NewOk.location = "275, 280"
$NewOk.Size = "$ShortBY, $ShortBX"
$NewOk.Text = "Ok"
$NewOk.FlatStyle = $Style 
$NewOk.ForeColor = $ButtionForeColor 
$NewOk.BackColor = $ButtionBackColor
$NewOk.FlatAppearance.BorderSize = 0
$NewOk.add_Click({  })

$NewForm.controls.AddRange(@( $1stLab, $LasLab, $1stPanl, $LasPanl, $UPNGB, $OUGB, $PasLabel, $PasPanl, $CopyTic, $UserGB, $NewCan, $NewOk  ))
[void]$NewForm.ShowDialog()
}


Function Accountinfo {
        
    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }
    
    Else {
        Try {
            $OutputTB.Clear()
            $userAccount = $AccCB.Text.ToString()
            $OutputTB.text = get-aduser $userAccount -Properties * | Format-List |  Out-String -Width 2147483647
                
        } catch { OutError }
    }
}



Function Allusersinfo {
        
            Try {
            $OutputTB.Clear()
            $OutputTB.text = Get-ADUser -filter * -Properties * | ForEach-Object {
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
        
    if ($GroCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No Group Selected") 
    }
    Else {
        Try {
            $GroupOBJ = $GroCB.Text.ToString()
            $OutputTB.Clear()
            Get-ADUser -Filter * | select  Name, SamAccountName, mail | Out-GridView  -Title "Select Member(s) to add to $GroupOBJ" -PassThru | Tee-Object -Variable "Members"
            Add-ADGroupMember -Identity $GroCB.text -Members $Members.SamAccountName -ErrorAction Stop 
            $OutputTB.AppendText("Members added to $GroupOBJ Group")
        } Catch { OutError }
    }
}
    
function RemoveMember {
    if ($GroCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No Group Selected") 
    }
    Else {
        Try {$OutputTB.Clear()
            $GroupOBJ = $GroCB.Text.ToString()
            Get-ADGroup $GroCB.SelectedItem | Get-ADGroupMember | Select-Object -ExpandProperty Name | Out-GridView -Title "Select Member(s) to remove from $GroupOBJ" -PassThru | Tee-Object -Variable "Members"
            Get-ADGroup $GroCB.SelectedItem | remove-adgroupmember -Members $members -Confirm:$false -ErrorAction Stop
            $OutputTB.AppendText("Removed members from $GroupOBJ Group")
        } Catch { OutError }
    }
}

function GroupInfo {
    if ($GroCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No Group Selected") 
    }
    Else {
        Try {$OutputTB.Clear()
            $OutputTB.text = Get-ADGroup $GroCB.SelectedItem -Properties * | FL| Out-String -Width 2147483647
        } Catch { OutError }
    }
}


function ShowallGroups {
        
    Try {$OutputTB.text =  Get-ADGroup -Filter * -Properties * | ForEach-Object {
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
        $OutputTB.Clear()
        $OutputTB.AppendText("Exported user acconunts to $SavePath")
    } Else  {
        $OutputTB.Clear()
        $OutputTB.AppendText(" Exported canceled")
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
        $OutputTB.Clear()
        $OutputTB.AppendText("Exported Groups acconunts to $SavePath")
    } Else {
        $OutputTB.Clear()
        $OutputTB.AppendText("Exported canceled")
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
        $OutputTB.Clear()
        $OutputTB.AppendText("Exported Computer acconunts to $SavePath")
    } Else {
        $OutputTB.Clear()
        $OutputTB.AppendText("Exported canceled")
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

        $ExpUser = New-Object System.Windows.Forms.ToolStripMenuItem
        $ExpUser.Text = "&Export Users to CSV"
        $ExpUser.Add_Click({ AdUserExport })

        $ExpGro = New-Object System.Windows.Forms.ToolStripMenuItem
        $ExpGro.Text = "&Export Group to CSV"
        $ExpGro.Add_Click({ GroupsExport })

        $ExpCPU = New-Object System.Windows.Forms.ToolStripMenuItem
        $ExpCPU.Text = "&Export Computers to CSV"
        $ExpCPU.Add_Click({ ComputerExport })

[void]$MenuExport.DropDownItems.AddRange(@( $ExpUser, $ExpGro, $ExpCPU  ))


$MenuHelp  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuHelp.Text = "&Help"
    [void]$Menu.Items.Add($MenuHelp)

        $MenuAbout  = New-Object System.Windows.Forms.ToolStripMenuItem
        $MenuAbout.Text = "&About"
        $MenuAbout.Add_Click({ [System.Windows.Forms.MessageBox]::Show("$About","    About") })
[void]$MenuHelp.DropDownItems.Add($MenuAbout)
    

# User accounts GroupBox 
#===========================================================
  
$AccGB = New-Object System.Windows.Forms.GroupBox
$AccGB.location = "20,40"
$AccGB.Size = "409, 167"
$AccGB.Text = "User Account Actions"
$AccGB.ForeColor = $GroupTextColor

$AccCB = New-Object System.Windows.Forms.ComboBox 
$AccCB.location = "7, 14"
$AccCB.DropDownStyle = "DropDown"
$AccCB.Width = 394
$AccCB.FlatStyle = $Style
ForEach ($user in $ADusers) {[void]$AccCB.Items.Add($user)}
$AccCB.AutoCompleteSource = "CustomSource" 
$AccCB.AutoCompleteMode = "SuggestAppend"
$ADusers | % {[void]$AccCB.AutoCompleteCustomSource.Add($_)}

$NewUser = New-Object System.Windows.Forms.Button
$NewUser.location = "7, 42"
$NewUser.Size = "$ShortBY, $ShortBX"
$NewUser.Text = "New User"
$NewUser.FlatStyle = $Style 
$NewUser.ForeColor = $ButtionForeColor 
$NewUser.BackColor = $ButtionBackColor
$NewUser.FlatAppearance.BorderSize = 0
$NewUser.add_Click({ NewUser })
   
$Unlock = New-Object System.Windows.Forms.Button
$Unlock.location = "140, 42"
$Unlock.Size = "$ShortBY, $ShortBX"
$Unlock.Text = "Unlock"
$Unlock.FlatStyle = $Style 
$Unlock.ForeColor = $ButtionForeColor 
$Unlock.BackColor = $ButtionBackColor
$Unlock.FlatAppearance.BorderSize = 0
$Unlock.add_Click({ Unlock })
    
$Disable = New-Object System.Windows.Forms.Button
$Disable.location = "273, 42"
$Disable.Size = "$ShortBY, $ShortBX"
$Disable.Text = "Disable/Enable Account"
$Disable.FlatStyle = $Style 
$Disable.ForeColor = $ButtionForeColor 
$Disable.BackColor = $ButtionBackColor
$Disable.FlatAppearance.BorderSize = 0
$Disable.add_Click({ Disable-Enable })
 
$RePass = New-Object System.Windows.Forms.Button
$RePass.location = "7, 82"
$RePass.Size = "$ShortBY, $ShortBX"
$RePass.Text = "Reset Password"
$RePass.FlatStyle = $Style 
$RePass.ForeColor = $ButtionForeColor 
$RePass.BackColor = $ButtionBackColor
$RePass.FlatAppearance.BorderSize = 0
$RePass.add_Click({ ResetPass })
    
$PasNvrEx = New-Object System.Windows.Forms.Button
$PasNvrEx.location = "140, 82"
$PasNvrEx.Size = "$ShortBY, $ShortBX"
$PasNvrEx.Text = "Password never Expires"
$PasNvrEx.FlatStyle = $Style 
$PasNvrEx.ForeColor = $ButtionForeColor 
$PasNvrEx.BackColor = $ButtionBackColor
$PasNvrEx.FlatAppearance.BorderSize = 0
$PasNvrEx.add_Click({ Passneverexp })

$PasNoChg = New-Object System.Windows.Forms.Button
$PasNoChg.location = "273, 82"
$PasNoChg.Size = "$ShortBY, $ShortBX"
$PasNoChg.Text = "Cannot change Password"
$PasNoChg.FlatStyle = $Style 
$PasNoChg.ForeColor = $ButtionForeColor 
$PasNoChg.BackColor = $ButtionBackColor
$PasNoChg.FlatAppearance.BorderSize = 0
$PasNoChg.add_Click({ PasscantChange })

$MoveOU = New-Object System.Windows.Forms.Button
$MoveOU.location = "7, 122"
$MoveOU.Size = "$ShortBY, $ShortBX"
$MoveOU.Text = "Move to OU"
$MoveOU.FlatStyle = $Style 
$MoveOU.ForeColor = $ButtionForeColor 
$MoveOU.BackColor = $ButtionBackColor
$MoveOU.FlatAppearance.BorderSize = 0
$MoveOU.add_Click({  })

$AllUser = New-Object System.Windows.Forms.Button
$AllUser.location = "140, 122"
$AllUser.Size = "$ShortBY, $ShortBX"
$AllUser.Text = "All users info"
$AllUser.FlatStyle = $Style 
$AllUser.ForeColor = $ButtionForeColor 
$AllUser.BackColor = $ButtionBackColor
$AllUser.FlatAppearance.BorderSize = 0
$AllUser.add_Click({ Allusersinfo })
    
$Userinfo = New-Object System.Windows.Forms.Button
$Userinfo.location = "273, 122"
$Userinfo.Size = "$ShortBY, $ShortBX"
$Userinfo.Text = "Account info"
$Userinfo.FlatStyle = $Style 
$Userinfo.ForeColor = $ButtionForeColor 
$Userinfo.BackColor = $ButtionBackColor
$Userinfo.FlatAppearance.BorderSize = 0
$Userinfo.add_Click({ Accountinfo })

$AccGB.Controls.AddRange(@( $AccCB, $NewUser, $Unlock, $Disable, $RePass, $PasNvrEx, $PasNoChg ,$MoveOU, $AllUser, $Userinfo ))
    
# Group GroupBox
#===========================================================
    
$GroGB = New-Object System.Windows.Forms.GroupBox
$GroGB.location = "20, 220"
$GroGB.Size = "409, 167"
$GroGB.Text = "Select Group"
$GroGB.ForeColor = $GroupTextColor
    
$GroCB = New-Object System.Windows.Forms.ComboBox 
$GroCB.location = "7, 14"
$GroCB.Width = 394
$GroCB.DropDownStyle = "DropDown"
$GroCB.FlatStyle = $Style 
ForEach ($Group in $Groups) {[void]$GroCB.Items.Add($Group)}
$GroCB.AutoCompleteSource = "CustomSource" 
$GroCB.AutoCompleteMode = "SuggestAppend"
$Groups | % {[void]$GroCB.AutoCompleteCustomSource.Add($_)}
    
 
$AddMem = New-Object System.Windows.Forms.Button
$AddMem.location = "7, 42"
$AddMem.Size = "$LongBY, $LongBX"
$AddMem.Text = "Add Members"
$AddMem.FlatStyle = $Style 
$AddMem.ForeColor = $ButtionForeColor 
$AddMem.BackColor = $ButtionBackColor
$AddMem.FlatAppearance.BorderSize = 0
$AddMem.add_Click({ AddMember })
    
$DelMem = New-Object System.Windows.Forms.Button
$DelMem.location = "207, 42"
$DelMem.Size = "$LongBY, $LongBX"
$DelMem.Text = "Remove Members"
$DelMem.FlatStyle = $Style 
$DelMem.ForeColor = $ButtionForeColor 
$DelMem.BackColor = $ButtionBackColor
$DelMem.FlatAppearance.BorderSize = 0
$DelMem.add_Click({ RemoveMember })

$NewGro = New-Object System.Windows.Forms.Button
$NewGro.location = "7, 82"
$NewGro.Size = "$LongBY, $LongBX"
$NewGro.Text = "New Group"
$NewGro.FlatStyle = $Style 
$NewGro.ForeColor = $ButtionForeColor 
$NewGro.BackColor = $ButtionBackColor
$NewGro.FlatAppearance.BorderSize = 0
$NewGro.add_Click({  })

$DelGro = New-Object System.Windows.Forms.Button
$DelGro.location = "207, 82"
$DelGro.Size = "$LongBY, $LongBX"
$DelGro.Text = "Delete Group"
$DelGro.FlatStyle = $Style 
$DelGro.ForeColor = $ButtionForeColor 
$DelGro.BackColor = $ButtionBackColor
$DelGro.FlatAppearance.BorderSize = 0
$DelGro.add_Click({  })
    
$AllGro = New-Object System.Windows.Forms.Button
$AllGro.location = "7, 122"
$AllGro.Size = "$LongBY, $LongBX"
$AllGro.Text = "All Groups and Members"
$AllGro.FlatStyle = $Style 
$AllGro.ForeColor = $ButtionForeColor 
$AllGro.BackColor = $ButtionBackColor
$AllGro.FlatAppearance.BorderSize = 0
$AllGro.add_Click({ ShowallGroups })

$InfGro = New-Object System.Windows.Forms.Button
$InfGro.location = "207, 122"
$InfGro.Size = "$LongBY, $LongBX"
$InfGro.Text = "Group info"
$InfGro.FlatStyle = $Style 
$InfGro.ForeColor = $ButtionForeColor 
$InfGro.BackColor = $ButtionBackColor
$InfGro.FlatAppearance.BorderSize = 0
$InfGro.add_Click({ GroupInfo })

$GroGB.Controls.AddRange(@( $GroCB, $AddMem, $DelMem, $NewGro, $DelGro, $AllGro, $InfGro )) 
# CPU GroupBox
#===========================================================
    
$CPUGB= New-Object System.Windows.Forms.GroupBox
$CPUGB.location = "20, 400"
$CPUGB.Size = "409, 127"
$CPUGB.Text = "Select Comouter"
$CPUGB.ForeColor = $GroupTextColor
    
$CPUCB = New-Object System.Windows.Forms.ComboBox 
$CPUCB.location = "7, 14"
$CPUCB.Width = 394
$CPUCB.DropDownStyle = "DropDown"
$CPUCB.FlatStyle = $Style 
ForEach ($CPU in $Computers) {[void]$CPUCB.Items.Add($CPU)}
$CPUCB.AutoCompleteSource = "CustomSource" 
$CPUCB.AutoCompleteMode = "SuggestAppend"
$Computers | % {[void]$CPUCB.AutoCompleteCustomSource.Add($_)}

$DisCPU = New-Object System.Windows.Forms.Button
$DisCPU.location = "7, 42"
$DisCPU.Size = "$LongBY, $LongBX"
$DisCPU.Text = "Disable/Enable Computer"
$DisCPU.FlatStyle = $Style 
$DisCPU.ForeColor = $ButtionForeColor 
$DisCPU.BackColor = $ButtionBackColor
$DisCPU.FlatAppearance.BorderSize = 0
$DisCPU.add_Click({  })

$DelCPU = New-Object System.Windows.Forms.Button
$DelCPU.location = "207, 42"
$DelCPU.Size = "$LongBY, $LongBX"
$DelCPU.Text = "Delete Computer"
$DelCPU.FlatStyle = $Style 
$DelCPU.ForeColor = $ButtionForeColor 
$DelCPU.BackColor = $ButtionBackColor
$DelCPU.FlatAppearance.BorderSize = 0
$DelCPU.add_Click({  })
    
$AllCPU = New-Object System.Windows.Forms.Button
$AllCPU.location = "7, 82"
$AllCPU.Size = "$LongBY, $LongBX"
$AllCPU.Text = "All Computers info"
$AllCPU.FlatStyle = $Style 
$AllCPU.ForeColor = $ButtionForeColor 
$AllCPU.BackColor = $ButtionBackColor
$AllCPU.FlatAppearance.BorderSize = 0
$AllCPU.add_Click({  })

$InfoCPU = New-Object System.Windows.Forms.Button
$InfoCPU.location = "207, 82"
$InfoCPU.Size = "$LongBY, $LongBX"
$InfoCPU.Text = "Computer info"
$InfoCPU.FlatStyle = $Style 
$InfoCPU.ForeColor = $ButtionForeColor 
$InfoCPU.BackColor = $ButtionBackColor
$InfoCPU.FlatAppearance.BorderSize = 0
$InfoCPU.add_Click({  })

$CPUGB.Controls.AddRange(@( $CPUCB, $DisCPU, $DelCPU, $AllCPU, $InfoCPU ))

    
# Output GroupBox
#===========================================================
    
$OutputGB = New-Object System.Windows.Forms.GroupBox
$OutputGB.location = "450,40"
$OutputGB.Size = "609, 530"
$OutputGB.Text = "Output"
$OutputGB.ForeColor = $GroupTextColor
$OutputGB.Anchor = "Top, Bottom, Left, Right"
    
$OutputTB = New-Object System.Windows.Forms.RichTextBox
$OutputTB.location = "7, 14"
$OutputTB.Size = "595, 507"
$OutputTB.ScrollBars = "both"
$OutputTB.Multiline = $true
$OutputTB.BackColor = "White"
$OutputTB.BorderStyle = $Style 
$OutputTB.WordWrap = $false
$OutputTB.Anchor = "Top, Bottom, Left, Right"
$OutputTB.Font = "lucida console"
$OutputTB.RightToLeft = "No"
$OutputTB.Cursor = "IBeam"

$OutputGB.Controls.AddRange(@( $OutputTB ))
    
# Controls
#===========================================================

$Form.controls.AddRange(@($Menu, $AccGB, $GroGB, $OutputGB, $CPUGB ))
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


<# Move this to reset password output
$NextLoginCheckBox = New-Object System.Windows.Forms.CheckBox
$NextLoginCheckBox.location = New-Object System.Drawing.Point(7, 78)
$NextLoginCheckBox.Text = "Reset password at next login"
$NextLoginCheckBox.Width = 170
$NextLoginCheckBox.Checked = "Checked"
#>