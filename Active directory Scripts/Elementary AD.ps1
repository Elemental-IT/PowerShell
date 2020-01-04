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
$GroupTextColor = "White"
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
$OUs = Get-ADOrganizationalUnit -Filter *   
$Domains = (Get-ADForest).rootdomain + (Get-ADForest).UPNSuffixes 

# help message
$About = @'
Elementary AD - Active Directory Administration Tool 

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
                Set-ADuser -Identity $AccCB.SelectedItem -ChangePasswordAtLogon $True
                $OutputTB.AppendText("$userAccount's password has been reset to $Seasons$Num and must be changed at next logon")
            }
            Else { 
                $OutputTB.Clear()
                $OutputTB.AppendText("Reset Password action canceled")
            }
        }
        catch { OutError }
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
        }
        catch { OutError }
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

            }
            catch { OutError }
        }
        else { 
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

            }
            catch { OutError }
        }
    }
}


Function Passneverexp {

    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }

    Else {
            if ((Get-ADUser $AccCB.SelectedItem -Properties *).PasswordNeverExpires -eq $false ) {

                Try {
                    $OutputTB.Clear()
                    $userAccount = $AccCB.Text.ToString()
                    set-aduser $AccCB.SelectedItem -PasswordNeverExpires:$true 
                    $OutputTB.AppendText("$userAccount``s account is set to 'Password Never Expires'")
        
                }
                catch { OutError }
            }
        
            else {
                Try {
                    $OutputTB.Clear()
                    $userAccount = $AccCB.Text.ToString()
                    set-aduser $AccCB.SelectedItem -PasswordNeverExpires:$false 
                    $OutputTB.AppendText("$userAccount's Password is now expired and must be changed")
                }
                catch { OutError }
            }
    }
}

Function PasscantChange {

    if ($AccCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No User Account Selected") 
    }

    Else {
            if ((Get-ADUser $AccCB.SelectedItem -Properties *).CannotChangePassword -eq $false ) {

                Try {
                    $OutputTB.Clear()
                    $userAccount = $AccCB.Text.ToString()
                    set-aduser $AccCB.SelectedItem -CannotChangePassword:$true
                    $OutputTB.AppendText("$userAccount's account is set to 'Cannot Change Password'")
        
                }
                catch { OutError }
            }
        
            else {
                Try {
                    $OutputTB.Clear()
                    $userAccount = $AccCB.Text.ToString()
                    set-aduser $AccCB.SelectedItem -CannotChangePassword:$false 
                    $OutputTB.AppendText("$userAccount's Password can now be changed by user")
                }
                catch { OutError }
            }
    }
}

Function  NewUser { 

    $Seasons = Get-Random @('Spring', 'Summer', 'Autumn', 'Winter') 
    $Num = Get-Random @(10..99)
    $UserName = $UPNTB.text.ToString()
    $CreatedOU = $OUCB.SelectedItem.ToString()

    $NewUser = @{

        'Name'              = $UPNTB.text.ToString()
        'GivenName'         = $1stNameTB.text.ToString()
        'Surname'           = $LasTB.text.ToString()
        'DisplayName'       = $1stNameTB.text.ToString() + '.' + $LasTB.text.ToString()
        'SamAccountName'    = $UPNTB.text.ToString()
        'UserPrincipalName' = $UPNTB.text.ToString() + $UPNCB.SelectedItem.ToString()
        'Path'              = $OUCB.SelectedItem
        'Enabled'           = $true
        'AccountPassword'   = $Seasons+$Num | ConvertTo-SecureString -AsPlainText -Force

    }   

        Try { 
            $OutputTB.Clear()
            New-ADUser @NewUser -ErrorAction Stop
            
            
            $OutputTB.AppendText(
"$UserName account has been successfully 

Accccount created in $CreatedOU

Password is $Seasons$Num and must be chagned at next login.")
    
            $Script:ADusers = (Get-ADUser -Filter *).SamAccountName

        }
        catch { OutError }
}

Function  NewUserUI {  

$NewForm = New-Object Windows.forms.form
$NewForm.text = "User creation tool"
$NewForm.size = "550, 310"
$NewForm.TopMost = $false
$NewForm.ShowIcon = $false
$NewForm.ShowInTaskbar = $False
$NewForm.MinimizeBox = $False
$NewForm.MaximizeBox = $False
$NewForm.FormBorderStyle = 3

$1stLab = New-Object System.Windows.Forms.Label
$1stLab.location = "15, 15"
$1stLab.Width = 80
$1stLab.Text = "First name:"
$1stLab.Font = $Font

$1stNameTB = New-Object System.Windows.Forms.TextBox
$1stNameTB.Location = "93, 11"
$1stNameTB.Width = 430
$1stNameTB.Font = $Font
$NewForm.Controls.Add( $1stNameTB )

$LasLab = New-Object System.Windows.Forms.Label
$LasLab.location = "15, 45"
$LasLab.Width = 80
$LasLab.Text = "Last name:"
$LasLab.Font = $Font

$LasTB = New-Object System.Windows.Forms.TextBox
$LasTB.Location = "93, 41"
$LasTB.Width = 430
$LasTB.Font = $Font
$NewForm.Controls.Add( $LasTB  )

$UPNGB = New-Object System.Windows.Forms.GroupBox
$UPNGB.location = "15, 70"
$UPNGB.Size = "508, 45"
$UPNGB.Text = "User Name"

$UPNTB = New-Object System.Windows.Forms.TextBox
$UPNTB.Location = "8, 15"
$UPNTB.Width = 340
$UPNTB.Font = $Font
$NewForm.Controls.Add( $UPNTB )

$UPNCB = New-Object System.Windows.Forms.ComboBox 
$UPNCB.location = "350, 15"
$UPNCB.Width = 150
$UPNCB.Font = 'Microsoft Sans Serif,9'
ForEach ($Domain in $Domains) { [void]$UPNCB.Items.Add("@$Domain") }
$UPNCB.AutoCompleteSource = "CustomSource" 
$UPNCB.AutoCompleteMode = "SuggestAppend"
$Domains | ForEach-Object { [void]$UPNCB.AutoCompleteCustomSource.Add($_) }
$UserGB.Controls.Add($UserCB)

$UPNGB.Controls.AddRange(@( $UPNTB, $UPNCB ))

$OUGB = New-Object System.Windows.Forms.GroupBox
$OUGB.location = "15, 120" 
$OUGB.Size = "508, 45"
$OUGB.Text = "Select OU"

$OUCB = New-Object System.Windows.Forms.ComboBox 
$OUCB.location = "8, 14"
$OUCB.Width = 493
ForEach ($OU in $OUs) { [void]$OUCB.Items.Add($OU) }
$OUCB.AutoCompleteSource = "CustomSource" 
$OUCB.AutoCompleteMode = "SuggestAppend"
$OUs | ForEach-Object { [void]$OUCB.AutoCompleteCustomSource.Add($_) }
$OUGB.Controls.Add($OUCB)

$UserGB = New-Object System.Windows.Forms.GroupBox
$UserGB.location = "15,170"
$UserGB.Size = "508, 45"
$UserGB.Text = "User account to copy ( not requered )"

$UserCB = New-Object System.Windows.Forms.ComboBox 
$UserCB.location = "7, 14"
$UserCB.DropDownStyle = "DropDown"
$UserCB.Width = 493
ForEach ($user in $ADusers) { [void]$UserCB.Items.Add($user) }
$UserCB.AutoCompleteSource = "CustomSource" 
$UserCB.AutoCompleteMode = "SuggestAppend"
$ADusers | ForEach-Object { [void]$UserCB.AutoCompleteCustomSource.Add($_) }
$UserGB.Controls.Add($UserCB)

$NewCan = New-Object System.Windows.Forms.Button
$NewCan.location = "140, 225"
$NewCan.Size = "$ShortBY, $ShortBX"
$NewCan.Text = "Cancel"
$NewCan.add_Click( { $NewForm.Close(); $NewForm.Dispose() })

$NewOk = New-Object System.Windows.Forms.Button
$NewOk.location = "275, 225"
$NewOk.Size = "$ShortBY, $ShortBX"
$NewOk.Text = "Ok"
$NewOk.add_Click( { NewUser })

$NewForm.controls.AddRange(@( $1stLab, $LasLab, $1stPanl, $LasPanl, $UPNGB, $OUGB, $CopyTic, $UserGB, $NewCan, $NewOk  ))
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
            $OutputTB.text = get-aduser $userAccount -Properties * | Format-List | Out-String -Width 2147483647
            
        }
        catch { OutError }
    }
}



Function Allusersinfo {

Try {
    $OutputTB.Clear()
    $OutputTB.text = Get-ADUser -filter * -Properties * | ForEach-Object {
        New-Object PSObject -Property @{

            UserName      = $_.SamAccountName
            Name          = $_.name
            Email         = $_.mail
            Groups        = ($_.memberof | Get-ADGroup | Select-Object -ExpandProperty Name) -join ", "
            Enabled       = $_.Enabled
            Created       = $_.whenCreated
            LastLogonDate = $_.LastLogonDate 
        }

    } | Select-Object UserName, Name, Email, Groups, Enabled, Created, LastLogonDate | Out-String -Width 2147483647
        
}
catch { OutError }
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
            $Members = Get-ADUser -Filter * | select  Name, SamAccountName, mail | Out-GridView  -Title "Select Member(s) to add to $GroupOBJ" -PassThru 
            Add-ADGroupMember -Identity $GroCB.text -Members $Members.SamAccountName -ErrorAction Stop 
            $OutputTB.AppendText("Members added to $GroupOBJ Group")
        }
        Catch { OutError }
    }
}

function RemoveMember {
    if ($GroCB.SelectedItem -eq $null) {
        $OutputTB.Clear()
        $OutputTB.AppendText("No Group Selected") 
    }
    Else {
        Try {
            $OutputTB.Clear()
            $GroupOBJ = $GroCB.Text.ToString()
            $Members = Get-ADGroup $GroCB.SelectedItem | Get-ADGroupMember | Select-Object -ExpandProperty Name | Out-GridView -Title "Select Member(s) to remove from $GroupOBJ" -PassThru
            Get-ADGroup $GroCB.SelectedItem | remove-adgroupmember -Members $members -Confirm:$false -ErrorAction Stop
            $OutputTB.AppendText("Removed members from $GroupOBJ Group")
        }
        Catch { OutError }
    }
}

function GroupInfo {
if ($GroCB.SelectedItem -eq $null) {
    $OutputTB.Clear()
    $OutputTB.AppendText("No Group Selected") 
}
Else {
    Try {
        $OutputTB.Clear()
        $OutputTB.text = Get-ADGroup $GroCB.SelectedItem -Properties * | FL | Out-String -Width 2147483647
    }
    Catch { OutError }
}
}


function ShowallGroups {

Try {
    $OutputTB.text = Get-ADGroup -Filter * -Properties * | ForEach-Object {
        New-Object PSObject -Property @{

            GroupName = $_.Name
            Type      = $_.groupcategory
            Members   = ($_.Name | Get-ADGroupMember | Select-Object -ExpandProperty Name) -join ", " 
        }

    } | Select-Object GroupName, Type, Members | Out-String -Width 2147483647
}
Catch { OutError }
}

# Computer functions
#===========================================================

function CPUinfo {
IF ($CPUCB.SelectedItem -eq $null) {
    $OutputTB.Clear()
    $OutputTB.AppendText("No Computer Selected") 
}
Else {
    Try {
        $OutputTB.Clear()
        $OutputTB.text = Get-ADComputer $CPUCB.SelectedItem -Properties * | Select-Object Name, Created, Enabled, OperatingSystem, OperatingSystemVersion, IPv4Address, LastLogonDate, logonCount | Out-String -Width 2147483647 
    }
    Catch { OutError }
}
}

function  AllCPU {
Try {
    $OutputTB.text = Get-ADComputer -Filter * -Properties * | Select-Object Name, Created, Enabled, OperatingSystem, OperatingSystemVersion, IPv4Address, LastLogonDate, logonCount | Out-String -Width 2147483647 
}
Catch { OutError }
}



# AD Export functions
#===========================================================

function CSVAdUserExport {
$AdUserExport = Get-ADUser -filter * -Properties * | ForEach-Object {
    New-Object PSObject -Property @{

        UserName      = $_.SamAccountName
        Name          = $_.name
        Email         = $_.mail
        Groups        = ($_.memberof | Get-ADGroup | Select-Object -ExpandProperty Name) -join ", "
        Enabled       = $_.Enabled
        Created       = $_.whenCreated
        LastLogonDate = $_.LastLogonDate 
    }
} 

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
        }

        Else {
            $OutputTB.Clear()
            $OutputTB.AppendText(" Exported canceled")
        }
}

function CSVGroupsExport {
$GroupsExport = Get-ADGroup -Filter * -Properties * | ForEach-Object {
    New-Object PSObject -Property @{

        GroupName = $_.Name
        Type      = $_.groupcategory
        Members   = ($_.Name | Get-ADGroupMember | Select-Object -ExpandProperty Name) -join ", " 
    }
} 

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
        }
        
        Else {
            $OutputTB.Clear()
            $OutputTB.AppendText("Exported canceled")
    
        }
}

function CSVComputerExport { 
$ComputersExport = Get-ADComputer -Filter * -Properties *
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
        }

        Else {
            $OutputTB.Clear()
            $OutputTB.AppendText("Exported canceled")
        }
}

# Base From & menus
#===========================================================

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '1080,590'
$Form.text = "     Elementary AD - Active Directory Administration Tool"
$Form.MinimumSize = '815,570'
$Form.TopMost = $false
$Form.ShowIcon = $false

$Menu = New-Object System.Windows.Forms.MenuStrip

$MenuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuFile.Text = "&File"
[void]$Menu.Items.Add($MenuFile)

$MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuExit.Text = "&Exit"
$menuExit.Add_Click({ $Form.close() })
[void]$MenuFile.DropDownItems.Add($MenuExit)

$MenuExport = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuExport.Text = "&Export"
[void]$Menu.Items.Add($MenuExport)

$CSVUser = New-Object System.Windows.Forms.ToolStripMenuItem
$CSVUser.Text = "&Export Users to CSV"
$CSVUser.Add_Click({ CSVAdUserExport })

$CSVGro = New-Object System.Windows.Forms.ToolStripMenuItem
$CSVGro.Text = "&Export Group to CSV"
$CSVGro.Add_Click({ CSVGroupsExport })

$CSVCPU = New-Object System.Windows.Forms.ToolStripMenuItem
$CSVCPU.Text = "&Export Computers to CSV"
$CSVCPU.Add_Click({ CSVComputerExport })

[void]$MenuExport.DropDownItems.AddRange(@( $CSVUser, $CSVGro, $CSVCPU ))

$MenuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuHelp.Text = "&Help"
[void]$Menu.Items.Add($MenuHelp)

$MenuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAbout.Text = "&About"
$MenuAbout.Add_Click({ [System.Windows.Forms.MessageBox]::Show("$About", "    About") })
[void]$MenuHelp.DropDownItems.Add($MenuAbout)

# User accounts GroupBox 
#===========================================================

$AccGB = New-Object System.Windows.Forms.GroupBox
$AccGB.location = "20,40"
$AccGB.Size = "409, 167"
$AccGB.Text = "Select Account"

$AccCB = New-Object System.Windows.Forms.ComboBox 
$AccCB.location = "7, 14"
$AccCB.DropDownStyle = "DropDown"
$AccCB.Width = 394
ForEach ($user in $ADusers) { [void]$AccCB.Items.Add($user) }
$AccCB.AutoCompleteSource = "CustomSource" 
$AccCB.AutoCompleteMode = "SuggestAppend"
$ADusers | ForEach-Object { [void]$AccCB.AutoCompleteCustomSource.Add($_) }

$NewUser = New-Object System.Windows.Forms.Button
$NewUser.location = "7, 42"
$NewUser.Size = "$ShortBY, $ShortBX"
$NewUser.Text = "New User"
$NewUser.add_Click( { NewUserUI })

$Unlock = New-Object System.Windows.Forms.Button
$Unlock.location = "140, 42"
$Unlock.Size = "$ShortBY, $ShortBX"
$Unlock.Text = "Unlock"
$Unlock.add_Click( { Unlock })

$Disable = New-Object System.Windows.Forms.Button
$Disable.location = "273, 42"
$Disable.Size = "$ShortBY, $ShortBX"
$Disable.Text = "Disable/Enable Account"
$Disable.add_Click( { Disable-Enable })

$RePass = New-Object System.Windows.Forms.Button
$RePass.location = "7, 82"
$RePass.Size = "$ShortBY, $ShortBX"
$RePass.Text = "Reset Password"
$RePass.add_Click( { ResetPass })

$PasNvrEx = New-Object System.Windows.Forms.Button
$PasNvrEx.location = "140, 82"
$PasNvrEx.Size = "$ShortBY, $ShortBX"
$PasNvrEx.Text = "Password never Expires"
$PasNvrEx.add_Click( { Passneverexp })

$PasNoChg = New-Object System.Windows.Forms.Button
$PasNoChg.location = "273, 82"
$PasNoChg.Size = "$ShortBY, $ShortBX"
$PasNoChg.Text = "Cannot change Password"
$PasNoChg.add_Click( { PasscantChange })

$MoveOU = New-Object System.Windows.Forms.Button
$MoveOU.location = "7, 122"
$MoveOU.Size = "$ShortBY, $ShortBX"
$MoveOU.Text = "Move to OU"
$MoveOU.add_Click( { })

$AllUser = New-Object System.Windows.Forms.Button
$AllUser.location = "140, 122"
$AllUser.Size = "$ShortBY, $ShortBX"
$AllUser.Text = "All users info"
$AllUser.add_Click({ Allusersinfo })

$Userinfo = New-Object System.Windows.Forms.Button
$Userinfo.location = "273, 122"
$Userinfo.Size = "$ShortBY, $ShortBX"
$Userinfo.Text = "Account info"
$Userinfo.add_Click({ Accountinfo })

$AccGB.Controls.AddRange(@( $AccCB, $NewUser, $Unlock, $Disable, $RePass, $PasNvrEx, $PasNoChg , $MoveOU, $AllUser, $Userinfo ))

# Group GroupBox
#===========================================================

$GroGB = New-Object System.Windows.Forms.GroupBox
$GroGB.location = "20, 220"
$GroGB.Size = "409, 128"
$GroGB.Text = "Select Group"

$GroCB = New-Object System.Windows.Forms.ComboBox 
$GroCB.location = "7, 14"
$GroCB.Width = 394
$GroCB.DropDownStyle = "DropDown"
ForEach ($Group in $Groups) { [void]$GroCB.Items.Add($Group) }
$GroCB.AutoCompleteSource = "CustomSource" 
$GroCB.AutoCompleteMode = "SuggestAppend"
$Groups | ForEach-Object { [void]$GroCB.AutoCompleteCustomSource.Add($_) }

$AddMem = New-Object System.Windows.Forms.Button
$AddMem.location = "7, 42"
$AddMem.Size = "$LongBY, $LongBX"
$AddMem.Text = "Add Members"
$AddMem.add_Click({ AddMember })

$DelMem = New-Object System.Windows.Forms.Button
$DelMem.location = "207, 42"
$DelMem.Size = "$LongBY, $LongBX"
$DelMem.Text = "Remove Members"
$DelMem.add_Click({ RemoveMember })

$AllGro = New-Object System.Windows.Forms.Button
$AllGro.location = "7, 82"
$AllGro.Size = "$LongBY, $LongBX"
$AllGro.Text = "All Groups and Members"
$AllGro.add_Click({ ShowallGroups })

$InfGro = New-Object System.Windows.Forms.Button
$InfGro.location = "207, 82"
$InfGro.Size = "$LongBY, $LongBX"
$InfGro.Text = "Group info"
$InfGro.add_Click( { GroupInfo })

$GroGB.Controls.AddRange(@( $GroCB, $AddMem, $DelMem, $AllGro, $InfGro )) 

# CPU GroupBox
#===========================================================

$CPUGB = New-Object System.Windows.Forms.GroupBox
$CPUGB.location = "20, 360"
$CPUGB.Size = "409, 88"
$CPUGB.Text = "Select Computer"

$CPUCB = New-Object System.Windows.Forms.ComboBox 
$CPUCB.location = "7, 14"
$CPUCB.Width = 394
$CPUCB.DropDownStyle = "DropDown"
ForEach ($CPU in $Computers) { [void]$CPUCB.Items.Add($CPU) }
$CPUCB.AutoCompleteSource = "CustomSource" 
$CPUCB.AutoCompleteMode = "SuggestAppend"
$Computers | ForEach-Object { [void]$CPUCB.AutoCompleteCustomSource.Add($_) }

$AllCPU = New-Object System.Windows.Forms.Button
$AllCPU.location = "7, 42"
$AllCPU.Size = "$LongBY, $LongBX"
$AllCPU.Text = "All Computers info"
$AllCPU.add_Click({ AllCPU })

$InfoCPU = New-Object System.Windows.Forms.Button
$InfoCPU.location = "207, 42"
$InfoCPU.Size = "$LongBY, $LongBX"
$InfoCPU.Text = "Computer info"
$InfoCPU.add_Click({ CPUinfo })

$CPUGB.Controls.AddRange(@( $CPUCB, $AllCPU, $InfoCPU ))

# Output GroupBox
#===========================================================

$OutputGB = New-Object System.Windows.Forms.GroupBox
$OutputGB.location = "450,40"
$OutputGB.Size = "609, 530"
$OutputGB.Text = "Output"
$OutputGB.Anchor = "Top, Bottom, Left, Right"

$OutputTB = New-Object System.Windows.Forms.RichTextBox
$OutputTB.location = "7, 14"
$OutputTB.Size = "595, 507"
$OutputTB.ScrollBars = "both"
$OutputTB.Multiline = $true
$OutputTB.BackColor = "White"
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

# Eable when finished
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

