Function New-User { 

    $Seasons = Get-Random @('Spring', 'Summer', 'Autumn', 'Winter') 
    $Num = Get-Random @(10..99)
    $UserName = $UPNTB.text.ToString()
    $CreatedOU = $OUCB.SelectedItem.ToString()

    $Button_UserCreateNew = @{

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
            $TextBox_Output.Clear()
            New-ADUser @New-User -ErrorAction Stop
            
            
            $TextBox_Output.AppendText(
"$UserName account has been successfully 

Accccount created in $CreatedOU

Password is $Seasons$Num and must be chagned at next login.")
    
            $Script:ADusers = (Get-ADUser -Filter *).SamAccountName

        }
        catch { OutError }
}

Function  New-UserUI {  

$NewForm = New-Object Windows.forms.form
$NewForm.text = "User creation tool"
$NewForm.Size = "550, 310"
$NewForm.TopMost = $false
$NewForm.ShowIcon = $false
$NewForm.ShowInTaskbar = $False
$NewForm.MinimizeBox = $False
$NewForm.MaximizeBox = $False
$NewForm.FormBorderStyle = 3

$1stLab = New-Object System.Windows.Forms.Label
$1stLab.Location = "15, 15"
$1stLab.Width = 80
$1stLab.Text = "First name:"
$1stLab.Font = $Font

$1stNameTB = New-Object System.Windows.Forms.TextBox
$1stNameTB.Location = "93, 11"
$1stNameTB.Width = 430
$1stNameTB.Font = $Font
$NewForm.Controls.Add( $1stNameTB )

$LasLab = New-Object System.Windows.Forms.Label
$LasLab.Location = "15, 45"
$LasLab.Width = 80
$LasLab.Text = "Last name:"
$LasLab.Font = $Font

$LasTB = New-Object System.Windows.Forms.TextBox
$LasTB.Location = "93, 41"
$LasTB.Width = 430
$LasTB.Font = $Font
$NewForm.Controls.Add( $LasTB  )

$UPNGB = New-Object System.Windows.Forms.GroupBox
$UPNGB.Location = "15, 70"
$UPNGB.Size = "508, 45"
$UPNGB.Text = "User Name"

$UPNTB = New-Object System.Windows.Forms.TextBox
$UPNTB.Location = "8, 15"
$UPNTB.Width = 340
$UPNTB.Font = $Font
$NewForm.Controls.Add( $UPNTB )

$UPNCB = New-Object System.Windows.Forms.ComboBox 
$UPNCB.Location = "350, 15"
$UPNCB.Width = 150
$UPNCB.Font = 'Microsoft Sans Serif,9'
ForEach ($Domain in $Domain_UPN) { [void]$UPNCB.Items.Add("@$Domain") }
$UPNCB.AutoCompleteSource = "CustomSource" 
$UPNCB.AutoCompleteMode = "SuggestAppend"
$Domain_UPN | ForEach-Object { [void]$UPNCB.AutoCompleteCustomSource.Add($_) }
$UserGB.Controls.Add($UserCB)

$UPNGB.Controls.AddRange(@( $UPNTB, $UPNCB ))

$OUGB = New-Object System.Windows.Forms.GroupBox
$OUGB.Location = "15, 120" 
$OUGB.Size = "508, 45"
$OUGB.Text = "Select OU"

$OUCB = New-Object System.Windows.Forms.ComboBox 
$OUCB.Location = "8, 14"
$OUCB.Width = 493
ForEach ($OU in $OUs) { [void]$OUCB.Items.Add($OU) }
$OUCB.AutoCompleteSource = "CustomSource" 
$OUCB.AutoCompleteMode = "SuggestAppend"
$OUs | ForEach-Object { [void]$OUCB.AutoCompleteCustomSource.Add($_) }
$OUGB.Controls.Add($OUCB)

$UserGB = New-Object System.Windows.Forms.GroupBox
$UserGB.Location = "15,170"
$UserGB.Size = "508, 45"
$UserGB.Text = "User account to copy ( not requered )"

$UserCB = New-Object System.Windows.Forms.ComboBox 
$UserCB.Location = "7, 14"
$UserCB.DropDownStyle = "DropDown"
$UserCB.Width = 493
ForEach ($user in $ADusers) { [void]$UserCB.Items.Add($user) }
$UserCB.AutoCompleteSource = "CustomSource" 
$UserCB.AutoCompleteMode = "SuggestAppend"
$ADusers | ForEach-Object { [void]$UserCB.AutoCompleteCustomSource.Add($_) }
$UserGB.Controls.Add($UserCB)

$NewCan = New-Object System.Windows.Forms.Button
$NewCan.Location = "140, 225"
$NewCan.Size = "128,35"
$NewCan.Text = "Cancel"
$NewCan.add_Click( { $NewForm.Close(); $NewForm.Dispose() })

$NewOk = New-Object System.Windows.Forms.Button
$NewOk.Location = "275, 225"
$NewOk.Size = "128,35"
$NewOk.Text = "Ok"
$NewOk.add_Click( { New-User })

$NewForm.controls.AddRange(@( $1stLab, $LasLab, $1stPanl, $LasPanl, $UPNGB, $OUGB, $CopyTic, $UserGB, $NewCan, $NewOk  ))
[void]$NewForm.ShowDialog()

}
