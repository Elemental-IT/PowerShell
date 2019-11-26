
# Assembly
#==========================================

Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# VariableS
#===========================================================

$style =  "flat"
$FormBackColor = "DodgerBlue"
$ButtionForeColor = "Black"
$ButtionBackColor = "Aqua"

$About = @'
  Text to Wave
   
  Version: 1.0
  Github: https://github.com/Bedlem55/PowerShell
  Author: Theo bird (Bedlem55)
    
'@


# Base Form
#==========================================

Function PlaySound {

if ($SelectVoiceCB.SelectedItem  -eq $null){
    [System.Windows.Forms.MessageBox]::Show("No voice selected","    Warning:") 
 } Else {
    $speak.SetOutputToDefaultAudioDevice() ; 
    $Speak.Rate = ($speed.Value)
    $Speak.Volume = $Volume.Value 
    $speak.SelectVoice($SelectVoiceCB.Text) 
    $speak.Speak($SpeakTextBox.Text)
    } 
}


Function SaveSound {
if ($SelectVoiceCB.SelectedItem  -eq $null){
      [System.Windows.Forms.MessageBox]::Show("No voice selected","    Warning:") 
    } else {
      $SaveChooser = New-Object -TypeName System.Windows.Forms.SaveFileDialog
      $SaveChooser.Title = "Save text to Wav file"
      $SaveChooser.FileName = "SpeechSynthesizer"
      $SaveChooser.Filter = 'Wave file (.wav) | *.wav'
        $Answer = $SaveChooser.ShowDialog(); $Answer

      if( $Answer -eq "OK" ) {
           $speak.SetOutputToDefaultAudioDevice() ; 
        $Script:Speak.Rate = ($speed.Value)
        $Speak.Volume = $Volume.Value 
        $speak.SelectVoice($SelectVoiceCB.Text) 
                $speak.SetOutputToWaveFile($SaveChooser.Filename)
          $speak.Speak($SpeakTextBox.Text)
          $speak.SetOutputToNull()
          $speak.SpeakAsyncCancelAll()
          }
     }
}

# Base Form
#==========================================


$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '798,525'
$Form.MinimumSize  = '815,570'
$Form.text = "Text to Wave"
$Form.BackColor = $FormBackColor
$Form.ShowIcon = $false
$Form.TopMost = $false

# Menu
#==========================================

$Menu  = New-Object System.Windows.Forms.MenuStrip

$MenuFile  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuFile.Text = "&File"
    [void]$Menu.Items.Add($MenuFile)

        $MenuExit  = New-Object System.Windows.Forms.ToolStripMenuItem
        $MenuExit.Text = "&Exit"
        $menuExit.Add_Click({ $Form.close() })
[void]$MenuFile.DropDownItems.Add($MenuExit)


$MenuVoices  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuVoices.Text = "&Install"
    [void]$Menu.Items.Add($MenuVoices)

        $InstallVoices  = New-Object System.Windows.Forms.ToolStripMenuItem
        $InstallVoices.Text = "&Install Other Voices"
        $InstallVoices.Add_Click({  })
[void]$MenuVoices.DropDownItems.Add($InstallVoices)


$MenuHelp  = New-Object System.Windows.Forms.ToolStripMenuItem
    $MenuHelp.Text = "&Help"
    [void]$Menu.Items.Add($MenuHelp)

        $MenuAbout  = New-Object System.Windows.Forms.ToolStripMenuItem
        $MenuAbout.Text = "&About"
        $MenuAbout.Add_Click({ [System.Windows.Forms.MessageBox]::Show("$About","    About") })
[void]$MenuHelp.DropDownItems.Add($MenuAbout)

# Base Form Objects
#==========================================

$SpeakTextBox = New-Object System.Windows.Forms.RichTextBox
$SpeakTextBox.location = "12, 35"
$SpeakTextBox.Size = "775, 350"
$SpeakTextBox.Anchor = "Top, Bottom, Left, Right"
$SpeakTextBox.Text = "Enter text here"
$speakTextbox.AllowDrop = $true
$speakTextbox.EnableAutoDragDrop = $true
$SpeakTextBox.multiline = $true
$SpeakTextBox.AcceptsTab = $true
$SpeakTextBox.ScrollBars = "both"
$SpeakTextBox.BorderStyle = "None"
$SpeakTextBox.Font = 'Microsoft Sans Serif,10'
$SpeakTextBox.Cursor = "IBeam"
$Form.Controls.Add($SpeakTextBox)

$SpeakButtion = New-Object system.Windows.Forms.Button
$SpeakButtion.location = "660, 400"
$SpeakButtion.Size = "127, 45"
$SpeakButtion.Anchor = "Bottom"
$SpeakButtion.text = "Play"
$SpeakButtion.BackColor = $ButtionBackColor
$SpeakButtion.FlatStyle = $style
$SpeakButtion.FlatAppearance.BorderSize = '0'
$SpeakButtion.Font = 'Microsoft Sans Serif,10'
$SpeakButtion.add_Click({ PlaySound })

$SaveButtion = New-Object system.Windows.Forms.Button
$SaveButtion.location = "660, 465"
$SaveButtion.Size = "127, 45"
$SaveButtion.Anchor = "Bottom"
$SaveButtion.text = "Save"
$SaveButtion.BackColor = $ButtionBackColor
$SaveButtion.FlatStyle = $style
$SaveButtion.FlatAppearance.BorderSize = '0'
$SaveButtion.Font = 'Microsoft Sans Serif,10'
$SaveButtion.add_Click({ SaveSound })

# Select Group Box
#==========================================

$SelectGB = New-Object system.Windows.Forms.Groupbox
$SelectGB.location = "11, 395"
$SelectGB.Size = "640, 50"
$SelectGB.Anchor = "Bottom"
$SelectGB.text = "Select Voice"
$SelectGB.ForeColor = "white"

$SelectVoiceCB = New-Object system.Windows.Forms.ComboBox
$SelectVoiceCB.location = "11, 15"
$SelectVoiceCB.Size = "618,24"
$SelectVoiceCB.Text = $speak.Voice.Name
$SelectVoiceCB.FlatStyle = $style
$SelectVoiceCB.DropDownStyle = 'DropDownList'

$SelectVoiceCB.Font = 'Microsoft Sans Serif,10'
$Voices = ($speak.GetInstalledVoices() | ForEach-Object { $_.voiceinfo }).Name
foreach ($Voice in $Voices) {
  [void]$SelectVoiceCB.Items.add($voice) 
}
$SelectGB.Controls.Add($SelectVoiceCB)

# Speed Group Box
#==========================================

$SpeedGB = New-Object system.Windows.Forms.Groupbox
$SpeedGB.location = "11, 450"
$SpeedGB.Size = "310,62"
$SpeedGB.Anchor = "Bottom"
$SpeedGB.text = "Speed"
$SpeedGB.ForeColor = "white"

$Speed = New-Object Windows.Forms.TrackBar
$Speed.Orientation = "Horizontal"
$Speed.location  = "5,15"
$Speed.Size = "300,40"
$Speed.TickStyle = "TopLeft"
$Speed.SetRange(-10,10)
$SpeedGB.Controls.Add( $Speed )

# Volume Group Box
#==========================================

$VolumeGB = New-Object system.Windows.Forms.Groupbox
$VolumeGB.location = "340, 450"
$VolumeGB.Size = "312,62"
$VolumeGB.Anchor = "Bottom"
$VolumeGB.text = "Volume"
$VolumeGB.ForeColor = "white"

$Volume = New-Object Windows.Forms.TrackBar
$Volume.Orientation = "Horizontal"
$Volume.location  = "5,15"
$Volume.Size = "300,40"
$Volume.TickStyle = "TopLeft"
$Volume.TickFrequency = 10
$Volume.SetRange(10,100)
$Volume.Value = 100
$VolumeGB.Controls.Add( $Volume )

# Controls
#==========================================

$Form.controls.AddRange(@( $Menu, $SpeechGB, $SpeakTextBox, $SpeakButtion, $SaveButtion, $SelectGB, $SpeedGB, $VolumeGB ))

[void]$form.ShowDialog()