<#
.SYNOPSIS
  Text to voice app
.DESCRIPTION
  A GUI for window speechsynthesis  
.NOTES
  Author Theo bird
#>

# Assembly
#==========================================

Add-Type -AssemblyName System.speech
$Speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

# Variables
#===========================================================

$Admin = [bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")

$About = @'
  Text to Wave
   
  Version: 1.0
  Github: https://github.com/Bedlem55/PowerShell
  Author: Theo bird (Bedlem55)
    
'@

$Eva = @'
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens\MSTTS_V110_enUS_EvaM]
@="Microsoft Eva Mobile - English (United States)"
"LangDataPath"="%windir%\\Speech_OneCore\\Engines\\TTS\\en-US\\MSTTSLocenUS.dat"
"LangUpdateDataDirectory"="%SystemDrive%\\Data\\SharedData\\Speech_OneCore\\Engines\\TTS\\en-US"
"VoicePath"="%windir%\\Speech_OneCore\\Engines\\TTS\\en-US\\M1033Eva"
"VoiceUpdateDataDirectory"="%SystemDrive%\\Data\\SharedData\\Speech_OneCore\\Engines\\TTS\\en-US"
"409"="Microsoft Eva Mobile - English (United States)"
"CLSID"="{179F3D56-1B0B-42B2-A962-59B7EF59FE1B}"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech_OneCore\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes]
"Version"="11.0"
"Language"="409"
"Gender"="Female"
"Age"="Adult"
"DataVersion"="11.0.2013.1022"
"SharedPronunciation"=""
"Name"="Microsoft Eva Mobile"
"Vendor"="Microsoft"
"PersonalAssistant"="1"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSTTS_V110_enUS_EvaM]
@="Microsoft Eva Mobile - English (United States)"
"LangDataPath"="%windir%\\Speech_OneCore\\Engines\\TTS\\en-US\\MSTTSLocenUS.dat"
"LangUpdateDataDirectory"="%SystemDrive%\\Data\\SharedData\\Speech_OneCore\\Engines\\TTS\\en-US"
"VoicePath"="%windir%\\Speech_OneCore\\Engines\\TTS\\en-US\\M1033Eva"
"VoiceUpdateDataDirectory"="%SystemDrive%\\Data\\SharedData\\Speech_OneCore\\Engines\\TTS\\en-US"
"409"="Microsoft Eva Mobile - English (United States)"
"CLSID"="{179F3D56-1B0B-42B2-A962-59B7EF59FE1B}"

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes]
"Version"="11.0"
"Language"="409"
"Gender"="Female"
"Age"="Adult"
"DataVersion"="11.0.2013.1022"
"SharedPronunciation"=""
"Name"="Microsoft Eva Mobile"
"Vendor"="Microsoft"
"PersonalAssistant"="1"

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\SPEECH\Voices\Tokens\MSTTS_V110_enUS_EvaM]
@="Microsoft Eva Mobile - English (United States)"
"LangDataPath"="%windir%\\Speech_OneCore\\Engines\\TTS\\en-US\\MSTTSLocenUS.dat"
"LangUpdateDataDirectory"="%SystemDrive%\\Data\\SharedData\\Speech_OneCore\\Engines\\TTS\\en-US"
"VoicePath"="%windir%\\Speech_OneCore\\Engines\\TTS\\en-US\\M1033Eva"
"VoiceUpdateDataDirectory"="%SystemDrive%\\Data\\SharedData\\Speech_OneCore\\Engines\\TTS\\en-US"
"409"="Microsoft Eva Mobile - English (United States)"
"CLSID"="{179F3D56-1B0B-42B2-A962-59B7EF59FE1B}"

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\SPEECH\Voices\Tokens\MSTTS_V110_enUS_EvaM\Attributes]
"Version"="11.0"
"Language"="409"
"Gender"="Female"
"Age"="Adult"
"DataVersion"="11.0.2013.1022"
"SharedPronunciation"=""
"Name"="Microsoft Eva Mobile"
"Vendor"="Microsoft"
"PersonalAssistant"="1"
'@

$Mark = @'
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSTTS_V110_enUS_MarkM]
@="Microsoft Mark - English (United States)"
"409"="Microsoft Mark - English (United States)"
"CLSID"="{179F3D56-1B0B-42B2-A962-59B7EF59FE1B}"
"LangDataPath"=hex(2):25,00,77,00,69,00,6e,00,64,00,69,00,72,00,25,00,5c,00,53,\
  00,70,00,65,00,65,00,63,00,68,00,5f,00,4f,00,6e,00,65,00,43,00,6f,00,72,00,\
  65,00,5c,00,45,00,6e,00,67,00,69,00,6e,00,65,00,73,00,5c,00,54,00,54,00,53,\
  00,5c,00,65,00,6e,00,2d,00,55,00,53,00,5c,00,4d,00,53,00,54,00,54,00,53,00,\
  4c,00,6f,00,63,00,65,00,6e,00,55,00,53,00,2e,00,64,00,61,00,74,00,00,00
"VoicePath"=hex(2):25,00,77,00,69,00,6e,00,64,00,69,00,72,00,25,00,5c,00,53,00,\
  70,00,65,00,65,00,63,00,68,00,5f,00,4f,00,6e,00,65,00,43,00,6f,00,72,00,65,\
  00,5c,00,45,00,6e,00,67,00,69,00,6e,00,65,00,73,00,5c,00,54,00,54,00,53,00,\
  5c,00,65,00,6e,00,2d,00,55,00,53,00,5c,00,4d,00,31,00,30,00,33,00,33,00,4d,\
  00,61,00,72,00,6b,00,00,00

[HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Speech\Voices\Tokens\MSTTS_V110_enUS_MarkM\Attributes]
"Age"="Adult"
"DataVersion"="11.0.2013.1022"
"Gender"="Male"
"Language"="409"
"Name"="Microsoft Mark"
"SharedPronunciation"=""
"Vendor"="Microsoft"
"Version"="11.0"

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\SPEECH\Voices\Tokens\MSTTS_V110_enUS_MarkM]
@="Microsoft Mark - English (United States)"
"409"="Microsoft Mark - English (United States)"
"CLSID"="{179F3D56-1B0B-42B2-A962-59B7EF59FE1B}"
"LangDataPath"=hex(2):25,00,77,00,69,00,6e,00,64,00,69,00,72,00,25,00,5c,00,53,\
  00,70,00,65,00,65,00,63,00,68,00,5f,00,4f,00,6e,00,65,00,43,00,6f,00,72,00,\
  65,00,5c,00,45,00,6e,00,67,00,69,00,6e,00,65,00,73,00,5c,00,54,00,54,00,53,\
  00,5c,00,65,00,6e,00,2d,00,55,00,53,00,5c,00,4d,00,53,00,54,00,54,00,53,00,\
  4c,00,6f,00,63,00,65,00,6e,00,55,00,53,00,2e,00,64,00,61,00,74,00,00,00
"VoicePath"=hex(2):25,00,77,00,69,00,6e,00,64,00,69,00,72,00,25,00,5c,00,53,00,\
  70,00,65,00,65,00,63,00,68,00,5f,00,4f,00,6e,00,65,00,43,00,6f,00,72,00,65,\
  00,5c,00,45,00,6e,00,67,00,69,00,6e,00,65,00,73,00,5c,00,54,00,54,00,53,00,\
  5c,00,65,00,6e,00,2d,00,55,00,53,00,5c,00,4d,00,31,00,30,00,33,00,33,00,4d,\
  00,61,00,72,00,6b,00,00,00

[HKEY_LOCAL_MACHINE\SOFTWARE\WOW6432Node\Microsoft\SPEECH\Voices\Tokens\MSTTS_V110_enUS_MarkM\Attributes]
"Age"="Adult"
"DataVersion"="11.0.2013.1022"
"Gender"="Male"
"Language"="409"
"Name"="Microsoft Mark"
"SharedPronunciation"=""
"Vendor"="Microsoft"
"Version"="11.0"
'@

$Message = @'
Enable Eva and Mark system voices? 

Warning: this will modify the system registry.
'@

$OS = @'
OS does not meet requirements:

Windows 10 or Server 2016 and higher is required.
'@

$AdminMeg = @'
Requires elevation to enable. 

Run text to wave as administrator 
'@

$Restart = @'
Restart required, restart computer now?
'@

# Base Form
#==========================================

Function PlaySound {

  if ($null -eq $SelectVoiceCB.SelectedItem) {
    [System.Windows.Forms.MessageBox]::Show("No voice selected", "Warning:",0,48) 
  }
  Else {
    $Speak.SetOutputToDefaultAudioDevice() ; 
    $Speak.Rate = ($speed.Value)
    $Speak.Volume = $Volume.Value 
    $Speak.SelectVoice($SelectVoiceCB.Text) 
    $Speak.Speak($SpeakTextBox.Text)
  } 
}

Function SaveSound {
  if ($null -eq $SelectVoiceCB.SelectedItem) {
    [System.Windows.Forms.MessageBox]::Show("No voice selected", "Warning:",0,48) 
  }
  else {
    $SaveChooser = New-Object -TypeName System.Windows.Forms.SaveFileDialog
    $SaveChooser.Title = "Save text to Wav file"
    $SaveChooser.FileName = "SpeechSynthesizer"
    $SaveChooser.Filter = 'Wave file (.wav) | *.wav'
    $Answer = $SaveChooser.ShowDialog(); $Answer

    if ( $Answer -eq "OK" ) {
      $Speak.SetOutputToDefaultAudioDevice() ; 
      $Speak.Rate = ($speed.Value)
      $Speak.Volume = $Volume.Value 
      $Speak.SelectVoice($SelectVoiceCB.Text) 
      $Speak.SetOutputToWaveFile($SaveChooser.Filename)
      $Speak.Speak($SpeakTextBox.Text)
      $Speak.SetOutputToNull()
      $Speak.SpeakAsyncCancelAll()
    }
  }
}

Function EnableMarkandEva { 

  if (-not(Get-WmiObject -Class win32_operatingsystem).version.remove(2) -eq 10 ) { 
    [System.Windows.Forms.MessageBox]::Show("$OS","Warning:",0,48) 
  }

  else {
    if ($Admin -eq $true) {

    $UserPrompt = new-object -comobject wscript.shell
    $Answer = $UserPrompt.popup($Message, 0, "Enable system Voices", 4)

      If ($Answer -eq 6) {
        New-Item -Value $eva -Path $env:SystemDrive\Eva.reg
        New-Item -Value $Mark -Path $env:SystemDrive\Mark.reg
        Start-Process regedit.exe -ArgumentList  /s, $env:SystemDrive\Eva.reg -Wait  
        Start-Process regedit.exe -ArgumentList  /s, $env:SystemDrive\Mark.reg -Wait
        Remove-Item $env:SystemDrive\Mark.reg -Force
        Remove-Item $env:SystemDrive\Eva.reg  -Force

        $UserPrompt = new-object -comobject wscript.shell
        $Answer = $UserPrompt.popup($Restart, 0, "Restart prompt", 4)
          If ($Answer -eq 6) { Restart-Computer -Force }

      } 
    }   Else { [System.Windows.Forms.MessageBox]::Show("$AdminMeg","Warning:",0,48) } 
  }
}

# Base Form
#==========================================

$Form = New-Object system.Windows.Forms.Form
$Form.ClientSize = '798,525'
$Form.MinimumSize = '815,570'
$Form.text = "Text to Wave"
$Form.ShowIcon = $false
$Form.TopMost = $false

# Menu
#==========================================

$Menu = New-Object System.Windows.Forms.MenuStrip

$MenuFile = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuFile.Text = "&File"
[void]$Menu.Items.Add($MenuFile)

$MenuExit = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuExit.Text = "&Exit"
$menuExit.Add_Click( { $Form.close() })
[void]$MenuFile.DropDownItems.Add($MenuExit)


$MenuVoices = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuVoices.Text = "&Voice"
[void]$Menu.Items.Add($MenuVoices)

$InstallVoices = New-Object System.Windows.Forms.ToolStripMenuItem
$InstallVoices.Text = "&Enable MarkandEva"
$InstallVoices.Add_Click( { EnableMarkandEva })
[void]$MenuVoices.DropDownItems.Add($InstallVoices)

$MenuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuHelp.Text = "&Help"
[void]$Menu.Items.Add($MenuHelp)

$MenuAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$MenuAbout.Text = "&About"
$MenuAbout.Add_Click( { [System.Windows.Forms.MessageBox]::Show("$About", "About",0,64) })
[void]$MenuHelp.DropDownItems.Add($MenuAbout)

$SpeakButtion = New-Object system.Windows.Forms.Button
$SpeakButtion.location = "660, 401"
$SpeakButtion.Size = "127, 43"
$SpeakButtion.Anchor = "Bottom"
$SpeakButtion.text = "Play"
$SpeakButtion.Font = 'Microsoft Sans Serif,10'
$SpeakButtion.add_Click( { PlaySound })

$SaveButtion = New-Object system.Windows.Forms.Button
$SaveButtion.location = "660, 456"
$SaveButtion.Size = "127, 55"
$SaveButtion.Anchor = "Bottom"
$SaveButtion.text = "Save"
$SaveButtion.Font = 'Microsoft Sans Serif,10'
$SaveButtion.add_Click( { SaveSound })

# Text Group Box
#==========================================

$TextGB = New-Object system.Windows.Forms.Groupbox
$TextGB.Anchor = "Top, Bottom, Left, Right"
$TextGB.location = "10, 35"
$TextGB.Size = "775, 350"
$TextGB.text = "Enter or drag text here"

$SpeakTextBox = New-Object System.Windows.Forms.RichTextBox
$SpeakTextBox.location = "10, 15"
$SpeakTextBox.Size = "755, 325"
$SpeakTextBox.Anchor = "Top, Bottom, Left, Right"
$SpeakTextBox.Text = "Hello World"
$speakTextbox.AllowDrop = $true
$speakTextbox.EnableAutoDragDrop = $true
$SpeakTextBox.multiline = $true
$SpeakTextBox.AcceptsTab = $true
$SpeakTextBox.ScrollBars = "both"
$SpeakTextBox.Font = 'Microsoft Sans Serif,10'
$SpeakTextBox.Cursor = "IBeam"
$TextGB.Controls.Add( $SpeakTextBox )

# Select Group Box
#==========================================

$SelectGB = New-Object system.Windows.Forms.Groupbox
$SelectGB.location = "11, 395"
$SelectGB.Size = "640, 50"
$SelectGB.Anchor = "Bottom"
$SelectGB.text = "Select Voice"

$SelectVoiceCB = New-Object system.Windows.Forms.ComboBox
$SelectVoiceCB.location = "11, 15"
$SelectVoiceCB.Size = "618,24"
$SelectVoiceCB.Text = $speak.Voice.Name
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

$Speed = New-Object Windows.Forms.TrackBar
$Speed.Orientation = "Horizontal"
$Speed.location = "5,15"
$Speed.Size = "300,40"
$Speed.TickStyle = "TopLeft"
$Speed.SetRange(-10, 10)
$SpeedGB.Controls.Add( $Speed )

# Volume Group Box
#==========================================

$VolumeGB = New-Object system.Windows.Forms.Groupbox
$VolumeGB.location = "340, 450"
$VolumeGB.Size = "311,62"
$VolumeGB.Anchor = "Bottom"
$VolumeGB.text = "Volume"

$Volume = New-Object Windows.Forms.TrackBar
$Volume.Orientation = "Horizontal"
$Volume.location = "5,15"
$Volume.Size = "300,40"
$Volume.TickStyle = "TopLeft"
$Volume.TickFrequency = 10
$Volume.SetRange(10, 100)
$Volume.Value = 100
$VolumeGB.Controls.Add( $Volume )

# Controls
#==========================================

$Form.controls.AddRange(@( $Menu, $SpeechGB, $SpeakButtion, $SaveButtion, $SelectGB, $SpeedGB, $VolumeGB, $TextGB ))

[void]$form.ShowDialog()
