Add-Type -AssemblyName System.speech
$speak = New-Object System.Speech.Synthesis.SpeechSynthesizer

Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

Function SaveSound {
$SaveChooser = New-Object -TypeName System.Windows.Forms.SaveFileDialog
$SaveChooser.ShowDialog()
$speak.SetOutputToWaveFile($SaveChooser.Filename)
$speak.Speak($SpeakTextBox.Text)
}

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = '723,135'
$Form.text                       = "    Speech Synthesizerer"
$Form.BackColor                  = "#ffffff"
$Form.TopMost                    = $false
$Form.ShowIcon                   = $false
$Form.FormBorderStyle            = "fixeddialog"

$SelectGB                        = New-Object system.Windows.Forms.Groupbox
$SelectGB.height                 = 48
$SelectGB.width                  = 449
$SelectGB.text                   = "Select Voice"
$SelectGB.location               = New-Object System.Drawing.Point(11,13)

$SelectVoiceCB                   = New-Object system.Windows.Forms.ComboBox
$SelectVoiceCB.Text              = $speak.Voice.Name
$SelectVoiceCB.width             = 341
$SelectVoiceCB.height            = 85
$SelectVoiceCB.location          = New-Object System.Drawing.Point(11,15)
$SelectVoiceCB.Font              = 'Microsoft Sans Serif,10'
$Voices = ($speak.GetInstalledVoices() | ForEach-Object {$_.voiceinfo}).Name
foreach($Voice in $Voices)
{
  $SelectVoiceCB.Items.add($voice) | Out-Null
}
$Form.Controls.Add($SelectVoiceCB)

$VoiceSelectButtion              = New-Object system.Windows.Forms.Button
$VoiceSelectButtion.text         = "Select"
$VoiceSelectButtion.width        = 78
$VoiceSelectButtion.height       = 26
$VoiceSelectButtion.location     = New-Object System.Drawing.Point(361,14)
$VoiceSelectButtion.Font         = 'Microsoft Sans Serif,10'
$VoiceSelectButtion.add_Click({$speak.SelectVoice($SelectVoiceCB.Text)})

$SpeechGB                        = New-Object system.Windows.Forms.Groupbox
$SpeechGB.height                 = 48
$SpeechGB.width                  = 698
$SpeechGB.text                   = "Type text here"
$SpeechGB.location               = New-Object System.Drawing.Point(12,73)

$SpeakTextBox                    = New-Object system.Windows.Forms.TextBox
$SpeakTextBox.Text               = "Hello World"
$SpeakTextBox.multiline          = $false
$SpeakTextBox.width              = 500
$SpeakTextBox.height             = 20
$SpeakTextBox.location           = New-Object System.Drawing.Point(12,16)
$SpeakTextBox.Font               = 'Microsoft Sans Serif,10'
$Form.Controls.Add($SpeakTextBox)

$SpeakButtion                    = New-Object system.Windows.Forms.Button
$SpeakButtion.text               = "Speak"
$SpeakButtion.width              = 86
$SpeakButtion.height             = 25
$SpeakButtion.location           = New-Object System.Drawing.Point(515,15)
$SpeakButtion.Font               = 'Microsoft Sans Serif,10'
$SpeakButtion.add_Click({$speak.Speak($SpeakTextBox.Text)})

$SaveButtion                    = New-Object system.Windows.Forms.Button
$SaveButtion.text               = "Save"
$SaveButtion.width              = 86
$SaveButtion.height             = 25
$SaveButtion.location           = New-Object System.Drawing.Point(603,15)
$SaveButtion.Font               = 'Microsoft Sans Serif,10'
$SaveButtion.add_Click({SaveSound})

$SpeedGB.controls.AddRange(@($RadioButtonSpeedNeg10,$RadioButtonSpeedNeg5,$RadioButtonSpeed0,$RadioButtonSpeed5,$RadioButtonSpeed10))
$SelectGB.controls.AddRange(@($SelectVoiceCB,$VoiceSelectButtion))
$Form.controls.AddRange(@($SelectGB,$SpeechGB,$SpeedGB))
$SpeechGB.controls.AddRange(@($SpeakTextBox,$SpeakButtion,$SaveButtion))

[void]$form.ShowDialog()