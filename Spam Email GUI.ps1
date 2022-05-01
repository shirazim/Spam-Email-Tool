<#
Author: Michael Shirazi
Date: 11/20/2021
Description: This script loads a Winform GUI and allows you to send emails to a large number of recipients. This code requires some slight adjustments for it to be useable.
   

#>


$secpasswd = ConvertTo-SecureString $TextBox1.Text -AsPlainText -Force
$cred = New-Object System.Management.Automation.PSCredential (<#Your Gmail#>, $secpasswd)


Function Get-FileName($initialDirectory)
{  
 [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) |
 Out-Null

 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = “All files (*.*)| *.*”
 $OpenFileDialog.ShowDialog() | Out-Null

 $SelectedFile = $OpenFileDialog.FileName
 Return $SelectedFile

} 

function send-email{

    $SendButton.text = "Sending"

    if($AttText.text)
    {
        Send-MailMessage -SmtpServer smtp.gmail.com -Port 587 -UseSsl -From <#Your Gmail#> -To <#Recipient#> -Subject $SubjectText.Text -Body $Bodytext.Text -Credential $cred -Attachments $AttText.Text -Cc $sender
    }
    else{
        Send-MailMessage -SmtpServer smtp.gmail.com -Port 587 -UseSsl -From <#Your Gmail#> -To <#Recipient#> -Subject $SubjectText.Text -Body $Bodytext.Text -Credential $cred
    }


    $SendButton.text = "Sent"

}


Add-Type -AssemblyName System.Windows.Forms
[System.Windows.Forms.Application]::EnableVisualStyles()

$Form                            = New-Object system.Windows.Forms.Form
$Form.ClientSize                 = New-Object System.Drawing.Point(621,530)
$Form.text                       = "Tiem 2 spam"
$Form.TopMost                    = $false

$Label1                          = New-Object system.Windows.Forms.Label
$Label1.text                     = "SPAM EMAIL APPLICATION"
$Label1.AutoSize                 = $true
$Label1.width                    = 25
$Label1.height                   = 10
$Label1.location                 = New-Object System.Drawing.Point(194,33)
$Label1.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',13)
$Label1.ForeColor                = [System.Drawing.ColorTranslator]::FromHtml("#000000")

$Label3                          = New-Object system.Windows.Forms.Label
$Label3.text                     = "Location List"
$Label3.AutoSize                 = $true
$Label3.width                    = 25
$Label3.height                   = 10
$Label3.location                 = New-Object System.Drawing.Point(303,87)
$Label3.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox2                        = New-Object system.Windows.Forms.TextBox
$TextBox2.multiline              = $false
$TextBox2.width                  = 133
$TextBox2.height                 = 20
$TextBox2.location               = New-Object System.Drawing.Point(385,83)
$TextBox2.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label2                          = New-Object system.Windows.Forms.Label
$Label2.text                     = "Password"
$Label2.AutoSize                 = $true
$Label2.width                    = 25
$Label2.height                   = 10
$Label2.location                 = New-Object System.Drawing.Point(23,87)
$Label2.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Label4                          = New-Object system.Windows.Forms.Label
$Label4.text                     = "Subject"
$Label4.AutoSize                 = $true
$Label4.width                    = 25
$Label4.height                   = 10
$Label4.location                 = New-Object System.Drawing.Point(23,131)
$Label4.Font                     = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$TextBox1                        = New-Object system.Windows.Forms.TextBox
$TextBox1.multiline              = $false
$TextBox1.width                  = 137
$TextBox1.height                 = 20
$TextBox1.Text                   = ""
$TextBox1.location               = New-Object System.Drawing.Point(89,83)
$TextBox1.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SubjectText                     = New-Object system.Windows.Forms.TextBox
$SubjectText.multiline           = $false
$SubjectText.width               = 430
$SubjectText.height              = 20
$SubjectText.location            = New-Object System.Drawing.Point(88,128)
$SubjectText.Font                = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$SubjectText.Text                = ""

$BodyLabel                       = New-Object system.Windows.Forms.Label
$BodyLabel.text                  = "Body"
$BodyLabel.AutoSize              = $true
$BodyLabel.width                 = 25
$BodyLabel.height                = 10
$BodyLabel.location              = New-Object System.Drawing.Point(23,175)
$BodyLabel.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$Bodytext                        = New-Object system.Windows.Forms.TextBox
$Bodytext.multiline              = $true
$Bodytext.width                  = 430
$Bodytext.height                 = 262
$Bodytext.location               = New-Object System.Drawing.Point(88,174)
$Bodytext.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$Bodytext.Text                   = ""

$AttLabel                        = New-Object system.Windows.Forms.Label
$AttLabel.text                   = "Att"
$AttLabel.AutoSize               = $true
$AttLabel.width                  = 25
$AttLabel.height                 = 10
$AttLabel.location               = New-Object System.Drawing.Point(23,459)
$AttLabel.Font                   = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$AttText                         = New-Object system.Windows.Forms.TextBox
$AttText.multiline               = $false
$AttText.width                   = 205
$AttText.height                  = 20
$AttText.location                = New-Object System.Drawing.Point(88,456)
$AttText.Font                    = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$SendButton                      = New-Object system.Windows.Forms.Button
$SendButton.text                 = "Send"
$SendButton.width                = 150
$SendButton.height               = 33
$SendButton.location             = New-Object System.Drawing.Point(372,452)
$SendButton.Font                 = New-Object System.Drawing.Font('Microsoft Sans Serif',10)
$SendButton.add_click({send-email})

$AttButton                       = New-Object system.Windows.Forms.Button
$AttButton.text                  = ":::"
$AttButton.add_click({$AttText.text = Get-FileName})
$AttButton.width                 = 30
$AttButton.height                = 20
$AttButton.location              = New-Object System.Drawing.Point(295,459)
$AttButton.Font                  = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

$LocationButton                  = New-Object system.Windows.Forms.Button
$LocationButton.add_click({$TextBox2.text = Get-FileName})
$LocationButton.text             = ":::"
$LocationButton.width            = 21
$LocationButton.height           = 14
$LocationButton.location         = New-Object System.Drawing.Point(520,89)
$LocationButton.Font             = New-Object System.Drawing.Font('Microsoft Sans Serif',10)

<#$SuccessLabel                    = New-Object system.Windows.Forms.Label
$SuccessLabel.text               = ""
$SuccessLabel.AutoSize           = $true
$SuccessLabel.width              = 25
$SuccessLabel.height             = 10
$SuccessLabel.location           = New-Object System.Drawing.Point(209,502)
$SuccessLabel.Font               = New-Object System.Drawing.Font('Microsoft Sans Serif',11)
$SuccessLabel.ForeColor          = [System.Drawing.ColorTranslator]::FromHtml("#31500a")
#>

$Form.controls.AddRange(@($Label1,$Label3,$TextBox2,$Label2,$TextBox1,$Label4, $SubjectText, $BodyLabel, $Bodytext, $AttLabel, $AttText, $SendButton, $AttButton, $LocationButton))




#Write your logic code here

[void]$Form.ShowDialog()