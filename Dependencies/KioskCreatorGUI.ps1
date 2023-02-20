# Kiosk Creator Script
# ~Script By SPC Burgess & SPC Santiago 02/16/2023
# MOS: 25B & 25U
<#
#####################################################
    Big thanks to Reddit Friends / Sources
 for making this script possible. The goal here
 is to make things easier for IMO's. If you get
 a moment feel free to check out this code. If 
 I am still in the Army apon you reading this,
 feel free to reach out with any feedback. 
       PURGED OF ALL CUI DATA FOR PUBLIC USE.
#####################################################
#>

# Once you learn to use PSEXEC + Powershell everything else falls into place.. -SPC BURGESS




Function KioskCreator {
    cls
    Sleep 1
    Write-Host "Kiosk Creator Ready For User Interaction" -ForegroundColor Green
    Sleep 1
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.AnchorStyles")
    [void][System.Windows.Forms.Application]::EnableVisualStyles()


    # Draw Form
    $ZipExtractorForm = New-Object System.Windows.Forms.Form
    $ZipExtractorForm.Text = "[ADMIN] Kiosk Creator v4.0"
    $ZipExtractorForm.ClientSize = New-Object System.Drawing.Size(400, 185)
    $ZipExtractorForm.BackColor = "LightGray"
    $ZipExtractorForm.StartPosition = "CenterScreen"
    $StopScript = $Script:CANCELED=$True
    $ZipExtractorForm.MaximizeBox = $false
    $ZipExtractorForm.AccessibleDescription = "Simple GUI based PS Script to create local user accounts."

    # Draw Icon
    $iconConverted2Base64 = [Convert]::ToBase64String((Get-Content "C:\temp\Launcher\Dependencies\icon\NewPanda.ico" -Encoding Byte))
    $iconBase64           = $iconConverted2Base64
    $iconBytes            = [Convert]::FromBase64String($iconBase64)
    $stream               = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    $iconImage            = [System.Drawing.Image]::FromStream($stream, $true)
    $ZipExtractorForm.Icon    = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
    # ico converter : https://cloudconvert.com/png-to-ico

    # This is defines what Enter does when pressed
    $ZipExtractorForm.KeyPreview = $True
    $ZipExtractorForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {
        # Currently Nothing..

    }})

    # creates the label title
    $titleLabel = New-Object System.Windows.Forms.Label
    $titleLabel.Location = New-Object System.Drawing.Point(190, 25)
    $titleLabel.Size = New-Object System.Drawing.Size(198, 20)
    $titleLabel.ForeColor = "Black"
    $titleLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $titleLabel.Font = New-Object System.Drawing.Font("Lucida Console",14,[System.Drawing.FontStyle]::Regular)
    $titleLabel.Text = "Kiosk Script V4"
    $ZipExtractorForm.Controls.Add($titleLabel)

    # This creates the TextBox for password input
    $passwordTextBox = New-Object System.Windows.Forms.TextBox 
    $passwordTextBox.Location = New-Object System.Drawing.Point(190,114)
    $passwordTextBox.Size = New-Object System.Drawing.Size(160,20) 
    $passwordTextBox.UseSystemPasswordChar = $true
    $ZipExtractorForm.Controls.Add($passwordTextBox)

    # Password input label.
    $PasswordInputlabel = New-Object System.Windows.Forms.Label
    $PasswordInputlabel.Location = New-Object System.Drawing.Point(190, 100)
    $PasswordInputlabel.Size = New-Object System.Drawing.Size(198,20)
    $PasswordInputlabel.ForeColor = [System.Drawing.Color]::FromKnownColor("Red")
    [String]$MandatoryWrite = "*" 
    $PasswordInputlabel.Text = "Enter Account Password $MandatoryWrite"
    $ZipExtractorForm.Controls.Add($PasswordInputlabel)

    # Password input example.
    $PasswordInputlabelexample = New-Object System.Windows.Forms.Label
    $PasswordInputlabelexample.Location = New-Object System.Drawing.Point(190, 134)
    $PasswordInputlabelexample.Size = New-Object System.Drawing.Size(198,20)
    $PasswordInputlabelexample.ForeColor = [System.Drawing.Color]::FromKnownColor("Blue")
    $PasswordInputlabelexample.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    [String]$MandatoryWrite = "*" 
    $PasswordInputlabelexample.Text = "Example: S6H3lpD3sk!@#$"
    $ZipExtractorForm.Controls.Add($PasswordInputlabelexample)

    # Expiration Check Box 
    $expirationCheckbox = New-Object System.Windows.Forms.CheckBox
    $expirationCheckbox.Location = New-Object System.Drawing.Point(190,50)
    $expirationCheckbox.Size = New-Object System.Drawing.Size(198,20)
    $expirationCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $expirationCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $expirationCheckbox.Text = "Account Never Expires"
    $ZipExtractorForm.Controls.Add($expirationCheckbox)
    
    # Password Change check box
    $PasswordChangeCheckbox = New-Object System.Windows.Forms.CheckBox
    $PasswordChangeCheckbox.Location = New-Object System.Drawing.Point(190,70)
    $PasswordChangeCheckbox.size = New-Object System.Drawing.Size(198,20)
    $PasswordChangeCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $PasswordChangeCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $PasswordChangeCheckbox.Text = "Can Change Password"
    $ZipExtractorForm.Controls.Add($PasswordChangeCheckbox) 

    #This creates the TextBox for hostname / ip addr
    $objZipExtractorTextBox1 = New-Object System.Windows.Forms.TextBox 
    $objZipExtractorTextBox1.Location = New-Object System.Drawing.Size(10,33) 
    $objZipExtractorTextBox1.Size = New-Object System.Drawing.Size(160,20)
    $objZipExtractorTextBox1.TabIndex = 0 
    $ZipExtractorForm.Controls.Add($objZipExtractorTextBox1)

    #This creates a label for the objLabelZipExtractor
    $objLabelZipExtractorHN = New-Object System.Windows.Forms.Label
    $objLabelZipExtractorHN.Location = New-Object System.Drawing.Size(10,20) 
    $objLabelZipExtractorHN.Size = New-Object System.Drawing.Size(120,20)
    [String]$MandatoryWrite = "*" 
    $objLabelZipExtractorHN.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $objLabelZipExtractorHN.ForeColor = [System.Drawing.Color]::FromName("Red")
    $objLabelZipExtractorHN.Text = "Enter New Hostname $MandatoryWrite"
    $ZipExtractorForm.Controls.Add($objLabelZipExtractorHN)

    #DisplayNameBox
    $DisplayNameBox = New-Object System.Windows.Forms.TextBox 
    $DisplayNameBox.Location = New-Object System.Drawing.Size(10,66) 
    $DisplayNameBox.Size = New-Object System.Drawing.Size(160,20)
    $DisplayNameBox.TabIndex = 1 
    $ZipExtractorForm.Controls.Add($DisplayNameBox)

    
    #This creates a label for the File PAth Label
    $objLabelZipExtractorPathlabel = New-Object System.Windows.Forms.Label
    $objLabelZipExtractorPathlabel.Location = New-Object System.Drawing.Size(10,53) 
    $objLabelZipExtractorPathlabel.Size = New-Object System.Drawing.Size(130,20)
    [String]$MandatoryWrite = "*" 
    $objLabelZipExtractorPathlabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $objLabelZipExtractorPathlabel.ForeColor = [System.Drawing.Color]::FromName("Red")
    $objLabelZipExtractorPathlabel.Text = "Enter Display Name $MandatoryWrite"
    $ZipExtractorForm.Controls.Add($objLabelZipExtractorPathlabel)

    # ExpirationBox
    $ExpirationBox = New-Object System.Windows.Forms.TextBox 
    $ExpirationBox.Location = New-Object System.Drawing.Size(10,114) 
    $ExpirationBox.Size = New-Object System.Drawing.Size(160,20)
    $ExpirationBox.TabIndex = 2 
    $ZipExtractorForm.Controls.Add($ExpirationBox)

    #This creates a label for the file name label
    $objLabelZipExtractorfilenamelabel = New-Object System.Windows.Forms.Label
    $objLabelZipExtractorfilenamelabel.Location = New-Object System.Drawing.Size(10,101) 
    $objLabelZipExtractorfilenamelabel.Size = New-Object System.Drawing.Size(120,20)
    [String]$MandatoryWrite = "*" 
    $objLabelZipExtractorfilenamelabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $objLabelZipExtractorfilenamelabel.ForeColor = [System.Drawing.Color]::FromName("Red")
    $objLabelZipExtractorfilenamelabel.Text = "Expiration Date $MandatoryWrite"
    $ZipExtractorForm.Controls.Add($objLabelZipExtractorfilenamelabel)

    #This creates a label for the File PAth Label example
    $objLabelZipExtractorPathexample = New-Object System.Windows.Forms.Label
    $objLabelZipExtractorPathexample.Location = New-Object System.Drawing.Size(10,86) 
    $objLabelZipExtractorPathexample.Size = New-Object System.Drawing.Size(300,20) 
    $objLabelZipExtractorPathexample.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $objLabelZipExtractorPathexample.ForeColor = [System.Drawing.Color]::FromName("Blue")
    $objLabelZipExtractorPathexample.Text = "Example: 2-3 FA S6 Kiosk"
    $ZipExtractorForm.Controls.Add($objLabelZipExtractorPathexample)
    
    #This creates a label for the File name example.
    $objLabelZipExtractorPathexample = New-Object System.Windows.Forms.Label
    $objLabelZipExtractorPathexample.Location = New-Object System.Drawing.Size(10,134) 
    $objLabelZipExtractorPathexample.Size = New-Object System.Drawing.Size(300,20) 
    $objLabelZipExtractorPathexample.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $objLabelZipExtractorPathexample.ForeColor = [System.Drawing.Color]::FromName("Blue")
    $objLabelZipExtractorPathexample.Text = "Example: mm/dd/yyyy"
    $ZipExtractorForm.Controls.Add($objLabelZipExtractorPathexample)

    

    # Start Button
    $ButtonStart = New-Object System.Windows.Forms.Button
    $ButtonStart.Location = New-Object System.Drawing.Point(2, 160)
    $ButtonStart.Size = New-Object System.Drawing.Size(98, 23)
    $ButtonStart.BackColor = "LightGray"
    $ButtonStart.Text = "SEND IT"
    $ButtonStart.Add_Click({
    cls
    # Set Global Vars
    $HNorIPV4=$objZipExtractorTextBox1.Text
    $Script:CANCELED=$False

    # Continue on
    If ($Script:CANCELED -eq $True) {
        Write-Host "`nScript Canceled" -ForegroundColor Yellow
        Sleep 1
        Write-Host "`nScript Exiting Current Process" -ForegroundColor Yellow
        Sleep 2
        Write-Host "`nAll Modules Exited" -ForegroundColor Green
    }
    else {
        If ($HNorIPV4 -and $null) {
            Write-Host "`nNo Hostname entered" -ForegroundColor Yellow
            Sleep 1
            Write-Host "`nScript Exiting Current Process" -ForegroundColor Yellow
            Sleep 2
            Write-Host "`nAll Modules Exited" -ForegroundColor Green
            $Script:CANCELED=$True
        }
        else {
            [String]$Computers = $HNorIPV4 # Get-InputBox "ZIP EXTRACTOR v1.0" "Enter Destination Computer Name"
            cls
            $Scriptpasswordbox 
            Write-Host "************************************************************`n`n              Script Created By SPC Burgess              `n`n************************************************************" -ForegroundColor Cyan      
            $ExpirationDate = $ExpirationBox.Text
            $DisplayName    = $DisplayNameBox.Text
            $TargetComputer = $HNorIPV4
            $THEIPADDRESSONLY = Test-Connection -ComputerName $TargetComputer -Count 1 -Protocol WSMan -ErrorAction Ignore | Format-Wide IPV4Address # Format-Wide gives the object by itself.. perfect.. ~SPC BURGESS
            If ($THEIPADDRESSONLY -cne $null) { 

                Invoke-Command -ComputerName $TargetComputer -ScriptBlock { # WinRM Required for this service... Do not need PSEXEC anymore boys wooohooo!! ~SPC BURGESS

                Write-Host "`nRemote Host is Online!" -ForegroundColor Green -BackgroundColor Black
      
                Write-Host "`nSetting up Kiosk Account on remote host" -ForegroundColor Cyan -BackgroundColor Black

                If ($passwordTextBox.Text -eq $null) {
                    $pt = [System.Text.Encoding]::UTF8.GetString([System.Convert]::FromBase64String("J1M2SDNscEQzc2shQCMkJw=="))
                    [System.Security.SecureString]$SecurePassword = ConvertTo-SecureString -String $pt -AsPlainText -Force
                    New-LocalUser -Name "Kiosk" -Password $SecurePassword
                }
                else {
                    $Password               = $passwordTextBox.Text
                    [System.Security.SecureString]$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force
                    New-LocalUser -Name "Kiosk" -Password $SecurePassword
                }
                
                Write-Host "`nCreating account..." -ForegroundColor Cyan -BackgroundColor Black
                
                If ($PasswordChangeCheckbox.Checked) {
                    Set-LocalUser -Name "Kiosk" -UserMayChangePassword $True
                }
                else {
                    Set-LocalUser -Name "Kiosk" -UserMayChangePassword $False
                }
                
                Write-Host "`nApplying Permissions to account on remote host" -ForegroundColor Cyan -BackgroundColor Black
                Add-LocalGroupMember -Group Users -Member Kiosk -ErrorAction Ignore
                
                If ($ExpirationDate -cne $null) { 
                    Set-LocalUser -Name "Kiosk" -AccountExpires $ExpirationDate
                }
                elseif ($expirationCheckbox.Checked -eq $True) {
                    Set-LocalUser -Name "Kiosk" -AccountNeverExpires
                }
                else {
                    Write-Host "`nAccount Expiration Set!" -BackgroundColor Black -ForegroundColor Green
                }

                If ($DisplayName -cne $null) {
                    Set-LocalUser -Name "Kiosk" -FullName $DisplayName
                    Write-Host "`nCustom Display Name Set" -ForegroundColor Cyan -BackgroundColor Black
                }
                else {
                    Set-LocalUser -Name "Kiosk" -FullName "KIOSK"
                    Write-Host "`nDefault Display Name Set" -ForegroundColor Green -BackgroundColor Black
                }

                Write-Host "`nFinished setting up Kiosk on remote host" -ForegroundColor Green -BackgroundColor Black

                Enable-LocalUser -Name "Kiosk"

                $users = Get-LocalUser | Select-Object Name | Out-String

                Write-Host "`nRemote User Accounts:`n$users" -ForegroundColor Green

                }
           }
           else {
            Write-Host "`nCould not reach host`nWinRM Service failed to connect.." -BackgroundColor Black -ForegroundColor Red
        }
     }
  }
})
$ZipExtractorForm.Controls.Add($ButtonStart)



    # Cancel Button
    $ButtonStop = New-Object System.Windows.Forms.Button
    $ButtonStop.Location = New-Object System.Drawing.Point(102, 160)
    $ButtonStop.Size = New-Object System.Drawing.Size(98, 23)
    $ButtonStop.BackColor = "LightGray"
    $ButtonStop.Text = "CANCEL"
    $ButtonStop.Add_Click({write-Host "`nScript Canceled" -ForegroundColor Yellow;Sleep 1;write-host "`nScript Closing Safely" -ForegroundColor Cyan;Sleep 1;$Script:CANCELED=$True;cls;$ZipExtractorForm.Close()})
    $ZipExtractorForm.Controls.Add($ButtonStop)

    # Finally Render it all.
    $ZipExtractorForm.Add_Shown({$ZipExtractorForm.Activate()})
    $ZipExtractorForm.ShowDialog() | Out-Null
    $ZipExtractorForm.Dispose() | Out-Null

}

KioskCreator
