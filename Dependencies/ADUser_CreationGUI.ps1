# ADUser Creation Script
# ~Script By SPC Burgess & SPC Santiago 2-3 FA S6 03/02/2021
# MOS: 25B
<#
#####################################################
    Big thanks to Reddit Friends / Sources
 for making this script possible. The goal here
 is to make things easier for IMO's. If you get
 a moment feel free to check out this code. If 
 I am still in the Army apon you reading this,
 feel free to reach out with any feedback. 

            Contact DSN: 915-741-4627
#####################################################
#>
Function UserCreator{
cls
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[System.Windows.Forms.Application]::EnableVisualStyles()
Sleep 1
Import-Module ActiveDirectory | Out-Null
Sleep 1
try { "C:\temp\Launcher\Dependencies\Get-SmartCardCredential.ps1" }
catch { Write-Host "Get-SmartCardCredential.ps1 Not Found" -ForegroundColor Yellow }
cls
Sleep 1
Write-Host "All Active Directory Modules Loaded" -ForegroundColor Cyan
Sleep 1
Write-Host "Script Ready For User Interaction" -ForegroundColor Green
Sleep 1
#creates window
$ADForm = New-Object System.Windows.Forms.Form
$ADForm.Text = 'AD User Creation [US Military Edition]'
$ADForm.Width = 600
$ADForm.Height = 400
$ADForm.BackColor = "White"
$ADForm.StartPosition = "CenterScreen"
$ADForm.Location = New-Object System.Drawing.Size(80,495)
$ADForm.FormBorderStyle = 'Fixed3D'
$ADForm.MaximizeBox = $false

$iconConverted2Base64 = [Convert]::ToBase64String((Get-Content "C:\temp\Launcher\Dependencies\icon\NewPanda.ico" -Encoding Byte))
$iconBase64           = $iconConverted2Base64
$iconBytes            = [Convert]::FromBase64String($iconBase64)
$stream               = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage            = [System.Drawing.Image]::FromStream($stream, $true)
$ADForm.Icon    = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
# ico converter : https://cloudconvert.com/png-to-ico

# Draws Logo
$img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\Panda\NewPanda.png')
$ADForm.BackgroundImage = $img
$ADForm.BackgroundImageLayout = "Center"

#***********************************************************************
#This creates a label for the TextBox Last Name
$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(20,20) 
$objLabel1.Size = New-Object System.Drawing.Size(280,20)
[String]$MandatoryWrite = "*" 
$objLabel1.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel1.Text = "Last Name"
$ADForm.Controls.Add($objLabel1) 

#This creates the TextBox Last Name
$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(20,40) 
$objTextBox1.Size = New-Object System.Drawing.Size(260,20)
$objTextBox1.TabIndex = 0 
$ADForm.Controls.Add($objTextBox1)
#***************************************************************************
#This creates a label for the TextBox First Name
$objLabel2 = New-Object System.Windows.Forms.Label
$objLabel2.Location = New-Object System.Drawing.Size(300,20) 
$objLabel2.Size = New-Object System.Drawing.Size(280,20)
$objLabel2.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel2.Text = "First Name"
$ADForm.Controls.Add($objLabel2) 

#This creates the TextBox Firstname
$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(300,40) 
$objTextBox2.Size = New-Object System.Drawing.Size(260,20)
$objTextBox2.TabIndex = 1 
$ADForm.Controls.Add($objTextBox2)
#*************************************************************************
#This creates a label for the TextBox Middle Initial
$objLabel3 = New-Object System.Windows.Forms.Label
$objLabel3.Location = New-Object System.Drawing.Size(20,70) 
$objLabel3.Size = New-Object System.Drawing.Size(75,20)
[String]$MandatoryWrite = "*" 
$objLabel3.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel3.Text = "Middle Initial"
$ADForm.Controls.Add($objLabel3) 

#This creates the TextBox Middle Initial
$objTextBox3 = New-Object System.Windows.Forms.TextBox 
$objTextBox3.Location = New-Object System.Drawing.Size(20,90) 
$objTextBox3.Size = New-Object System.Drawing.Size(100,20)
$objTextBox3.TabIndex = 2 
$ADForm.Controls.Add($objTextBox3)
#***************************************************************************
#This creates a label for the TextBox DODID
$objLabel4 = New-Object System.Windows.Forms.Label
$objLabel4.Location = New-Object System.Drawing.Size(150,70) 
$objLabel4.Size = New-Object System.Drawing.Size(100,20)# (280,20)
[String]$MandatoryWrite = "*" 
$objLabel4.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel4.Text = "DODID"
$ADForm.Controls.Add($objLabel4) 

#This creates the TextBox DODID
$objTextBox4 = New-Object System.Windows.Forms.TextBox 
$objTextBox4.Location = New-Object System.Drawing.Size(150,90) 
$objTextBox4.Size = New-Object System.Drawing.Size(130,20)
$objTextBox4.TabIndex = 3 
$ADForm.Controls.Add($objTextBox4)
#***************************************************************************
#This creates a label for the TextBox Email
$objLabel5 = New-Object System.Windows.Forms.Label
$objLabel5.Location = New-Object System.Drawing.Size(300,70) 
$objLabel5.Size = New-Object System.Drawing.Size(280,20)
[String]$MandatoryWrite = "*" 
$objLabel5.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel5.Text = "Email"
$ADForm.Controls.Add($objLabel5) 

#This creates the TextBox Email
$objTextBox5 = New-Object System.Windows.Forms.TextBox 
$objTextBox5.Location = New-Object System.Drawing.Size(300,90) 
$objTextBox5.Size = New-Object System.Drawing.Size(260,20)
$objTextBox5.TabIndex = 4 
$ADForm.Controls.Add($objTextBox5)
#***************************************************************************
#This creates a label for the TextBox Rank
$objLabel6 = New-Object System.Windows.Forms.Label
$objLabel6.Location = New-Object System.Drawing.Size(20,120) 
$objLabel6.Size = New-Object System.Drawing.Size(45,20)
[String]$MandatoryWrite = "*" 
$objLabel6.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel6.Text = "Rank"
$ADForm.Controls.Add($objLabel6) 

#This creates the TextBox Rank
$objTextBox6 = New-Object System.Windows.Forms.TextBox 
$objTextBox6.Location = New-Object System.Drawing.Size(20,140) 
$objTextBox6.Size = New-Object System.Drawing.Size(100,20)
$objTextBox6.TabIndex = 5 
$ADForm.Controls.Add($objTextBox6)
#***************************************************************************
#This creates a label for the TextBox DSN
$objLabel7 = New-Object System.Windows.Forms.Label
$objLabel7.Location = New-Object System.Drawing.Size(150,120) 
$objLabel7.Size = New-Object System.Drawing.Size(45,20)
[String]$MandatoryWrite = "*" 
$objLabel7.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel7.Text = "DSN"
$ADForm.Controls.Add($objLabel7)

#This creates the TextBox DSN
$objTextBox7 = New-Object System.Windows.Forms.TextBox 
$objTextBox7.Location = New-Object System.Drawing.Size(150,140) 
$objTextBox7.Size = New-Object System.Drawing.Size(130,20)
$objTextBox7.TabIndex = 6 
$ADForm.Controls.Add($objTextBox7)
#***************************************************************************
#This creates a label for the office 
$objLabel8 = New-Object System.Windows.Forms.Label
$objLabel8.Location = New-Object System.Drawing.Size(20,170) 
$objLabel8.Size = New-Object System.Drawing.Size(75,20) 
$objLabel8.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel8.Text = "Office"
$ADForm.Controls.Add($objLabel8)

#This creates the Textbox for office
$objTextBox8 = New-Object System.Windows.Forms.TextBox
$objTextBox8.Location = New-Object System.Drawing.Point(20,190)
$objTextBox8.Size = New-Object System.Drawing.Size(260,20)
$objTextBox8.TabIndex = 6
$ADForm.Controls.Add($objTextBox8)
#******************************************************************************
#This creates a label for the TextBox Path
$objLabel9 = New-Object System.Windows.Forms.Label
$objLabel9.Location = New-Object System.Drawing.Size(20,220) 
$objLabel9.Size = New-Object System.Drawing.Size(45,20)
$objLabel9.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel9.Text = "Path"
$ADForm.Controls.Add($objLabel9)

#This creates the TextBox Path
$objTextBox9 = New-Object System.Windows.Forms.TextBox 
$objTextBox9.Location = New-Object System.Drawing.Size(20,240) 
$objTextBox9.Size = New-Object System.Drawing.Size(540,20)
$objTextBox9.TabIndex = 6 # remember to fix tabs
$ADForm.Controls.Add($objTextBox9)

#This creates a label for the Example path
$objLabel10 = New-Object System.Windows.Forms.Label
$objLabel10.Location = New-Object System.Drawing.Size(20,270) 
$objLabel10.Size = New-Object System.Drawing.Size(600,40) 
$objLabel10.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel10.Text = "Example: 'OU=HHB,OU=Laptops,OU=2-3 FA,OU=1-1 Armored Div,OU=1st Armor Div,OU=FORSCOM,OU=Bliss,OU=Installations,DC=nasw,DC=ds,DC=army,DC=mil'"
$ADForm.Controls.Add($objLabel10)

#****************************************************************************
#Check box section
#****************************************************************************

#This creates a checkbox 
$objChangePWCheckbox = New-Object System.Windows.Forms.Checkbox 
$objChangePWCheckbox.Location = New-Object System.Drawing.Size(300,140) 
$objChangePWCheckbox.Size = New-Object System.Drawing.Size(300,20)
$objChangePWCheckbox.ForeColor = [System.Drawing.Color]::FromName("Black")
$objChangePWCheckbox.Text = "User Must Change Password on next login"
$objChangePWCheckbox.TabIndex = 7
$ADForm.Controls.Add($objChangePWCheckbox)

#This creates a checkbox 
$objNoChgPWCheckbox = New-Object System.Windows.Forms.Checkbox 
$objNoChgPWCheckbox.Location = New-Object System.Drawing.Size(300,160) 
$objNoChgPWCheckbox.Size = New-Object System.Drawing.Size(300,20)
$objNoChgPWCheckbox.ForeColor = [System.Drawing.Color]::FromName("Black")
$objNoChgPWCheckbox.Text = "User cannot change password"
$objNoChgPWCheckbox.TabIndex = 8
$ADForm.Controls.Add($objNoChgPWCheckbox)

#This creates a checkbox 
$objNoPWExpCheckbox = New-Object System.Windows.Forms.Checkbox 
$objNoPWExpCheckbox.Location = New-Object System.Drawing.Size(300,180) 
$objNoPWExpCheckbox.Size = New-Object System.Drawing.Size(300,20)
$objNoPWExpCheckbox.ForeColor = [System.Drawing.Color]::FromName("Black")
$objNoPWExpCheckbox.Text = "Password never expires"
$objNoPWExpCheckbox.TabIndex = 9
$ADForm.Controls.Add($objNoPWExpCheckbox)

#This creates a checkbox 
$objAccountDisabledCheckbox = New-Object System.Windows.Forms.Checkbox 
$objAccountDisabledCheckbox.Location = New-Object System.Drawing.Size(300,200) 
$objAccountDisabledCheckbox.Size = New-Object System.Drawing.Size(300,20)
$objAccountDisabledCheckbox.ForeColor = [System.Drawing.Color]::FromName("Black")
$objAccountDisabledCheckbox.Text = "Account is disabled"
$objAccountDisabledCheckbox.TabIndex = 10
$ADForm.Controls.Add($objAccountDisabledCheckbox)

#******************************************************************************
#Buttons
#******************************************************************************

#This Creates Button Create
$CREATEButton = New-Object System.Windows.Forms.Button
$CREATEButton.Location = New-Object System.Drawing.Size(225,310)
$CREATEButton.Size = New-Object System.Drawing.Size(75,23)
$CREATEButton.BackColor = "LightGray"
$CREATEButton.Text = "CREATE"
$CREATEButton.Anchor = 'right,bottom'
$CREATEButton.Add_Click({$Script:CANCELED=$False;$ADForm.Close()})
$ADForm.Controls.Add($CREATEButton)

#This Creates Button Cancel
$CANCELButton = New-Object System.Windows.Forms.Button
$CANCELButton.Location = New-Object System.Drawing.Size(300,310)
$CANCELButton.Size = New-Object System.Drawing.Size(75,23)
$CANCELButton.BackColor = "LightGray"
$CANCELButton.Anchor = 'right,bottom'
$CANCELButton.Text = "CANCEL"
$CANCELButton.Add_Click({$Script:CANCELED=$True;$ADForm.Close()})
$ADForm.Controls.Add($CANCELButton)

#This creates a label for the Credits
$objLabel11 = New-Object System.Windows.Forms.Label
$objLabel11.Location = New-Object System.Drawing.Size(220,340) 
$objLabel11.Size = New-Object System.Drawing.Size(200,20)
$objLabel11.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel11.Text = "By: SPC Burgess, SPC Santiago"
$ADForm.Controls.Add($objLabel11) 

###### FONT SIZE CHANGE:
$objLabel4.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel4.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")

$ADForm.TopMost = $True # Set Window to open in front of all apps.
$ADForm.Add_Shown({$ADForm.Activate()})
$ADForm.ShowDialog() | Out-Null

# Back End here

$Fullname = '$Lastname $Firstname'
$Displayname = '$Fullname $Rank MIL USA FORSCOM'

If ($objTextBox1.Text -cne $null) { $Lastname = $objTextBox1.Text }
else { Write-Host "must have last name to add user to AD"; $DontContinue=$True }

If ($objTextBox2.Text -cne $null) { $Firstname = $objTextBox2.Text }
else { Write-Host "must have first name to add user to AD"; $DontContinue=$True }

If ($objTextBox3.Text -cne $null) { $Middle = $objTextBox3.Text }
else { Write-Host "must have middle initial to add user to AD"; $DontContinue=$True }

$Samaccountname = $objTextBox4.Text
$AddMIL = '.mil'
$FinalSamName = '$Samaccountname$AddMIL'

If ($objTextBox4.Text -cne $null) { $Samaccountname=$objTextBox1.Text;$AddMIL='.mil';$FinalSamName='$Samaccountname$AddMIL' }
else { Write-Host "must have DODID to add user to AD"; $DontContinue=$True }

If ($objTextBox5.Text -cne $null) { $Email = $objTextBox5.Text }
else { Write-Host "must have Email to add user to AD"; $DontContinue=$True }

If ($objTextBox6.Text -cne $null) { $Rank = $objTextBox6.Text }
else { Write-Host "must have Rank to add user to AD"; $DontContinue=$True }

If ($objTextBox7.Text -cne $null) { $DSN = $objTextBox7.Text }
else { Write-Host "must have Rank to add user to AD"; $DontContinue=$True }

If ($objTextBox8.Text -cne $null) { $Office = $objTextBox8.Text }
else { Write-Host "must have office to add user to AD"; $DontContinue=$True }

If ($objTextBox9.Text -cne $null) { $paath = $objTextBox9.Text }
else { Write-Host "must have DSN to add user to AD"; $DontContinue=$True }

If ($objChangePWCheckbox.Checked -cne $null) { $MustChPW = $objChangePWCheckbox.checked }
else { $MustChPW = $False;  }

If ($objNoChgPWCheckbox.Checked -cne $null) { $objNoChgPW = $objNoChgPWCheckbox.checked }
else { $objNoChgPW = $False;  }

If ($objNoPWExpCheckbox.Checked -cne $null) { $objNoPWExp = $objNoPWExpCheckbox.checked }
else { $objNoPWExp = $False;  }

If ($objAccountDisabledCheckbox.Checked -cne $null) { $objAccountDisabled = $objAccountDisabledCheckbox.checked }
else { $objAccountDisabled = $False;  }

If ($Script:CANCELED -eq $True) { Write-Host "`nStopping Script" -ForegroundColor Yellow; Sleep 1; return "`nFinished Unloading All Active Directory Modules"; }
else {
$DontContinue = $True
If ($DontContinue -eq $False) {
    Write-Host "Somethings missing"
}
else { 
    #try { "C:\temp\Launcher\Dependencies\Get-SmartCardCredential.ps1" }
    #catch { Write-host "Failed to locate File Get-SmartCardCredential" -ForegroundColor Yellow }
    #$TheCredential = Get-SmartCardCert
    New-ADUser -Name $Fullname -GivenName $Firstname -Surname $Lastname -SamAccountName $FinalSamName -UserPrincipalName ($FinalSamName + '@' + $AddMIL) -CannotChangePassword $objNoChgPW -PasswordNeverExpires $objNoPWExp -Path $paath -ChangePasswordAtLogon $MustChPW -Enabled $MustChPW -DisplayName $Displayname -Office $Office -OfficePhone $DSN -EmailAddress $Email -Credential $TheCredential
  }
 }
}
UserCreator

