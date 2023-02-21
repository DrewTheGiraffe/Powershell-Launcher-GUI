# Launcher Script v3.0
# ~Script By SPC Burgess & SPC Santiago 02/16/2023
<#
#####################################################
    Big thanks to Reddit Friends / Sources
 for making this script possible. The goal here
 is to make things easier for IMO's. If you get
 a moment feel free to check out this code. If 
 I am still in the Army apon you reading this,
 feel free to reach out with any feedback. 
       ALL CUI DATA PURGED FROM THIS SOURCE.
#####################################################
#>

Function MainLauncher {
Add-Type -AssemblyName System
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
[void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
Add-Type -AssemblyName System.Security
[System.Windows.Forms.Application]::EnableVisualStyles()
[Bool]$Setupbackend

$LauncherForm = New-Object System.Windows.Forms.Form
$LauncherForm.Text = "Script Launcher"
$LauncherForm.ClientSize = New-Object System.Drawing.Size(407, 200)
$LauncherForm.BackColor = "White"
$LauncherForm.StartPosition = "CenterScreen"
$LauncherForm.FormBorderStyle = 'Fixed3D'
$LauncherForm.MaximizeBox = $false

$iconConverted2Base64 = [Convert]::ToBase64String((Get-Content "C:\temp\Launcher\Dependencies\icon\NewPanda.ico" -Encoding Byte))
$iconBase64           = $iconConverted2Base64
$iconBytes            = [Convert]::FromBase64String($iconBase64)
$stream               = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage            = [System.Drawing.Image]::FromStream($stream, $true)
$LauncherForm.Icon    = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
# ico converter : https://cloudconvert.com/png-to-ico

$img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\Panda\NewPanda.png')
$LauncherForm.BackgroundImage = $img
$LauncherForm.BackgroundImageLayout = "Center"

$LaunchButton = New-Object System.Windows.Forms.Button
$LaunchButton.Location = New-Object System.Drawing.Point(26, 40)
$LaunchButton.Size = New-Object System.Drawing.Size(350, 23)
$LaunchButton.BackColor = "LightGray"
$LaunchButton.ForeColor = "Black"
$LaunchButton.Text = "LAUNCH"
$LaunchButton.add_Click({$Script:CANCELED=$False;$info=$comboBox1.SelectedItem;$LauncherForm.Close()})
$LauncherForm.Controls.Add($LaunchButton)

$UpdateButton = New-Object System.Windows.Forms.Button
$UpdateButton.Location = New-Object System.Drawing.Point(69, 63)
$UpdateButton.Size = New-Object System.Drawing.Size(265.5, 23)
$UpdateButton.BackColor = "LightGray"
$UpdateButton.ForeColor = "Black"
$UpdateButton.Text = "CHECK FOR UPDATES"
$UpdateButton.add_Click({
    $LAUNCHERSUITEDownloadLocation = "C:\temp\Download\"
    $Get_Date = Get-Content -Path "C:\temp\Launcher\Logs\Suite_date.txt" -Force
    
    If (Test-Path -LiteralPath "C:\temp\Download\Launcher-$Get_Date.zip" -PathType Leaf) {
        
        Write-Host "`nScript Launcher Suite is up to date!" -ForegroundColor Green

        Start-Sleep -Milliseconds 100

        Write-Host "`nLauncher Standing By..." -ForegroundColor Green
    }
    Else {
    
        Write-Host "`nUpdated script package found!" -ForegroundColor Yellow

        Start-Sleep -Milliseconds 100

        Write-Host "`nInitiating Update..." -ForegroundColor Cyan
        
        C:\Windows\System32\xcopy.exe "$LAUNCHERSUITEDownloadLocation" "C:\temp\Launcher\Updates" /H /Y | Out-Host

        Start-Sleep -Milliseconds 100

        Get-ChildItem 'C:\temp\Launcher\Updates' -Filter *.zip | Expand-Archive -DestinationPath 'C:\temp\Launcher\Updates' -Force

        Start-Sleep -Seconds 2

        Get-ChildItem -Path "C:\temp\Launcher\Updates" -Include *.zip -Recurse | Remove-Item

        Start-Sleep -Milliseconds 100

        Write-Host "`nUpdated Script Files Can be found @C:\temp\Launcher\Updates`nClose the Launcher and overwrite the current Launcher folder with the new one. ~Drew Burgess" -ForegroundColor Yellow
    }

})
$LauncherForm.Controls.Add($UpdateButton)

#region Credits
$objLogoText2Name = New-Object System.Windows.Forms.Label
$objLogoText2Name.Location = New-Object System.Drawing.Size(12,175) 
$objLogoText2Name.Size = New-Object System.Drawing.Size(105,15)
$objLogoText2Name.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLogoText2Name.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
$objLogoText2Name.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$objLogoText2Name.Text = "COLLECTION"
$LauncherForm.Controls.Add($objLogoText2Name)

$objLogoText1Name = New-Object System.Windows.Forms.Label
$objLogoText1Name.Location = New-Object System.Drawing.Size(2,150) 
$objLogoText1Name.Size = New-Object System.Drawing.Size(65,15)
$objLogoText1Name.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLogoText1Name.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
$objLogoText1Name.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$objLogoText1Name.Text = "SCRIPT"
$LauncherForm.Controls.Add($objLogoText1Name)

$objLogoText1Name = New-Object System.Windows.Forms.Label
$objLogoText1Name.Location = New-Object System.Drawing.Size(265,160) 
$objLogoText1Name.Size = New-Object System.Drawing.Size(115,15)
$objLogoText1Name.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLogoText1Name.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
$objLogoText1Name.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$objLogoText1Name.Text = "SPC BURGESS"
$LauncherForm.Controls.Add($objLogoText1Name)

$objLogoText1Name = New-Object System.Windows.Forms.Label
$objLogoText1Name.Location = New-Object System.Drawing.Size(280,180) 
$objLogoText1Name.Size = New-Object System.Drawing.Size(125,15)
$objLogoText1Name.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLogoText1Name.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
$objLogoText1Name.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$objLogoText1Name.Text = "SPC SANTIAGO"
$LauncherForm.Controls.Add($objLogoText1Name)

$objLogoText1Name = New-Object System.Windows.Forms.Label
$objLogoText1Name.Location = New-Object System.Drawing.Size(255,140) 
$objLogoText1Name.Size = New-Object System.Drawing.Size(35,15)
$objLogoText1Name.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLogoText1Name.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
$objLogoText1Name.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
$objLogoText1Name.Text = "By"
$LauncherForm.Controls.Add($objLogoText1Name)
#endregion

$option1 = 'Applications          [ADMIN]'
$option2 = 'Account Auditor    [Read Only]'
$option3 = 'Hostname Creator [ADMIN]'
$option4 = 'ADUser Creator    [ADMIN]'
$option5 = 'General Tech       [ADMIN]'
$option6 = 'Zip Script             ADMIN]'
$option7 = 'Master Tool          [ADMIN]'
$option8 = 'Kiosk Creator        [ADMIN]'
$option9 = 'Print Script            [ADMIN]'
$option10= 'Tracker Script      [ADMIN]'



$Choices = @($option1,$option2,$option3,$option4,$option5,$option6,$option7,$option8,$option9,$option10)
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(27, 15)
$comboBox1.Size = New-Object System.Drawing.Size(350, 310)
$LauncherForm.Controls.Add($comboBox1)

foreach($Selectedoption in $Choices)
{
  $comboBox1.Items.add($Selectedoption) 
} 

cls
Sleep 1
Write-Host "`nLauncher Ready`n" -ForegroundColor Cyan

$LauncherForm.KeyPreview = $True
$LauncherForm.Add_KeyDown({if ($_.KeyCode -eq "Enter"){$info=$comboBox1.SelectedItem;$LauncherForm.Close()}})
$LauncherForm.Add_KeyDown({if ($_.KeyCode -eq "Escape"){$Script:CANCELED=$True;$LauncherForm.Close()}})

$LauncherForm.TopMost = $True 
$LauncherForm.Add_Shown({$LauncherForm.Activate()}) | Out-Null
$LauncherForm.ShowDialog() | Out-Null
$LauncherForm.Dispose() | Out-Null



If ($Script:CANCELED -cne $True) { 

    If ($info -and $null) { 
        # Forever NUll basically this should never get called. 
    }
    else {
    If ($comboBox1.SelectedItem -eq $option1) {
        $ArgList = ".\ApplicationsGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -verb RunAs -Wait 
        Write-Host "Applications Script Finished" -ForegroundColor Green
        return;
    }
    If ($comboBox1.SelectedItem -eq $option2) {
        #$ArgList = ".\AccountAuditorGUI.ps1"
        #Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs 
        Write-Host "Account Auditor Script Retired Permanently" -ForegroundColor Red
        return;
    }
    If ($comboBox1.SelectedItem -eq $option3)
    {
        $ArgList = ".\HostnameGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "Hostname Creator Script Finished" -ForegroundColor Green
        return;
    }
    If ($comboBox1.SelectedItem -eq $option4) 
    {
        $ArgList = ".\ADUser_CreationGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "ADUser Creation Script Finished" -ForegroundColor Green
        #Write-Host "AD User Creation Script Coming Soon" -ForegroundColor Yellow
        return;
    }
    If ($comboBox1.SelectedItem -eq $option5)
    {
        $ArgList = ".\GeneralTechGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "General Technician Script Finished" -ForegroundColor Green
        return;
    }
    If ($comboBox1.SelectedItem -eq $option6)
    {
        $ArgList = ".\ZipExtractorGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "Remote Zip Script V2 Finished" -ForegroundColor Green
        return;
    }
    If ($comboBox1.SelectedItem -eq $option8)
    {
        $ArgList = ".\KioskCreatorGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "Kiosk Creator Script Finished" -ForegroundColor Green
        return;
    }
    If ($comboBox1.SelectedItem -eq $option9)
    {
        $ArgList = ".\PrintScriptGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "Print Script Finished" -ForegroundColor Green
        return;
    }
    If ($comboBox1.SelectedItem -eq $option10)
    {
        $ArgList = ".\TrackerGUI.ps1"
        Start-Process Powershell -WorkingDirectory C:\temp\Launcher\Dependencies $ArgList -Verb RunAs -Wait 
        Write-Host "Tracker Script Finished" -ForegroundColor Green
        return;
    }
        else { Write-Host "No Option selected" -ForegroundColor Yellow }
    }
 }
 else { Write-Host "Script Exited Successfully" }        
}
MainLauncher
