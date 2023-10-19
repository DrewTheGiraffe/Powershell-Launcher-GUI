# ~Script By SPC Burgess & Jonathan Santiago 02/20/2023
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
Function GeneralTool{

[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
#[system.windows.forms]::jitDebugging("false") 
#[System.Windows.Forms.Application]::EnableVisualStyles()

try { . ("C:\temp\Launcher\Dependencies\Get-InputBox.ps1") }
catch { Write-Host -f Yellow "Unable to Located File" }




# Globals 
$PSTOOLSDownloadLocation = Get-Content -LiteralPath "C:\temp\Launcher\Logs\PSEXEC_INSTALL.txt" -Force
$exactadminfile = "\\network\path\to\PS_Tools\PsExec.exe" 
$userfile = "C:\Windows\System32" 
$FinalFileString = "$exactadminfile`n$userfile"
$LocalHostName = $env:COMPUTERNAME # returns HOSTNAME only.
#[System.Net.DNS]::GetHostByName($null).HostName; # returns : TheHostname.domain.com

#creates window
$GForm = New-Object System.Windows.Forms.Form
$GForm.Text = '[SA/WA] General Tech v5.0'
$GForm.Width = 800
$GForm.Height = 420
$GForm.BackColor = "White"
$GForm.StartPosition = "CenterScreen"
$GForm.Location = New-Object System.Drawing.Size(80,495)
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$GForm.AutoSize = $true
$GForm.FormBorderStyle = 'Fixed3D'
$GForm.MaximizeBox = $False
$GForm.MinimizeBox = $False

#Correct the initial state of the form to prevent the .Net maximized form issue
$OnLoadForm_StateCorrection={$GForm.WindowState=$InitialFormWindowState}

#Init the OnLoad event to correct the initial state of the form
$GForm.add_Load($OnLoadForm_StateCorrection)

$iconConverted2Base64 = [Convert]::ToBase64String((Get-Content "C:\temp\Launcher\Dependencies\icon\NewPanda.ico" -Encoding Byte))
$iconBase64           = $iconConverted2Base64
$iconBytes            = [Convert]::FromBase64String($iconBase64)
$stream               = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage            = [System.Drawing.Image]::FromStream($stream, $true)
$GForm.Icon    = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
# ico converter : https://cloudconvert.com/png-to-ico

# Draws Logo
$img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\Panda\NewPanda.png')
$GForm.BackgroundImage = $img
$GForm.BackgroundImageLayout = "Center"

#This creates a label for the TextBox Hostname/IP
$objLabel1 = New-Object System.Windows.Forms.Label
$objLabel1.Location = New-Object System.Drawing.Size(125,20) 
$objLabel1.Size = New-Object System.Drawing.Size(280,20)
$objLabel1.ForeColor = [System.Drawing.Color]::FromName("Red")
$objLabel1.Text = "Enter hostname or IP address"
$GForm.Controls.Add($objLabel1) 

#This creates the TextBox Hostname/IP
$objTextBox1 = New-Object System.Windows.Forms.TextBox 
$objTextBox1.Location = New-Object System.Drawing.Size(80,40) 
$objTextBox1.Size = New-Object System.Drawing.Size(260,20)
$objTextBox1.TabIndex = 0 
$GForm.Controls.Add($objTextBox1)
$TheComputer=$objTextBox1.Text # User Input

#This Creates Button Ping
$PingButton = New-Object System.Windows.Forms.Button
$PingButton.Location = New-Object System.Drawing.Size(125,70)
$PingButton.Size = New-Object System.Drawing.Size(75,23)
$PingButton.BackColor = "LightGray"
$PingButton.Text = "PING"
$PingButton.Add_Click({

# Write Hostname to File then read it in.. 
# Create Dir
Sleep 2
New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
Sleep 2
New-Item -Path "C:\temp\Launcher\Logs" -Name "PING_OUTPUT.log" -ItemType "file" -Force
Sleep 1
Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
Sleep 1

Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
Sleep 2
    $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
    Foreach ($filename in $filenames) 
    {
        if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
        {   
            Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
            Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList '"\\network\path\to\PS_Tools\*" "C:\Windows\System32" /H /Y' 
            Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
            break;
        } 
        else 
        {
           If ($objTextBox1.Text -cne $null) 
           {
             sleep 1   
             Write-Host "`nPinging Remote Host" -ForegroundColor Cyan
             $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
             C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s cmd /c "ping $theinfo" > C:\temp\Launcher\Logs\PING_OUTPUT.log
             Sleep 5
             Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\PING_OUTPUT.log"
             Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
             break;
           }
           else 
           {
             sleep 1
             Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
             break;
           }
            sleep 1
            Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
            break;
        }
          
    }# End of for loop

})
$GForm.Controls.Add($PingButton)

#region Later dude

#This Creates Button IP Info
$IPCONFIGButton = New-Object System.Windows.Forms.Button
$IPCONFIGButton.Location = New-Object System.Drawing.Size(200,70)
$IPCONFIGButton.Size = New-Object System.Drawing.Size(75,23)
$IPCONFIGButton.BackColor = "LightGray"
$IPCONFIGButton.Text = "IP INFO"
$IPCONFIGButton.Add_Click(
{

# Write Hostname to File then read it in.. 
# Create Dir
Sleep 2
New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
Sleep 2
New-Item -Path "C:\temp\Launcher\Logs" -Name "IPCONFIG_OUTPUT.log" -ItemType "file" -Force
Sleep 1
Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
Sleep 1

Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
Sleep 2
    $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
    Foreach ($filename in $filenames) 
    {
        if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
        {   
            Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
            Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
            Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
            break;
        } 
        else 
        {
           If ($objTextBox1.Text -cne $null) 
           {
             sleep 1   
             Write-Host "`nQuerying IP Configuration" -ForegroundColor Cyan
             $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
             C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s cmd /c "ipconfig" > C:\temp\Launcher\Logs\IPCONFIG_OUTPUT.log
             Sleep 5
             Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\IPCONFIG_OUTPUT.log"
             Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
             break;
           }
           else 
           {
             sleep 1
             Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
             break;
           }
            sleep 1
            Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
            break;
        }
          
    }# End of for loop

})
$GForm.Controls.Add($IPCONFIGButton)

#This Creates Delete all users button
$QueryUsersButton = New-Object System.Windows.Forms.Button
$QueryUsersButton.Location = New-Object System.Drawing.Size(150,95)
$QueryUsersButton.Size = New-Object System.Drawing.Size(100,23)
$QueryUsersButton.BackColor = "LightGray"
$QueryUsersButton.Text = "Query Users"
$GForm.Controls.Add($QueryUsersButton)

# Available 
$option1 = 'Active Logged in Users'
$option2 = 'Local Users'
$option3 = 'Account Profiles'

$Choices = @($option1,$option2,$option3)
$comboBox1 = New-Object System.Windows.Forms.ComboBox
$comboBox1.Location = New-Object System.Drawing.Point(80,120)
$comboBox1.Size = New-Object System.Drawing.Size(260,20)
$GForm.Controls.Add($comboBox1)

foreach($Selectedoption in $Choices)
{
  $comboBox1.Items.add($Selectedoption)
  
} 
$QueryUsersButton.Add_Click({

    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green 
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
   # Exception error happens when logging query user
   If ($comboBox1.SelectedItem -and $null) { 
       Write-Host "Choose an option from the drop down" -ForegroundColor Yellow 
   }
   else {
    If ($comboBox1.SelectedItem -eq $option1) {
        Sleep 2
        Write-Host "Querying Actively Logged In Users" -ForegroundColor Green
        Sleep 1
        New-Item -Path "C:\temp\Launcher\Logs" -Name "ACTIVE_USER_QUERY_OUTPUT.log" -ItemType "file" -Force
        Sleep 1
        $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
        C:\Windows\System32\PsExec.exe \\$theinfo -s cmd /c query user > C:\temp\Launcher\Logs\ACTIVE_USER_QUERY_OUTPUT.log
        Start-Process "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\ACTIVE_USER_QUERY_OUTPUT.log"
        break;
    }
    If ($comboBox1.SelectedItem -eq $option2) {
        Sleep 2
        Write-Host "Querying Local Users" -ForegroundColor Green
        sleep 1
        New-Item -Path "C:\temp\Launcher\Logs" -Name "LOCAL_USER_QUERY_OUTPUT.log" -ItemType "file" -Force
        Sleep 1
        $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
        C:\Windows\System32\PsExec.exe \\$theinfo -s net user > C:\temp\Launcher\Logs\LOCAL_USER_QUERY_OUTPUT.log
        Start-Process "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\LOCAL_USER_QUERY_OUTPUT.log"
        break;
    }
    If ($comboBox1.SelectedItem -eq $option3) {
        Sleep 2
        Write-Host "Querying Network Profiles" -ForegroundColor Green
        sleep 1
        New-Item -Path "C:\temp\Launcher\Logs" -Name "ACCOUNT_PROFILES_QUERY_OUTPUT.log" -ItemType "file" -Force
        sleep 1
        $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
        C:\Windows\System32\PsExec.exe \\$theinfo -s cmd /c dir C:\Users > C:\temp\Launcher\Logs\ACCOUNT_PROFILES_QUERY_OUTPUT.log
        Start-Process "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\ACCOUNT_PROFILES_QUERY_OUTPUT.log"
        break;
    }
  }
})

cls
sleep 3
Write-Host "Script ready for user interaction`n`n" -ForegroundColor Green
#This Creates Delete all users button
$DelAllButton = New-Object System.Windows.Forms.Button
$DelAllButton.Location = New-Object System.Drawing.Size(150,150)
$DelAllButton.Size = New-Object System.Drawing.Size(100,23)
$DelAllButton.BackColor = "LightGray"
$DelAllButton.Text = "Delete All Users"
$DelAllButton.Add_Click({

# Write Hostname to File then read it in.. 
# Create Dir
Sleep 2
New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
Sleep 2
New-Item -Path "C:\temp\Launcher\Logs" -Name "DeleteALLUsers_OUTPUT.log" -ItemType "file" -Force
Sleep 1
Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
Sleep 1

Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
Sleep 2
    $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
    Foreach ($filename in $filenames) 
    {
        if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
        {   
            Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
            Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
            Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
            break;
        } 
        else 
        {
           If ($objTextBox1.Text -cne $null) 
           {
             sleep 1   
             Write-Host "`nRemoving all user profiles from remote host: $theinfo" -ForegroundColor Cyan
             $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
             C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s cmd /c "del C:\Users\* /q /s /f" > C:\temp\Launcher\Logs\DeleteALLUsers_OUTPUT.log
             Sleep 5
             Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\DeleteALLUsers_OUTPUT.log"
             Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
             break;
           }
           else 
           {
             sleep 1
             Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
             break;
           }
            sleep 1
            Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
            break;
        }
          
    }# End of for loop

})
$GForm.Controls.Add($DelAllButton)

#This Creates Delete all users button
$DeleteUserButton = New-Object System.Windows.Forms.Button
$DeleteUserButton.Location = New-Object System.Drawing.Size(150,180)
$DeleteUserButton.Size = New-Object System.Drawing.Size(100,23)
$DeleteUserButton.BackColor = "LightGray"
$DeleteUserButton.Text = "Delete A User"
$DeleteUserButton.Add_Click({
 # Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "Deleted_User.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox2.Text -cne $null) 
            {
                sleep 1   
                Write-Host "`nUser is being delete" -ForegroundColor Cyan
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                $Custominputuser = $objTextBox2.Text 
                C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s cmd /c "del C:\Users\$Custominputuser /q /s /f" > C:\temp\Launcher\Logs\Deleted_User.log
                Sleep 5
                Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\Delete_User.log"
                Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    })
$GForm.Controls.Add($DeleteUserButton)

#This creates the TextBox for deleting a profile
$objTextBox2 = New-Object System.Windows.Forms.TextBox 
$objTextBox2.Location = New-Object System.Drawing.Size(80,210) 
$objTextBox2.Size = New-Object System.Drawing.Size(260,20)
$objTextBox2.TabIndex = 0 
$GForm.Controls.Add($objTextBox2)

#This Creates Button Restart
$RestartButton = New-Object System.Windows.Forms.Button
$RestartButton.Location = New-Object System.Drawing.Size(125,240)
$RestartButton.Size = New-Object System.Drawing.Size(75,23)
$RestartButton.BackColor = "LightGray"
$RestartButton.Text = "Restart"
$RestartButton.Add_Click({
 # Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox2.Text -cne $null) 
            {
                sleep 1   
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force 
                Write-Host "`nRestarting Remote Host : $theinfo" -ForegroundColor Cyan
                C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s shutdown -r -t 0
                Sleep 5
                Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    })
$GForm.Controls.Add($RestartButton)

#This Creates Button Shutdown
$ShutdownButton = New-Object System.Windows.Forms.Button
$ShutdownButton.Location = New-Object System.Drawing.Size(200,240)
$ShutdownButton.Size = New-Object System.Drawing.Size(75,23)
$ShutdownButton.BackColor = "LightGray"
$ShutdownButton.Text = "Shutdown"
$ShutdownButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
            {
                sleep 1   
                Write-Host "`nRemote Host is being shutdown" -ForegroundColor Cyan
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s shutdown -s -t 0
                write-host "Remote Host Successfully Shutting Down!" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    }) 
$GForm.Controls.Add($ShutdownButton)

#This Creates Button CMD
$CMDButton = New-Object System.Windows.Forms.Button
$CMDButton.Location = New-Object System.Drawing.Size(125,270)
$CMDButton.Size = New-Object System.Drawing.Size(75,23)
$CMDButton.BackColor = "LightGray"
$CMDButton.Text = "CMD"
$CMDButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
            {
                sleep 1   
                Write-Host "`nSending CMD Window to remote host`nNOTE: Close CMD on remote client to continue with General Tech Application" -ForegroundColor Cyan
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s -i cmd
                Write-Host "Process finished on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    }) 
$GForm.Controls.Add($CMDButton)

#This Creates Button Powershell
$PSButton = New-Object System.Windows.Forms.Button
$PSButton.Location = New-Object System.Drawing.Size(200,270)
$PSButton.Size = New-Object System.Drawing.Size(75,23)
$PSButton.BackColor = "LightGray"
$PSButton.Text = "Powershell"
$PSButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
            {
                sleep 1   
                Write-Host "`nSending Powershell Window to remote host`nNOTE: Close PS on remote client to continue with General Tech Application" -ForegroundColor Cyan
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s -i Powershell
                Write-Host "Process finished on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    }) 
$GForm.Controls.Add($PSButton)

#This Creates Button Enables local admin
$ELocalButton = New-Object System.Windows.Forms.Button
$ELocalButton.Location = New-Object System.Drawing.Size(125,300)
$ELocalButton.Size = New-Object System.Drawing.Size(150,23)
$ELocalButton.BackColor = "LightGray"
$ELocalButton.Text = "Enable Local Admin"
$ELocalButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1
    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
    $Password = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String('MwA3AEoAZQBrACEAKgBUADQAZwAzADcASgBlAGsA'))
    $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
    
            If ($objTextBox1.Text -cne $null) 
            {
                # Function Globals

                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force -ErrorAction Ignore

                # Process Start

                Write-Host "`nCreating Administrator Local Administrator on remote Computer" -ForegroundColor Cyan

                Invoke-Command -ComputerName $theinfo -ScriptBlock {
    
                    $Password = [System.Text.Encoding]::Unicode.GetString([System.Convert]::FromBase64String('MwA3AEoAZQBrACEAKgBUADQAZwAzADcASgBlAGsA'))
    
                    [System.Security.SecureString]$SecurePassword = ConvertTo-SecureString -String $Password -AsPlainText -Force

                    New-LocalUser -Name "joe.smith" -Password $SecurePassword -ErrorAction Ignore

                    Set-LocalUser -Name "Administrator" -AccountNeverExpires -Description "Local Administrator Account 20H2+ PXE Image" -FullName "Administrator" -PasswordNeverExpires $True -UserMayChangePassword $False -ErrorAction Ignore

                    If (Get-LocalUser -Name "Administrator" -ErrorAction Ignore) {
        
                        Write-Host "`nCreated Administrator local administrator" -ForegroundColor Green

                        Write-Host "`nApplying Permissions to Administrator local administrator" -ForegroundColor Cyan

                    If (Get-LocalGroupMember -Group "Administrators" -Member "Administrator" -ErrorAction Ignore)
                    {
            
                        Write-Host "`nAdministrator is already an administrator..Moving on!" -ForegroundColor Green
        
                    }
                    else {
            
                        Write-Host "`nApplying administrator privileges to Joe Smith" -ForegroundColor Cyan
                        Add-LocalGroupMember -Group "Administrators" -Member "Administrator" -ErrorAction Ignore

                    }

                    Write-Host "`nActivating Joe Smith local administrator" -ForegroundColor Cyan

                    Enable-LocalUser -Name "Administrator" -ErrorAction Ignore

        
                } 
                else {

                    Write-Host "`nFailed to locate Administrator on remote computer" -ForegroundColor Yellow -BackgroundColor Black

                    Write-Host "`nScript Stopped" -ForegroundColor Red -BackgroundColor Black

                    Break;
    
                }
              
                Write-Host "`nFinished Administrator process on remote host" -ForegroundColor Green
   
            }
            
            $oFile = New-Object System.IO.FileInfo "C:\temp\Launcher\Logs\Administrator.txt"

            if ((Test-Path -Path "C:\temp\Launcher\Logs\Administrator.txt" -ErrorAction Ignore) -eq $false) {
        
                Write-Host "`nCreating Joe Smith Login Credential log" -ForegroundColor Cyan

                New-Item -Path "C:\temp\Launcher\Logs" -Name "JoeSmith.txt" -ItemType "file" -Force -ErrorAction Ignore

                $JoeSmithCreds = "USERNAME: .\Administrator`nPASSWORD: $Password`n`nHostname: $theinfo"

                Set-Content -Path "C:\temp\Launcher\Logs\Administrator.txt" -Value ($JoeSmithCreds)

                Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\Administrator.txt"

                Start-Sleep -Seconds 4

                $oStream = $oFile.Open([System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)

                if ($oStream) {

            
                    $oStream.Close()

                    
        
                }

            }
            Else {
        
                    Set-Content -Path "C:\temp\Launcher\Logs\Administrator.txt" -Value ($null)

                    $JoeSmithCreds = "USERNAME: .\Administrator`nPASSWORD: $Password`n`nHostname: $theinfo"
               
                    Set-Content -Path "C:\temp\Launcher\Logs\Administrator.txt" -Value ($JoeSmithCreds)

                    Start-sleep -Seconds 2

                    Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\Administrator.txt"

            }
            
            }

            Else {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
            
            # break out after first run..
            break;
            
          
        }# End of for loop

        # Remove log for security       
        Remove-Item -Path "C:\temp\Launcher\Logs\Administrator.txt" -Force

})
$GForm.Controls.Add($ELocalButton)

#This Creates Button Deletes local admin
$DELocalButton = New-Object System.Windows.Forms.Button
$DELocalButton.Location = New-Object System.Drawing.Size(125,330)
$DELocalButton.Size = New-Object System.Drawing.Size(150,23)
$DELocalButton.BackColor = "LightGray"
$DELocalButton.Text = "Delete Local Admin"
$DELocalButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
            {
                sleep 1   
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                Write-Host "Disabling Administrator Local Administrator on remote Computer : $theinfo" -ForegroundColor Cyan
                C:\Windows\System32\PsExec.exe \\$theinfo -s net user joe.smith /Active:no /delete
                Sleep 2
                Write-Host "`nFinished Removing Administrator on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    })
$GForm.Controls.Add($DELocalButton)

#This Creates Button Monitor Bitlocker
$MBButton = New-Object System.Windows.Forms.Button
$MBButton.Location = New-Object System.Drawing.Size(400,40)
$MBButton.Size = New-Object System.Drawing.Size(150,23)
$MBButton.BackColor = "LightGray"
$MBButton.Text = "Monitor Bitlocker Status"
$MBButton.Add_Click({
    # Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "MANAGE_BITLOCKER_STATUS.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
                {
                    sleep 1   
                    Write-Host "`nWriting Bitlocker status to file" -ForegroundColor Cyan
                    $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                    C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s manage-bde -status > C:\temp\Launcher\Logs\MANAGE_BITLOCKER_STATUS.log
                    Sleep 5
                    Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\MANAGE_BITLOCKER_STATUS.log"
                    Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                    break;
                }
            }
        }
    })
$GForm.Controls.Add($MBButton)

#This Creates Button Disable bitlocker
$DisBButton = New-Object System.Windows.Forms.Button
$DisBButton.Location = New-Object System.Drawing.Size(550,40)
$DisBButton.Size = New-Object System.Drawing.Size(150,23)
$DisBButton.BackColor = "LightGray"
$DisBButton.Text = "Disable Bitlocker"
$DisBButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "MANAGE_BITLOCKER_STATUS.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
                {
                    sleep 1   
                    $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                    Write-Host "`nDisabling Bitlocker Encryption on Remote Host : $theinfo" -ForegroundColor Cyan
                    C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s manage-bde -off C:
                    Sleep 5
                    Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\MANAGE_BITLOCKER_STATUS.log"
                    Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                    break;
                }
            }
        }
    })
$GForm.Controls.Add($DisBButton)

#This Creates Button Query Bitlocker Key
$QueryKeyButton = New-Object System.Windows.Forms.Button
$QueryKeyButton.Location = New-Object System.Drawing.Size(400,65) 
$QueryKeyButton.Size = New-Object System.Drawing.Size(300,23)
$QueryKeyButton.BackColor = "LightGray"
$QueryKeyButton.Text = "Query Bitlocker Key & Backup to AD"
$QueryKeyButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "BITLOCKER_KEY.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1
    $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
    Sleep 2
    If ($theinfo -and !$null) 
    {
        sleep 1 
        Write-Host "`nQuerying Bitlocker Key Info..." -ForegroundColor Cyan 
        $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
        Invoke-Command -ComputerName $theinfo -ScriptBlock {
        
            # Smart Alternative right here.. this will just grab the numericalpassword by itself without querying the entire volume like i do below... ~ DREW BURGESS
            # (Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume").GetKeyProtectors(3).volumekeyprotectorID

            $Bitlocker = Get-BitLockerVolume -MountPoint 'C:'
    
            foreach ($blv in $Bitlocker) {

                 Backup-BitLockerKeyProtector –MountPoint $blv.MountPoint –KeyProtectorId (($blv.KeyProtector)<#[2]#> | Select-Object –ExpandProperty KeyProtectorID)

            }
            
            Write-Host "`nRecovery information was successfully backed up to Active Directory on remote host." -ForegroundColor Green

        } -ErrorAction Ignore
                    <# 

                    Retired code relies on PSEXEC SUITE ~ DREW BURGESS
                    
                    Write-Host "`nWriting Bitlocker key to file" -ForegroundColor Cyan
                    $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                    C:\Windows\System32\PsExec.exe -accepteula \\$theinfo manage-bde -protectors -get C: > C:\temp\Launcher\Logs\BITLOCKER_KEY.log
                    Write-Host "Copy Numerical Password ID and close Notepad`nExample: {EA70CF76-XXXX-XXXX-XXXX-9EDF86339DF7}"
                    Sleep 5
                    Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\BITLOCKER_KEY.log"
                    Write-Host "`nFinishing Process on remote host : $theinfo" -ForegroundColor Green
                    [String]$KeyBackup = Get-InputBox "Numerical Password Input" "Example: `n{EA70CF76-XXXX-XXXX-XXXX-9EDF86339DF7}" 
                    C:\Windows\System32\PsExec.exe -accepteula \\$theinfo manage-bde -protectors -adbackup C: -id $KeyBackup
                    Write-Host "Successfully backed up Bitlocker Key to Active Directory" -ForegroundColor Green
                    
                    #>
                   # break;
   }
   Else {
   
        Write-Host "`nNo Hostname Detected.. Backing up local host bitlocker info..." -ForegroundColor DarkYellow
        $NumericalPassword = (Get-WmiObject -Namespace "Root\cimv2\Security\MicrosoftVolumeEncryption" -Class "Win32_EncryptableVolume").GetKeyProtectors(3).volumekeyprotectorID
        manage-bde -protectors -adbackup C: -id "$NumericalPassword" | Out-Null
        Start-Sleep -Milliseconds 100
        Write-Host "`nRecovery information was successfully backed up to Active Directory on local host : $env:COMPUTERNAME" -ForegroundColor Green
   }
})
$GForm.Controls.Add($QueryKeyButton)

#This Creates Button Uses TPM
$TPMButton = New-Object System.Windows.Forms.Button
$TPMButton.Location = New-Object System.Drawing.Size(460,90)
$TPMButton.Size = New-Object System.Drawing.Size(185,23)
$TPMButton.BackColor = "LightGray"
$TPMButton.Text = "Disable Bitlocker PIN"
$TPMButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
            {
                sleep 1   
                Write-Host "`nDisabling Bitlocker Pin on remote host" -ForegroundColor Cyan
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                Write-Host "Disabling Joe.Smith Local Administrator on remote Computer : $theinfo" -ForegroundColor Cyan
                C:\Windows\System32\PsExec.exe \\$theinfo -s cmd /c manage-bde -protectors -add C: -tpm
                Sleep 2
                Write-Host "`nFinished Removing Bitlocker Pin on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {
                sleep 1
                Write-Host "`nNO HostName or IP Address Entered" -ForegroundColor Yellow
                break;
            }
                sleep 1
                Write-Host "`nPSEXEC could not be found at the current path : \\network\path\to\PS_Tools\" -ForegroundColor Yellow
                break;
            }
          
        }# End of for loop

    })
$GForm.Controls.Add($TPMButton)
#endregion
#This Creates Button Install PSEXEC
$PSEXECButton = New-Object System.Windows.Forms.Button
$PSEXECButton.Location = New-Object System.Drawing.Size(460,115)
$PSEXECButton.Size = New-Object System.Drawing.Size(185,23)
$PSEXECButton.BackColor = "LightGray"
$PSEXECButton.Text = "Install PSEXEC (Local Host)"
$PSEXECButton.Add_Click(
{
Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
Sleep 2
    $filenames = Get-Content "C:\temp\Launcher\Logs\PSEXEC_INSTALL.txt"; # Reading the names of the files to test the existance in one of the locations
    if ((Test-Path "C:\Windows\System32\PSExec.exe" -PathType Leaf) -eq $False) #if psexec is not on the localhost
    {   
        Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
        C:\Windows\System32\xcopy.exe "$PSTOOLSDownloadLocation" "C:\Windows\System32" /H /Y | Out-Host
        Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
    } 
    else 
    {
        Write-Host "`nPSTools is already installed on this computer" -ForegroundColor Green
    }
    Write-Host "`nGeneralTech Finished`nStanding By!" -ForegroundColor Green
}# End of for loop
)
$GForm.Controls.Add($PSEXECButton)

#region BG
Function Set-WallPaper {
 
<#
 
    .SYNOPSIS
    Applies a specified wallpaper to the current user's desktop
    
    .PARAMETER Image
    Provide the exact path to the image
 
    .PARAMETER Style
    Provide wallpaper style (Example: Fill, Fit, Stretch, Tile, Center, or Span)
  
    .EXAMPLE
    Set-WallPaper -Image "C:\Wallpaper\Default.jpg"
    Set-WallPaper -Image "C:\Wallpaper\Background.jpg" -Style Fit
  
#>
 
param (
    [parameter(Mandatory=$True)]
    # Provide path to image
    [string]$Image,
    # Provide wallpaper style that you would like applied
    [parameter(Mandatory=$False)]
    [ValidateSet('Fill', 'Fit', 'Stretch', 'Tile', 'Center', 'Span')]
    [string]$Style
)
 
$WallpaperStyle = Switch ($Style) {
  
    "Fill" {"10"}
    "Fit" {"6"}
    "Stretch" {"2"}
    "Tile" {"0"}
    "Center" {"0"}
    "Span" {"22"}
  
}
 
If($Style -eq "Tile") {
 
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 1 -Force
 
}
Else {
 
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name WallpaperStyle -PropertyType String -Value $WallpaperStyle -Force
    New-ItemProperty -Path "HKCU:\Control Panel\Desktop" -Name TileWallpaper -PropertyType String -Value 0 -Force
 
}
 
Add-Type -TypeDefinition @" 
using System; 
using System.Runtime.InteropServices;
  
public class Params
{ 
    [DllImport("User32.dll",CharSet=CharSet.Unicode)] 
    public static extern int SystemParametersInfo (Int32 uAction, 
                                                   Int32 uParam, 
                                                   String lpvParam, 
                                                   Int32 fuWinIni);
}
"@ 
  
    $SPI_SETDESKWALLPAPER = 0x0014
    $UpdateIniFile = 0x01
    $SendChangeEvent = 0x02
  
    $fWinIni = $UpdateIniFile -bor $SendChangeEvent
  
    $ret = [Params]::SystemParametersInfo($SPI_SETDESKWALLPAPER, 0, $Image, $fWinIni)
}
 # Source : https://www.joseespitia.com/2017/09/15/set-wallpaper-powershell-function/

#endregion

#This Creates Button Execute Post Image Commands Remotely
$DesktopBackroundButton = New-Object System.Windows.Forms.Button
$DesktopBackroundButton.Location = New-Object System.Drawing.Size(460,140)
$DesktopBackroundButton.Size = New-Object System.Drawing.Size(185,23)
$DesktopBackroundButton.BackColor = "LightGray"
$DesktopBackroundButton.Text = "Setup Desktop (Local Host)"
$DesktopBackroundButton.Add_Click({
    # Write Hostname to File then read it in.. 
    # Create Dir
    sleep 1   
    $localhost = $env:COMPUTERNAME
    Write-Host "`nSetting Backround to $localhost" -ForegroundColor Cyan
    Set-WallPaper -Image "C:\temp\Launcher\Dependencies\icon\DesktopBg\noplacelikehome.jpg" -Style Fit
    Sleep 3
    Write-Host "`nDesktop Backround Set!`nEnjoy!" -ForegroundColor Green
})
$GForm.Controls.Add($DesktopBackroundButton) 

#This Creates Button Query installed apps
$QAppsButton = New-Object System.Windows.Forms.Button
$QAppsButton.Location = New-Object System.Drawing.Size(480,165)
$QAppsButton.Size = New-Object System.Drawing.Size(150,23)
$QAppsButton.BackColor = "LightGray"
$QAppsButton.Text = "Query Installed Apps"
$QAppsButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "QUERY_APPLICATIONS.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
                {
                    sleep 1   
                    Write-Host "`nQuerying Installed Applications" -ForegroundColor Cyan
                    $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                    C:\Windows\System32\Psinfo.exe -accepteula \\$theinfo -S > C:\temp\Launcher\Logs\QUERY_APPLICATIONS.log
                    Sleep 5
                    Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\QUERY_APPLICATIONS.log"
                    Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                    break;
                }
            }
        }
    })
$GForm.Controls.Add($QAppsButton)


#This Creates Button Query SN
$SNButton = New-Object System.Windows.Forms.Button
$SNButton.Location = New-Object System.Drawing.Size(480,190) 
$SNButton.Size = New-Object System.Drawing.Size(150,23)
$SNButton.BackColor = "LightGray"
$SNButton.Text = "Query Serial Number"
$SNButton.Add_Click({
# Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "QUERY_SERIAL_NUMBER.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 1

    Write-Host "`nChecking if psexec exists, Do not Spam!" -ForegroundColor Cyan
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            if ((Test-Path $exactadminfile\$filename) -and !(Test-Path $userfile\$filename)) #if the file is in share drive but not in Win\Sys32 folder
            {   
                Write-Host "`nBeginning Download of PS_Tools`nPlease Be Patient" -ForegroundColor Cyan # Change this directory to point to your NECs sharedrive w/ PSTools
                Start-Process -Wait -PSPath "C:\Windows\System32\xcopy.exe" -ArgumentList "$PSTOOLSDownloadLocation C:\Windows\System32 /H /Y" 
                Write-Host "`nFinished Downloading PS_Tools" -ForegroundColor Green
                break;
            } 
            else 
            {
            If ($objTextBox1.Text -cne $null) 
                {
                    sleep 1   
                    Write-Host "`nQuerying Serial Number" -ForegroundColor Cyan
                    $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                    C:\Windows\System32\PsExec.exe -accepteula \\$theinfo -s wmic bios get SerialNumber > C:\temp\Launcher\Logs\QUERY_SERIAL_NUMBER.log
                    Sleep 5
                    Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\QUERY_SERIAL_NUMBER.log"
                    Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                    break;
                }
            }
        }
    })
$GForm.Controls.Add($SNButton)

#This Creates Button Enable PS Remoting Remotely or Locally
$EnablePSRemotingButton = New-Object System.Windows.Forms.Button
$EnablePSRemotingButton.Location = New-Object System.Drawing.Size(480,215)
$EnablePSRemotingButton.Size = New-Object System.Drawing.Size(150,23)
$EnablePSRemotingButton.BackColor = "LightGray"
$EnablePSRemotingButton.Text = "Enable PS Remoting"
$EnablePSRemotingButton.Add_Click({
    # Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "ENABLE-PSREMOTING.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 2
        $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
            If ($objTextBox1.Text -cne $null) 
            {
                cls
                
                try { . ("C:\temp\Launcher\Dependencies\Enable-PSRemotingRemotely.ps1") }
                catch { Write-Host -ForegroundColor Yellow "Unable to Locate PSRemoting Script" }
                
                sleep 1   
                
                Write-Host "`nAttempting to Enable PS Remoting Remotely" -ForegroundColor Cyan
                
                $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
                
                Enable-PSRemotingRemotely -ComputerNames $theinfo > C:\temp\Launcher\Logs\ENABLE-PSREMOTING.log
                
                Start-Sleep -Milliseconds 100

                $Confirm = Read-Host -Prompt "Would you like to do an in-depth verification on all services? (yes/no)"

                If (($Confirm -eq "yes") -or ($Confirm -eq "Yes") -or ($Confirm -eq "YES")) {
                
                    # Wired AutoConfig Service Status
                    $WiredAutConfigService = Get-Service -Name "dot3svc" -ComputerName $theinfo -ErrorAction Ignore | Select-Object -ExpandProperty "Status"

                    # Wired AutoConfig Dependents
                    $RpcSc_Status =   Get-Service -Name "Wired AutoConfig" -RequiredServices -ComputerName $theinfo -ErrorAction Ignore | Select-Object -ExpandProperty 'Status' -Property "RpcSs" -First 1
                    $Eaphost_Status = Get-Service -Name "Wired AutoConfig" -RequiredServices -ComputerName $theinfo -ErrorAction Ignore | Select-Object -ExpandProperty 'Status' -Property "Eaphost" -First 1
                    $Ndisuio_Status = Get-service -Name "Wired AutoConfig" -RequiredServices -ComputerName $theinfo -ErrorAction Ignore | Select-Object -ExpandProperty 'Status' -Property "Ndisuio" -First 1

                    Invoke-Command -ComputerName $theinfo -ScriptBlock {
        
                        If ($WiredAutConfigService -eq "Running") {
        
                            Write-Host "`nDetected dot3svc..." -ForegroundColor Cyan
        
                            Start-Sleep -Milliseconds 100
        
                            Write-Host "`nChecking for linked services..." -ForegroundColor Cyan
        
                            Start-Sleep -Milliseconds 100
        
                            # Check for RpcSs
                            If ($RpcSc_Status -eq "Running") { Write-Host "`nDetected RpcSs..." -ForegroundColor Cyan }
                            else { Start-Service -Name "RpcSs"; Write-Host "`nStarted RpcSs..." -ForegroundColor Yellow }

                            # Check for Eaphost
                            If ($Eaphost_Status -eq "Running") { Write-Host "`nDetected Eaphost..." -ForegroundColor Cyan }
                            else { Start-Service -Name "Eaphost"; Write-Host "`nStarted Eaphost..." -ForegroundColor Yellow }

                            # Check for Ndisuio
                            If ($Ndisuio_Status -eq "Running") { Write-Host "`nDetected Ndisuio..." -ForegroundColor Cyan }
                            else { Start-Service -Name "Ndisuio"; Write-Host "`nStarted Ndisuio..." -ForegroundColor Yellow }

                        }
                        else {
      
                            Write-Host "`nUnabled to Detect dot3svc..." -ForegroundColor Yellow

                            Start-Sleep -Milliseconds 100

                            Write-Host "`nEnabling dot3svc..." -ForegroundColor Cyan

                            Restart-Service -Name "dot3svc" -Force

                            Start-Sleep -Milliseconds 100

                            Write-Host "`nChecking for linked services..." -ForegroundColor Cyan

                            Start-Sleep -Milliseconds 100
        
                            # Check for RpcSs
                            If ($RpcSc_Status -eq "Running") { Write-Host "`nDetected RpcSs..." -ForegroundColor Cyan }
                            else { Start-Service -Name "RpcSs"; Write-Host "`nStarted RpcSs..." -ForegroundColor Yellow }

                            # Check for Eaphost
                            If ($Eaphost_Status -eq "Running") { Write-Host "`nDetected Eaphost..." -ForegroundColor Cyan }
                            else { Start-Service -Name "Eaphost"; Write-Host "`nStarted Eaphost..." -ForegroundColor Yellow }

                            # Check for Ndisuio
                            If ($Ndisuio_Status -eq "Running") { Write-Host "`nDetected Ndisuio..." -ForegroundColor Cyan }
                            else { Start-Service -Name "Ndisuio"; Write-Host "`nStarted Ndisuio..." -ForegroundColor Yellow }

                        }
                    } -ErrorAction Ignore

                    Start-Sleep -Milliseconds 100

                    Write-Host "`nFinished dot3svc Verification" -ForegroundColor Green

                    Start-Sleep -Milliseconds 100
                
                }

                Sleep 5
                
                cls
                
                Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\ENABLE-PSREMOTING.log"
                
                Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                
                break;
            }
            else 
            {   # Local Host if Hostname is Null
                cls
                Write-Host "`nNo Hostname or IP Address entered.`nGeneral Tech Ready." -ForegroundColor Green
                break;
            }
            break;
        }

})
$GForm.Controls.Add($EnablePSRemotingButton)


#This Creates Button SMS CLIENT Refresh Remotely
$ReFreshSoftwareCenterCommandButton = New-Object System.Windows.Forms.Button
$ReFreshSoftwareCenterCommandButton.Location = New-Object System.Drawing.Size(480,240) 
$ReFreshSoftwareCenterCommandButton.Size = New-Object System.Drawing.Size(150,23)
$ReFreshSoftwareCenterCommandButton.BackColor = "LightGray"
$ReFreshSoftwareCenterCommandButton.Text = "Refresh Software Center"
$ReFreshSoftwareCenterCommandButton.Add_Click({
    # Write Hostname to File then read it in.. 
    # Create Dir
    Remove-Item -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force # Delete File
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "GeneralTechLog.txt" -ItemType "file" -Force # Re-Create File
    Write-Host "`nGeneral Tech Log Created" -ForegroundColor Green
    Sleep 2
    New-Item -Path "C:\temp\Launcher\Logs" -Name "ENABLE-PSREMOTING.log" -ItemType "file" -Force
    Sleep 1
    Set-Content -Path "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Value ($objTextBox1.Text)
    Write-Host "`nContents Finished Writing to file `nat Path : C:\temp\Launcher\Logs\GeneralTechLog.txt" -ForegroundColor Green
    Sleep 2
    $filenames=Get-Content "C:\temp\Launcher\Logs\FileCheckLog.txt"; # Reading the names of the files to test the existance in one of the locations
        Foreach ($filename in $filenames) 
        {
        $theinfo = Get-Content -LiteralPath "C:\temp\Launcher\Logs\GeneralTechLog.txt" -Force
            If ($theinfo -and !$null) 
            {
                cls
                sleep 1   
                Write-Host "`nAttempting to Refresh Software Center Remotely..." -ForegroundColor Cyan
                Sleep 1
                # Hardware Inventory Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}"
                Write-Host "`nHardware Inventory Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Software Inventory Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000002}"
                Write-Host "`nSoftware Inventory Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Discovery Data Collection Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000003}"
                Write-Host "`nDiscovery Data Collection Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # File Collection Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000010}"
                Write-Host "`nFile Collection Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Machine Policy Retrieval Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"
                Write-Host "`nMachine Policy Retrieval Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Machine Policy Evaluation Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}"
                Write-Host "`nMachine Policy Evaluation Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Refresh Default MP Task Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000023}"
                Write-Host "`nRefresh Default MP Task Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # User Policy Retrieval Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000026}" -ErrorAction Ignore # Seems like some machines cant run this action..
                Write-Host "`nUser Policy Retrieval Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # User Policy Evaluation Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000027}" -ErrorAction Ignore # Seems like some machines cant run this action..
                Write-Host "`nUser Policy Evaluation Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Windows Installers Source List Update Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000032}"
                Write-Host "`nWindows Installers Source List Update Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Machine Policy Agent Cleanup
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000040}" 
                Write-Host "`nMachine Policy Agent Cleanup Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # User Policy Agent Cleanup
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000041}"-ErrorAction Ignore # Seems like some machines cant run this action..
                Write-Host "`nUser Policy Agent Cleanup Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # IDMIF Collection Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000105}" -ErrorAction Ignore # Seems like some machines cant run this action..
                Write-Host "`nIDMIF Collection Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Software Updates Assignments Evaluation Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000108}"
                Write-Host "`nSoftware Updates Assignments Evaluation Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # State Message Refresh
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000111}"
                Write-Host "`nState Message Refresh Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Software Update Scan Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}"
                Write-Host "`nSoftware Update Scan Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Software Update Deployment Evaluation Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000114}"
                Write-Host "`nSoftware Update Deployment Evaluation Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Application Deployment Evaluation Cycle
                Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}"
                Write-Host "`nApplication Deployment Evaluation Cycle Triggered...." -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                # Source Command List : https://docs.microsoft.com/en-us/mem/configmgr/develop/reference/core/clients/client-classes/triggerschedule-method-in-class-sms_client
                # Refresh all computer policies / workstation certificates
                Invoke-Command -ComputerName $theinfo -ScriptBlock { gpupdate /force /target:computer /wait:0 }
                $msg = "This computer is scheduled for a restart in 20 minutes, please save all work."
                Start-Sleep -Milliseconds 100
                $CurrentTime = Get-date
                $DateNTime = $CurrentTime.ToLocalTime()
                $FormattedTime = "$DateNTime"
                Write-Host "`nEstimated time to completion...`n15 minutes after: $FormattedTime" -ForegroundColor Cyan
                Start-Sleep -Milliseconds 100
                shutdown -m \\$theinfo -r -f -c $msg -t 1200 <#20 minute delay to give SCCM Client time#> 
                Start-Sleep -Milliseconds 100
                Write-Host "`nRemote Host Scheduled to restart...`n20 minutes after: $FormattedTime" -ForegroundColor Cyan
                Write-Host "`nFinished Process on remote host : $theinfo" -ForegroundColor Green
                break;
            }
            else 
            {   
                cls
                $Confirm = Read-Host -Prompt "Would you like to Refresh SCCM Client on $LocalHostName (yes/no)"
                If ($Confirm -eq "yes") {
                    sleep 1   
                    Write-Host "`nAttempting to Refresh Software Center Locally..." -ForegroundColor Cyan
                    Sleep 1
                    # Hardware Inventory Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000001}"
                    Write-Host "`nHardware Inventory Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Software Inventory Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000002}"
                    Write-Host "`nSoftware Inventory Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Discovery Data Collection Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000003}"
                    Write-Host "`nDiscovery Data Collection Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # File Collection Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000010}"
                    Write-Host "`nFile Collection Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Machine Policy Retrieval Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000021}"
                    Write-Host "`nMachine Policy Retrieval Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Machine Policy Evaluation Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000022}"
                    Write-Host "`nMachine Policy Evaluation Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Refresh Default MP Task Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000023}"
                    Write-Host "`nRefresh Default MP Task Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # User Policy Retrieval Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000026}" -ErrorAction Ignore # Seems like some machines cant run this action..
                    Write-Host "`nUser Policy Retrieval Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # User Policy Evaluation Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000027}" -ErrorAction Ignore # Seems like some machines cant run this action..
                    Write-Host "`nUser Policy Evaluation Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Windows Installers Source List Update Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000032}"
                    Write-Host "`nWindows Installers Source List Update Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Machine Policy Agent Cleanup
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000040}" 
                    Write-Host "`nMachine Policy Agent Cleanup Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # User Policy Agent Cleanup
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000041}"-ErrorAction Ignore # Seems like some machines cant run this action..
                    Write-Host "`nUser Policy Agent Cleanup Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # IDMIF Collection Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000105}" -ErrorAction Ignore # Seems like some machines cant run this action..
                    Write-Host "`nIDMIF Collection Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Software Updates Assignments Evaluation Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000108}"
                    Write-Host "`nSoftware Updates Assignments Evaluation Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # State Message Refresh
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000111}"
                    Write-Host "`nState Message Refresh Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Software Update Scan Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000113}"
                    Write-Host "`nSoftware Update Scan Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Software Update Deployment Evaluation Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000114}"
                    Write-Host "`nSoftware Update Deployment Evaluation Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Application Deployment Evaluation Cycle
                    Invoke-WmiMethod -ComputerName $theinfo -Namespace root\ccm -Class SMS_CLIENT -Name TriggerSchedule "{00000000-0000-0000-0000-000000000121}"
                    Write-Host "`nApplication Deployment Evaluation Cycle Triggered...." -ForegroundColor Cyan
                    Start-Sleep -Milliseconds 100
                    # Source Command List : https://docs.microsoft.com/en-us/mem/configmgr/develop/reference/core/clients/client-classes/triggerschedule-method-in-class-sms_client
                    # Refresh all computer policies / workstation certificates
                    Invoke-Command -ComputerName $theinfo -ScriptBlock { gpupdate /force /target:computer /wait:0 }
                    $msg = "This computer is scheduled for a restart in 20 minutes, please save all work."
                    $CurrentTime = Get-date
                    $DateNTime = $CurrentTime.ToLocalTime()
                    $FormattedTime = "$DateNTime"
                    Start-Sleep -Milliseconds 100
                    Write-Host "`nEstimated time to completion...`n20 minutes after: $FormattedTime" -ForegroundColor Cyan
                    shutdown -m \\$theinfo -r -f -c $msg -t 1200 <#20 minute delay to give SCCM Client time#> 
                    Start-Sleep -Milliseconds 100
                    Write-Host "`nLocal host scheduled to restart...`n20 minutes after: $FormattedTime" -ForegroundColor Cyan 
                    Start-Sleep -Milliseconds 100
                    Write-Host "`nFinished Process on local host : $LocalHostName" -ForegroundColor Green
                    break;
                }
                Else {
                    write-Host "`nLocal host denied by user and no other clients found..." -ForegroundColor DarkYellow
                    Start-Sleep -Seconds 1
                    Write-Host "`nFinished Process!" -ForegroundColor Green
                    break;
                }
            }
        }

})
$GForm.Controls.Add($ReFreshSoftwareCenterCommandButton)

#This Creates Button Clear Output
$ClearScreenButton = New-Object System.Windows.Forms.Button
$ClearScreenButton.Location = New-Object System.Drawing.Size(480,265)
$ClearScreenButton.Size = New-Object System.Drawing.Size(150,23)
$ClearScreenButton.BackColor = "LightGray"
$ClearScreenButton.Text = "Clear Console Output"
$ClearScreenButton.Add_Click({cls})
$GForm.Controls.Add($ClearScreenButton) 

#This creates a label for the Credits
$objLabel4 = New-Object System.Windows.Forms.Label
$objLabel4.Location = New-Object System.Drawing.Size(360,300) 
$objLabel4.Size = New-Object System.Drawing.Size(400,65)
$objLabel4.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel4.Text = "Development Team`n                SPC Burgess`n                           SPC Santiago"
$GForm.Controls.Add($objLabel4) 

###### FONT SIZE CHANGE:
$objLabel4.ForeColor = [System.Drawing.Color]::FromName("Black")
$objLabel4.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
$objLabel4.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
#endregion


$GForm.Add_Shown({$GForm.Activate()})
$GForm.ShowDialog() | Out-Null
$GForm.Dispose() | Out-Null

}
GeneralTool
