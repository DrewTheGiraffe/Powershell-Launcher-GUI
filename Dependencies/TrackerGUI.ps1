<# Script By ~ SPC BURGESS 2-3 FA S6
**************************************

    Tracker Script DOD Edition Only

Do NOT DISTROBUTE CODE OUTSIDE OF DOD
ORGANIZATIONS, ALL INFORMATION ON THIS
PAGE IS SUBJECT TO SEARCH & REVIEW BY
FORT BLISS NETWORK ENTERPRISE CENTER
PERSONEL AT ANY AND ALL TIMES. 

**************************************
#>

# Once you learn to use PSEXEC + Powershell everything else falls into place.. -SPC BURGESS




Function TrackerScript {
    cls
    Sleep 1
    Write-Host "Tracker Script Ready For User Interaction" -ForegroundColor Green
    Sleep 1
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.AnchorStyles")
    [void][System.Windows.Forms.Application]::EnableVisualStyles()


    # Draw Form
    $TRACKERForm = New-Object System.Windows.Forms.Form
    $TRACKERForm.Text = "[SA/WA] Tracker Script v4.0"
    $TRACKERForm.ClientSize = New-Object System.Drawing.Size(1200, 600)
    $TRACKERForm.BackColor = "LightGray"
    $TRACKERForm.StartPosition = "CenterScreen"
    $StopScript = $Script:CANCELED=$True
    $TRACKERForm.MaximizeBox = $false
    $TRACKERForm.FormBorderStyle = 'FixedDialog'
    $TRACKERForm.AccessibleDescription = "Simple GUI based PS Script to create local user accounts."

    # Draw Icon
    $iconConverted2Base64 = [Convert]::ToBase64String((Get-Content "C:\temp\Launcher\Dependencies\icon\NewPanda.ico" -Encoding Byte))
    $iconBase64           = $iconConverted2Base64
    $iconBytes            = [Convert]::FromBase64String($iconBase64)
    $stream               = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
    $stream.Write($iconBytes, 0, $iconBytes.Length);
    $iconImage            = [System.Drawing.Image]::FromStream($stream, $true)
    $TRACKERForm.Icon    = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
    # ico converter : https://cloudconvert.com/png-to-ico

    # This is defines what Enter does when pressed
    $TRACKERForm.KeyPreview = $True
    $TRACKERForm.Add_KeyDown({if ($_.KeyCode -eq "Enter") {
        # Currently Nothing..

    }})
    
    #This creates a label for the TRACKERForm
    $objLabel1 = New-Object System.Windows.Forms.Label
    $objLabel1.Location = New-Object System.Drawing.Size(5,2) 
    $objLabel1.Size = New-Object System.Drawing.Size(70,15)
    [String]$MandatoryWrite = "*" 
    $objLabel1.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $objLabel1.ForeColor = [System.Drawing.Color]::FromName("Blue")
    $objLabel1.Text = "OPTIONS $MandatoryWrite"
    $TRACKERForm.Controls.Add($objLabel1)

    # checkbox 1
    $Checkbox1 = New-Object System.Windows.Forms.CheckBox
    $Checkbox1.Location = New-Object System.Drawing.Point(145,22)
    $Checkbox1.Size = New-Object System.Drawing.Size(110,20)
    $Checkbox1.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $Checkbox1.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $Checkbox1.Text = "NOT SUPPORTED YET"
    $TRACKERForm.Controls.Add($Checkbox1)
    
    # checkbox General Compliance
    $CheckboxGeneralCompliance = New-Object System.Windows.Forms.CheckBox
    $CheckboxGeneralCompliance.Location = New-Object System.Drawing.Point(145,5)
    $CheckboxGeneralCompliance.size = New-Object System.Drawing.Size(205,20)
    $CheckboxGeneralCompliance.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CheckboxGeneralCompliance.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CheckboxGeneralCompliance.Text = "General Compliance"
    $TRACKERForm.Controls.Add($CheckboxGeneralCompliance) 

    # checkbox COMPUTER DESCRIPTIONS
    $CheckboxDESCRIPTIONS = New-Object System.Windows.Forms.CheckBox
    $CheckboxDESCRIPTIONS.Location = New-Object System.Drawing.Point(810,5)
    $CheckboxDESCRIPTIONS.size = New-Object System.Drawing.Size(200,20)
    $CheckboxDESCRIPTIONS.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CheckboxDESCRIPTIONS.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CheckboxDESCRIPTIONS.Text = "COMPUTER DESCRIPTIONS"
    $TRACKERForm.Controls.Add($CheckboxDESCRIPTIONS) 

    # checkbox COMPUTER IPV4 ADDRESSES
    $CheckboxIPV4 = New-Object System.Windows.Forms.CheckBox
    $CheckboxIPV4.Location = New-Object System.Drawing.Point(810,22)
    $CheckboxIPV4.Size = New-Object System.Drawing.Size(215,20)
    $CheckboxIPV4.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CheckboxIPV4.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CheckboxIPV4.Text = "COMPUTER IPV4 ADDRESSES"
    $TRACKERForm.Controls.Add($CheckboxIPV4)

    # Credits
    $bottomCREDITSLabel = New-Object System.Windows.Forms.Label
    $bottomCREDITSLabel.Location = New-Object System.Drawing.Point(915, 580) # 520,580 is bottom center @ 14 size font.
    $bottomCREDITSLabel.Size = New-Object System.Drawing.Size(275, 20)
    $bottomCREDITSLabel.ForeColor = "Black"
    $bottomCREDITSLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $bottomCREDITSLabel.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $bottomCREDITSLabel.Text = "AUTHOR: SPC DREW J. BURGESS"
    $TRACKERForm.Controls.Add($bottomCREDITSLabel)

    # creates the label title (BOTTOM CENTER)
    $bottomtitleLabel = New-Object System.Windows.Forms.Label
    $bottomtitleLabel.Location = New-Object System.Drawing.Point(490, 580) # 520,580 is bottom center @ 14 size font.
    $bottomtitleLabel.Size = New-Object System.Drawing.Size(205, 20)
    $bottomtitleLabel.ForeColor = "Black"
    $bottomtitleLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $bottomtitleLabel.Font = New-Object System.Drawing.Font("Lucida Console",14,[System.Drawing.FontStyle]::Regular)
    $bottomtitleLabel.Text = "Tracker Script V4"
    $TRACKERForm.Controls.Add($bottomtitleLabel)

    # STATIC STATUS LABEL
    $bottomSTATICSTATUSLabel = New-Object System.Windows.Forms.Label
    $bottomSTATICSTATUSLabel.Location = New-Object System.Drawing.Point(5, 580) # 520,580 is bottom center @ 14 size font.
    $bottomSTATICSTATUSLabel.Size = New-Object System.Drawing.Size(75, 20)
    $bottomSTATICSTATUSLabel.ForeColor = "Black"
    $bottomSTATICSTATUSLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $bottomSTATICSTATUSLabel.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $bottomSTATICSTATUSLabel.Text = "STATUS:"
    $TRACKERForm.Controls.Add($bottomSTATICSTATUSLabel)

    # DYNAMIC STATUS LABEL
    $bottomDYNAMICSTATUSLabel = New-Object System.Windows.Forms.Label
    $bottomDYNAMICSTATUSLabel.Location = New-Object System.Drawing.Point(80, 580) # 520,580 is bottom center @ 14 size font.
    $bottomDYNAMICSTATUSLabel.Size = New-Object System.Drawing.Size(150, 20)
    $bottomDYNAMICSTATUSLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $bottomDYNAMICSTATUSLabel.Font = New-Object System.Drawing.Font("Lucida Console",12,[System.Drawing.FontStyle]::Regular)
    $bottomDYNAMICSTATUSLabel.ForeColor = "Green"
    $bottomDYNAMICSTATUSLabel.Text = "READY!"
    $TRACKERForm.Controls.Add($bottomDYNAMICSTATUSLabel)

    # AD COMPUTER PATH QUERY BUTTON
    $ADPATHCOMPUTERBUTTON = New-Object System.Windows.Forms.Button
    $ADPATHCOMPUTERBUTTON.Location = New-Object System.Drawing.Point(600,25)
    $ADPATHCOMPUTERBUTTON.Size = New-Object System.Drawing.Size(200,23)
    $ADPATHCOMPUTERBUTTON.BackColor = "LightGray"
    $ADPATHCOMPUTERBUTTON.Text = "EDIT AD COMPUTER PATH"
    $TRACKERForm.Controls.Add($ADPATHCOMPUTERBUTTON) 

    # AD USER PATH QUERY BUTTON
    $ADPATHUSERBUTTON = New-Object System.Windows.Forms.Button
    $ADPATHUSERBUTTON.Location = New-Object System.Drawing.Point(350, 25)
    $ADPATHUSERBUTTON.Size = New-Object System.Drawing.Size(200, 23)
    $ADPATHUSERBUTTON.BackColor = "LightGray"
    $ADPATHUSERBUTTON.Text = "EDIT AD USER PATH"
    $TRACKERForm.Controls.Add($ADPATHUSERBUTTON)

    # USER Button
    $ButtonUSER = New-Object System.Windows.Forms.Button
    $ButtonUSER.Location = New-Object System.Drawing.Point(350, 2)
    $ButtonUSER.Size = New-Object System.Drawing.Size(200, 23)
    $ButtonUSER.BackColor = "LightGray"
    $ButtonUSER.Text = "QUERY USER INFO"
    $TRACKERForm.Controls.Add($ButtonUSER)

    # COMPUTER BUTTON
    $ButtonCOMPUTER = New-Object System.Windows.Forms.Button
    $ButtonCOMPUTER.Location = New-Object System.Drawing.Point(600, 2)
    $ButtonCOMPUTER.Size = New-Object System.Drawing.Size(200, 23)
    $ButtonCOMPUTER.BackColor = "LightGray"
    $ButtonCOMPUTER.Text = "QUERY COMPUTER INFO"
    $TRACKERForm.Controls.Add($ButtonCOMPUTER)

    # Computer output
    $outputComputerbox = New-Object System.Windows.Forms.TextBox
    $outputComputerbox.Location = New-Object System.Drawing.Point(575,75) 
    $outputComputerbox.Size = New-object System.Drawing.Size(600,500)
    $outputComputerbox.Multiline = $True
    $outputComputerbox.ScrollBars = "Vertical"
    $outputComputerbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $outputComputerbox.BackColor = "Black"
    $outputComputerbox.ForeColor = "White"
    $TRACKERForm.Controls.Add($outputComputerbox)

    # User output
    $outputuserbox = New-Object System.Windows.Forms.TextBox
    $outputuserbox.Location = New-Object System.Drawing.Point(25,75)
    $outputuserbox.Size = New-object System.Drawing.Size(550,500) 
    $outputuserbox.Multiline = $True
    $outputuserbox.ScrollBars = "Vertical"
    $outputuserbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $outputuserbox.BackColor = "Black"
    $outputuserbox.ForeColor = "White"
    $TRACKERForm.Controls.Add($outputuserbox)

    # Append data to outputbox
    Function Add-OutboxlineUser {
        Param ($NewData)
        $outputuserbox.AppendText("`r`n$NewData")
        $outputuserbox.Refresh()
        $outputuserbox.ScrollToCaret()
    }

    # Append data to outputbox
    Function Add-OutboxlineComputer {
        Param ($NewData)
        $outputComputerbox.AppendText("`r`n$NewData")
        $outputComputerbox.Refresh()
        $outputComputerbox.ScrollToCaret()
    }

    $ADPATHCOMPUTERBUTTON.Add_Click({
        Start-Sleep -Milliseconds 100
        Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\ADCOMPUTERPATHQUERY.txt"
    })

    $ADPATHUSERBUTTON.Add_Click({
        Start-Sleep -Milliseconds 100
        Start-Process -Wait -PSPath "notepad.exe" -ArgumentList "C:\temp\Launcher\Logs\ADUSERPATHQUERY.txt"
    })

    # USER Button CLICK (MAIN BACKEND CODE)
    $ButtonUSER.Add_Click({
        $bottomDYNAMICSTATUSLabel.ForeColor = "DarkRed"
        $bottomDYNAMICSTATUSLabel.Text = "IN PROGRESS..."
        Sleep 1
        Add-OutboxlineUser -NewData "***************** Script Started DO NOT SPAM!! *****************"
        $PathofUsers = Get-Content -Path "C:\temp\Launcher\Logs\ADUSERPATHQUERY.txt" -Force
        If ($CheckboxGeneralCompliance.Checked -eq $True) {
            $ADINFO = ((Get-ADUser -Filter '*' -SearchBase $PathofUsers -Properties * | Select -Property GivenName,Surname,EmailAddress,DisplayName,Enabled,SamAccountName,LastLogonDate | out-string) -replace "\s+:","") -split "\r\n"
        }
        Else {
            $ADINFO = $null
        }
        If ($ADINFO -cne $null) {
            
            Add-OutboxlineUser -NewData "******************** START LOG ********************"
            
            Foreach ($Line in $ADINFO) {
        
                Add-OutboxlineUser -NewData $Line
        
                Start-sleep -Milliseconds 100
        
            }
            
            Add-OutboxlineUser -NewData "********************* END LOG *********************"
        
        }
        Else {
            Add-OutboxlineUser -NewData "************************************************************
               
    Script Created By SPC Burgess 2-3 FA S6 Fort Bliss TX       
               ************************************************************"
            Add-OutboxlineUser -NewData "*************** Script Finished ***************"
        }
        $bottomDYNAMICSTATUSLabel.ForeColor = "Green"
        $bottomDYNAMICSTATUSLabel.Text = "READY!"
    })

    # COMPUTER Button CLICK (MAIN BACKEND CODE)
    $ButtonCOMPUTER.Add_Click({
        $bottomDYNAMICSTATUSLabel.ForeColor = "DarkRed"
        $bottomDYNAMICSTATUSLabel.Text = "IN PROGRESS..."
        Sleep 1
        Add-OutboxlineComputer -NewData "******************** Script Started DO NOT SPAM!! ********************"
        $PathofcOMPUTERS = Get-Content -Path "C:\temp\Launcher\Logs\ADCOMPUTERPATHQUERY.txt" -Force
        If ($CheckboxDESCRIPTIONS.Checked -eq $True) {
            $ADINFO = ((Get-ADComputer -Filter '*' -SearchBase $PathofcOMPUTERS -prop description | select name,description | out-string) -replace "\s+:","") -split "\r\n"
        }
        Elseif ($CheckboxIPV4.Checked -eq $True) {
            $ADINFO = ((Get-ADComputer -Filter '*' -SearchBase $PathofcOMPUTERS -prop IPv4Address | select name,ipv4address | out-string) -replace "\s+:","") -split "\r\n"
        }
        Else {
            $ADINFO = $null
        }
        If ($ADINFO -cne $null) {
            
            Add-OutboxlineComputer -NewData "******************** START LOG ********************"
            
            Foreach ($Line in $ADINFO) {
        
                Add-OutboxlineComputer -NewData $Line
        
                Start-sleep -Milliseconds 100
        
            }
            
            Add-OutboxlineComputer -NewData "********************* END LOG *********************"
        
        }
        Else {
            Add-OutboxlineComputer -NewData "************************************************************
               
    Script Created By SPC Burgess 2-3 FA S6 Fort Bliss TX       
               ************************************************************"
            Add-OutboxlineComputer -NewData "*************** Script Finished ***************"
        }
        $bottomDYNAMICSTATUSLabel.ForeColor = "Green"
        $bottomDYNAMICSTATUSLabel.Text = "READY!"
    })
    

    # Finally Render it all.
    $TRACKERForm.Add_Shown({$TRACKERForm.Activate()})
    $TRACKERForm.ShowDialog() | Out-Null
    $TRACKERForm.Dispose() | Out-Null

}

TrackerScript

# Sources : https://community.spiceworks.com/topic/2127610-redirect-all-script-output-into-forms-textbox