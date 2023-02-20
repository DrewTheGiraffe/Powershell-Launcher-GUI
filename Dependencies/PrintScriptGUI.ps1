# Printer Script By Drew Burgess 02/20/2023
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

Function Get-FileName($initialDirectory, $FileExtension)
{  
 [System.Reflection.Assembly]::LoadWithPartialName(“System.windows.forms”) | Out-Null

 $FileType = "$FileExtension"
 $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
 $OpenFileDialog.initialDirectory = $initialDirectory
 $OpenFileDialog.filter = “(*.$FileType)|*.$FileType*" # All files (*.*)| *.*”
 $OpenFileDialog.ShowDialog() | Out-Null
 $OpenFileDialog.filename
 $FileLoc = $OpenFileDialog.OpenFile() | Select-String -SimpleMatch "Name"
 $FileLoc | Out-String
} #Source : https://devblogs.microsoft.com/scripting/hey-scripting-guy-can-i-open-a-file-dialog-box-with-windows-powershell/
# EG: CALL IT LIKE THIS: Get-FileName -initialDirectory "C:\" -FileExtension "INF"


Function PrinterScript {
    cls
    Sleep 1
    Write-Host "Print Script Ready For User Interaction" -ForegroundColor Green
    Sleep 1
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
    [void][System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms.AnchorStyles")
    [void][System.Windows.Forms.Application]::EnableVisualStyles()


    # Draw Form
    $ZipExtractorForm = New-Object System.Windows.Forms.Form
    $ZipExtractorForm.Text = "[SA/WA] Print Script v4.0"
    $ZipExtractorForm.ClientSize = New-Object System.Drawing.Size(400, 280) # 270 without custom printer elements
    $ZipExtractorForm.BackColor = "LightGray"
    $ZipExtractorForm.StartPosition = "CenterScreen"
    $StopScript = $Script:CANCELED=$True
    $ZipExtractorForm.MaximizeBox = $false
    $ZipExtractorForm.AccessibleDescription = "Simple GUI based PS Script to add printers to computers remotely."


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
    $titleLabel.Text = "Print Script V4"
    $ZipExtractorForm.Controls.Add($titleLabel)

    # CUSTOM
    $CUSTOMPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $CUSTOMPrinterCheckbox.Location = New-Object System.Drawing.Point(10,235)
    $CUSTOMPrinterCheckbox.size = New-Object System.Drawing.Size(210,20)
    $CUSTOMPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CUSTOMPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CUSTOMPrinterCheckbox.Text = "CUSTOM Printer"
    $ZipExtractorForm.Controls.Add($CUSTOMPrinterCheckbox)

    # CUSTOM LABEL ADDON
    $CUSTOMPRINTERLABELADDON = New-Object System.Windows.Forms.Label
    $CUSTOMPRINTERLABELADDON.Location = New-Object System.Drawing.Point(215,235)
    $CUSTOMPRINTERLABELADDON.Size = New-Object System.Drawing.Size(160,20)
    $CUSTOMPRINTERLABELADDON.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CUSTOMPRINTERLABELADDON.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CUSTOMPRINTERLABELADDON.Text = "[Click Refresh]"
    $ZipExtractorForm.Controls.Add($CUSTOMPRINTERLABELADDON)

    # CUSTOM DRIVER BUTTON
    $CUSTOMPRINTERDRIVERBUTTON = New-Object System.Windows.Forms.CheckBox
    $CUSTOMPRINTERDRIVERBUTTON.Location = New-Object System.Drawing.Point(220,235)
    $CUSTOMPRINTERDRIVERBUTTON.Size = New-Object System.Drawing.Size(180,20)
    $CUSTOMPRINTERDRIVERBUTTON.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CUSTOMPRINTERDRIVERBUTTON.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CUSTOMPRINTERDRIVERBUTTON.Text = "Import Driver"

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

    # Custom Printer Name Label
    $CustomPrinterNameLabel = New-Object System.Windows.Forms.Label
    $CustomPrinterNameLabel.Size = New-Object System.Drawing.Size(230,20)
    $CustomPrinterNameLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CustomPrinterNameLabel.ForeColor = [System.Drawing.Color]::FromName("Red")
    $CustomPrinterNameLabel.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CustomPrinterNameLabel.Text = "Enter Printer Name Here $MandatoryWrite"

    # Custom Printer Name Textbox
    $CustomPrinterNameTextbox = New-Object System.Windows.Forms.TextBox
    $CustomPrinterNameTextbox.Size = New-Object System.Drawing.Size(200,20)
    
    # Custom Printer IP Address
    $CustomPrinterIPLabel = New-Object System.Windows.Forms.Label
    $CustomPrinterIPLabel.Size = New-Object System.Drawing.Size(290,20)
    $CustomPrinterIPLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CustomPrinterIPLabel.ForeColor = [System.Drawing.Color]::FromName("Red")
    $CustomPrinterIPLabel.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CustomPrinterIPLabel.Text = "Enter Printer IPV4 Address Here $MandatoryWrite"

    # Custom Printer IP Textbox
    $CustomPrinterIPTextbox = New-Object System.Windows.Forms.TextBox
    $CustomPrinterIPTextbox.Size = New-Object System.Drawing.Size(120,20)

    # Custom Printer Port ID Name Address
    $CustomPrinterPortIdNameLabel = New-Object System.Windows.Forms.Label
    $CustomPrinterPortIdNameLabel.Size = New-Object System.Drawing.Size(290,20)
    $CustomPrinterPortIdNameLabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CustomPrinterPortIdNameLabel.ForeColor = [System.Drawing.Color]::FromName("Red")
    $CustomPrinterPortIdNameLabel.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CustomPrinterPortIdNameLabel.Text = "Enter Printer Port ID Name Here $MandatoryWrite"

    # Custom Printer Port ID Textbox
    $CustomPrinterPortIdNameTextbox = New-Object System.Windows.Forms.TextBox
    $CustomPrinterPortIdNameTextbox.Size = New-Object System.Drawing.Size(200,20)

    # Custom Printer Port ID Name Address
    $CustomPrinterNOTELabel = New-Object System.Windows.Forms.Label
    $CustomPrinterNOTELabel.Size = New-Object System.Drawing.Size(290,40)
    $CustomPrinterNOTELabel.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CustomPrinterNOTELabel.ForeColor = [System.Drawing.Color]::FromName("Blue")
    $CustomPrinterNOTELabel.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CustomPrinterNOTELabel.Text = "$MandatoryWrite NOTE: Do not name Port ID Name the same as Printer name or script will throw lots of errors"
    $ButtonSizeXcoordinate = 88
    
    # Start Button
    $ButtonStart = New-Object System.Windows.Forms.Button
    $ButtonStart.Location = New-Object System.Drawing.Point(2, 255) # + 40 for y coord from f battery checkbox
    $ButtonStart.Size = New-Object System.Drawing.Size($ButtonSizeXcoordinate, 23)
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
        $Script:CANCELED=$True
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
            Write-Host "************************************************************`n`n   Script Created By Drew Burgess   `n`n************************************************************" -ForegroundColor Cyan      
            
                # ID NAMES (Print Server Aliase)
                $CustomPortID      = $CustomPrinterPortIdNameTextbox.Text
                 
                # Port Addresses (IPV4 Addresses)
                $CustomPort = $CustomPrinterIPTextbox.Text
                
                # Printer Names (Front End Aliases)
                $CustomPrintername = $CustomPrinterNameTextbox.Text
                
            If ($CUSTOMPrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $CustomPortID -ComputerName $Computers -ErrorAction Ignore)) {
                    Add-PrinterPort -Name $CustomPortID -PrinterHostAddress $CustomPort -ComputerName $Computers
                    Start-Sleep -Seconds 1
                }
                
                If ($CUSTOMPRINTERDRIVERBUTTON.Checked -eq $True) {
                    $HostnameInput = $objZipExtractorTextBox1.Text
                    $CustomDriverLocation = Get-FileName -initialDirectory "C:\" -FileExtension "INF"
                    PnPUtil /add-driver "$CustomDriverLocation"
                    Start-Sleep -Seconds 2 
                    If (!(Get-PrinterDriver -ComputerName $Computers -Name $CustomPrintername -ErrorAction Ignore)) {
                    #This whole code needs to just be re-worked.. Best option would be to add either a queried list in a text box of all installed drivers or make the user input the specific driver name to use in coordination with inf file...
                        $DriverName = PnPUtil /add-driver "$CustomDriverLocation" /install #Get-ChildItem -Path "$CustomDriverLocation" -Recurse -Filter "*.INF" | ForEach-Object { PnPUtil /add-driver $_.FullName /install } # Kudos to Stack Overflow for the PnPUtil object help...
                        Add-Printer -ComputerName $Computers -Name $CustomPrintername -PortName $CustomPortID -DriverName $DriverName #$DriverName # Ugh I absolutely hate printers.. this might work but also might not..
                    } 
                }
                Else {
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep -Seconds 2
                    If (!(Get-Printer -ComputerName $Computers -Name $CustomPrintername -ErrorAction Ignore)) {
                        Add-Printer -ComputerName $Computers -Name $CustomPrintername -PortName $CustomPortID -DriverName "Microsoft PCL6 Class Driver"
                    }
                }
            }
        }
  }
})
    $ZipExtractorForm.Controls.Add($ButtonStart)

    # Cancel Button
    $ButtonStop = New-Object System.Windows.Forms.Button
    $ButtonStop.Location = New-Object System.Drawing.Point(92, 255) # + 40 for y coord from f battery checkbox
    $ButtonStop.Size = New-Object System.Drawing.Size($ButtonSizeXcoordinate, 23)
    $ButtonStop.BackColor = "LightGray"
    $ButtonStop.Text = "CANCEL"
    $ButtonStop.Add_Click({write-Host "`nScript Canceled" -ForegroundColor Yellow;Sleep 1;write-host "`nScript Closing Safely" -ForegroundColor Cyan;Sleep 1;$Script:CANCELED=$True;cls;$ZipExtractorForm.Close()})
    $ZipExtractorForm.Controls.Add($ButtonStop)

    # Refresh Button
    $RefreshButton = New-Object System.Windows.Forms.Button
    $RefreshButton.Location = New-Object System.Drawing.Point(182,255)
    $RefreshButton.Size = New-Object System.Drawing.Point($ButtonSizeXcoordinate,23)
    $RefreshButton.BackColor = "LightGray"
    $RefreshButton.Text = "REFRESH"
    $RefreshButton.Add_Click({
        # Draw new elems
        If ($CUSTOMPrinterCheckbox.Checked -eq $True) {
            
            # Remove Label 
            $ZipExtractorForm.Controls.Remove($CUSTOMPRINTERLABELADDON)

            # New State elems
            $ZipExtractorForm.ClientSize = New-Object System.Drawing.Size(400, 450)
            $ButtonStop.Location = New-Object System.Drawing.Point(92, 425)
            $ButtonStart.Location = New-Object System.Drawing.Point(2, 425)
            $RefreshButton.Location = New-Object System.Drawing.Point(182,425) 
            $ZipExtractorForm.Controls.Add($CUSTOMPRINTERDRIVERBUTTON)
            $RemovePrinterButton.Location = New-Object System.Drawing.Point(270,425)

            # Printer Name Label & Textbox
            $CustomPrinterNameLabel.Location = New-Object System.Drawing.Size(10,255)
            $ZipExtractorForm.Controls.Add($CustomPrinterNameLabel)

            $CustomPrinterNameTextbox.Location = New-Object System.Drawing.Size(10,275)
            $ZipExtractorForm.Controls.Add($CustomPrinterNameTextbox)

            $CustomPrinterIPLabel.Location = New-Object System.Drawing.Point(10,295)
            $ZipExtractorForm.Controls.Add($CustomPrinterIPLabel)

            $CustomPrinterIPTextbox.Location = New-Object System.Drawing.Point(10,315)
            $ZipExtractorForm.Controls.Add($CustomPrinterIPTextbox)

            $CustomPrinterPortIdNameLabel.Location = New-Object System.Drawing.Point(10,335)
            $ZipExtractorForm.Controls.Add($CustomPrinterPortIdNameLabel)

            $CustomPrinterPortIdNameTextbox.Location = New-Object System.Drawing.Point(10,355)
            $ZipExtractorForm.Controls.Add($CustomPrinterPortIdNameTextbox)

            $CustomPrinterNOTELabel.Location = New-Object System.Drawing.Point(10,375)
            $ZipExtractorForm.Controls.Add($CustomPrinterNOTELabel)
            
        }  
        Else {
            
            # Original State Elems
            $ZipExtractorForm.ClientSize = New-Object System.Drawing.Size(400, 280)
            $ButtonStop.Location = New-Object System.Drawing.Point(92, 255)
            $ButtonStart.Location = New-Object System.Drawing.Point(2, 255)
            $RefreshButton.Location = New-Object System.Drawing.Point(182,255)
            $RemovePrinterButton.Location = New-Object System.Drawing.Point(270,255)
            $ZipExtractorForm.Controls.Add($CUSTOMPRINTERLABELADDON)
            
            # Remove All new elements
            $ZipExtractorForm.Controls.Remove($CustomPrinterNameLabel)
            $ZipExtractorForm.Controls.Remove($CustomPrinterNameTextbox)
            $ZipExtractorForm.Controls.Remove($CustomPrinterIPLabel)
            $ZipExtractorForm.Controls.Remove($CustomPrinterIPTextbox)
            $ZipExtractorForm.Controls.Remove($CustomPrinterPortIdNameLabel)
            $ZipExtractorForm.Controls.Remove($CustomPrinterPortIdNameTextbox)
            $ZipExtractorForm.Controls.Remove($CustomPrinterNOTELabel)
            $ZipExtractorForm.Controls.Remove($CUSTOMPRINTERDRIVERBUTTON)

        }


        # refresh gui state
        $ZipExtractorForm.Refresh()
        [System.Windows.Forms.Application]::DoEvents() 
    })
    $ZipExtractorForm.Controls.Add($RefreshButton)

    # Delete Selected Printer Button
    $RemovePrinterButton = New-Object System.Windows.Forms.Button 
    $RemovePrinterButton.Location = New-Object System.Drawing.Point(270,255)
    $RemovePrinterButton.Size = New-Object System.Drawing.Size(127,23)
    $RemovePrinterButton.BackColor = "LightGray"
    $RemovePrinterButton.Text = "DELETE PRINTER"
    $RemovePrinterButton.Add_Click({
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
        $Script:CANCELED=$True
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
            Write-Host "************************************************************`n`n   Script Created By SPC Burgess 2-3 FA S6 Fort Bliss TX`n`n************************************************************" -ForegroundColor Cyan      
            
                # ID NAMES (Print Server Aliase)
                $CustomPortID      = $CustomPrinterPortIdNameTextbox.Text
                 
                # Port Addresses (IPV4 Addresses)
                $CustomPort = $CustomPrinterIPTextbox.Text
                
                # Printer Names (Front End Aliases)
                $CustomPrintername = $CustomPrinterNameTextbox.Text

            If ($CUSTOMPrinterCheckbox.Checked -eq $True) {
               If ((Get-Printer -ComputerName $Computers -Name $CustomPrintername -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $CustomPort -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $CustomPrintername
                    Write-Host "Finished Removing $CustomPrintername from remote host : $Computers" -ForegroundColor Green
               }                 
            }
        }
  }
}) 
    $ZipExtractorForm.Controls.Add($RemovePrinterButton)

    # Finally Render it all.
    $ZipExtractorForm.Add_Shown({$ZipExtractorForm.Activate()})
    $ZipExtractorForm.ShowDialog() | Out-Null
    $ZipExtractorForm.Dispose() | Out-Null
    

}

PrinterScript
