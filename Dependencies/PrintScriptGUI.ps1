<# Script By ~ SPC BURGESS 2-3 FA S6
**************************************

    Printer Script DOD Edition Only

Do NOT DISTROBUTE CODE OUTSIDE OF DOD
ORGANIZATIONS, ALL INFORMATION ON THIS
PAGE IS SUBJECT TO SEARCH & REVIEW BY
FORT BLISS NETWORK ENTERPRISE CENTER
PERSONEL AT ANY AND ALL TIMES. 

**************************************
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

    # CMD 
    $CMDTeamPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $CMDTeamPrinterCheckbox.Location = New-Object System.Drawing.Point(10,55)
    $CMDTeamPrinterCheckbox.Size = New-Object System.Drawing.Size(298,20)
    $CMDTeamPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CMDTeamPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CMDTeamPrinterCheckbox.Text = "2-3 FA CMD Printer"
    $ZipExtractorForm.Controls.Add($CMDTeamPrinterCheckbox)
    
    # S1
    $S1PrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $S1PrinterCheckbox.Location = New-Object System.Drawing.Point(10,75)
    $S1PrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $S1PrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $S1PrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $S1PrinterCheckbox.Text = "2-3 FA S1 Printer"
    $ZipExtractorForm.Controls.Add($S1PrinterCheckbox) 
   
    # S3
    $S3PrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $S3PrinterCheckbox.Location = New-Object System.Drawing.Point(10,95)
    $S3PrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $S3PrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $S3PrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $S3PrinterCheckbox.Text = "2-3 FA S3 Printer"
    $ZipExtractorForm.Controls.Add($S3PrinterCheckbox)

    # S6
    $S6PrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $S6PrinterCheckbox.Location = New-Object System.Drawing.Point(10,115)
    $S6PrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $S6PrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $S6PrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $S6PrinterCheckbox.Text = "2-3 FA S6 Printer"
    $ZipExtractorForm.Controls.Add($S6PrinterCheckbox)

    # HHB
    $HHBPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $HHBPrinterCheckbox.Location = New-Object System.Drawing.Point(10,135)
    $HHBPrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $HHBPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $HHBPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $HHBPrinterCheckbox.Text = "2-3 FA HHB Printer"
    $ZipExtractorForm.Controls.Add($HHBPrinterCheckbox)

    # ABAT
    $ABATPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $ABATPrinterCheckbox.Location = New-Object System.Drawing.Point(10,155)
    $ABATPrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $ABATPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $ABATPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $ABATPrinterCheckbox.Text = "2-3 FA ALPHA Printer"
    $ZipExtractorForm.Controls.Add($ABATPrinterCheckbox)

    # BBAT
    $BBATPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $BBATPrinterCheckbox.Location = New-Object System.Drawing.Point(10,175)
    $BBATPrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $BBATPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $BBATPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $BBATPrinterCheckbox.Text = "2-3 FA BRAVO Printer"
    $ZipExtractorForm.Controls.Add($BBATPrinterCheckbox)

    # CBAT
    $CBATPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $CBATPrinterCheckbox.Location = New-Object System.Drawing.Point(10,195)
    $CBATPrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $CBATPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $CBATPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $CBATPrinterCheckbox.Text = "2-3 FA CHARLIE Printer"
    $ZipExtractorForm.Controls.Add($CBATPrinterCheckbox)

    # FBAT
    $FBATPrinterCheckbox = New-Object System.Windows.Forms.CheckBox
    $FBATPrinterCheckbox.Location = New-Object System.Drawing.Point(10,215)
    $FBATPrinterCheckbox.size = New-Object System.Drawing.Size(298,20)
    $FBATPrinterCheckbox.BackColor = [System.Drawing.Color]::FromKnownColor("Transparent")
    $FBATPrinterCheckbox.Font = New-Object System.Drawing.Font("Lucida Console",10,[System.Drawing.FontStyle]::Regular)
    $FBATPrinterCheckbox.Text = "2-3 FA FOX Printer"
    $ZipExtractorForm.Controls.Add($FBATPrinterCheckbox)

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
            Write-Host "************************************************************`n`n   Script Created By SPC Burgess 2-3 FA S6 Fort Bliss TX`n`n************************************************************" -ForegroundColor Cyan      
            
                # ID NAMES
                $S3_Port_ID        = "S3 Printer"              
                $S1_Port_ID        = "S1 Printer"                                  
                $S6_Port_ID        = "S6 Printer"              
                $CMD_Port_ID       = "CMD Printer"
                $HHB_Port_ID       = "HHB Printer" 
                $Alpha_Port_ID     = "Alpha Printer" 
                $Bravo_Port_ID     = "Bravo Printer"
                $Charlie_Port_ID   = "Charlie Btry Printer"
                $Fox_Port_ID       = "Fox Printer"
                $CustomPortID      = $CustomPrinterPortIdNameTextbox.Text
                 
                # Port Addresses
                $S3_Port    = "143.78.2.200"               
                $S1_Port    = "143.78.2.53"                
                $S6_Port    = "143.78.2.222" 
                $HHB_Port   = "143.78.3.106"              
                $CMD_Port   = "143.78.3.126" 
                $ABTRY_Port = "143.78.2.50" 
                $BBTRY_Port = "143.78.2.157"
                $CBTRY_Port = "143.78.2.80" 
                $FBTRY_Port = "143.78.3.86"
                $CustomPort = $CustomPrinterIPTextbox.Text
                
                # Printer Names
                $S1                = "2-3 FA S1 Printer"
                $S3                = "2-3 FA S3 Printer"
                $S6               = "2-3 FA S6 Printer"
                $CMD               = "2-3 FA CMD Team Printer"
                $HHB               = "2-3 FA HHB Battery Printer"
                $Alpha             = "2-3 FA ALPHA Battery Printer"
                $Bravo             = "2-3 FA BRAVO Battery Printer"
                $Charlie           = "2-3 FA CHARLIE Battery Printer"
                $Fox               = "2-3 FA FOX Battery Printer" 
                $CustomPrintername = $CustomPrinterNameTextbox.Text

            If ($CMDTeamPrinterCheckbox.Checked -eq $True) {

                If (!(Get-PrinterPort -Name $CMD_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $CMD_Port_ID -PrinterHostAddress $CMD_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $CMD -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $CMD -PortName $CMD_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $CMD to remote host : $Computers" -ForegroundColor Green  
                }  
            }

            If ($S1PrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $S1_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $CMD_Port_ID -PrinterHostAddress $S1_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $S1 -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $S1 -PortName $S1_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $S1 to remote host : $Computers" -ForegroundColor Green  
                } 
            }
            
            If ($S3PrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $S3_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $S3_Port_ID -PrinterHostAddress $S3_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $S3 -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $S3 -PortName $S3_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $S3 to remote host : $Computers" -ForegroundColor Green  
                } 
            }
            
            If ($S6PrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $S6_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $S6_Port_ID -PrinterHostAddress $S6_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $S6 -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $S6 -PortName $S6_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $S6 to remote host : $Computers" -ForegroundColor Green  
                } 
            }

            If ($HHBPrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $HHB_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $HHB_Port_ID -PrinterHostAddress $HHB_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $HHB -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $HHB -PortName $HHB_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $HHB to remote host : $Computers" -ForegroundColor Green  
                } 
            }

            If ($ABATPrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $Alpha_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $Alpha_Port_ID -PrinterHostAddress $ABTRY_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $Alpha  -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $Alpha -PortName $Alpha_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $Alpha to remote host : $Computers" -ForegroundColor Green  
                }
            }

            If ($BBATPrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $Bravo_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $Bravo_Port_ID -PrinterHostAddress $BBTRY_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $Bravo -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $Bravo -PortName $Bravo_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $Bravo to remote host : $Computers" -ForegroundColor Green  
                }
            }

            If ($CBATPrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $Charlie_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $Charlie_Port_ID -PrinterHostAddress $CBTRY_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $Charlie -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $Charlie -PortName $Charlie_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $Charlie to remote host : $Computers" -ForegroundColor Green  
                }
            }

            If ($FBATPrinterCheckbox.Checked -eq $True) {
                If (!(Get-PrinterPort -Name $Fox_Port_ID -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorPortString = "Adding IP Address to corresponding port ID" 
                    Add-PrinterPort -Name $Fox_Port_ID -PrinterHostAddress $FBTRY_Port -ComputerName $Computers
                    Start-Sleep 1
                }
                If (!(Get-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers -ErrorAction Ignore)) {
                    $ErrorDriverString = "`nAdding Driver to DriverStore"
                    Add-PrinterDriver -Name "Microsoft PCL6 Class Driver" -ComputerName $Computers 
                    Start-Sleep 2
                }
                If (!(Get-Printer -ComputerName $Computers -Name $Fox -ErrorAction Ignore)) {
                    Add-Printer -computername $Computers -name $Fox -PortName $Fox_Port_ID -DriverName "Microsoft PCL6 Class Driver"
                    Write-Host "Finished Adding $Fox to remote host : $Computers" -ForegroundColor Green  
                }
            }
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
            
                # ID NAMES
                $S3_Port_ID        = "S3 Printer"              
                $S1_Port_ID        = "S1 Printer"                                  
                $S6_Port_ID        = "S6 Printer"              
                $CMD_Port_ID       = "CMD Printer"
                $HHB_Port_ID       = "HHB Printer" 
                $Alpha_Port_ID     = "Alpha Printer" 
                $Bravo_Port_ID     = "Bravo Printer"
                $Charlie_Port_ID   = "Charlie Btry Printer"
                $Fox_Port_ID       = "Fox Printer"
                $CustomPortID      = $CustomPrinterPortIdNameTextbox.Text
                 
                # Port Addresses
                $S3_Port    = "143.78.2.200"               
                $S1_Port    = "143.78.2.53"                
                $S6_Port    = "143.78.2.222" 
                $HHB_Port   = "143.78.3.106"              
                $CMD_Port   = "143.78.3.126" 
                $ABTRY_Port = "143.78.2.50" 
                $BBTRY_Port = "143.78.2.157"
                $CBTRY_Port = "143.78.2.80" 
                $FBTRY_Port = "143.78.3.86"
                $CustomPort = $CustomPrinterIPTextbox.Text
                
                # Printer Names
                $S1                = "2-3 FA S1 Printer"
                $S3                = "2-3 FA S3 Printer"
                $S6                = "2-3 FA S6 Printer"
                $CMD               = "2-3 FA CMD Team Printer"
                $HHB               = "2-3 FA HHB Battery Printer"
                $Alpha             = "2-3 FA ALPHA Battery Printer"
                $Bravo             = "2-3 FA BRAVO Battery Printer"
                $Charlie           = "2-3 FA CHARLIE Battery Printer"
                $Fox               = "2-3 FA FOX Battery Printer" 
                $CustomPrintername = $CustomPrinterNameTextbox.Text

            If ($CMDTeamPrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $CMD -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $CMD_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $CMD 
                    Write-Host "Finished Removing $CMD from remote host : $Computers" -ForegroundColor Green  
                }  
            }

            If ($S1PrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $S1 -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $S1_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $S1 
                    Write-Host "Finished Removing $S1 from remote host : $Computers" -ForegroundColor Green  
                } 
            }
            
            If ($S3PrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $S3 -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $S3_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $S3 
                    Write-Host "Finished Removing $S3 from remote host : $Computers" -ForegroundColor Green  
                } 
            }
            
            If ($S6PrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $S6 -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $S6_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $S6 
                    Write-Host "Finished Removing $S6 from remote host : $Computers" -ForegroundColor Green 
                } 
            }

            If ($HHBPrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $HHB -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $HHB_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $HHB 
                    Write-Host "Finished Removing $HHB from remote host : $Computers" -ForegroundColor Green 
                } 
            }

            If ($ABATPrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $Alpha  -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $Alpha_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $Alpha 
                    Write-Host "Finished Removing $Alpha from remote host : $Computers" -ForegroundColor Green  
                }
            }

            If ($BBATPrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $Bravo -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $Bravo_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $Bravo
                    Write-Host "Finished Removing $Bravo from remote host : $Computers" -ForegroundColor Green  
                }
            }

            If ($CBATPrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $Charlie -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $Charlie_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $Charlie
                    Write-Host "Finished Removing $Charlie from remote host : $Computers" -ForegroundColor Green  
                }
            }

            If ($FBATPrinterCheckbox.Checked -eq $True) {
                If ((Get-Printer -ComputerName $Computers -Name $Fox -ErrorAction Ignore)) {
                    Remove-PrinterPort -ComputerName $Computers -Name $Fox_Port_ID -ErrorAction Ignore
                    Remove-Printer -ComputerName $Computers -Name $Fox
                    Write-Host "Finished Removing $Fox from remote host : $Computers" -ForegroundColor Green  
                }
            }
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