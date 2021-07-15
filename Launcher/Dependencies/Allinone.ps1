function CreateForm {
#[reflection.assembly]::loadwithpartialname("System.Windows.Forms") | Out-Null
#[reflection.assembly]::loadwithpartialname("System.Drawing") | Out-Null

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.drawing

#Form Setup
$form1 = New-Object System.Windows.Forms.Form
#$anything = new-object System.Windows.Forms.Page
$outputBox = New-Object System.Windows.Forms.TextBox
$Gbutton = New-Object System.Windows.Forms.Button
$Hostnamebutton = New-Object System.Windows.Forms.Button
$ADUserbutton = New-Object System.Windows.Forms.Button
$Appbutton = New-Object System.Windows.Forms.Button
$Zipbutton = New-Object System.Windows.Forms.Button
$button1 = New-Object System.Windows.Forms.Button
$button2 = New-Object System.Windows.Forms.Button
$button3 = New-Object System.Windows.Forms.Button
$button4 = New-Object System.Windows.Forms.Button
$button5 = New-Object System.Windows.Forms.Button
$button6 = New-Object System.Windows.Forms.Button
$button7 = New-Object System.Windows.Forms.Button
$button8 = New-Object System.Windows.Forms.Button
$button9 = New-Object System.Windows.Forms.Button
$button1Bit = New-Object System.Windows.Forms.Button
$button2Bit = New-Object System.Windows.Forms.Button
$button3Bit = New-Object System.Windows.Forms.Button
$button4Bit = New-Object System.Windows.Forms.Button
$button5Bit = New-Object System.Windows.Forms.Button
$ADBackupBox = New-Object System.Windows.Forms.TextBox
$button1Users = New-Object System.Windows.Forms.Button
$button2Users = New-Object System.Windows.Forms.Button
$button3Users = New-Object System.Windows.Forms.Button
$button4Users = New-Object System.Windows.Forms.Button
$button5Users = New-Object System.Windows.Forms.Button
$DefaultThemeButton = New-Object System.Windows.Forms.Button
$DarkThemeButton = New-Object System.Windows.Forms.Button
$LightThemeButton = New-Object System.Windows.Forms.Button
$TabControl = New-object System.Windows.Forms.TabControl
$TroubleshootingPage = New-Object System.Windows.Forms.TabPage
$CPUPage = New-Object System.Windows.Forms.TabPage
$BitlockerPage = New-Object System.Windows.Forms.TabPage
$UsersPage = New-Object System.Windows.Forms.TabPage

$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState

#Form Parameter
$form1.Text = "General Tech"
$form1.Name = "form1"
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$form1.BackColor = "White"
$System_Drawing_Size.Width = 900
$System_Drawing_Size.Height = 700
$form1.FormBorderStyle = 'Fixed3D'
$form1.ClientSize = $System_Drawing_Size

#region logo

# Draws Logo
$img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\Panda\NewPanda.png')
$form1.BackgroundImage = $img
$form1.BackgroundImageLayout = "Center"

#endregion

#region tab_control
#Tab Control 
$tabControl.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 75
$System_Drawing_Point.Y = 85
$tabControl.Location = $System_Drawing_Point
$tabControl.Name = "tabControl"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Height = 315
$System_Drawing_Size.Width = 275
$tabControl.Size = $System_Drawing_Size
$form1.Controls.Add($tabControl)

#endregion

#region verbose_box

#Verbose Output box
$outputBox = New-Object System.Windows.Forms.TextBox 
$outputBox.Location = New-Object System.Drawing.Size(75,450) 
$outputBox.Size = New-Object System.Drawing.Size(750,200) 
$outputBox.MultiLine = $True 
$outputBox.ScrollBars = "Vertical"
$form1.Controls.Add($outputBox)

#endregion

#region General Tabs

#region tabs_setup

#Troubleshooting Page
$TroubleshootingPage.DataBindings.DefaultDataSourceUpdateMode = 0
$TroubleshootingPage.UseVisualStyleBackColor = $True
$TroubleshootingPage.Name = "TroubleshootingPage"
$TroubleshootingPage.Text = "Troubleshooting”
$tabControl.Controls.Add($TroubleshootingPage)

#Remote Page #########FINISH##############
$CPUPage.DataBindings.DefaultDataSourceUpdateMode = 0
$CPUPage.UseVisualStyleBackColor = $True
$CPUPage.Name = "CPUPage"
$CPUPage.Text = "Remote”
$tabControl.Controls.Add($CPUPage)

#Bitlocker Page #########FINISH##############
$BitlockerPage.DataBindings.DefaultDataSourceUpdateMode = 0
$BitlockerPage.UseVisualStyleBackColor = $True
$BitlockerPage.Name = "BitlockerPage"
$BitlockerPage.Text = "Bitlocker”
$tabControl.Controls.Add($BitlockerPage)

#Users Page #########FINISH##############
$UsersPage.DataBindings.DefaultDataSourceUpdateMode = 0
$UsersPage.UseVisualStyleBackColor = $True
$UsersPage.Name = "UsersPage"
$UsersPage.Text = "Users”
$tabControl.Controls.Add($UsersPage)

#Add Label and TextBox
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(175,20)  
$objLabel.Size = New-Object System.Drawing.Size(110,20)  
$objLabel.Text = "Enter Hostname"
$form1.Controls.Add($objLabel)
$objTextBox = New-Object System.Windows.Forms.TextBox 
$objTextBox.Location = New-Object System.Drawing.Size(120,45) 
$objTextBox.Size = New-Object System.Drawing.Size(200,20)  
$form1.Controls.Add($objTextBox) 

#endregion

#region Page_buttons
#General tech button
$Gbutton.TabIndex = 0
$Gbutton.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 35
$Gbutton.Size = $System_Drawing_Size
$Gbutton.UseVisualStyleBackColor = $True
$Gbutton.Text = "General Tech"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 400
$System_Drawing_Point.Y = 35
$Gbutton.Location = $System_Drawing_Point
$Gbutton.DataBindings.DefaultDataSourceUpdateMode = 0
$Gbutton.add_Click($Gbutton_RunOnClick)
$form1.Controls.Add($Gbutton)


#Hostname creator button
$Hostnamebutton.TabIndex = 0
$Hostnamebutton.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 35
$Hostnamebutton.Size = $System_Drawing_Size
$Hostnamebutton.UseVisualStyleBackColor = $True
$Hostnamebutton.Text = "Create a Hostname"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 475
$System_Drawing_Point.Y = 35
$Hostnamebutton.Location = $System_Drawing_Point
$Hostnamebutton.DataBindings.DefaultDataSourceUpdateMode = 0
$Hostnamebutton.add_Click($Hostnamebutton_RunOnClick)
$form1.Controls.Add($Hostnamebutton)


#Ad user creation button
$ADUserbutton.TabIndex = 0
$ADUserbutton.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 35
$ADUserbutton.Size = $System_Drawing_Size
$ADUserbutton.UseVisualStyleBackColor = $True
$ADUserbutton.Text = "Create User"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 550
$System_Drawing_Point.Y = 35
$ADUserbutton.Location = $System_Drawing_Point
$ADUserbutton.DataBindings.DefaultDataSourceUpdateMode = 0
$ADUserbutton.add_Click($ADUserbutton_RunOnClick)
$form1.Controls.Add($ADUserbutton)


#Application button
$Appbutton.TabIndex = 0
$Appbutton.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 35
$Appbutton.Size = $System_Drawing_Size
$Appbutton.UseVisualStyleBackColor = $True
$Appbutton.Text = "Install Apps"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 625
$System_Drawing_Point.Y = 35
$Appbutton.Location = $System_Drawing_Point
$Appbutton.DataBindings.DefaultDataSourceUpdateMode = 0
$Appbutton.add_Click($Appbutton_RunOnClick)
$form1.Controls.Add($Appbutton)


#Zip Extractor button
$Zipbutton.TabIndex = 0
$Zipbutton.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 35
$Zipbutton.Size = $System_Drawing_Size
$Zipbutton.UseVisualStyleBackColor = $True
$Zipbutton.Text = "Zip Extractor"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 700
$System_Drawing_Point.Y = 35
$Zipbutton.Location = $System_Drawing_Point
$Zipbutton.DataBindings.DefaultDataSourceUpdateMode = 0
$Zipbutton.add_Click($Zipbutton_RunOnClick)
$form1.Controls.Add($Zipbutton)
#endregion

#region troubleshooting_tab

#Troubleshooting Page###############################################################################
#Button 1 Ping
$button1.TabIndex = 0
$button1.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$button1.Size = $System_Drawing_Size
$button1.UseVisualStyleBackColor = $True
$button1.Text = "Ping"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 15
$button1.Location = $System_Drawing_Point
$button1.DataBindings.DefaultDataSourceUpdateMode = 0
$button1.add_Click($button1_RunOnClick)
$TroubleshootingPage.Controls.Add($button1)

#Button 2 IP Info
$button2.TabIndex = 1
$button2.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$button2.Size = $System_Drawing_Size
$button2.UseVisualStyleBackColor = $True
$button2.Text = "IP Info"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 15
$button2.Location = $System_Drawing_Point
$button2.DataBindings.DefaultDataSourceUpdateMode = 0
$button2.add_Click($button2_RunOnClick)
$TroubleshootingPage.Controls.Add($button2)

#Button 3 Restart
$button3.TabIndex = 2
$button3.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$button3.Size = $System_Drawing_Size
$button3.UseVisualStyleBackColor = $True
$button3.Text = "Restart"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 45
$button3.Location = $System_Drawing_Point
$button3.DataBindings.DefaultDataSourceUpdateMode = 0
$button3.add_Click($button3_RunOnClick)
$TroubleshootingPage.Controls.Add($button3)

#Button 4 Shutdown
$button4.TabIndex = 3
$button4.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$button4.Size = $System_Drawing_Size
$button4.UseVisualStyleBackColor = $True
$button4.Text = "Shutdown"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 45
$button4.Location = $System_Drawing_Point
$button4.DataBindings.DefaultDataSourceUpdateMode = 0
$button4.add_Click($button4_RunOnClick)
$TroubleshootingPage.Controls.Add($button4)

#Button 5 CMD
$button5.TabIndex = 4
$button5.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$button5.Size = $System_Drawing_Size
$button5.UseVisualStyleBackColor = $True
$button5.Text = "CMD"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 75
$button5.Location = $System_Drawing_Point
$button5.DataBindings.DefaultDataSourceUpdateMode = 0
$button5.add_Click($button5_RunOnClick)
$TroubleshootingPage.Controls.Add($button5)

#Button 6 PowerShell
$button6.TabIndex = 5
$button6.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$button6.Size = $System_Drawing_Size
$button6.UseVisualStyleBackColor = $True
$button6.Text = "Powershell"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 75
$button6.Location = $System_Drawing_Point
$button6.DataBindings.DefaultDataSourceUpdateMode = 0
$button6.add_Click($button6_RunOnClick)
$TroubleshootingPage.Controls.Add($button6)

#Button 7 Enable Local Admin
$button7.TabIndex = 6
$button7.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 175
$System_Drawing_Size.Height = 25
$button7.Size = $System_Drawing_Size
$button7.UseVisualStyleBackColor = $True
$button7.Text = "Enable Local Admin"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 105
$button7.Location = $System_Drawing_Point
$button7.DataBindings.DefaultDataSourceUpdateMode = 0
$button7.add_Click($button7_RunOnClick)
$TroubleshootingPage.Controls.Add($button7)

#Button 8 Derlete Local Admin
$button8.TabIndex = 7
$button8.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 175
$System_Drawing_Size.Height = 25
$button8.Size = $System_Drawing_Size
$button8.UseVisualStyleBackColor = $True
$button8.Text = "Delete Local Admin"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 135
$button8.Location = $System_Drawing_Point
$button8.DataBindings.DefaultDataSourceUpdateMode = 0
$button8.add_Click($button5_RunOnClick)
$TroubleshootingPage.Controls.Add($button8)

#Button 9 Delete Local Admin
$button9.TabIndex = 8
$button9.Name = "button2"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 175
$System_Drawing_Size.Height = 25
$button9.Size = $System_Drawing_Size
$button9.UseVisualStyleBackColor = $True
$button9.Text = "Install PSExec"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 165
$button9.Location = $System_Drawing_Point
$button9.DataBindings.DefaultDataSourceUpdateMode = 0
$button9.add_Click($button5_RunOnClick)
$TroubleshootingPage.Controls.Add($button9)

#endregion

#region Bitlocker_tab

$button1Bit.TabIndex = 0
$button1Bit.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button1Bit.Size = $System_Drawing_Size
$button1Bit.UseVisualStyleBackColor = $True
$button1Bit.Text = "Monitor Bitlocker Status"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 15
$button1Bit.Location = $System_Drawing_Point
$button1Bit.DataBindings.DefaultDataSourceUpdateMode = 0
$button1Bit.add_Click($button1Bit_RunOnClick)
$BitlockerPage.Controls.Add($button1Bit)

$button2Bit.TabIndex = 1
$button2Bit.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button2Bit.Size = $System_Drawing_Size
$button2Bit.UseVisualStyleBackColor = $True
$button2Bit.Text = "Disable Bitlocker"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 45
$button2Bit.Location = $System_Drawing_Point
$button2Bit.DataBindings.DefaultDataSourceUpdateMode = 0
$button2Bit.add_Click($button2Bit_RunOnClick)
$BitlockerPage.Controls.Add($button2Bit)

$button3Bit.TabIndex = 2
$button3Bit.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button3Bit.Size = $System_Drawing_Size
$button3Bit.UseVisualStyleBackColor = $True
$button3Bit.Text = "Query Bitlocker Key"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 75
$button3Bit.Location = $System_Drawing_Point
$button3Bit.DataBindings.DefaultDataSourceUpdateMode = 0
$button3Bit.add_Click($button3Bit_RunOnClick)
$BitlockerPage.Controls.Add($button3Bit)

$button4Bit.TabIndex = 3
$button4Bit.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button4Bit.Size = $System_Drawing_Size
$button4Bit.UseVisualStyleBackColor = $True
$button4Bit.Text = "Disable Bitlocker PIN"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 105
$button4Bit.Location = $System_Drawing_Point
$button4Bit.DataBindings.DefaultDataSourceUpdateMode = 0
$button4Bit.add_Click($button4Bit_RunOnClick)
$BitlockerPage.Controls.Add($button4Bit)

$button5Bit.TabIndex = 4
$button5Bit.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button5Bit.Size = $System_Drawing_Size
$button5Bit.UseVisualStyleBackColor = $True
$button5Bit.Text = "Backup to AD"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 135
$button5Bit.Location = $System_Drawing_Point
$button5Bit.DataBindings.DefaultDataSourceUpdateMode = 0
$button5Bit.add_Click($button5Bit_RunOnClick)
$BitlockerPage.Controls.Add($button5Bit)

#AD backup input box
$ADBackupBox = New-Object System.Windows.Forms.TextBox 
$ADBackupBox.Location = New-Object System.Drawing.Size(50,165) 
$ADBackupBox.Size = New-Object System.Drawing.Size(160,100) 
$ADBackupBox.MultiLine = $True 
#$ADBackupBox.ScrollBars = "Vertical"
$BitlockerPage.Controls.Add($ADBackupBox)

#endregion

#region Users_tab
$button1Users.TabIndex = 0
$button1Users.Name = "button1Users"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 40
$button1Users.Size = $System_Drawing_Size
$button1Users.UseVisualStyleBackColor = $True
$button1Users.Text = "Query Active Loged in Users"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 15
$button1Users.Location = $System_Drawing_Point
$button1Users.DataBindings.DefaultDataSourceUpdateMode = 0
$button1Users.add_Click($button1Users_RunOnClick)
$UsersPage.Controls.Add($button1Users)

$button2Users.TabIndex = 1
$button2Users.Name = "button2Users"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button2Users.Size = $System_Drawing_Size
$button2Users.UseVisualStyleBackColor = $True
$button2Users.Text = "Query Local Users"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 60
$button2Users.Location = $System_Drawing_Point
$button2Users.DataBindings.DefaultDataSourceUpdateMode = 0
$button2Users.add_Click($button2Users_RunOnClick)
$UsersPage.Controls.Add($button2Users)

$button3Users.TabIndex = 2
$button3Users.Name = "button3Users"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button3Users.Size = $System_Drawing_Size
$button3Users.UseVisualStyleBackColor = $True
$button3Users.Text = "Query Account Profiles"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 90
$button3Users.Location = $System_Drawing_Point
$button3Users.DataBindings.DefaultDataSourceUpdateMode = 0
$button3Users.add_Click($button3Users_RunOnClick)
$UsersPage.Controls.Add($button3Users)

$button4Users.TabIndex = 3
$button4Users.Name = "button4Users"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button4Users.Size = $System_Drawing_Size
$button4Users.UseVisualStyleBackColor = $True
$button4Users.Text = "Delete all Network Profiles"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 120
$button4Users.Location = $System_Drawing_Point
$button4Users.DataBindings.DefaultDataSourceUpdateMode = 0
$button4Users.add_Click($button4Users_RunOnClick)
$UsersPage.Controls.Add($button4Users)

$button5Users.TabIndex = 3
$button5Users.Name = "button5Users"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 25
$button5Users.Size = $System_Drawing_Size
$button5Users.UseVisualStyleBackColor = $True
$button5Users.Text = "Delete Specific Profile"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 55
$System_Drawing_Point.Y = 150
$button5Users.Location = $System_Drawing_Point
$button5Users.DataBindings.DefaultDataSourceUpdateMode = 0
$button5Users.add_Click($button5Users_RunOnClick)
$UsersPage.Controls.Add($button5Users)

#endregion

#endregion

#region Theme Tab
# Set Default Theme Button
$DefaultThemeButton.TabIndex = 0
$DefaultThemeButton.Name = "DefaultThemeButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 25
$DefaultThemeButton.Size = $System_Drawing_Size
$DefaultThemeButton.UseVisualStyleBackColor = $True
$DefaultThemeButton.BackColor = "LightGray"
$DefaultThemeButton.Text = "DEFAULT"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 400
$System_Drawing_Point.Y = 70
$DefaultThemeButton.Location = $System_Drawing_Point
$DefaultThemeButton.DataBindings.DefaultDataSourceUpdateMode = 0
$DefaultThemeButton.add_Click({

    $form1.BackColor = "LightGray"
    $outputBox.BackColor = "White"
    $button1.BackColor = "LightGray"
    $button2.BackColor = "LightGray"
    $button3.BackColor = "LightGray"
    $button4.BackColor = "LightGray"
    $button5.BackColor = "LightGray"
    $button6.BackColor = "LightGray"
    $button7.BackColor = "LightGray"
    $button8.BackColor = "LightGray"
    $button9.BackColor = "LightGray"
    

    $TabControl.BackColor = "White"
    $TroubleshootingPage.BackColor = "LightGray"
    $CPUPage.BackColor = "LightGray"
    $BitlockerPage.BackColor = "LightGray"
    $UsersPage.BackColor = "LightGray"

    
    $objTextBox.BackColor = "White"
    $DefaultThemeButton.BackColor = "LightGray"
    $LightThemeButton.BackColor = "LightGray"
    $DarkThemeButton.BackColor = "LightGray"
})
$form1.Controls.Add($DefaultThemeButton)

# Set Dark Theme Button
$DarkThemeButton = New-Object System.Windows.Forms.Button
$DarkThemeButton.Location = New-Object System.Drawing.Point(475,70)
$DarkThemeButton.Size = New-Object System.Drawing.Size(75, 25)
$DarkThemeButton.BackColor = "LightGray"
$DarkThemeButton.Text = "DARK"
$DarkThemeButton.Add_Click({

    $form1.BackColor = "Gray"
    $outputBox.BackColor = "LightGray"
    $button1.BackColor = "LightGray"
    $button2.BackColor = "LightGray"
    $button3.BackColor = "LightGray"
    $button4.BackColor = "LightGray"
    $button5.BackColor = "LightGray"
    $button6.BackColor = "LightGray"
    $button7.BackColor = "LightGray"
    $button8.BackColor = "LightGray"
    $button9.BackColor = "LightGray"
    

    $TabControl.BackColor = "Gray"
    $TroubleshootingPage.BackColor = "Gray"
    $CPUPage.BackColor = "Gray"
    $BitlockerPage.BackColor = "Gray"
    $UsersPage.BackColor = "Gray"

    $objTextBox.BackColor = "LightGray"
    $DefaultThemeButton.BackColor = "LightGray"
    $LightThemeButton.BackColor = "LightGray"
    $DarkThemeButton.BackColor = "LightGray"
})
$form1.Controls.Add($DarkThemeButton)

# Set LIGHT Theme Button
$LightThemeButton = New-Object System.Windows.Forms.Button
$LightThemeButton.Location = New-Object System.Drawing.Point(550,70)
$LightThemeButton.Size = New-Object System.Drawing.Size(75, 25)
$LightThemeButton.BackColor = "LightGray"
$LightThemeButton.Text = "LIGHT"
$LightThemeButton.Add_Click({

    $form1.BackColor = "White"
    $outputBox.BackColor = "White"
    $button1.BackColor = "LightGray"
    $button2.BackColor = "LightGray"
    $button3.BackColor = "LightGray"
    $button4.BackColor = "LightGray"
    $button5.BackColor = "LightGray"
    $button6.BackColor = "LightGray"
    $button7.BackColor = "LightGray"
    $button8.BackColor = "LightGray"
    $button9.BackColor = "LightGray"
    

    $TabControl.BackColor = "White"
    $TroubleshootingPage.BackColor = "White"
    $CPUPage.BackColor = "White"
    $BitlockerPage.BackColor = "White"
    $UsersPage.BackColor = "White"

    
    $objTextBox.BackColor = "White"
    $DefaultThemeButton.BackColor = "White"
    $LightThemeButton.BackColor = "White"
    $DarkThemeButton.BackColor = "White"
})
$form1.Controls.Add($LightThemeButton)
#endregion

#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null
} #End function CreateForm
 
 function Invoke-Sqlcmd3

{
    param(
    [string]$Query,             
    [string]$Database="tempdb",
    [Int32]$QueryTimeout=30
    )
    $conn=new-object System.Data.SqlClient.SQLConnection
    $conn.ConnectionString="Server={0};Database={1};Integrated Security=True" -f $Server,$Database
    $conn.Open()
    $cmd=new-object system.Data.SqlClient.SqlCommand($Query,$conn)
    $cmd.CommandTimeout=$QueryTimeout
    $ds=New-Object system.Data.DataSet
    $da=New-Object system.Data.SqlClient.SqlDataAdapter($cmd)
    [void]$da.fill($ds)
    $conn.Close()
    $ds.Tables[0]
}

 

Function SQLVersion
{
[string]$SQLVersion = @"
SELECT  @@Version
"@ 
 $Server = $objTextBox.text
Invoke-Sqlcmd3 -ServerInstance $Server -Database Master -Query $SQLVersion | Out-GridView -Title "$server SQL Server Version"
}

Function LastReboot
{
$Server = $objTextBox.text
$wmi = Get-WmiObject -Class Win32_OperatingSystem -Computer $server
$wmi.ConvertToDatetime($wmi.LastBootUpTime) | Select DateTime | Out-GridView -Title "$Server Last Reboot"
}

Function Requests
{
[string]$Requests = @"
SELECT
   db_name(r.database_id) as database_name, r.session_id AS SPID,r.status,s.host_name,
     r.start_time,(r.total_elapsed_time/1000) AS 'TotalElapsedTime Sec',
   r.wait_type as current_wait_type,r.wait_resource as current_wait_resource,
   r.blocking_session_id,r.logical_reads,r.reads,r.cpu_time as cpu_time_ms,r.writes,r.row_count,
   substring(st.text,r.statement_start_offset/2,
   (CASE WHEN r.statement_end_offset = -1 THEN len(convert(nvarchar(max), st.text)) * 2 ELSE r.statement_end_offset END - r.statement_start_offset)/2) as statement
FROM
   sys.dm_exec_requests r
      LEFT OUTER JOIN sys.dm_exec_sessions s on s.session_id = r.session_id
      LEFT OUTER JOIN sys.dm_exec_connections c on c.connection_id = r.connection_id       
      CROSS APPLY sys.dm_exec_sql_text(r.sql_handle) st 
WHERE r.status NOT IN ('background','sleeping')
"@ 
 $Server = $objTextBox.text
Invoke-Sqlcmd3 -ServerInstance $Server -Database Master -Query $Requests | Out-GridView -Title "$server Requests"
}



#Call the Function

CreateForm
