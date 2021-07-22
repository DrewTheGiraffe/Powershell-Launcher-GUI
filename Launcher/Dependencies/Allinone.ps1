# General Technician Script!
# ~Script By SPC Burgess & SPC Santiago 2-3 FA S6 07/22/2021
# MOS: 25B & 25U
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
$Themebutton = New-Object System.Windows.Forms.Button
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
$objLabel1 = New-Object System.Windows.Forms.Label
$objTextBox1 = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
$objChromeCheckbox = New-Object System.Windows.Forms.Checkbox
$objFireFoxCheckbox = New-Object System.Windows.Forms.Checkbox 
$objMSTEAMSCheckbox = New-Object System.Windows.Forms.Checkbox 
$objCitrixCheckbox = New-Object System.Windows.Forms.Checkbox 
$objDCAMCheckbox = New-Object System.Windows.Forms.Checkbox
$objWinGUICheckbox = New-Object System.Windows.Forms.Checkbox
$objAdobeDCPROCheckbox = New-Object System.Windows.Forms.Checkbox
$objSharePointDesigner2013Checkbox = New-Object System.Windows.Forms.Checkbox
$objJoeSmithCheckbox = New-Object System.Windows.Forms.CheckBox
$objGEarthCheckbox = New-Object System.Windows.Forms.CheckBox
$objDisableLogsCheckbox = New-Object System.Windows.Forms.CheckBox

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



#region verbose_box

#Verbose Output box
$outputBox = New-Object System.Windows.Forms.TextBox 
$outputBox.Location = New-Object System.Drawing.Size(75,450) 
$outputBox.Size = New-Object System.Drawing.Size(750,200) 
$outputBox.MultiLine = $True 
$outputBox.ScrollBars = "Vertical"
$form1.Controls.Add($outputBox)

#endregion

#region draws everything


#region Applications content

#This creates a label for the TextBox1
$objLabel1.Location = New-Object System.Drawing.Size(10,20) 
$objLabel1.Size = New-Object System.Drawing.Size(280,20)
[String]$MandatoryWrite = "*" 
$objLabel1.ForeColor = [System.Drawing.Color]::FromName("Red")
$objLabel1.Text = "Enter Hostname or IPV4 Address $MandatoryWrite"

#Creates the Textbox
$objTextBox1.Location = New-Object System.Drawing.Size(10,40) 
$objTextBox1.Size = New-Object System.Drawing.Size(260,20)
$objTextBox1.TabIndex = 0 


#This creates a checkbox 
$objChromeCheckbox.Location = New-Object System.Drawing.Size(30,110) 
$objChromeCheckbox.Size = New-Object System.Drawing.Size(350,20)
$objChromeCheckbox.Text = "Install Google Chrome"
$objChromeCheckbox.TabIndex = 2

#This creates a checkbox 
$objFireFoxCheckbox.Location = New-Object System.Drawing.Size(30,130) 
$objFireFoxCheckbox.Size = New-Object System.Drawing.Size(350,20)
$objFireFoxCheckbox.Text = "Install FireFox"
$objFireFoxCheckbox.TabIndex = 3

#This creates a checkbox 
$objMSTEAMSCheckbox.Location = New-Object System.Drawing.Size(30,150) 
$objMSTEAMSCheckbox.Size = New-Object System.Drawing.Size(350,20)
$objMSTEAMSCheckbox.Text = "Install MS Teams"
$objMSTEAMSCheckbox.TabIndex = 4

#This creates a checkbox 
$objCitrixCheckbox.Location = New-Object System.Drawing.Size(30,170) 
$objCitrixCheckbox.Size = New-Object System.Drawing.Size(350,20)
$objCitrixCheckbox.Text = "Install Citrix"
$objCitrixCheckbox.TabIndex = 4

#This creates a checkbox  
$objDCAMCheckbox.Location = New-Object System.Drawing.Size(30,190) 
$objDCAMCheckbox.Size = New-Object System.Drawing.Size(350,20)
$objDCAMCheckbox.Text = "Install DCAM"
$objDCAMCheckbox.TabIndex = 5

#This creates a checkbox  
$objWinGUICheckbox.Location = New-Object System.Drawing.Size(30,210) 
$objWinGUICheckbox.Size = New-Object System.Drawing.Size(350,20)
$objWinGUICheckbox.Text = "Install GSS Army (WinGUI)"
$objWinGUICheckbox.TabIndex = 6

#This creates a checkbox  
$objAdobeDCPROCheckbox.Location = New-Object System.Drawing.Size(30,230) 
$objAdobeDCPROCheckbox.Size = New-Object System.Drawing.Size(350,20)
$objAdobeDCPROCheckbox.Text = "Install Adobe DC Pro"
$objAdobeDCPROCheckbox.TabIndex = 7

#This creates a checkbox  
$objSharePointDesigner2013Checkbox.Location = New-Object System.Drawing.Size(30,250) 
$objSharePointDesigner2013Checkbox.Size = New-Object System.Drawing.Size(350,20)
$objSharePointDesigner2013Checkbox.Text = "Install Share Point Designer 2013"
$objSharePointDesigner2013Checkbox.TabIndex = 8

#Santiago's checkbox
$objJoeSmithCheckbox.Location = New-Object System.Drawing.Size(30,270)
$objJoeSmithCheckbox.size = New-Object System.Drawing.Size(350,20)
$objJoeSmithCheckbox.Text = "Install Joe.Smith Local Administrator"
$objJoeSmithCheckbox.TabIndex = 9

# Google Earth 
$objGEarthCheckbox.Location = New-Object System.Drawing.Size(30,290)
$objGEarthCheckbox.size = New-Object System.Drawing.Size(350,20)
$objGEarthCheckbox.Text = "Install Google Earth"
$objGEarthCheckbox.TabIndex = 10

$objDisableLogsCheckbox.Location = New-Object System.Drawing.Size(30,400)
$objDisableLogsCheckbox.size = New-Object System.Drawing.Size(350,20)
$objDisableLogsCheckbox.Text = "Disable Auto Logging [NOT RECOMMENDED]"
$objDisableLogsCheckbox.TabIndex = 11


#endregion

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
$Gbutton.add_Click({

$form1.Controls.Remove($objLabel1)
$form1.Controls.Remove($objTextBox1)
$form1.Controls.Remove($objChromeCheckbox)
$form1.Controls.Remove($objFireFoxCheckbox)
$form1.Controls.Remove($objMSTEAMSCheckbox)
$form1.Controls.Remove($objCitrixCheckbox)
$form1.Controls.Remove($objDCAMCheckbox)
$form1.Controls.Remove($objWinGUICheckbox)
$form1.Controls.Remove($objAdobeDCPROCheckbox)
$form1.Controls.Remove($objSharePointDesigner2013Checkbox)
$form1.Controls.Remove($objJoeSmithCheckbox)
$form1.Controls.Remove($objGEarthCheckbox)
$form1.Controls.Remove($objDisableLogsCheckbox)

#region General_tech

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
$form1.Controls.Remove($tabControl)
$form1.Controls.Add($tabControl)

#endregion

#region Enter Hostname

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

#region tabs_setup

#Troubleshooting Page
$TroubleshootingPage.DataBindings.DefaultDataSourceUpdateMode = 0
$TroubleshootingPage.UseVisualStyleBackColor = $True
$TroubleshootingPage.Name = "TroubleshootingPage"
$TroubleshootingPage.Text = "Troubleshooting"
$tabControl.Controls.Remove($TroubleshootingPage)
$tabControl.Controls.Add($TroubleshootingPage)

#Remote Page #########FINISH##############
$CPUPage.DataBindings.DefaultDataSourceUpdateMode = 0
$CPUPage.UseVisualStyleBackColor = $True
$CPUPage.Name = "CPUPage"
$CPUPage.Text = "Remote"
$tabControl.Controls.Remove($CPUPage)
$tabControl.Controls.Add($CPUPage)

#Bitlocker Page #########FINISH##############
$BitlockerPage.DataBindings.DefaultDataSourceUpdateMode = 0
$BitlockerPage.UseVisualStyleBackColor = $True
$BitlockerPage.Name = "BitlockerPage"
$BitlockerPage.Text = "Bitlocker"
$tabControl.Controls.Remove($BitlockerPage)
$tabControl.Controls.Add($BitlockerPage)

#Users Page #########FINISH##############
$UsersPage.DataBindings.DefaultDataSourceUpdateMode = 0
$UsersPage.UseVisualStyleBackColor = $True
$UsersPage.Name = "UsersPage"
$UsersPage.Text = "Users"
$tabControl.Controls.Remove($UsersPage)
$tabControl.Controls.Add($UsersPage)



#endregion

#region troubleshooting_tab
    #region TroubleshootingPage
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


})
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
$Appbutton.add_Click({





#region applications content
#This creates Textbox Label
$form1.Controls.Add($objLabel1) 

#This creates the TextBox1
$form1.Controls.Add($objTextBox1)





#region Action Check Boxes for Apps
$form1.Controls.Add($objChromeCheckbox)


$form1.Controls.Add($objFireFoxCheckbox)


$form1.Controls.Add($objMSTEAMSCheckbox)


$form1.Controls.Add($objCitrixCheckbox)


$form1.Controls.Add($objDCAMCheckbox)


$form1.Controls.Add($objWinGUICheckbox)


$form1.Controls.Add($objAdobeDCPROCheckbox)


$form1.Controls.Add($objSharePointDesigner2013Checkbox)


$form1.Controls.Add($objJoeSmithCheckbox)


$form1.Controls.Add($objGEarthCheckbox)


$form1.Controls.Add($objDisableLogsCheckbox)
#endregion
#endregion


})

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


#Theme button
$Themebutton.TabIndex = 0
$Themebutton.Name = "button1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 75
$System_Drawing_Size.Height = 35
$Themebutton.Size = $System_Drawing_Size
$Themebutton.UseVisualStyleBackColor = $True
$Themebutton.Text = "Change Theme"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 775
$System_Drawing_Point.Y = 35
$Themebutton.Location = $System_Drawing_Point
$Themebutton.DataBindings.DefaultDataSourceUpdateMode = 0
$Themebutton.add_Click({
#$TabControl.Dispose($TabControl)

#region Theme Tab
# Set Default Theme Button
$DefaultThemeButton.TabIndex = 0
$DefaultThemeButton.Name = "DefaultThemeButton"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 150
$System_Drawing_Size.Height = 125
$DefaultThemeButton.Size = $System_Drawing_Size
$DefaultThemeButton.UseVisualStyleBackColor = $True
$DefaultThemeButton.BackColor = "LightGray"
$DefaultThemeButton.Text = "DEFAULT"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 150
$System_Drawing_Point.Y = 125
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
$DarkThemeButton.Location = New-Object System.Drawing.Point(375,125)
$DarkThemeButton.Size = New-Object System.Drawing.Size(150, 125)
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
$LightThemeButton.Location = New-Object System.Drawing.Point(600,125)
$LightThemeButton.Size = New-Object System.Drawing.Size(150, 125)
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

})
$form1.Controls.Add($Themebutton)

#endregion



#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null
#$form1.DialogResult = "OK" | Out-Null
}
 #End function CreateForm

 


 




#Call the Function

CreateForm

