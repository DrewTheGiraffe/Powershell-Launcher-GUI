# All in One Script!
# ~Script By SPC Burgess & SPC Santiago 2-3 FA S6 07/15/2021
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

#region Definitions
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.drawing

#Form Setup
$form1 = New-Object System.Windows.Forms.Form
$outputBox = New-Object System.Windows.Forms.TextBox
$button1 = New-Object System.Windows.Forms.Button
$button2 = New-Object System.Windows.Forms.Button
$button3 = New-Object System.Windows.Forms.Button
$button4 = New-Object System.Windows.Forms.Button
$button5 = New-Object System.Windows.Forms.Button
$button6 = New-Object System.Windows.Forms.Button
$button7 = New-Object System.Windows.Forms.Button
$button8 = New-Object System.Windows.Forms.Button
$remotebutton1 = New-Object System.Windows.Forms.Button

$TabControl = New-object System.Windows.Forms.TabControl
$TroubleshootingPage = New-Object System.Windows.Forms.TabPage
$CPUPage = New-Object System.Windows.Forms.TabPage
$DiskPage = New-Object System.Windows.Forms.TabPage
$MemoryPage = New-Object System.Windows.Forms.TabPage

$tabControl2 = New-object System.Windows.Forms.TabControl
$Tab1Page = New-Object System.Windows.Forms.TabPage
$Tab2Page = New-Object System.Windows.Forms.TabPage
$Tab3Page = New-Object System.Windows.Forms.TabPage
$Tab4Page = New-Object System.Windows.Forms.TabPage

$DefaultThemeButton = New-Object System.Windows.Forms.Button
$DarkThemeButton = New-Object System.Windows.Forms.Button
$LightThemeButton = New-Object System.Windows.Forms.Button

$objTextBox = New-Object System.Windows.Forms.TextBox
$InitialFormWindowState = New-Object System.Windows.Forms.FormWindowState
#endregion

#region Form1
#Form Parameter
$form1.Text = "General Tech"
$form1.Name = "form1"
$form1.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Size = New-Object System.Drawing.Size
$form1.BackColor = "LightGray"
$System_Drawing_Size.Width = 900
$System_Drawing_Size.Height = 700
$form1.FormBorderStyle = 'Fixed3D'
$form1.MaximizeBox = $false
$form1.MinimizeBox = $false
$form1.ClientSize = $System_Drawing_Size

# Draws Icon
$iconConverted2Base64 = [Convert]::ToBase64String((Get-Content "C:\temp\Launcher\Dependencies\icon\NewPanda.ico" -Encoding Byte))
$iconBase64           = $iconConverted2Base64
$iconBytes            = [Convert]::FromBase64String($iconBase64)
$stream               = New-Object IO.MemoryStream($iconBytes, 0, $iconBytes.Length)
$stream.Write($iconBytes, 0, $iconBytes.Length);
$iconImage            = [System.Drawing.Image]::FromStream($stream, $true)
$form1.Icon           = [System.Drawing.Icon]::FromHandle((New-Object System.Drawing.Bitmap -Argument $stream).GetHIcon())
# ico converter : https://cloudconvert.com/png-to-ico

# Draws Logo
$img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\Panda\NewPanda.png')
$form1.BackgroundImage = $img
$form1.BackgroundImageLayout = "Center"
#endregion

#region Tab_frontend

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

#Troubleshooting Page
$TroubleshootingPage.DataBindings.DefaultDataSourceUpdateMode = 0
$TroubleshootingPage.UseVisualStyleBackColor = $True
$TroubleshootingPage.Name = "TroubleshootingPage"
$TroubleshootingPage.Text = "Troubleshooting”
$tabControl.Controls.Add($TroubleshootingPage)

#Remote Page
$CPUPage.DataBindings.DefaultDataSourceUpdateMode = 0
$CPUPage.UseVisualStyleBackColor = $True
$CPUPage.Name = "CPUPage"
$CPUPage.Text = "Remote”
$tabControl.Controls.Add($CPUPage)

#Bitlocker Page
$DiskPage.DataBindings.DefaultDataSourceUpdateMode = 0
$DiskPage.UseVisualStyleBackColor = $True
$DiskPage.Name = "DiskPage"
$DiskPage.Text = "Bitlocker”
$tabControl.Controls.Add($DiskPage)

#Users Page
$MemoryPage.DataBindings.DefaultDataSourceUpdateMode = 0
$MemoryPage.UseVisualStyleBackColor = $True
$MemoryPage.Name = "MemoryPage"
$MemoryPage.Text = "Users”
$tabControl.Controls.Add($MemoryPage)

#Tab Control 2
$tabControl2.DataBindings.DefaultDataSourceUpdateMode = 0
$System_Drawing_Point2 = New-Object System.Drawing.Point
$System_Drawing_Point2.X = 550
$System_Drawing_Point2.Y = 85
$tabControl2.Location = $System_Drawing_Point2
$tabControl2.Name = "tabControl2"
$System_Drawing_Size2 = New-Object System.Drawing.Size
$System_Drawing_Size2.Height = 315
$System_Drawing_Size2.Width = 275
$tabControl2.Size = $System_Drawing_Size2
$form1.Controls.Add($tabControl2)

# TabControl2 | Tab 1
$Tab1Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab1Page.UseVisualStyleBackColor = $True
$Tab1Page.Name = "Tab1Page"
$Tab1Page.Text = "Theme”
$tabControl2.Controls.Add($Tab1Page)

# TabControl2 | Tab 2
$Tab2Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab2Page.UseVisualStyleBackColor = $True
$Tab2Page.Name = "Tab2Page"
$Tab2Page.Text = "Tab2”
$tabControl2.Controls.Add($Tab2Page)

# TabControl2 | Tab 3
$Tab3Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab3Page.UseVisualStyleBackColor = $True
$Tab3Page.Name = "Tab3Page"
$Tab3Page.Text = "Tab3”
$tabControl2.Controls.Add($Tab3Page)

# TabControl2 | Tab 4
$Tab4Page.DataBindings.DefaultDataSourceUpdateMode = 0
$Tab4Page.UseVisualStyleBackColor = $True
$Tab4Page.Name = "Tab4Page"
$Tab4Page.Text = "Tab4”
$tabControl2.Controls.Add($Tab4Page)

#Verbose Output box
$outputBox = New-Object System.Windows.Forms.TextBox 
$outputBox.Location = New-Object System.Drawing.Size(10,500) 
$outputBox.Size = New-Object System.Drawing.Size(880,175) 
$outputBox.MultiLine = $True 
$outputBox.ScrollBars = "Vertical"
$form1.Controls.Add($outputBox)

#endregion

#region Hostname_Box
#Add Label and TextBox
$objLabel = New-Object System.Windows.Forms.Label
$objLabel.Location = New-Object System.Drawing.Size(175,20)  
$objLabel.Size = New-Object System.Drawing.Size(110,20)  
$objLabel.Text = "Enter Hostname"
$form1.Controls.Add($objLabel) 
$objTextBox.Location = New-Object System.Drawing.Size(120,45) 
$objTextBox.Size = New-Object System.Drawing.Size(200,20)  
$form1.Controls.Add($objTextBox) 
#endregion

#region Buttons_ControlTabs2

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
$System_Drawing_Point.X = 100
$System_Drawing_Point.Y = 100
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
    $remotebutton1.BackColor = "LightGray"

    $TabControl.BackColor = "White"
    $TroubleshootingPage.BackColor = "LightGray"
    $CPUPage.BackColor = "LightGray"
    $DiskPage.BackColor = "LightGray"
    $MemoryPage.BackColor = "LightGray"

    $tabControl2.BackColor = "White"
    $Tab1Page.BackColor = "LightGray"
    $Tab2Page.BackColor = "LightGray"
    $Tab3Page.BackColor = "LightGray"
    $Tab4Page.BackColor = "LightGray"
    $objTextBox.BackColor = "White"
    $DefaultThemeButton.BackColor = "LightGray"
    $LightThemeButton.BackColor = "LightGray"
    $DarkThemeButton.BackColor = "LightGray"
})
$Tab1Page.Controls.Add($DefaultThemeButton)

# Set Dark Theme Button
$DarkThemeButton = New-Object System.Windows.Forms.Button
$DarkThemeButton.Location = New-Object System.Drawing.Point(100,150)
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
    $remotebutton1.BackColor = "LightGray"

    $TabControl.BackColor = "Gray"
    $TroubleshootingPage.BackColor = "Gray"
    $CPUPage.BackColor = "Gray"
    $DiskPage.BackColor = "Gray"
    $MemoryPage.BackColor = "Gray"

    $tabControl2.BackColor = "Gray"
    $Tab1Page.BackColor = "Gray"
    $Tab2Page.BackColor = "Gray"
    $Tab3Page.BackColor = "Gray"
    $Tab4Page.BackColor = "Gray"
    $objTextBox.BackColor = "LightGray"
    $DefaultThemeButton.BackColor = "LightGray"
    $LightThemeButton.BackColor = "LightGray"
    $DarkThemeButton.BackColor = "LightGray"
})
$Tab1Page.Controls.Add($DarkThemeButton)

# Set LIGHT Theme Button
$LightThemeButton = New-Object System.Windows.Forms.Button
$LightThemeButton.Location = New-Object System.Drawing.Point(100,125)
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
    $remotebutton1.BackColor = "LightGray"

    $TabControl.BackColor = "White"
    $TroubleshootingPage.BackColor = "White"
    $CPUPage.BackColor = "White"
    $DiskPage.BackColor = "White"
    $MemoryPage.BackColor = "White"

    $tabControl2.BackColor = "White"
    $Tab1Page.BackColor = "White"
    $Tab2Page.BackColor = "White"
    $Tab3Page.BackColor = "White"
    $Tab4Page.BackColor = "White"
    $objTextBox.BackColor = "White"
    $DefaultThemeButton.BackColor = "White"
    $LightThemeButton.BackColor = "White"
    $DarkThemeButton.BackColor = "White"
})
$Tab1Page.Controls.Add($LightThemeButton)
#endregion

#region Logo Tab
# Custom Logo 1 Theme Button
$NewPandaPngButton = New-Object System.Windows.Forms.Button
$NewPandaPngButton.Location = New-Object System.Drawing.Point(100,150)
$NewPandaPngButton.Size = New-Object System.Drawing.Size(75, 25)
$NewPandaPngButton.BackColor = "LightGray"
$NewPandaPngButton.Text = "RiceBowl"
$NewPandaPngButton.Add_Click({
    # Draws Logo
    $img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\Panda\NewPanda.png')
    $form1.BackgroundImage = $img
    $form1.BackgroundImageLayout = "Center"
})
$Tab2Page.Controls.Add($NewPandaPngButton)

# Custom Logo 1 Theme Button
$AltPandaPngButton = New-Object System.Windows.Forms.Button
$AltPandaPngButton.Location = New-Object System.Drawing.Point(100,125)
$AltPandaPngButton.Size = New-Object System.Drawing.Size(75, 25)
$AltPandaPngButton.BackColor = "LightGray"
$AltPandaPngButton.Text = "Matrix"
$AltPandaPngButton.Add_Click({
    # Draws Logo
    $img = [System.Drawing.Image]::Fromfile('C:\temp\Launcher\Dependencies\icon\AltPanda\AltPanda.png')
    $form1.BackgroundImage = $img
    $form1.BackgroundImageLayout = "Center"
})
$Tab2Page.Controls.Add($AltPandaPngButton)

#endregion

#endregion

#region Buttons_ControlTabs1

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
$System_Drawing_Point.Y = 45
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
$System_Drawing_Point.Y = 45
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
$System_Drawing_Point.Y = 75
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
$System_Drawing_Point.Y = 75
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
$System_Drawing_Point.Y = 105
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
$System_Drawing_Point.Y = 105
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
$System_Drawing_Point.Y = 135
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
$System_Drawing_Point.Y = 165
$button8.Location = $System_Drawing_Point
$button8.DataBindings.DefaultDataSourceUpdateMode = 0
$button8.add_Click($button5_RunOnClick)
$TroubleshootingPage.Controls.Add($button8)

#Button 1 Remote | Test
$remotebutton1.TabIndex = 0
$remotebutton1.Name = "remotebutton1"
$System_Drawing_Size = New-Object System.Drawing.Size
$System_Drawing_Size.Width = 175
$System_Drawing_Size.Height = 25
$remotebutton1.Size = $System_Drawing_Size
$remotebutton1.UseVisualStyleBackColor = $True
$remotebutton1.Text = "Local Computer Name"
$System_Drawing_Point = New-Object System.Drawing.Point
$System_Drawing_Point.X = 50
$System_Drawing_Point.Y = 165
$remotebutton1.Location = $System_Drawing_Point
$remotebutton1.DataBindings.DefaultDataSourceUpdateMode = 0
$remotebutton1.Add_Click({
    $outputBox.Text = cmd /c "ping www.google.com -4"
})
$CPUPage.Controls.Add($remotebutton1)

#endregion


#Save the initial state of the form
$InitialFormWindowState = $form1.WindowState
#Init the OnLoad event to correct the initial state of the form
$form1.add_Load($OnLoadForm_StateCorrection)
#Show the Form
$form1.ShowDialog()| Out-Null
} #End function CreateForm
 
 #region Random_funcs
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
$Server = $env:COMPUTERNAME
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
#endregion



#Call the Function

CreateForm
