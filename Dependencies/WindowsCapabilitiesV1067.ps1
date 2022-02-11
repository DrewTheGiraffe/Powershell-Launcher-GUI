<#
    .Synopsis
       Description:
       This script should only be ran in PowerShell console mode, as it uses Windows Forms; DO NOT RUN in PowerShell ISE, as whis will cause random lockups of the ISE.

       This script presents a Graphical User Interface (GUI), in order to obtain a list of Windows Capabilities which are either Installed or Not Present on one or more computer systems.
       The user starts by indicating which computers he/she would like to assess, by entering the computernames individually, passing in as a CSV file, or looking up computer objects in AD.
       Once the computers have been identified, the script will check to see if the systems are online and if privleged access is feasible; if yes, the script returns a Windows Capabilities list for each computer, 
       from which the user then selects one or more Windows Capabilities to remove or install.  If the capability is "Installed" and the user selects it, it will be uninstalled; conversly, if the capability is "NotPresent"
       and the user selects it, it will be installed.

       This Script relies on several core functions in order work correctly:
       1.) The "Get-WindowsCapabilities" obtains the list of Windows Capabilites for each system and wheter they are Installed or NotPresent.
       2.) The "Get-WindowsCapabilitiesParametersForm" presents the user GUI to obtain the parameters passed to the "Get-WindowsCapabilites" function.
       3.) The "Select-ItemsFromListViewForm" presents the user GUI to present a list of Windows Capabilities (Installed and/or NotPresent) for Installing or Removing from the targeted computers.

       
       By default, this script will temporarily disable the use of WSUS or WindowsUpdate as the source repositories for the .CAB files that are used to install the features on demand.
       The variable $FODSource1809 specifies an orgnaizational path to the Windows Capabilities 1809 .CAB files, which get installed when the user opts to install features on demand.
       The variable $FODSource1903 specifies an orgnaizational path to the Windows Capabilities 1903 .CAB files, which get installed when the user opts to install features on demand.
       The variable $FODSource1909 specifies an orgnaizational path to the Windows Capabilities 1909 .CAB files, which get installed when the user opts to install features on demand.
       The variable $FODSource20H2 specifies an orgnaizational path to the Windows Capabilities 20H2 .CAB files, which get installed when the user opts to install features on demand.
       The variable $LocalFODSource specifies the local path ($env:ProgramData\armylocal\FOD) to the Windows Capabilities .CAB files, which get installed when the user opts to install features on demand.
       Look for the following in the script and modify as necessary to match your environment: $FODSource1809 = "\\blisw6syaaa7nec\IMO_Info\FOD\1809"
       Look for the following in the script and modify as necessary to match your environment: $FODSource1903 = "\\blisw6syaaa7nec\IMO_Info\FOD\1903"
       Look for the following in the script and modify as necessary to match your environment: $FODSource1909 = "\\blisw6syaaa7nec\IMO_Info\FOD\1909"
       Look for the following in the script and modify as necessary to match your environment: $FODSource20H2 = "\\blisw6syaaa7nec\IMO_Info\FOD\20H2"
       Look for the following in the script and modify as necessary to match your environment: $LocalFODSource = "$env:ProgramData\armylocal\FOD"
       You can obtain the .CAB files only from Microsoft's TechNet, which requires a subscription; search specificially for Windows 10 Features on Demand Part 1 Version 1809, 1903, or 1909, etc.
       The most recent ones we downloaded were "en_windows_10_features_on_demand_part_1_version_1809_updated_sept_2018_x64_dvd_a68fa301.iso" and "en_windows_10_features_on_demand_part_1_version_1903_x64_dvd_1076e85a"     

              
    .VERSION: 1.06 (9 MARCH 2020) - INITIAL RELEASE
        .UPDATE: 1.65 (26 JULY 2021) - UPDATES FOR 20H2 Version
        .UPDATE: 1.66 (26 JULY 2021) - ADDED LOGIC TO PULL FOD FILES FROM LOCAL SOURCE
         
    .EXAMPLE
       Run this script in the CONUS Forest on a computer that has the Admin Tool (RSAT) installed.  This ensures that PowerShell with the necessary ActiveDirectory modules can be loaded.

    .INPUTS
       User must select "WHERE" the computer objects are being passed in from (i.e. -ComputerNames, -ComputerListCSV, -SearchAD), and "WHICH" OU from he/she desires the search to begin (i.e. -targetOU, -LocateTargetOU).
       OPTIONALLY: 
            The Search Scope can be defined with -SearchScope parameter.
            Logging (HIGHLY RECOMMENDED) can be selected with -RunningResults2csv parameter.
            Logpath can be defined with -ResultsFilePath parameter.
            Capability Name Filtering can be defined with -CapabilityNameFilter parameter
            Only If Installed filtering can be defined with -OnlyIfInstalled parameter switch
            Only If Not Present filtering can be defined with -OnlyIfNotPresent parameter switch

    .OUTPUTS
       One XLSX Summary file is created in the "$env:SystemDrive\LOGS\WindowsCapabilities\" folder, with the name "C:\Logs\WindowsCapabilities\WindowsCapabilitiesSummary_RunningResults_MM-DD-YYYY_HHmmss.XLSX"
       One XLSX file is created per computer, in the "$env:SystemDrive\LOGS\WindowsCapabilities\" folder, with the name "C:\Logs\WindowsCapabilities\HOSTNAME_WindowsCapabilities_MM-DD-YYYY_HHmmss.XLSX"

    .AUTHOR 
        Michael D. Sloan
        Fort Bliss SANEC
        michael.d.sloan.civ@mail.mil
        DSN: 312-711-0744
        COMM: +1 915-741-0744

    #>


# Check for administrative rights
IF(-NOT([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Warning -Message "The script requires elevation"
    break
}

FUNCTION Select-ItemsFromListViewForm {
    [cmdletbinding()]    
    #Parameters passed into the funtion
    Param (
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
		[string]$FormTitle="Items Selection from ListView",
        [Parameter(Mandatory=$false)]
        [ValidateSet(640, 800, 960, 1120, 1280, 1440)]
		[int]$FormSizeHoriz=640,
        [Parameter(Mandatory=$false)]
        [ValidateSet(480, 520, 640, 800, 960, 1120, 1280, 1440)]
		[int]$FormSizeVert=480,
        [Parameter(Mandatory=$false)]
		[int]$MaximumSizeWidth=0,
        [Parameter(Mandatory=$false)]
		[int]$MaximumSizeHeight=0,
        [Parameter(Mandatory=$false)]
		[bool]$MaximizeBox=$true,
        [Parameter(Mandatory=$false)]
		[int]$MinimumSizeWidth=0,
        [Parameter(Mandatory=$false)]
		[int]$MinimumSizeHeight=0,
        [Parameter(Mandatory=$false)]
		[bool]$MinimizeBox=$true,
        [parameter(Mandatory=$false)]
        [ValidateSet('None', 'FixedSingle', 'Fixed3D', 'FixedDialog', 'Sizable', 'FixedToolWindow', 'SizableToolWindow')]
        [string]$FormBorderStyle='Sizable',
        [parameter(Mandatory=$false)]
        [ValidateSet('Manual', 'CenterScreen', 'WindowsDefaultLocation', 'WindowsDefaultBounds', 'CenterParent')]
        [string]$StartPosition='CenterScreen',
        [Parameter(Mandatory=$false)]
		[int]$FormLocationX=0,
        [Parameter(Mandatory=$false)]
		[int]$FormLocationY=0,
        [Parameter(Mandatory=$false)]
        [bool]$AutoSize,
        [Parameter(Mandatory=$false)]
        [bool]$TopMost=$true,
        [parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
        [string]$IconSourceFilePath,
        [parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
        [array]$IconIndex,
        
        [parameter(Mandatory=$false)]
        [switch]$IncludeMsgBxLbl,
        [parameter(Mandatory=$false)]
        [string]$Message,

        [parameter(Mandatory=$false)]
        [ValidateSet('OK', 'CANCEL', 'SELECTALL', 'DESELECTALL', 'CUSTOM1', 'CUSTOM2', 'YES', 'NO')]
        [string]$AcceptButton='OK',
        [parameter(Mandatory=$false)]
        [ValidateSet('OK', 'CANCEL', 'SELECTALL', 'DESELECTALL', 'CUSTOM1', 'CUSTOM2', 'YES', 'NO')]
        [string]$CancelButton='CANCEL',

        [parameter(Mandatory=$false)]
        [switch]$IncludeOkBtn,
        [parameter(Mandatory=$false)]
        [string]$OkBtnLbl="      OK      ",
        
        [parameter(Mandatory=$false)]
        [switch]$IncludeCancelBtn,
        [parameter(Mandatory=$false)]
        [string]$CancelBtnLbl=" CANCEL ",

        [parameter(Mandatory=$false)]
        [switch]$IncludeSelectAllBtn,
        [parameter(Mandatory=$false)]
        [string]$SelectAllBtnLbl="Select ALL",

        [parameter(Mandatory=$false)]
        [switch]$IncludeDeSelectAllBtn,
        [parameter(Mandatory=$false)]
        [string]$DeSelectAllBtnLbl="De-Select ALL",

        [parameter(Mandatory=$false)]
        [switch]$IncludeCustomBtn,
        [parameter(Mandatory=$false)]
        [string]$CustomBtnLabel = "CUSTOM BUTTON",
        [parameter(Mandatory=$false)]
        [switch]$IncludeCustom2Btn,
        [parameter(Mandatory=$false)]
        [string]$Custom2BtnLabel = "CUSTOM2 BUTTON",
        [parameter(Mandatory=$false)]
        [switch]$IncludeYesBtn,
        [parameter(Mandatory=$false)]
        [string]$YesBtnLbl="      YES      ",
        [parameter(Mandatory=$false)]
        [switch]$IncludeNoBtn,
        [parameter(Mandatory=$false)]
        [string]$NoBtnLbl="      NO      ",
        [parameter(Mandatory=$false)]
        $FormTimeout,
        [parameter(Mandatory=$false)]
        $ScriptBlock,

        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true)]
		[string]$ListLabelText="Please select one or more items from the list and then OK to continue, or Cancel to abort: ",
        [Parameter(Mandatory=$false)]
        [array]$ItemsListViewColumnNames,
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true,ParameterSetName="OneDim")]
        [array]$ItemsList,
        [Parameter(Mandatory=$false, ValueFromPipelineByPropertyName=$true,ParameterSetName="MultiDim")]
        [array]$ItemsListCollection,
        [Parameter(Mandatory=$false)]
        [bool]$MultiSelect = $true,
        [Parameter(Mandatory=$false)]
        [switch]$ListViewSortable,
        [Parameter(Mandatory=$false)]
        [switch]$ListViewSearchable,
        [Parameter(Mandatory=$false)]
        [switch]$ReturnObjects
    )
    
    BEGIN 
    {
        Add-Type -AssemblyName System.Windows.Forms
        Add-Type -AssemblyName System.Drawing

        FUNCTION Extract-Icons
        {
                 <#
            .SYNOPSIS
                Exports an ico and bmp file from a given source to a given destination
            .Description
                You need to set the Source and Destination locations. First version of a script, I found other examples 
                but all I wanted to do as grab and ico file from an exe but found getting a bmp useful. Others might find useful
    
            .EXAMPLE
                This will run with the default paramter values (i.e. Extracts the first icon from Shell32.dll ($IconIndexRange = 0),
                    and writes it to $env:TEMP\Icons\Shell32_0.ico)
                .\Extract-Icons.ps1
    
            .EXAMPLE
                this will default to shell32.dll automatically for -SourceFilePath, and extract the 239th Icon (238th index) in Shell32.dll
                .\Extract-Icons.ps1 -DestinationFolder 'C:\temp\MyIcons' -IconIndexRange 238
            .EXAMPLE
                This will give you a green tree icon (press F5 for windows to refresh Windows explorer) i.e. 'C:\temp\MyIcons\Shell32_41.ico'
                .\Extract-Icons.ps1 -SourceFilePath 'C:/Windows/system32/shell32.dll' -DestinationFolder 'C:\temp\MyIcons' -IconIndexRange 41

            .Notes
                Based on http://stackoverflow.com/questions/8435/how-do-you-get-the-icons-out-of-shell32-dll Version 1.1 2012.03.8
        
                New version: Version 1.0 2020.8.26 
            #>
    
            [CmdletBinding()]
            Param ( 
        
                [parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=0)]
                [string]$SourceFilePath,

                [parameter(Mandatory=$true,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=1)]
                [array]$IconIndexRange,

                [parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=2)]
                [switch]$largeIcon,

                [parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=2)]
                [ValidateSet('BitmapImage', 'Icon')]
                [string]$OutputFormat = 'Icon',

                [parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Position=3)]
                [switch]$ExtractToFolder
            )

            DynamicParam
            {
                #Create the RuntimeDefinedParameterDictionary 
                $RuntimeParameterDictionary = [System.Management.Automation.RuntimeDefinedParameterDictionary]::new()

                IF($ExtractToFolder)
                {
                    #Set the FIRST dynamic parameter's name
                    $ParamName_DestinationFolder = 'DestinationFolder'
                    #Create a new ParameterAttribute Object and set the parametrer's attributes        
                    $DestinationFolderParameterAttributes = [System.Management.Automation.ParameterAttribute]::new()
                    $DestinationFolderParameterAttributes.Position = 4
                    $DestinationFolderParameterAttributes.Mandatory = $false
                    $DestinationFolderParameterAttributes.ValueFromPipeline = $true
                    $DestinationFolderParameterAttributes.ValueFromPipelineByPropertyName = $true
                    $DestinationFolderParameterAttributes.ParameterSetName = '__AllParameterSets'
                    $DestinationFolderParameterAttributes.HelpMessage = "Select a destination folder where you would like your icons extracted to; DEFAULT LOCATION=$ENV:TEMP\Icons"
            
                    #Create an attributeCollection object for the dynamic parameter's attributes we just created above
                    $DestinationFolderAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]

                    #Add our dynamic parameter's attributes to the attributeCollection
                    $DestinationFolderAttributeCollection.Add($DestinationFolderParameterAttributes)

                    #Create and return the RuntimeDefinedParameter DestinationFolderParam
                    $DestinationFolderParam = [System.Management.Automation.RuntimeDefinedParameter]::new($ParamName_DestinationFolder, [string], $DestinationFolderAttributeCollection)
            
                    #Add the RuntimeParameter to the RuntimeDefinedParameterDictionary
                    $RuntimeParameterDictionary.Add($ParamName_DestinationFolder, $DestinationFolderParam)
            
                    #Set the SECOND dynamic parameter's name
                    $ParamName_AutoCreateSubFolders = 'AutoCreateSubFolders'
                    #Create a new ParameterAttribute Object and set the parametrer's attributes        
                    $AutoCreateSubFoldersParameterAttributes = [System.Management.Automation.ParameterAttribute]::new()
                    $AutoCreateSubFoldersParameterAttributes.Position = 5
                    $AutoCreateSubFoldersParameterAttributes.Mandatory = $false
                    $AutoCreateSubFoldersParameterAttributes.ValueFromPipeline = $true
                    $AutoCreateSubFoldersParameterAttributes.ValueFromPipelineByPropertyName = $true
                    $AutoCreateSubFoldersParameterAttributes.ParameterSetName = '__AllParameterSets'
                    $AutoCreateSubFoldersParameterAttributes.HelpMessage = "Optionally auto create sub folders in the destination folder, based on filename where icons were extracted from; EXAMPLES=$ENV:TEMP\Icons\Shell32\, $ENV:TEMP\Icons\explorer\"
            
                    #Create an attributeCollection object for the dynamic parameter's attributes we just created above
                    $AutoCreateSubFoldersAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]

                    #Add our dynamic parameter's attributes to the attributeCollection
                    $AutoCreateSubFoldersAttributeCollection.Add($AutoCreateSubFoldersParameterAttributes)

                    #Create and return the RuntimeDefinedParameter AutoCreateSubFoldersParam
                    $AutoCreateSubFoldersParam = [System.Management.Automation.RuntimeDefinedParameter]::new($ParamName_AutoCreateSubFolders, [switch], $AutoCreateSubFoldersAttributeCollection)
            
                    #Add the RuntimeParameter to the RuntimeDefinedParameterDictionary
                    $RuntimeParameterDictionary.Add($ParamName_AutoCreateSubFolders, $AutoCreateSubFoldersParam)
            
                } #END IF($ExtractToFolder)

                IF($OutputFormat -eq 'Icon')
                {
                    #Set the THIRD dynamic parameter's name
                    $ParamName_IconQuality = 'IconQuality'
                    #Create a new ParameterAttribute Object and set the parametrer's attributes        
                    $IconQualityParameterAttributes = [System.Management.Automation.ParameterAttribute]::new()
                    $IconQualityParameterAttributes.Position = 6
                    $IconQualityParameterAttributes.Mandatory = $false
                    $IconQualityParameterAttributes.ValueFromPipeline = $true
                    $IconQualityParameterAttributes.ValueFromPipelineByPropertyName = $true
                    $IconQualityParameterAttributes.ParameterSetName = '__AllParameterSets'
                    $IconQualityParameterAttributes.HelpMessage = "Optionally select the quality of the icon being exported, either `"Original`" or `"Bitmap`" quality."
            
                    #Create an attributeCollection object for the dynamic parameter's attributes we just created above
                    $IconQualityAttributeCollection = New-Object System.Collections.ObjectModel.Collection[System.Attribute]

                    #Add our dynamic parameter's attributes to the IconQualityAttributeCollection
                    $IconQualityAttributeCollection.Add($IconQualityParameterAttributes)

                    #Add a parameter ValidateSet to the the IconQualityAttributeCollection
                    $IconQualityAttributeValidateSet = [System.Management.Automation.ValidateSetAttribute]::new('Bitmap', 'Original')
                    $IconQualityAttributeCollection.Add($IconQualityAttributeValidateSet)

                    #Create and return the RuntimeDefinedParameter IconQualityParam
                    $IconQualityParam = [System.Management.Automation.RuntimeDefinedParameter]::new($ParamName_IconQuality, [string], $IconQualityAttributeCollection)
                               
                    #Add the RuntimeParameter to the RuntimeDefinedParameterDictionary
                    $RuntimeParameterDictionary.Add($ParamName_IconQuality, $IconQualityParam)
                }

                return $RuntimeParameterDictionary
            } #END DynamicParam
    
            BEGIN
            {
                # Store variable list in memory at start-up, excluding parameters supplied to the script
                Set-Variable VariablesStartup -Option ReadOnly -Value (Get-Variable -Scope Global | Where { $_.Attributes.TypeId.Name -notcontains "ParameterAttribute" } | Select -ExpandProperty Name)

                # De-reference variables used by the script
                Function Invoke-VariableCleanup {
	                $VariablesToBeRemoved = (Get-Variable -Scope Global | Select -ExpandProperty Name) | Where { $VariablesStartup -notcontains $_ }
	                ForEach ($Item in $VariablesToBeRemoved) { Remove-Variable -Name $Item -Scope Global -Force -ErrorAction SilentlyContinue }
                }

                IF(-not (Test-path -Path $SourceFilePath -ErrorAction SilentlyContinue))
                {
                    Throw "Source file [$SourceFilePath] does not exist!"
                }

                IF($ExtractToFolder)
                {
                    IF((!$PSBoundParameters.ContainsKey('DestinationFolder')) -or (($PSBoundParameters.ContainsKey('DestinationFolder')) -and ($PSBoundParameters[$ParamName_DestinationFolder]).Length -lt 4))
                    {
                        #Write-Host "DynamicParameter `"$ParamName_DestinationFolder`" was not specified or is invalid, populating with default value of `"$env:TEMP\Icons`"" -BackgroundColor Yellow -ForegroundColor Red 
                        $DestinationFolder = "$env:TEMP\Icons"
                    }
                    ELSE #DynamicParameter DestinationFolder was created and specified with a vaule having a length > 4
                    {
                        #Write-Host "DynamicParameter `"$ParamName_DestinationFolder`" was created at runtime, and contains $($PSBoundParameters[$ParamName_IconQuality])" -BackgroundColor DarkCyan -ForegroundColor White
                        $DestinationFolder = $PSBoundParameters[$ParamName_DestinationFolder]
                    }
            
                    #Remove "\" character from the $DestinationFolder if exists.
                    IF($DestinationFolder.EndsWith("\"))
                    {
                        $DestinationFolder = $DestinationFolder.Substring(0,$DestinationFolder.Length-1)
                    }

                    $IconNamePrefix = $SourceFilePath.Split("\")[-1].Split(".")[0]

                    IF($PSBoundParameters.ContainsKey('AutoCreateSubFolders'))
                    {
                        $DestinationFolder = "$DestinationFolder\$IconNamePrefix"
                    }

                    IF(-not (Test-path -Path $DestinationFolder -PathType Container -ErrorAction SilentlyContinue))
                    {
                        Try {New-Item -ItemType "directory" -Path $DestinationFolder -Force | Out-Null
                             IF(!(Get-Item -LiteralPath $DestinationFolder).PSIsContainer)
                             {
                                Throw "Target folder specified [$DestinationFolder] is not a folder!"
                             }
                             Write-Host "Successfully created directory `"$DestinationFolder`" as the DestinationFolder for your icons!" -ForegroundColor White -BackgroundColor DarkGreen
                        }
                        Catch {Write-Host "`"ErrorMessage:$($Error[0].Exception.Message)`"" -BackgroundColor Red -ForegroundColor Yellow; Throw "Unable to create target icon folder `"$DestinationFolder]!`""}
                    }
            
                } #END IF($ExtractToFolder)

                IF($OutputFormat -eq "Icon")
                {
                    IF(!$PSBoundParameters.ContainsKey('IconQuality'))
                    {
                        #Write-Host "DynamicParameter `"$ParamName_IconQuality`" was NOT created, populating with default value of `"Bitmap`"" -BackgroundColor Yellow -ForegroundColor Red
                        $IconQuality = "Bitmap"
                    }
                    ELSE 
                    {
                        #Write-Host "Dynamic Parameter `"$ParamName_IconQuality`" was created at runtime, and contains $($PSBoundParameters[$ParamName_IconQuality])" -BackgroundColor DarkCyan -ForegroundColor White
                        $IconQuality = $PSBoundParameters[$ParamName_IconQuality]
                    }
                }
    
            } #END BEGIN

            PROCESS
            {
                #Add required assemblies
                Add-Type -AssemblyName System.Drawing, WindowsBase, PresentationCore
    
                #Add Win32API.Icon Class and ExtractIconEx, DeleteObject FUNCTIONS if not present
                try{[Win32API.Icon]|Out-Null
                    #Write-Host "Win32API Already Loaded"    
                }
                catch
                {
                    #Add Win32API.Icon Class and ExtractIconEx, DeleteObject FUNCTIONS if not present
                    Add-Type -Namespace Win32API -Name Icon -MemberDefinition @'
                    [DllImport("Shell32.dll", SetLastError=true)]
                    public static extern int ExtractIconEx(string lpszFile, int nIconIndex, out IntPtr phiconLarge, out IntPtr phiconSmall, int nIcons);
 
                    [DllImport("gdi32.dll", SetLastError=true)]
                    public static extern bool DeleteObject(IntPtr hObject);
'@      }

                [array]$Icons = $null
                FOREACH ($Index in $IconIndexRange)
                {
                    $IconIndexNo = $Index 
            
                    Try #Extract the Icon from the file
                    {
                        IF(($SourceFilePath.ToLower().Contains(".dll")) -or $SourceFilePath.ToLower().Contains(".exe"))
                        {
                            #Initialize variables for reference conversion
                            $large,$small = 0,0

                            #Call Win32 API Function for handles
                            [Win32API.Icon]::ExtractIconEx($SourceFilePath, $IconIndexNo, [ref]$large, [ref]$small, 1) | Out-Null

                            #If large icon desired store large handle, default to small handle
                            $handle = IF($LargeIcon){$large}ELSE{$small}

                            #Get the icon from the handle
                            IF($handle)
                            {
                                $Icon = [System.Drawing.Icon]::FromHandle($handle)
                            }

                            #If the handles are valid, delete them for good memory practice
                            $large, $small, $handle | Where-Object {$_} | ForEach-Object {[Win32API.Icon]::DeleteObject($_)} | Out-Null
                        }
                        ELSEIF(($SourceFilePath.ToLower().Contains(".bmp")) -or $SourceFilePath.ToLower().Contains(".ico") -or $SourceFilePath.ToLower().Contains(".gif") -or $SourceFilePath.ToLower().Contains(".tiff") -or $SourceFilePath.ToLower().Contains(".jpg") -or $SourceFilePath.ToLower().Contains(".exif") -or $SourceFilePath.ToLower().Contains(".png"))
                        {
                            #FUTURE DEVELOPMENT : Import Bitmap files and export as Icons - NOT FULLY TESTED
                            $Icon = New-Object System.Windows.Media.Imaging.BitmapImage -ArgumentList $SourceFilePath
                        }
                        ELSE #Extract Associated Icon 
                        {
                            $Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($SourceFilePath)
                        }
                    } 
                    Catch 
                    {
                        Throw "Error extracting ICO file"
                    }
                        
                    #Convert the Icon to a BitmapImage if selected by $OutputFormat parameter
                    IF($OutputFormat -eq "BitmapImage" -and $Icon.GetType().Name -ne "BitmapImage")
                    {
                        $Icon = $Icon.ToBitmap()
                    }
            
                    IF($ExtractToFolder)
                                                                                                                                                                                                        {
                    IF($OutputFormat -eq "BitmapImage")
                    {
                        $IconName = "$($IconNamePrefix)_$IconIndexNo.bmp"
                    }
                    ELSE
                    {
                        $IconName = "$($IconNamePrefix)_$IconIndexNo.ico"
                    }
                
                    $TargetIconFilePath = "$DestinationFolder\$IconName"
                    Try 
                    {
                        IF(Test-Path -LiteralPath $TargetIconFilePath -PathType Leaf -ErrorAction SilentlyContinue)
                        {
                            rm $TargetIconFilePath
                        }

                        $stream = [System.IO.File]::OpenWrite("$($TargetIconFilePath)")
                    
                        IF($OutputFormat -eq "BitmapImage")
                        {
                            $Icon.save($stream, [Drawing.Imaging.ImageFormat]::Bmp)
                        }
                        ELSE #OutputFormat was Icon
                        {
                            #IconQuality defaults to Bitmap, even if the user doesn't specify the parameter
                            #This is done to ensure the full color gamut is included by converting the Icon to a Bitmap image first
                            IF($IconQuality -eq "Bitmap" -and $Icon.GetType().Name -ne "BitmapImage")
                            {
                                $Icon = $Icon.ToBitmap()
                                $Icon.save($stream, [Drawing.Imaging.ImageFormat]::Bmp)
                            }
                            ELSE #The user passed in "Original" to the IconQuality parameter, overriding the default of Bitmap; this may result in color distortion in the Icon
                            {
                                $Icon.save($stream)
                            }
                        }
                        Start-Sleep -Milliseconds 10
                        $stream.close()
                        Start-Sleep -Milliseconds 10
                    }
                    Catch
                    {
                        Throw "Error saving ICO file [$TargetIconFilePath]"
                    }
                }
                    ELSE {$Icons+=$Icon}
                } #END FOREACH ($Index in $IconIndexRange)

                IF($ExtractToFolder)
                {
                    Write-Host "Your $OutputFormat(s) were extracted in the following directory: `"$DestinationFolder`"" -BackgroundColor DarkGreen -ForegroundColor White
                    Set-Location -LiteralPath $DestinationFolder
                    Get-ChildItem -LiteralPath $DestinationFolder
                }
                ELSE {RETURN $Icons}
            } #END PROCESS
            END{
                Invoke-VariableCleanup
            } #END END
        } #END FUNCTION Extract-Icons #Version 1.05
                
        ## Event handler
        FUNCTION SortListView
        {
            param([parameter(Position=0)][UInt32]$Column)
 
            $Numeric = $true # determine how to sort
 
            # If the user clicked the same column that was clicked last time, reverse its sort order. otherwise, reset for normal ascending sort
            IF($Script:LastColumnClicked -eq $Column)
            {
                $Script:LastColumnAscending = -not $Script:LastColumnAscending
            }
            ELSE
            {
                $Script:LastColumnAscending = $true
            }

            $Script:LastColumnClicked = $Column
            $ListItems = @(@(@())) # three-dimensional array; column 1 indexes the other columns, column 2 is the value to be sorted on, and column 3 is the System.Windows.Forms.ListViewItem object
 
            FOREACH($ListItem in $ListView.Items)
            {
                # If all items are numeric, can use a numeric sort
                IF($Numeric -ne $false) # nothing can set this back to true, so don't process unnecessarily
                {
                    try
                    {
                        $Test = [Double]$ListItem.SubItems[[int]$Column].Text
                    }
                    catch
                    {
                        $Numeric = $false # a non-numeric item was found, so sort will occur as a string
                    }
                }
                $ListItems += ,@($ListItem.SubItems[[int]$Column].Text,$ListItem)
            }
 
            # create the expression that will be evaluated for sorting
            $EvalExpression = {
                if($Numeric)
                { return [Double]$_[0] }
                else
                { return [String]$_[0] }
            }
 
            # all information is gathered; perform the sort
            $ListItems = $ListItems | Sort-Object -Property @{Expression=$EvalExpression; Ascending=$Script:LastColumnAscending}
 
            ## the list is sorted; display it in the listview
            $ListView.BeginUpdate()
            $ListView.Items.Clear()

            foreach($ListItem in $ListItems)
            {
                $ListView.Items.Add($ListItem[1])
            }
            $ListView.EndUpdate()
        }

        ## Event handler
        FUNCTION SearchListViewItems {
           param(
                [System.Windows.Forms.ListView]$ListView,
                [string]$SearchString
           )

           FOREACH($Item in $ListView.Items)
           {
               $MatchFound = $false
               $Index = $Item.Index
               
               IF($SearchString.Length -gt 0)
               {
                   FOREACH($SubItem in $Item.SubItems)
                   {
                       IF($SubItem.Text -match $SearchString)
                       {
                           $MatchFound = $true
                           IF($Script:FirstItemFound -eq $null)
                           {
                               $Script:FirstItemFound = $Item.Index
                           }
                       }
                       
                   }
               }
               ELSE
               {
                   $Script:FirstItemFound = 0
               }

               IF($MatchFound)
               {
                   $ListView.Items[$Item.Index].BackColor = [System.Drawing.Color]::FromName("PowderBlue")
               }
               ELSE
               {
                   $ListView.Items[$Item.Index].BackColor = [System.Drawing.Color]::FromName("White")
               }
           }
           $ListView.EnsureVisible($Script:FirstItemFound)
           $Script:FirstItemFound = $null
        }
                
        FUNCTION ClearAndClose()
        {
            $Timer.Stop(); 
            $form.Close(); 
            $form.Dispose();
            $Timer.Dispose();
            $Script:FormTimeout=$null
        }

        FUNCTION Timer_Tick()
        {
            $LBL_MessageBox.Text = "$Message`n`n`Form will auto-close in $Script:FormTimeout seconds"
            --$Script:FormTimeout
            IF($Script:FormTimeout -lt 0)
            {
                ClearAndClose
            }
        }

        IF($FormTimeout -ne $null)
        {
            $Script:FormTimeout = $FormTimeout
            $Timer = [System.Windows.Forms.Timer]::new()
            $Timer.Interval = 1000 #ms
            $Timer.Add_Tick({Timer_Tick})
        }

        FUNCTION OnClick_BTN_OK {
            $Script:ButtonPressed = $OkBtnLbl
        }

        FUNCTION OnClick_BTN_CANCEL {
            $Script:ButtonPressed = $CancelBtnLbl
        }

        FUNCTION OnClick_BTN_SelectAll {
            $Script:ButtonPressed = $SelectAllBtnLbl
        }

        FUNCTION OnClick_BTN_DeSelectAll {
            $Script:ButtonPressed = $DeSelectAllBtnLbl
        }

        FUNCTION OnClick_BTN_CUSTOM {
            $Script:ButtonPressed = $CustomBtnLabel
        }

        FUNCTION OnClick_BTN_CUSTOM2 {
            $Script:ButtonPressed = $Custom2BtnLabel
        }

        FUNCTION OnClick_BTN_YES {
            $Script:ButtonPressed = $YesBtnLbl
        }

        FUNCTION OnClick_BTN_NO {
            $Script:ButtonPressed = $NoBtnLbl
        }

        $form = [System.Windows.Forms.Form]::new()
        $form.Text = $FormTitle
        IF($AutoSize)
        {
            $form.AutoSize = $AutoSize
            $form.AutoSizeMode = "GrowAndShrink"
        }
        $form.Size = [System.Drawing.Size]::new($FormSizeHoriz,$FormSizeVert) 
        $form.StartPosition = $StartPosition
        IF($StartPosition -eq 'Manual')
        {
            $form.Location = [System.Drawing.Point]::new($FormLocationX,$FormLocationY)
        }
        $form.FormBorderStyle = $FormBorderStyle
        $form.MaximizeBox = $MaximizeBox
        $form.MinimizeBox = $MinimizeBox
        $form.Name = "FRM_ListView"
        IF($IconSourceFilePath.Length -ne 0 -and $IconIndex -ne $null)
        {
            $FormIcon = Extract-Icons -SourceFilePath $IconSourceFilePath -IconIndexRange $IconIndex -largeIcon -OutputFormat Icon -IconQuality Bitmap
            $form.Icon = $FormIcon
        }

        IF($IncludeCancelBtn)
        {
            $BTN_Cancel = [System.Windows.Forms.Button]::new()
            $BTN_Cancel.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_Cancel.Text = $CancelBtnLbl
            $BTN_Cancel.Size = [System.Drawing.Size]::new(100,30)
            $BTN_Cancel.AutoSize = $true
            $BTN_Cancel.AutoSizeMode = "GrowAndShrink"

            $BTN_Cancel.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($BTN_Cancel.Size.Width+15)),($FormSizeVert-($BTN_Cancel.Size.Height+55)))
            $BTN_Cancel.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_Cancel.Name = "BTN_CANCEL"
            
            $BTN_Cancel.Text = $CancelBtnLbl
            $BTN_Cancel.DialogResult = [System.Windows.Forms.DialogResult]::CANCEL
            $BTN_Cancel.Add_Click({OnClick_BTN_CANCEL})
            IF($AcceptButton -eq "CANCEL")
            {
                $form.AcceptButton = $BTN_Cancel
            }
            ELSEIF($CancelButton -eq "CANCEL")
            {
                $form.CancelButton = $BTN_Cancel
            }
            $BTN_Cancel.TabIndex = 100
            $form.Controls.Add($BTN_Cancel)
        }
                        
        IF($IncludeOkBtn)
        {
            $BTN_OK = [System.Windows.Forms.Button]::new()
            $BTN_OK.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_OK.Text = $OkBtnLbl
            $BTN_OK.Size = [System.Drawing.Size]::new(100,30)
            $BTN_OK.AutoSize = $true
            $BTN_OK.AutoSizeMode = "GrowAndShrink"

            $BTN_OK.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($BTN_Cancel.Size.Width + $BTN_OK.Size.Width + 30)),($FormSizeVert-($BTN_OK.Size.Height+55)))
            $BTN_OK.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_OK.Name = "BTN_OK"            
            
            $BTN_OK.DialogResult = [System.Windows.Forms.DialogResult]::OK
            $BTN_OK.Add_Click({OnClick_BTN_OK})
            IF($AcceptButton -eq "OK")
            {
                $form.AcceptButton = $BTN_OK
            }
            ELSEIF($CancelButton -eq "OK")
            {
                $form.CancelButton = $BTN_OK
            }
            $BTN_OK.TabIndex = 10
            $form.Controls.Add($BTN_OK)
        }
        
        IF($IncludeNoBtn)
        {
            $BTN_No = [System.Windows.Forms.Button]::new()
            $BTN_No.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_No.Text = $NoBtnLbl
            $BTN_No.Size = [System.Drawing.Size]::new(100,30)
            $BTN_No.AutoSize = $true
            $BTN_No.AutoSizeMode = "GrowAndShrink"
            IF($IncludeOkBtn -OR $IncludeCancelBtn){$Height2Add = 102}ELSE{$Height2Add = 55}
            $BTN_No.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($BTN_No.Size.Width+15)),($FormSizeVert-($BTN_No.Size.Height+$Height2Add)))
            $BTN_No.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_No.Name = 'BTN_NO'
            
            $BTN_No.DialogResult = [System.Windows.Forms.DialogResult]::NO
            $BTN_No.Add_Click({OnClick_BTN_NO})
            IF($AcceptButton -eq "NO")
            {
                $form.AcceptButton = $BTN_No
            }
            ELSEIF($CancelButton -eq "NO")
            {
                $form.CancelButton = $BTN_No
            }
            $BTN_No.TabIndex = 90
            $form.Controls.Add($BTN_No)
        }

        IF($IncludeYesBtn)
        {
            $BTN_Yes = [System.Windows.Forms.Button]::new()
            $BTN_Yes.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_Yes.Text = $YesBtnLbl
            $BTN_Yes.Size = [System.Drawing.Size]::new(100,30)
            $BTN_Yes.AutoSize = $true
            $BTN_Yes.AutoSizeMode = "GrowAndShrink"
            IF($IncludeOkBtn -OR $IncludeCancelBtn){$Height2Add = 102}ELSE{$Height2Add = 55}
            $BTN_Yes.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($BTN_No.Size.Width + $BTN_Yes.Size.Width + 30)),($FormSizeVert-($BTN_Yes.Size.Height+$Height2Add)))
            $BTN_Yes.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_Yes.Name = 'BTN_YES'
                        
            $BTN_Yes.DialogResult = [System.Windows.Forms.DialogResult]::YES
            $BTN_Yes.Add_Click({OnClick_BTN_YES})
            IF($AcceptButton -eq "YES")
            {
                $form.AcceptButton = $BTN_Yes
            }
            ELSEIF($CancelButton -eq "YES")
            {
                $form.CancelButton = $BTN_Yes
            }
            $BTN_Yes.TabIndex = 20
            $form.Controls.Add($BTN_Yes)
        }
                
        IF($IncludeCustomBtn)
        {
            $BTN_Custom = [System.Windows.Forms.Button]::new()
            $BTN_Custom.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_Custom.Text = $CustomBtnLabel
            $BTN_Custom.Size = [System.Drawing.Size]::new(100,30)
            $BTN_Custom.AutoSize = $true
            $BTN_Custom.AutoSizeMode = "GrowAndShrink"
            
            $BTN_Custom.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($FormSizeHoriz-15)),($FormSizeVert-($BTN_Custom.Size.Height+55)))
            $BTN_Custom.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_Custom.Name = "BTN_CUSTOM"
            
            $BTN_Custom.DialogResult = [System.Windows.Forms.DialogResult]::Retry
            $BTN_Custom.Add_Click({OnClick_BTN_CUSTOM})
            IF($AcceptButton -eq "CUSTOM1")
            {
                $form.AcceptButton = $BTN_Custom
            }
            ELSEIF($CancelButton -eq "CUSTOM1")
            {
                $form.CancelButton = $BTN_Custom
            }
            $BTN_Custom.TabIndex = 40
            $form.Controls.Add($BTN_Custom)
        }

        IF($IncludeCustom2Btn)
        {
            $BTN_Custom2 = [System.Windows.Forms.Button]::new()
            $BTN_Custom2.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_Custom2.Text = $Custom2BtnLabel
            $BTN_Custom2.Size = [System.Drawing.Size]::new(100,30)
            $BTN_Custom2.AutoSize = $true
            $BTN_Custom2.AutoSizeMode = "GrowAndShrink"
            
            $BTN_Custom2.Location = [System.Drawing.Point]::new($FormSizeHoriz-($FormSizeHoriz-($BTN_Custom.Size.Width+30)),($FormSizeVert-($BTN_Custom2.Size.Height+55)))
            $BTN_Custom2.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_Custom2.Name = "BTN_CUSTOM2"
            
            $BTN_Custom2.DialogResult = [System.Windows.Forms.DialogResult]::Abort
            $BTN_Custom2.Add_Click({OnClick_BTN_CUSTOM2})
            IF($AcceptButton -eq "CUSTOM2")
            {
                $form.AcceptButton = $BTN_Custom2
            }
            ELSEIF($CancelButton -eq "CUSTOM2")
            {
                $form.CancelButton = $BTN_Custom2
            }
            $BTN_Custom2.TabIndex = 70
            $form.Controls.Add($BTN_Custom2)
        }

        IF($IncludeSelectAllBtn)
        {
            $BTN_ALL = [System.Windows.Forms.Button]::new()
            $BTN_ALL.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_ALL.Text = $SelectAllBtnLbl
            $BTN_ALL.Size = [System.Drawing.Size]::new(100,30)
            $BTN_ALL.AutoSize = $true
            $BTN_ALL.AutoSizeMode = "GrowAndShrink"
            IF($IncludeCustomBtn -OR $IncludeCustom2Btn){$Height2Add = 102}ELSE{$Height2Add = 55}
            $BTN_ALL.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($FormSizeHoriz-15)),($FormSizeVert-($BTN_ALL.Size.Height+$Height2Add)))
            $BTN_ALL.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_ALL.Name = 'BTN_SelectAll'
                        
            $BTN_ALL.DialogResult = [System.Windows.Forms.DialogResult]::Yes
            $BTN_ALL.Add_Click({OnClick_BTN_SelectAll})
            IF($AcceptButton -eq "SELECTALL")
            {
                $form.AcceptButton = $BTN_ALL
            }
            ELSEIF($CancelButton -eq "SELECTALL")
            {
                $form.CancelButton = $BTN_ALL
            }
            $BTN_ALL.TabIndex = 30
            $form.Controls.Add($BTN_ALL)
        }

        IF($IncludeDeSelectAllBtn)
        {
            $BTN_NONE = [System.Windows.Forms.Button]::new()
            $BTN_NONE.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $BTN_NONE.Text = $DeSelectAllBtnLbl
            $BTN_NONE.Size = [System.Drawing.Size]::new(100,30)
            $BTN_NONE.AutoSize = $true
            $BTN_NONE.AutoSizeMode = "GrowAndShrink"
            IF($IncludeCustomBtn -OR $IncludeCustom2Btn){$Height2Add = 102}ELSE{$Height2Add = 55}
            $BTN_NONE.Location = [System.Drawing.Point]::new(($FormSizeHoriz-($FormSizeHoriz-($BTN_ALL.Size.Width+30))),($FormSizeVert-($BTN_NONE.size.Height+$Height2Add)))
            $BTN_NONE.Margin = [System.Windows.Forms.Padding]::new(2)
            $BTN_NONE.Name = 'BTN_DeSelectAll'
                        
            $BTN_NONE.DialogResult = [System.Windows.Forms.DialogResult]::No
            $BTN_NONE.Add_Click({OnClick_BTN_DeSelectAll})
            IF($AcceptButton -eq "DESELECTALL")
            {
                $form.AcceptButton = $BTN_NONE
            }
            ELSEIF($CancelButton -eq "DESELECTALL")
            {
                $form.CancelButton = $BTN_NONE
            }
            $BTN_NONE.TabIndex = 80
            $form.Controls.Add($BTN_NONE)
        }

        $LBL_ListHorizSz = $FormSizeHoriz - 35
        $LBL_ListVertSz = 35
        $LBL_List = [System.Windows.Forms.Label]::new()
        $LBL_List.Location = [System.Drawing.Point]::new(10,10) 
        $LBL_List.Size = [System.Drawing.Size]::new($LBL_ListHorizSz,$LBL_ListVertSz) 
        $LBL_List.Text = $ListLabelText
        $LBL_List.BackColor = "BLUE"
        $LBL_List.ForeColor = "WHITE"
        $form.Controls.Add($LBL_List)
                    
        $listView = [System.Windows.Forms.ListView]::new()
        $listView.View = "Details"
        $listView.Text = "ListView"
        $listView.Location = [System.Drawing.Point]::new($LBL_List.Location.X,($LBL_List.Location.Y + $LBL_ListVertSz)) 

        $ListViewHorizSz = $FormSizeHoriz - 35
        
        $SizeToSubtract = 110        
        IF(($IncludeCustomBtn -OR $IncludeCustom2Btn -OR $IncludeOkBtn -OR $IncludeCancelBtn) -AND `
         -NOT ($IncludeSelectAllBtn -OR $IncludeDeSelectAllBtn -OR $IncludeYesBtn -OR $IncludeNoBtn)){$SizeToSubtract += 85}
        ELSEIF(($IncludeSelectAllBtn -OR $IncludeDeSelectAllBtn -OR $IncludeYesBtn -OR $IncludeNoBtn) -AND `
         -NOT ($IncludeCustomBtn -OR $IncludeCustom2Btn -OR $IncludeOkBtn -OR $IncludeCancelBtn)){$SizeToSubtract += 85}
        ELSEIF(($IncludeCustomBtn -OR $IncludeCustom2Btn -OR $IncludeOkBtn -OR $IncludeCancelBtn) -AND  `
         ($IncludeSelectAllBtn -OR $IncludeDeSelectAllBtn -OR $IncludeYesBtn -OR $IncludeNoBtn)){$SizeToSubtract += 140}
        IF($ListViewSearchable){$SizeToSubtract += 30}
        IF($IncludeMsgBxLbl){$SizeToSubtract += 100}
        $ListViewVertSz = $FormSizeVert - $SizeToSubtract
        $listView.Size = [System.Drawing.Size]::new($ListViewHorizSz,$ListViewVertSz) 
        $listView.Scrollable = $true
        $listView.FullRowSelect = $true
        $listView.GridLines = $true
        $listView.HideSelection = $false
        $listView.MultiSelect = $MultiSelect

        $listView.Add_Click({
            $listView.Focus();
            $listView.Select();
        })

        IF($ListViewSortable)
        {
            ## Set up the event handler
            $listView.add_ColumnClick({SortListView $_.Column})
        }

        #Adding Columns to the ListView
        IF($ItemsListViewColumnNames.COUNT -ge 1)
        {
            FOREACH ($ItemsListViewColumnName IN $ItemsListViewColumnNames)
            {
                $listView.Columns.Add($ItemsListViewColumnName).Width = 133
                #Write-Host "Adding Column:$ItemsListViewColumnName from function param: -ItemsListViewColumnNames" -BackgroundColor Blue -ForegroundColor White
            }
        }
        ELSEIF($ItemsList.COUNT -ge 1)
        {
            $listView.Columns.Add("Column1").Width = 133
        }
        ELSEIF($ItemsListCollection.COUNT -ge 1)
        {
            [array]$ItemsListViewColumnNames = ($ItemsListCollection[0] | Get-Member -MemberType Property).Name
            FOREACH ($ItemsListViewColumnName IN $ItemsListViewColumnNames)
            {
                $listView.Columns.Add($ItemsListViewColumnName).Width = 133
            }
        }

        #Adding Items to the Columns
        IF (($ItemsList -ne $null) -or ($ItemsListCollection -ne $null))
        {
            #Populate the list box with items to choose from
            IF($ItemsList.Count -gt 0)
            {
                foreach ($Item in $ItemsList) 
                {
                    $ListView_Item = ([System.Windows.Forms.ListViewItem]::new($Item)).ToString()
                    [void]$listView.Items.AddRange($ListView_Item)
                }
            }
            ELSEIF($ItemsListCollection.Count -gt 0)
            {
                $ColumnCount = $ItemsListViewColumnNames.Count - 1
                foreach ($item in $ItemsListCollection)
                {   
                    $ColumnIndex = 0
                    $ColumnName = $ItemsListViewColumnNames[$ColumnIndex]
                                
                    IF($Item.$ColumnName -eq $null)
                    {
                        #Write-Host "NULL Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor white -BackgroundColor Magenta
                        $ColumnName = ""
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "DateTime")
                    {
                        #Write-Host "DateTime Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        $ColumnName = $Item.$ColumnName.ToString() 
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)       
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "ADPropertyValueCollection")
                    {
                        #Write-Host "ADPropertyValueCollection Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        $ADPropertyValueCollection = @($($Item.$ColumnName))
                        [string]$ColumnName = $null
                        FOREACH ($Value in $ADPropertyValueCollection)
                        {
                            $ColumnName += $Value.ToString() + ";"
                        }
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "ActiveDirectorySecurity")
                    {
                        #Write-Host "ActiveDirectorySecurity Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = "System.DirectoryServices.ActiveDirectorySecurity"
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "Guid")
                    {
                        #Write-Host "Guid Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName.Guid).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName) 
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "SecurityIdentifier")
                    {
                        #Write-Host "SecurityIdentifier Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName.Value).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)       
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "Boolean")
                    {
                        #Write-Host "Boolean Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)  
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "PackageFeatureState")
                    {
                        #Write-Host "PackageFeatureState Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "UInt32")
                    {
                        #Write-Host "UInt32 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "Int32")
                    {
                        #Write-Host "Int32 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "UInt64")
                    {
                        #Write-Host "UInt64 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "Int64")
                    {
                        #Write-Host "Int64 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        [string]$ColumnName = ($Item.$ColumnName).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                    ELSEIF($Item.$ColumnName.GetType().Name -eq "String")
                    {
                        #Write-Host "String Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($Item.$ColumnName)   
                    }
                    ELSE
                    {
                        Write-Host "Value Found:$(($Item.$ColumnName).ToString()) in $($ItemsListViewColumnNames[$ColumnIndex]) of DataType:$(($Item.$ColumnName).GetType().Name)" -ForegroundColor Yellow -BackgroundColor Red
                        [string]$ColumnName = ($Item.$ColumnName).ToString()
                        $ListView_Item = [System.Windows.Forms.ListViewItem]::new($ColumnName)
                    }
                                                
                    #Add individual subitems to each ListView_Item (i.e. column values)
                    While ($ColumnIndex -lt $ColumnCount)
                    {
                        $ColumnIndex++
                        #Write-Host "ColumnIndex:$ColumnIndex and ColumnCount:$ColumnCount" -ForegroundColor White -BackgroundColor DarkCyan
                                    
                            $ColumnValue = $ItemsListViewColumnNames[$ColumnIndex]
                                        
                            IF($Item.$ColumnValue -eq $null)
                            {
                                #Write-Host "NULL Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor white -BackgroundColor Blue
                                $ColumnValue = ""
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "DateTime")
                            {
                                #Write-Host "DateTime Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                $ColumnValue = $Item.$ColumnValue.ToString() 
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "ADPropertyValueCollection")
                            {
                                #Write-Host "ADPropertyValueCollection Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                $ADPropertyValueCollection = @($($Item.$ColumnValue))
                                [string]$ColumnValue = $null
                                FOREACH ($Value in $ADPropertyValueCollection)
                                {
                                    $ColumnValue += $Value.ToString() + ";"
                                }
                                [void]$ListView_Item.SubItems.Add($ColumnValue)       
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "ActiveDirectorySecurity")
                            {
                                #Write-Host "ActiveDirectorySecurity Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = "System.DirectoryServices.ActiveDirectorySecurity"
                                [void]$ListView_Item.SubItems.Add($ColumnValue)      
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "Guid")
                            {
                                #Write-Host "Guid Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue.Guid).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)      
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "SecurityIdentifier")
                            {
                                #Write-Host "SecurityIdentifier Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue.Value).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)      
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "Boolean")
                            {
                                #Write-Host "Boolean Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)      
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "PackageFeatureState")
                            {
                                #Write-Host "PackageFeatureState Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "UInt32")
                            {
                                #Write-Host "UInt32 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "Int32")
                            {
                                #Write-Host "Int32 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "UInt64")
                            {
                                #Write-Host "UInt64 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "Int64")
                            {
                                #Write-Host "Int64 Found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor Red -BackgroundColor Yellow
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                            ELSEIF($Item.$ColumnValue.GetType().Name -eq "String")
                            {
                                #Write-Host "$(($Item.$ColumnValue).ToString()) found in $($ItemsListViewColumnNames[$ColumnIndex])" -ForegroundColor White -BackgroundColor DarkGreen
                                [void]$ListView_Item.SubItems.Add($Item.$ColumnValue)
                            }
                            ELSE #IF($Item.$ColumnValue.GetType().Name -eq "String")
                            {                                            
                                Write-Host "Value Found:$(($Item.$ColumnValue).ToString()) in $($ItemsListViewColumnNames[$ColumnIndex]) of DataType:$(($Item.$ColumnValue).GetType().Name)" -ForegroundColor Yellow -BackgroundColor Red
                                [string]$ColumnValue = ($Item.$ColumnValue).ToString()
                                [void]$ListView_Item.SubItems.Add($ColumnValue)
                            }
                    } #end WHILE
                    [void]$ListView.items.AddRange($ListView_Item)
                } #END foreach ($item in $ItemsListCollection)
            } #END ELSEIF($ItemsListCollection.Count -gt 0)
            [void]$ListView.AutoResizeColumns(1)
            $ListView.TabIndex = 0
            $form.Controls.Add($listView) 
            $form.TopMost = $TopMost
            $form.BringToFront()
        }
        
        #Adding Search Capabilities if -ListViewSearchable Parameter passed
        IF($ListViewSearchable)
        {
            $LBL_Search = [System.Windows.Forms.Label]::new()
            $LBL_Search.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $LBL_Search.Text = "Search:"
            $LBL_Search.Size = [System.Drawing.Size]::new(100,30)
            $LBL_Search.AutoSize = $true
            $LBL_Search.Location = [System.Drawing.Point]::new($listView.Location.X, ($LBL_ListView.Location.Y + $ListViewVertSz + 55))
            $LBL_Search.Margin = [System.Windows.Forms.Padding]::new(2)
            $LBL_Search.Name = 'LBL_Search' 
            $form.Controls.Add($LBL_Search)
            
            $TXTBX_SearchSize = ($FormSizeHoriz - $LBL_Search.Size.Width - 45)
            $TXTBX_Search = [System.Windows.Forms.TextBox]::new()
            $TXTBX_Search.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $TXTBX_Search.Size = [System.Drawing.Size]::new($TXTBX_SearchSize,30)
            $TXTBX_Search.Location = [System.Drawing.Point]::new(($LBL_Search.Location.X + $LBL_Search.Size.Width + 10),$LBL_Search.Location.Y)
            $TXTBX_Search.Margin = [System.Windows.Forms.Padding]::new(2)
            $TXTBX_Search.Name = 'TXTBX_Search' 
            $TXTBX_Search.add_TextChanged({SearchListViewItems -ListView $ListView -SearchString $TXTBX_Search.Text})  
            $TXTBX_Search.TabIndex = 1
            $form.Controls.Add($TXTBX_Search)
        }
        
        #Adding a Label for Error Messages
        $LBL_ErrorMessage = [System.Windows.Forms.Label]::new()
        $LBL_ErrorMessage.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
        $LBL_ErrorMessageHorizSz = $FormSizeHoriz - 35
        $LBL_ErrorMessageVertSz = 35
        $LBL_ErrorMessage.Size = [System.Drawing.Size]::new($LBL_ErrorMessageHorizSz,$LBL_ErrorMessageVertSz) 

        IF($ListViewSearchable)
        {
            $LBL_ErrorMessage.Location = [System.Drawing.Point]::new($LBL_Search.Location.X, ($TXTBX_Search.Location.Y + $TXTBX_Search.Size.Height + 10))
        }
        ELSE
        {
            $LBL_ErrorMessage.Location = [System.Drawing.Point]::new($listView.Location.X, ($LBL_ListView.Location.Y + $ListViewVertSz + 55))
        }
        $LBL_ErrorMessage.Margin = [System.Windows.Forms.Padding]::new(2)


        IF($IncludeMsgBxLbl)
        {
            $LBL_MessageBox = [System.Windows.Forms.Label]::new()
            $LBL_MessageBox.Font = [System.Drawing.Font]::new("Microsoft Sans Serif", 9.75,[System.Drawing.FontStyle]::Regular,[System.Drawing.GraphicsUnit]::Point, 0)
            $LBL_MessageBox.BackColor = [System.Drawing.Color]::FromName("PowderBlue")
            $LBL_MessageBox.Text = $Message
            $LBL_MessageBox.Size = [System.Drawing.Size]::new(($FormSizeHoriz-35),100)
            $LBL_MessageBox.Location = [System.Drawing.Point]::new($LBL_ErrorMessage.Location.X, ($LBL_ErrorMessage.Location.Y + $LBL_ErrorMessageVertSz +10))
            $LBL_MessageBox.Name = "LBL_MessageBox"
            $form.Controls.Add($LBL_MessageBox)
        }
    }

    PROCESS 
    {   
        IF(($ItemsList -eq $null) -and ($ItemsListCollection -eq $null)) #Stops a potential looping issue, when no items are passed into the function.
        {
            RETURN
            $form.Close()
        }

        [ARRAY]$itemsSelected = $null
        do
        {
            IF ($Message.Length -gt 0)
            {
                $LBL_ErrorMessage.Text = $Message
                $LBL_ErrorMessage.BackColor = "yellow"
                $LBL_ErrorMessage.ForeColor = "red"
                $form.TopMost = $TopMost
                $form.BringToFront()
                $Message = $null
                $form.Controls.Add($LBL_ErrorMessage)
            }

            $result = $form.ShowDialog()

            if ($result -eq [System.Windows.Forms.DialogResult]::OK) #USER SELECTED "OK" Button on the Form
            {   
                IF($ReturnObjects)
                {
                    $itemsSelectedIndices = ($listView.SelectedItems).Index
                    FOREACH ($itemSelectedIndex in $itemsSelectedIndices)
                    {
                        $itemsSelected += $ItemsListCollection[$itemSelectedIndex]
                    }
                }
                ELSE
                {
                    $itemsSelected = @(($listView.SelectedItems).Text)
                }
                                               

                If ($itemsSelected -ne $null)
                {
                    $Finished = $true
                }
                ELSE
                {
                    $Message = "MESSAGE: You failed to select one or more items from the list above, either select items and then OK to continue, or select Cancel to abort:"
                }
            } 
            elseif ($result -eq [System.Windows.Forms.DialogResult]::Cancel)  #USER SELECTED "Cancel" Button on the Form
            {
                Write-Host "User CANCELED the request..." -BackgroundColor Yellow -ForegroundColor Red
                $Finished = $true
                $itemsSelected = "CANCELED"
            }
            elseif ($result -eq [System.Windows.Forms.DialogResult]::Yes)  #USER SELECTED "Select ALL" Button on the Form
            {
                $listViewCount = $listView.items.count -1
                $counter = 0
                $listview.BeginUpdate();
                while ($counter -le $listViewCount)
                {
                    $listView.items[$Counter].Selected = $true
                    $counter++
                }
                $listview.EndUpdate();
                $listView.Select();
            }
            elseif ($result -eq [System.Windows.Forms.DialogResult]::No)  #USER SELECTED "De-Select ALL" Button on the Form
            {
                $listViewCount = $listView.items.count -1
                $counter = 0
                $listview.BeginUpdate();
                while ($counter -le $listViewCount)
                {
                    $listView.Items[$Counter].Selected = $false #($counter, $false)
                    $counter++
                }
                $listview.EndUpdate();
                $listView.Select();
            }
        }WHILE ($itemsSelected -eq $null)   
        
        RETURN $itemsSelected
        $form.Close()
    } #END PROCESS
    END {} #END END
} #END FUNCTION Select-ItemsFromListViewForm #Version 2.02

FUNCTION Get-WindowsCapabilities #VERSION 1.063
{
    <#
    .Synopsis
       Description:
       This script retrieves a list of the installed Windows Capabilities (Features On Demand) for all Computer Objects passed into it via a CSV, array, or looking up in AD
       There are three parameters which can be used to pass in Computer Objects for processing.
       1.) Computer Objects can be passed in as either CN or sAMAccount Name using the -ComputerNames parameter option.
            EXAMPLE: -ComputerNames "ComputerName1", "ComputerName2", "ComputerName3"  

       2.) Computer Objects can be pulled in from a CSV file by using the -ComputerListCSV parameter option; this works as long as the CSV file contains a "ComputerName" column, or a "sAMAccountName" column.
            EXAMPLE: -ComputerListCSV "C:\Path2File\NameOfTheCSVFile2Process.csv"

       3.) Computer Objects can be pulled in from Active Directory by using the -SearchAD parameter option, by connecting to the same domain that authenticated the user currently running the script.
            EXAMPLE: -SearchAD 


       Regardless of parameter option selected above (i.e. -ComputerNames, -ComputerListCSV, or -SearchAD), you must specify the targetOU that you want the script to use for the SearchBase in AD.
       This is done by selecting one of the two following parameters.
       1.) -targetOU parameter with the name of the Distinguished Name path to the OU in AD where you want the search to begin.
            EXAMPLE: -targetOU "OU=BLISS,OU=INSTALLATIONS,DC=NASW,DC=DS,DC=ARMY,DC=MIL"

       2.) -LocateTargetOU parameter will pop open a Windows Form, allowing you to drill down into the AD OU heirarchy where you want to begin your search.
            EXAMPLE: -LocateTargetOU


       After selecting one of the targetOU selection paramters above, you can optionally specify the -SearchScope parameter to define the scope of the search in AD (i.e. Base, OneLevel, and Subtree)
       1.) EXAMPLE: -SearchScope Base        This will ONLY search the current OU that was specified by the targetOU from the parameters above.
       2.) EXAMPLE: -SearchScope OneLevel    This will ONLY search OneLevel below the current OU that was specified by the targetOU from the parameters above.
       3.) EXAMPLE: -SearchScope Subtree     This will search all sub-OUs, to include the current OU that was specified by the targetOU from the parameters above.
       

       In order to create a log of your actions, select the optional -RunningResults2csv parameter.  By default, this will create a log directory at "$env:SystemDrive\LOGS\StaleADComputerObjects\" if needed.
       1.) EXAMPLE: -RunningResults2csv


       In order to specify an alternate location for your log files, select the optional -ResultsFilePath parameter.
       1.) EXAMPLE: -ResultsFilePath "C:\Your\Own\Path\"


       If desired, you can specify a Capability Name Filter by selecting the optional -CapabilityNameFilter parameter.  
       This will filter the Installed Windows Capabilities to match the name provided.  
       or change the description field on the AD Computer Objects.
       1.) EXAMPLE: -CapabilityNameFilter RSAT

       Additionally, you can filter for ONLY Capabilities that are INSTALLED, or those Capabilities that are NOT Present, using the optional parameters, OnlyIfInstalled and OnlyIfNotPresent
       1.) EXAMPLE: -OnlyIfInstalled
       2.) EXAMPLE: -OnlyIfNotPresent

              
       If the -RunningResults2csv parameter is selected, the script outputs a single XLSX workbook for each AD Computer Object with several key attributes, including:
            AD_ComputerName, AD_OperatingSystem, CapabilityName, CapabilityState and AD_CononicalName path, in the C:\Logs\WindowsCapabilities directory. 
       Additionally, it will create a summary file (XLSX) containing several key attributes, including:
            AD_ComputerName, AD_OperatingSystem, AD_Created, AD_Enabled, AD_LastLogonTimeStamp, AD_IPv4Addresss, AD_CononicalName, AD_Description, 
            PS1_VerifiedComputer, FOD_CAPABILITIES_Installed_COUNT, FOD_CAPABILITIES_NotPresent_COUNT, in the C:\Logs\WindowsCapabilities directory.

       
       By default the Summary LOG results are then written to the "C:\Logs\WindowsCapabilities\WindowsCapabilitiesSummary_RunningResults_MM-DD-YYYY_HHmmss.XLSX" file, 
       while the individual computer results are written to the "C:\Logs\WindowsCapabilities\HOSTNAME_WindowsCapabilities_MM-DD-YYYY_HHmmss.XLSX" file.

              
    .VERSION: 1.0 (9 MARCH 2020)
         
    .EXAMPLE
       Run this script in the CONUS Forest on a computer that has the Admin Tool (RSAT) installed.  This ensures that PowerShell with the necessary ActiveDirectory modules can be loaded.

    .INPUTS
       User must select "WHERE" the computer objects are being passed in from (i.e. -ComputerNames, -ComputerListCSV, -SearchAD), and "WHICH" OU from he/she desires the search to begin (i.e. -targetOU, -LocateTargetOU).
       OPTIONALLY: 
            The Search Scope can be defined with -SearchScope parameter.
            Logging (HIGHLY RECOMMENDED) can be selected with -RunningResults2csv parameter.
            Logpath can be defined with -ResultsFilePath parameter.
            Capability Name Filtering can be defined with -CapabilityNameFilter parameter
            Only If Installed filtering can be defined with -OnlyIfInstalled parameter switch
            Only If Not Present filtering can be defined with -OnlyIfNotPresent parameter switch

    .OUTPUTS
       One XLSX Summary file is created in the "$env:SystemDrive\LOGS\WindowsCapabilities\" folder, with the name "C:\Logs\WindowsCapabilities\WindowsCapabilitiesSummary_RunningResults_MM-DD-YYYY_HHmmss.XLSX"
       One XLSX file is created per computer, in the "$env:SystemDrive\LOGS\WindowsCapabilities\" folder, with the name "C:\Logs\WindowsCapabilities\HOSTNAME_WindowsCapabilities_MM-DD-YYYY_HHmmss.XLSX"

    .AUTHOR 
        Michael D. Sloan
        Fort Bliss SANEC
        michael.d.sloan.civ@mail.mil
        DSN: 312-711-0744
        COMM: +1 915-741-0744

    #>


    #Asseses Target computer or computers passed in as a parameter, for Pingability, Remote Registry access and Windows Remote Management, returns $ComputerInfoResults
    [cmdletbinding()]    
    #Parameters passed into the funtion
    Param (
        [Parameter(Mandatory=$false,ValueFromPipelineByPropertyName=$true,HelpMessage="Provide one or more ComputerNames to lookup in ActiveDirectory")]
        [Alias("COMPUTER")]
        [ARRAY]$ComputerNames,
    
        [Parameter(Mandatory=$false,HelpMessage="Optionally provide a filepath to a list of COMPUTERS in CSV file format, in order to import for the AD Lookup.")]
        [string]$ComputerListCSV,

        [Parameter(Mandatory=$false,HelpMessage="Select this option to manually search AD base on an OU passed into -targetOU parameter, or -LocateTargetOU")]
        [switch]$SearchAD,

        [Parameter(Mandatory=$false,HelpMessage="Select this parameter option in order to use a GUI to locate the target OU in AD where you want to begin your search.")]
        [switch]$LocateTargetOU,

        [Parameter(Mandatory=$false,HelpMessage="Provide the targetOU as your AD search base, i.e. OU=InstallationName,OU=Installations,DC=Region,DC=DS,DC=ARMY,DC=MIL")]
        [string]$targetOU,

        [Parameter(Mandatory=$false,HelpMessage="Select `"Base`" to search only the current TargetOU path, `"OneLevel`" to search the immediate children of the TargetOU path, `"Subtree`" to search the current TargetOU path, and all children of that path.",Position=2)]
        [ValidateSet('Base', 'OneLevel', 'Subtree')]
        [string]$SearchScope='Subtree',

        [Parameter(Mandatory=$false,HelpMessage="Select this option to write a CSV file, while running...useful for large queries where 1000's of computers will be processed", Position=3)]
        [switch]$RunningResults2csv,
                
        [Parameter(Mandatory=$false,HelpMessage="Optionally provide the filepath where you want the results to be written to.",Position=4)]
        [string]$ResultsFilePath = "$env:SystemDrive\LOGS\WindowsCapabilities\", #Set Root Directory for the CSV Output File

        [Parameter(Mandatory=$false,HelpMessage="Optionally provide a filter for the capability name.",Position=5)]
        [string]$CapabilityNameFilter,

        [Parameter(Mandatory=$false,HelpMessage="Optionally select only installed capabilities.",Position=6)]
        [switch]$OnlyIfInstalled,

        [Parameter(Mandatory=$false,HelpMessage="Optionally select only staged capabilities.",Position=7)]
        [switch]$OnlyIfStaged,

        [Parameter(Mandatory=$false,HelpMessage="Optionally select only capabilities which are Not Present.",Position=8)]
        [switch]$OnlyIfNotPresent

    ) #END OF PARAMTER DEFINITIONS

    BEGIN
    {

        [float]$MinVersion = 2.34
        $PsExec64LocalPath = "$env:Windir\system32\PsExec64.exe"
        $PsExec64RemoteShare = "\\blisw6syaaa7nec\IMO_Info\Available Software & Fixes\PS_Tools\PsExec64.exe"

        IF (!(Test-Path $PsExec64LocalPath))
        {
            try {
                    [float]$ProductVersionOnShare = (Get-ChildItem -LiteralPath $PsExec64RemoteShare | Select-Object -Property VersionInfo).VersionInfo.ProductVersion
                    IF($ProductVersionOnShare -lt $MinVersion)
                    {
                        throw "The version of PSExec64.exe on the file share `"$PsExec64RemoteShare`" is not greater than or equal to the minimum required version of `"$MinVersion`"`
                        Please update your file share with the latest copy from Microsoft."
                    }

                    COPY $PsExec64RemoteShare -Destination $PsExec64LocalPath -Force
                    IF($INTERACTIVE)
                    {
                        Write-Host "Successfully copied `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor DarkGreen -ForegroundColor White
                    }
                }
            catch [System.Exception]
                {
                    IF($INTERACTIVE)
                    {
                        Write-Host "Unable to copy `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor Red -ForegroundColor White
                    }
                    Write-Warning -Message $_.Exception.Message
                    break
                }
        }
        ELSE
        {
            [float]$ProductVersion = (Get-ChildItem -LiteralPath $PsExec64LocalPath | Select-Object -Property VersionInfo).VersionInfo.ProductVersion
            IF($ProductVersion -lt $MinVersion)
            {
                try {
                        [float]$ProductVersionOnShare = (Get-ChildItem -LiteralPath $PsExec64RemoteShare | Select-Object -Property VersionInfo).VersionInfo.ProductVersion
                        IF($ProductVersionOnShare -lt $MinVersion)
                        {
                            throw "The version of PSExec64.exe on the file share `"$PsExec64RemoteShare`" is not greater than or equal to the minimum required version of `"$MinVersion`"`
                            Please update your file share with the latest copy from Microsoft."
                        }

                        COPY $PsExec64RemoteShare -Destination $PsExec64LocalPath -Force
                        IF($INTERACTIVE)
                        {
                            Write-Host "Successfully copied `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor DarkGreen -ForegroundColor White
                        }
                    }
                catch [System.Exception]
                    {
                        IF($INTERACTIVE)
                        {
                            Write-Host "Unable to copy `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor Red -ForegroundColor White
                        }
                        Write-Warning -Message $_.Exception.Message
                        break
                    }
            } #END IF($ProductVersion -lt $MinVersion)
        } #END ELSE

        FUNCTION Choose-ADOUForm #Version 1.15
        {
            param(
                [switch]$MultiSelect
            )

            BEGIN
            {
                ####################################### GET YOUR LOCAL DOMAIN NFO #######################################################################################
                $DomainName = $env:USERDOMAIN
                $Root = [ADSI]"LDAP://RootDSE"
                Try { $DomainRoot = ($Root.Get("rootDomainNamingContext")).ToUpper() }
                Catch { Write-Warning "Unable to contact Active Directory because $($Error[0]); aborting script!"; Start-Sleep -Seconds 5; Exit }
                $DomainDN = "DC=$DomainName,$DomainRoot"
                ####################################### END: GET YOUR LOCAL DOMAIN NFO ##################################################################################


                ####################################### Set $InstallationsDN path, based on domain ###########################################################################
                $InstallationsDN = "OU=Installations," + $DomainDN
                ####################################### END: Set $InstallationsDN path, based on domain ######################################################################
        
        
                ####################################### GET YOUR LOCAL SITE NAME ############################################################################################
                $sAMAccountName = "$env:ComputerName`$"
                $searcher = [adsisearcher]"(&(objectClass=Computer)(sAMAccountName=$sAMAccountName))"
                $searcher.SearchRoot = "LDAP://$($InstallationsDN)"
                $searcher.SearchScope = 'Subtree'
                $searcher.PropertiesToLoad.Add('CanonicalName') | Out-Null
                $Site = ($searcher.Findall()).properties.canonicalname.Split('/')[2].ToUpper() 
                ####################################### END: GET YOUR LOCAL SITE NAME #######################################################################################


                ####################################### SET YOUR LOCAL baseDN PATH ############################################################################################
                $baseDN = "OU=$Site,$InstallationsDN"
                ####################################### END: SET YOUR LOCAL baseDN PATH ############################################################################################
                
                FUNCTION Get-CheckedNodesDNs {
                    param($Nodes)
            
                    FOREACH($Node in $Nodes){
                        IF($Node.Nodes.Count -gt 0) # RECURSION
                        {
                            Get-CheckedNodesDNs $Node.Nodes
                        }
                        IF($Node.Checked){
                            $NodeOU = [adsi]$Node.Name
                            $Script:NodesNames += ($NodeOU.DistinguishedName).ToString()
                        }
                    }
                }               

                #region Import the Assemblies 
                [reflection.assembly]::loadwithpartialname("System.Windows.Forms") | out-Null 
                [reflection.assembly]::loadwithpartialname("System.Drawing") | out-Null 
 
                #region Generated Form objects 
                $form1 = New-object System.Windows.Forms.Form 
                $form1.StartPosition = "CenterScreen"
                $treeView1 = New-object System.Windows.Forms.TreeView
                $imagelist1 = New-Object System.Windows.Forms.ImageList
                $label1 = New-object System.Windows.Forms.Label
                $textbox1 = New-object System.Windows.Forms.TextBox
                $InitialFormWindowState = New-object System.Windows.Forms.FormWindowState 
        
                #region   OK Button
                $OKButton = New-Object System.Windows.Forms.Button
                $OKButton.Location = '125,460'
                $OKButton.Size = '75,23'
                $OKButton.Text = "OK"

                # Got rid of the Click event for OK Button, and instead just assigned its DialogResult property to OK.
                $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

                $form1.Controls.Add($OKButton)

                # Setting the form's AcceptButton property causes it to automatically intercept the Enter keystroke and
                # treat it as clicking the OK button (without having to write your own KeyDown events).
                $form1.AcceptButton = $OKButton
                #endregion 
        
                #region Cancel Button
                $CancelButton = New-Object System.Windows.Forms.Button
                $CancelButton.Location = '215,460'
                $CancelButton.Size = '75,23'
                $CancelButton.Text = "Cancel"

                # Got rid of the Click event for Cancel Button, and instead just assigned its DialogResult property to Cancel.
                $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

                $form1.Controls.Add($CancelButton)

                # Setting the form's CancelButton property causes it to automatically intercept the Escape keystroke and
                # treat it as clicking the OK button (without having to write your own KeyDown events).
                $form1.CancelButton = $CancelButton
                #endregion 


                #
                # imagelist1
                #
                $Formatter_binaryFomatter = New-Object System.Runtime.Serialization.Formatters.Binary.BinaryFormatter
                #region Binary Data
                $System_IO_MemoryStream = New-Object System.IO.MemoryStream ( , [byte[]][System.Convert]::FromBase64String('
                AAEAAAD/////AQAAAAAAAAAMAgAAAFdTeXN0ZW0uV2luZG93cy5Gb3JtcywgVmVyc2lvbj00LjAu
                MC4wLCBDdWx0dXJlPW5ldXRyYWwsIFB1YmxpY0tleVRva2VuPWI3N2E1YzU2MTkzNGUwODkFAQAA
                ACZTeXN0ZW0uV2luZG93cy5Gb3Jtcy5JbWFnZUxpc3RTdHJlYW1lcgEAAAAERGF0YQcCAgAAAAkD
                AAAADwMAAABwCgAAAk1TRnQBSQFMAgEBBAEAAUABAAFAAQABEAEAARABAAT/AQkBAAj/AUIBTQE2
                AQQGAAE2AQQCAAEoAwABQAMAASADAAEBAQABCAYAAQgYAAGAAgABgAMAAoABAAGAAwABgAEAAYAB
                AAKAAgADwAEAAcAB3AHAAQAB8AHKAaYBAAEzBQABMwEAATMBAAEzAQACMwIAAxYBAAMcAQADIgEA
                AykBAANVAQADTQEAA0IBAAM5AQABgAF8Af8BAAJQAf8BAAGTAQAB1gEAAf8B7AHMAQABxgHWAe8B
                AAHWAucBAAGQAakBrQIAAf8BMwMAAWYDAAGZAwABzAIAATMDAAIzAgABMwFmAgABMwGZAgABMwHM
                AgABMwH/AgABZgMAAWYBMwIAAmYCAAFmAZkCAAFmAcwCAAFmAf8CAAGZAwABmQEzAgABmQFmAgAC
                mQIAAZkBzAIAAZkB/wIAAcwDAAHMATMCAAHMAWYCAAHMAZkCAALMAgABzAH/AgAB/wFmAgAB/wGZ
                AgAB/wHMAQABMwH/AgAB/wEAATMBAAEzAQABZgEAATMBAAGZAQABMwEAAcwBAAEzAQAB/wEAAf8B
                MwIAAzMBAAIzAWYBAAIzAZkBAAIzAcwBAAIzAf8BAAEzAWYCAAEzAWYBMwEAATMCZgEAATMBZgGZ
                AQABMwFmAcwBAAEzAWYB/wEAATMBmQIAATMBmQEzAQABMwGZAWYBAAEzApkBAAEzAZkBzAEAATMB
                mQH/AQABMwHMAgABMwHMATMBAAEzAcwBZgEAATMBzAGZAQABMwLMAQABMwHMAf8BAAEzAf8BMwEA
                ATMB/wFmAQABMwH/AZkBAAEzAf8BzAEAATMC/wEAAWYDAAFmAQABMwEAAWYBAAFmAQABZgEAAZkB
                AAFmAQABzAEAAWYBAAH/AQABZgEzAgABZgIzAQABZgEzAWYBAAFmATMBmQEAAWYBMwHMAQABZgEz
                Af8BAAJmAgACZgEzAQADZgEAAmYBmQEAAmYBzAEAAWYBmQIAAWYBmQEzAQABZgGZAWYBAAFmApkB
                AAFmAZkBzAEAAWYBmQH/AQABZgHMAgABZgHMATMBAAFmAcwBmQEAAWYCzAEAAWYBzAH/AQABZgH/
                AgABZgH/ATMBAAFmAf8BmQEAAWYB/wHMAQABzAEAAf8BAAH/AQABzAEAApkCAAGZATMBmQEAAZkB
                AAGZAQABmQEAAcwBAAGZAwABmQIzAQABmQEAAWYBAAGZATMBzAEAAZkBAAH/AQABmQFmAgABmQFm
                ATMBAAGZATMBZgEAAZkBZgGZAQABmQFmAcwBAAGZATMB/wEAApkBMwEAApkBZgEAA5kBAAKZAcwB
                AAKZAf8BAAGZAcwCAAGZAcwBMwEAAWYBzAFmAQABmQHMAZkBAAGZAswBAAGZAcwB/wEAAZkB/wIA
                AZkB/wEzAQABmQHMAWYBAAGZAf8BmQEAAZkB/wHMAQABmQL/AQABzAMAAZkBAAEzAQABzAEAAWYB
                AAHMAQABmQEAAcwBAAHMAQABmQEzAgABzAIzAQABzAEzAWYBAAHMATMBmQEAAcwBMwHMAQABzAEz
                Af8BAAHMAWYCAAHMAWYBMwEAAZkCZgEAAcwBZgGZAQABzAFmAcwBAAGZAWYB/wEAAcwBmQIAAcwB
                mQEzAQABzAGZAWYBAAHMApkBAAHMAZkBzAEAAcwBmQH/AQACzAIAAswBMwEAAswBZgEAAswBmQEA
                A8wBAALMAf8BAAHMAf8CAAHMAf8BMwEAAZkB/wFmAQABzAH/AZkBAAHMAf8BzAEAAcwC/wEAAcwB
                AAEzAQAB/wEAAWYBAAH/AQABmQEAAcwBMwIAAf8CMwEAAf8BMwFmAQAB/wEzAZkBAAH/ATMBzAEA
                Af8BMwH/AQAB/wFmAgAB/wFmATMBAAHMAmYBAAH/AWYBmQEAAf8BZgHMAQABzAFmAf8BAAH/AZkC
                AAH/AZkBMwEAAf8BmQFmAQAB/wKZAQAB/wGZAcwBAAH/AZkB/wEAAf8BzAIAAf8BzAEzAQAB/wHM
                AWYBAAH/AcwBmQEAAf8CzAEAAf8BzAH/AQAC/wEzAQABzAH/AWYBAAL/AZkBAAL/AcwBAAJmAf8B
                AAFmAf8BZgEAAWYC/wEAAf8CZgEAAf8BZgH/AQAC/wFmAQABIQEAAaUBAANfAQADdwEAA4YBAAOW
                AQADywEAA7IBAAPXAQAD3QEAA+MBAAPqAQAD8QEAA/gBAAHwAfsB/wEAAaQCoAEAA4ADAAH/AgAB
                /wMAAv8BAAH/AwAB/wEAAf8BAAL/AgAD//8A/wD/AP8ABQAk/wL0Bv8D9AP/DPQD/wp0BHMC/wp0
                BHMC/wPsAesBbQH3Av8BBwLsAesBcgFtAfQC/wwqA/8BdAGaA3kBegd5AXMC/wF0AZoDeQF6B3kB
                cwL/AfcBBwGYATQBVgH3Av8BvAHvAQcBVgE5AXIB9AL/AVEBHAF0A3MFUQEqA/8BeQKaBUsFmgF0
                Av8BeQyaAXQC/wHvAQcB7wJ4AZIC8QMHAXgBWAHrAfQC/wF0ApkCeQN0A1IBKgP/AXkCmgFLA1EB
                KgWaAXQC/wF5DJoBdAL/Ae8CBwHvAZIC7AFyAe0CBwLvAewB9AL/AZkCGgGgBJoCegF5AVID/wF5
                AaABmgF5AZkCeQFRBZoBdAL/AXkBoAuaAXQC/wEHAe8C9wLtAXgBNQF4Ae8D9wHsAfQC/wGZAhoB
                oASaAnoBeQFSA/8BeQGgAZoCmQGgAXkBUgWaAXQC/wF5AaALmgF0Av8BBwPvAfcB7QGYAXgBmQEH
                A+8B7AP/AZkCGgGgBJoCegF5AVID/wGZAaABmgGZAXkBmgF5AVIFmgF0Av8BmQGgC5oBdAL/AbwD
                8wG8AZIBBwHvAQcB8QLzAfIB7QP/AZkCGgGgBJoCegF5AVID/wGZAaABmgJ5AXQCUgWaAXQC/wGZ
                AaALmgF0Av8BvAEHAu8B9wPtAe8CBwLvAe0D/wGZARoBmgKZBnkBUgP/AZkBwwGaBHQBeQGgBJoB
                dAL/AZkBwwaaAaAEmgF0Av8CvAIHAvcCBwO8AgcBkgP/AZkBGgGZAxoDmgFSAXkBUgP/AZkBwwOa
                AqABmQWaAXQC/wGZAcMDmgKgAZkFmgF0Av8CvAHrAewCBwLzAfABvAHtAW0BBwH3A/8BmQEaAZkC
                9gTDAVIBeQFSA/8BmQWgAZoCdAV5Av8BmQWgAZoCdAV5Av8BvAEHApIB7wH3ApIB7wG8Ae8BkgHv
                AfcD/wGZAhoC9gTDAVgBeQFSA/8BmQGaBBoBdAOaApkBmgF5Av8BeQGaBBoBdAOaApkBmgF5Av8D
                9AHyAbwB8QK8Ae8B8AT0A/8BmQMaApkDeQFYAXkBUgP/ARsGeQGaAvYB1gG0AZoBmQL/AZkGeQGa
                AvYB1gG0AZoBeQX/AfQBvAH3ARIB7AHvAfAH/wFRARwBeQN0AVIEUQEqCf8BwwZ5AcMI/wGaBnkB
                mgX/AfQBvAEHAu8B9wHxB/8MUUL/AUIBTQE+BwABPgMAASgDAAFAAwABIAMAAQEBAAEBBgABARYA
                A///AAIACw=='))
                #endregion
                $imagelist1.ImageStream = $Formatter_binaryFomatter.Deserialize($System_IO_MemoryStream)
                $Formatter_binaryFomatter = $null
                $System_IO_MemoryStream = $null
                $imagelist1.TransparentColor = 'Transparent'

        
                #endregion Generated Form objects 
 
                #---------------------------------------------- 
                #Generated Event Script Blocks 
                #---------------------------------------------- 
                $onLoadForm_StateCorrection= 
                {
                    $FormLoaded = $false
                    $rootCN=[adsi]''
                    $nodeName=$baseDN.Split(",")[0].Replace("OU=", "")
                    $key="LDAP://$($baseDN)"
            
                    $BaseNode = $System_Windows_Forms_TreeNode_1.Nodes.Add($key,$nodeName)
                    $BaseNode.ImageIndex = 2
                    $BaseNode.SelectedImageIndex = 2
            
                    $thisOU = [adsi]$BaseNode.Name
                    if (-not $BaseNode.Nodes) 
                    {
                        $textbox1.Text = $thisOU.DistinguishedName
                        $Script:Result = $textbox1.Text
                        $searcher = [adsisearcher]'objectClass=organizationalunit'
                        $searcher.SearchRoot = $BaseNode.Name
                        $searcher.SearchScope = 'OneLevel'
                        $searcher.PropertiesToLoad.Add('name')
                        $OUs = $searcher.Findall()
                
                        foreach ($ou in $OUs)
                        {
                            $key = $ou.Path
                            $nodeName = $ou.Properties['name']
                            $newNode = new-object System.Windows.Forms.TreeNode
                            $newNode.Name = $key
                            $newNode.Text = $nodeName
                            $BaseNode.Nodes.Add($newNode)
                            $newNode.ImageIndex = 1
                            $newNode.SelectedImageIndex = 1
                        }
                    }
                    $System_Windows_Forms_TreeNode_1.Expand()
                    $BaseNode.Expand()
        
                    $form1.Controls.Add($treeView1)
                }
     
                #---------------------------------------------- 
                #region Generated Form Code 
                $form1.Text = "Choose Active Directory OU" 
                $form1.Name = "form1" 
                $form1.DataBindings.DefaultDataSourceUpdateMode = 0 
                $form1.ClientSize = New-object System.Drawing.Size(400,500) 
        
                IF($MultiSelect) #NEW in 1.11
                {    
                    $treeView1.CheckBoxes = $true
                }
                $treeView1.Size = New-object System.Drawing.Size(350,375)
                $treeView1.Margin = New-Object System.Windows.Forms.Padding(2) #NEW in 1.11
                $treeView1.Name = "treeView1" 
                $treeView1.Location = New-object System.Drawing.Size(15,15)
                $treeView1.DataBindings.DefaultDataSourceUpdateMode = 0 
                $treeview1.ImageIndex = 0
                $treeview1.ImageList = $imagelist1
                $treeview1.SelectedImageIndex = 0
                $treeview1.TabIndex = 1

                <#IF(-not $MultiSelect)
                {
                    $treeView1.ShowPlusMinus = $false
                }#>
                $System_Windows_Forms_TreeNode_1 = New-Object 'System.Windows.Forms.TreeNode' ("Active Directory Hierarchy", 3, 3)
                $System_Windows_Forms_TreeNode_1.ImageIndex = 3
                $System_Windows_Forms_TreeNode_1.Name = "Active Directory Hierarchy"
                $System_Windows_Forms_TreeNode_1.SelectedImageIndex = 3
                $System_Windows_Forms_TreeNode_1.Tag = "root"
                $System_Windows_Forms_TreeNode_1.Text = "Active Directory Hierarchy"
                [void]$treeview1.Nodes.Add($System_Windows_Forms_TreeNode_1)

                [array]$Script:PrevNodeHandles = $null
        
                $treeView1.Add_NodeMouseClick({
                    $CurrentNodeHandle = $_.Node.Handle

                    IF(-not $_.Node.IsExpanded)
                    {
                        IF($_.Node.Name -ne "Active Directory Hierarchy")
                        {
                            $thisOU = [adsi]$_.Node.Name
                            $textbox1.Text = $thisOU.DistinguishedName
                            $Script:Result = $textbox1.Text
                            #write-host $($_.Node.Name)
                            IF(-not $_.Node.Nodes) 
                            {
                                $searcher = [adsisearcher]'objectClass=organizationalunit'
                                $searcher.SearchRoot = $_.Node.Name
                                $searcher.SearchScope = 'OneLevel'
                                $searcher.PropertiesToLoad.Add('name')
                                $OUs = $searcher.Findall()
                                IF($OUs.Count -ge 1)
                                {
                                    $_.Node.ImageIndex = 0
                                    $_.Node.SelectedImageIndex = 0
                            
                                }

                                foreach ($ou in $OUs)
                                {
                                    $key = $ou.Path
                                    $nodeName = $ou.Properties['name']
                                    $newNode = new-object System.Windows.Forms.TreeNode
                                    $newNode.Name = $key
                                    $newNode.Text = $nodeName
                                    $_.Node.Nodes.Add($newNode)
                                    $newNode.ImageIndex = 1
                                    $newNode.SelectedImageIndex = 1
                                }

                                $_.Node.Expand()
                            }
                        }

                        IF($CurrentNodeHandle -notin $Script:PrevNodeHandles)
                        {
                            $_.Node.Expand()
                        }
                    } 
                    ELSE
                    {
                        IF($CurrentNodeHandle -notin $Script:PrevNodeHandles)
                        {
                            $_.Node.Collapse()
                        }
                        $thisOU = [adsi]$_.Node.Name
                        $textbox1.Text = $thisOU.DistinguishedName
                    }
                    $Script:PrevNodeHandles += $CurrentNodeHandle
                }) #END $treeView1.Add_NodeMouseClick
         
                $label1.Name = "label1" 
                $label1.Location = New-object System.Drawing.Size(15,400)
                $label1.Size = New-object System.Drawing.Size(350,20)
                $label1.Text = "Selected Value:"
                $form1.Controls.Add($label1) 

                $textbox1.Name = "textbox1" 
                $textbox1.Location = New-object System.Drawing.Size(15,425)
                $textbox1.Size = New-object System.Drawing.Size(350,20)
                $textbox1.Text = ""
                $form1.Controls.Add($textbox1) 
                #endregion Generated Form Code 
 
                #Save the initial state of the form 
                $InitialFormWindowState = $form1.WindowState 
                #Init the onLoad event to correct the initial state of the form 
                $form1.add_Load($onLoadForm_StateCorrection) 
                $FormLoaded = $true
                #Show the Form 
                $Results = $form1.ShowDialog()

            } #END BEGIN

            PROCESS
            {
                IF($Results -eq [System.Windows.Forms.DialogResult]::OK) #USER SELECTED "OK" Button on the Form
                {
                    $form1.Dispose()
                }
                ELSEIF ($Results -eq [System.Windows.Forms.DialogResult]::Cancel)  #USER SELECTED "Cancel" Button on the Form
                {
                    $Script:Result = "CANCELED"
                }

                IF($MultiSelect)
                {
                    [array]$Script:NodesNames = $null
                    Get-CheckedNodesDNs -Nodes $treeView1.Nodes
                    RETURN $Script:NodesNames;
                }
                ELSE
                {
                    RETURN $Script:Result;
                }

            } #END PROCESS

            END
            {
                $form1.Dispose()
            } #END END
 
        } #END FUNCTION Choose-ADOUForm #Version 1.15
        
        FUNCTION PSO_WindowsCapabilities {
	        $obj = New-Object PSObject
	        $obj | Add-Member -MemberType NoteProperty -Name "AD_ComputerName" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_OperatingSystem" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "FOD_CAPABILITY_Name" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty  -Name "FOD_CAPABILITY_State" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_CanonicalName" -Value $null #""
            return $obj
        } #END PSO_WindowsCapabilities FUNCTION

        FUNCTION PSO_WindowsCapabilitiesSummary {
	        $obj = New-Object PSObject
	        $obj | Add-Member -MemberType NoteProperty -Name "AD_ComputerName" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_OperatingSystem" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_Created" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_Enabled" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_LastLogonTimeStamp" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_IPv4Address" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_CanonicalName" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "AD_Description" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "PS1_VerifiedComputer" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "FOD_CAPABILITIES_Installed_COUNT" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "FOD_CAPABILITIES_Staged_COUNT" -Value $null #""
            $obj | Add-Member -MemberType NoteProperty -Name "FOD_CAPABILITIES_NotPresent_COUNT" -Value $null #""
            return $obj
        } #END PSO_WindowsCapabilitiesSummary FUNCTION

        FUNCTION Export-CSVToXLS #Version 1.1 
        {
            Param( 
                [String]$CsvFileLocation
                ,[String]$ExcelFilePath
            )
            If (Test-Path $ExcelFilePath )
            {
                Remove-Item -Path $ExcelFilePath
            }

            $xlFileFormatTypes = Add-Type -AssemblyName 'Microsoft.Office.Interop.Excel' -Passthru
            #$xlFileFormat = $xlFileFormatTypes | Where {$_.Name -like "XlFileFormat"}
            #[Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
            
            $FixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault #$xlFileFormat #'xlOpenXMLWorkbook'
            $Excel = New-Object -ComObject excel.application 

            #Exit script if Excel not found (check last execution status variable)
            If (-Not $?) {

                #Write a custom error 
                Write-Error "Excel not installed. Script execution stopped."
                Exit 2

            }   #End of If (-Not $?)...

            $Excel.visible = $false 
            $Excel.Workbooks.OpenText($CsvFileLocation)
            $Excel.ActiveWorkbook.SaveAs($ExcelFilePath,$FixedFormat) 
            $Excel.Quit() 
            Remove-Variable -Name Excel 
            [gc]::collect() 
            [gc]::WaitForPendingFinalizers()
        } #END FUNCTION Export-CSVToXLS #Version 1.1

        FUNCTION Verify-ComputerName { #version 1.45
            #Verifies that you are connecting to the host you think you are; helps to identify DNS issues.
            [cmdletbinding()]

            Param (
                    [Parameter(Mandatory=$false,ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,
                    HelpMessage="Enter one or more computer names, separated by commas, or passed in through the pipeline.")]
                    [array]$ComputerNames = $env:COMPUTERNAME,

                    # Timeout in milliseconds
                    [Parameter(Mandatory=$false,HelpMessage="Enter the timeout in milliseconds before a fail occurs.")]
                    [ValidateRange(100,50000)]
                    [int]$TimeOut=2000,

                    # Number of attempts
                    [Parameter(Mandatory=$false,HelpMessage="Enter the number of ping attempts before giving up.")]
                    [ValidateRange(1,100)]
                    [int]$Attempts=1,

                    [Parameter(Mandatory=$false,
                    HelpMessage="Select this option to only validate ping.")]
                    [switch]$PingOnly,

                    [Parameter(Mandatory=$false,
                    HelpMessage="Select this option to get verbose output, as the function runs.")]
                    [switch]$INTERACTIVE
                    )  

            BEGIN
            { 
                FUNCTION ConvertFrom-NSlookup {
                [CmdletBinding()]
                Param(
                    [Parameter(Mandatory,Position=0)]
                    [array]$NSlookupReply
                    )
                    $Hash = [ordered]@{}
                    [int]$AddressCount = 0
                    foreach($line in $NSlookupReply){
                        if ($line -notmatch "======================================================================="){
                            switch -Regex ($line){
                                '\s*Server\s*:\s+(?<Server>.*$)'{$Hash.Server = $Matches.Server;Break}
                                '\s*Address\s*:\s+(?<Address>.*$)'{  
                                                                    IF(($Hash.Server -ne $null) -and ($AddressCount -eq 0))
                                                                    {
                                                                        $Hash.SVR_Address = $Matches.Address; $AddressCount++; Break
                                                                    }
                                                                    ELSEIF(($Hash.Server -ne $null) -and ($AddressCount -eq 1))
                                                                    {
                                                                        $Hash.HOST_Address = $Matches.Address; Break
                                                                    }
                                                                    }
                                '\s*Name\s*:\s+(?<Name>.*$)' {$Hash.HostName = $Matches.Name;Break}
                                default {Break}
                            }
                        }
                    }
                    RETURN $Hash; 
                } #END FUNCTION ConvertFrom-NSlookup

                FUNCTION CreatePSObjectForVerifyComputerName {
                    $obj = New-Object PSObject
                    $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -value $null
                    $obj | Add-Member -MemberType NoteProperty -Name "PINGTEST" -value $null
                    $obj | Add-Member -MemberType NoteProperty -Name "NSLOOKUP" -value $null
                    $obj | Add-Member -MemberType NoteProperty -Name "RemoteHostNameViaPSEXEC" -value $null
                    $obj | Add-Member -MemberType NoteProperty -Name "RemoteHostNameViaInvkCmd" -value $null
                    $obj | Add-Member -MemberType NoteProperty -Name "ComputerNameVerified" -value $null
                    return $obj
                }

                [float]$MinVersion = 2.34
                $PsExec64LocalPath = "$env:Windir\system32\PsExec64.exe"
                $PsExec64RemoteShare = "\\blisw6syaaa7nec\IMO_Info\Available Software & Fixes\PS_Tools\PsExec64.exe"
                                      

                IF (!(Test-Path $PsExec64LocalPath))
                {
                    try {
                            [float]$ProductVersionOnShare = (Get-ChildItem -LiteralPath $PsExec64RemoteShare | Select-Object -Property VersionInfo).VersionInfo.ProductVersion
                            IF($ProductVersionOnShare -lt $MinVersion)
                            {
                                throw "The version of PSExec64.exe on the file share `"$PsExec64RemoteShare`" is not greater than or equal to the minimum required version of `"$MinVersion`"`
                                Please update your file share with the latest copy from Microsoft."
                            }

                            COPY $PsExec64RemoteShare -Destination $PsExec64LocalPath -Force
                            IF($INTERACTIVE)
                            {
                                Write-Host "Successfully copied `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor DarkGreen -ForegroundColor White
                            }
                        }
                    catch [System.Exception]
                        {
                            IF($INTERACTIVE)
                            {
                                Write-Host "Unable to copy `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor Red -ForegroundColor White
                            }
                            Write-Warning -Message $_.Exception.Message
                            break
                        }
                }
                ELSE
                {
                    [float]$ProductVersion = (Get-ChildItem -LiteralPath $PsExec64LocalPath | Select-Object -Property VersionInfo).VersionInfo.ProductVersion
                    IF($ProductVersion -lt $MinVersion)
                    {
                        try {
                                [float]$ProductVersionOnShare = (Get-ChildItem -LiteralPath $PsExec64RemoteShare | Select-Object -Property VersionInfo).VersionInfo.ProductVersion
                                IF($ProductVersionOnShare -lt $MinVersion)
                                {
                                    throw "The version of PSExec64.exe on the file share `"$PsExec64RemoteShare`" is not greater than or equal to the minimum required version of `"$MinVersion`"`
                                    Please update your file share with the latest copy from Microsoft."
                                }

                                COPY $PsExec64RemoteShare -Destination $PsExec64LocalPath -Force
                                IF($INTERACTIVE)
                                {
                                    Write-Host "Successfully copied `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor DarkGreen -ForegroundColor White
                                }
                            }
                        catch [System.Exception]
                            {
                                IF($INTERACTIVE)
                                {
                                    Write-Host "Unable to copy `"$PsExec64RemoteShare`" to `"$PsExec64LocalPath`"" -BackgroundColor Red -ForegroundColor White
                                }
                                Write-Warning -Message $_.Exception.Message
                                break
                            }
                    } #END IF($ProductVersion -lt $MinVersion)
                } #END ELSE


            } #END BEGIN

            PROCESS
            {
                [ARRAY]$VeryifyComputerNameResults = $NULL
                FOREACH ($ComputerName in $ComputerNames)
                {        
                    $ping = new-object System.Net.NetworkInformation.Ping
            
                    $objVerifyComputerName = CreatePSObjectForVerifyComputerName
                    $objVerifyComputerName.ComputerName = $ComputerName

                    $Stop = $false
                    [int]$AttemptCount = 0
            
                    while(($Stop -ne $true) -and $AttemptCount -le $Attempts)
                    {
                        TRY {
                                IF($ComputerName -ne $env:COMPUTERNAME)
                                {
                                    $reply = $ping.send($ComputerName,$TimeOut)
                                }
                                ELSE
                                {
                                    $reply = "Success"
                                }
                    
                                IF ($reply -eq "Success" -or $reply.status -eq "Success") 
                                {
                                    $objVerifyComputerName.PINGTEST = "SUCCESS"         
                                    IF ($INTERACTIVE) {
                                                    Write-Host "SUCCESS: PINGTEST: Able to ping an IP that resolves to $ComputerName" -BackgroundColor DarkGreen -ForegroundColor Yellow
                                                    } #Write-Log -Path $LogFile -Message "SUCCESS: PINGTEST: Able to ping $ComputerName" -Level Info
                        
                                    IF(!$PingOnly)
                                    {
                                        #RESET THE FOLLOWING VARIABLES FOR FOLLOWING IF-THEN LOGIC PROCESSES
                                        $NSLookupRemoteHostAddress = $null
                                        $RemoteHostName = $null
                        
                                        $NSLookupRemoteHostAddress = (ConvertFrom-NSlookup (nslookup $ComputerName)).HOST_Address
                                        IF ($NSLookupRemoteHostAddress -ne $null) #If NSLookup returned an IP, execute the "hostname" command using psexec64.exe, against the IP that was resolved in DNS, in order to grab the RemoteHostName, as seen from the remote computer's perspective.
                                        {   
                                            IF ($INTERACTIVE) {
                                                            $Message = "SUCCESS: NSLOOKUP: Resolved $ComputerName in DNS to: $NSLookupRemoteHostAddress"
                                                            Write-Host $Message -BackgroundColor DarkGreen -ForegroundColor Yellow
                                                            } #Write-Log -Path $LogFile -Message $Message -Level Info
                                            $objVerifyComputerName.NSLOOKUP = $NSLookupRemoteHostAddress
                                                                        
                                            $PSExecScriptBlock = {param ($NSLookupRemoteHostAddress)
                                                                    $RemoteHostName = (& psexec64.exe \\$NSLookupRemoteHostAddress -s -nobanner -accepteula powershell -executionpolicy bypass -command "hostname" 2>$null)
                                                                    IF(($RemoteHostName.gettype()).BaseType.Name -eq "Array"){$RemoteHostName = $RemoteHostName[0]}
                                                                    return $RemoteHostName
                                                                }
                                            $PSExecJob = Start-Job -ScriptBlock $PSExecScriptBlock -ArgumentList $NSLookupRemoteHostAddress
                                            $PSExecJob | Wait-Job -Timeout 30 -ErrorAction SilentlyContinue | Out-Null
                                            IF($PSExecJob.state -eq "Running")
                                            {
                                                WRITE-HOST "The job timer ran out for psexec64.exe, forcibly stopping and removing the job!" -ForegroundColor yellow -BackgroundColor Red
                                                $PSExecJob | Stop-Job
                                                $PSExecJob | Remove-Job
                                            }
                                            ELSE
                                            {
                                                $RemoteHostName = $PSExecJob | Receive-Job

                                                IF($RemoteHostName.Contains("Connecting to "))
                                                {
                                                    $RemoteHostName = $RemoteHostName.Substring(0,$RemoteHostName.IndexOf("Connecting to "))
                                                }

                                                $PSExecJob | Remove-Job
                                            }
                                    
                                            IF ($RemoteHostName -ne $null) #Compare $ComputerHostName against $RemoteHostName
                                            {
                                
                                                IF ($INTERACTIVE) {
                                                                $Message = "SUCCESS: RemoteHostName: $RemoteHostName was obtainined on the remote IP $NSLookupRemoteHostAddress using PSEXEC64"
                                                                Write-Host $Message -BackgroundColor DarkGreen -ForegroundColor Yellow
                                                                } #Write-Log -Path $LogFile -Message $Message -Level Info
                                                $objVerifyComputerName.RemoteHostNameViaPSEXEC = $RemoteHostName
                                            }
                                            ELSE
                                            {
                                                IF ($INTERACTIVE) {
                                                                $Message = "FAILURE: RemoteHostName: Unable to obtain the hostname on the remote IP $NSLookupRemoteHostAddress using PSEXEC64, attempting Invoke-Command"
                                                                Write-Host $Message -BackgroundColor Red -ForegroundColor Yellow
                                                                } #Write-Log -Path $LogFile -Message $Message -Level Warn
                                                $objVerifyComputerName.RemoteHostNameViaPSEXEC = "FAILED"

                                                #$Command = {hostname}
                                                #$RemoteHostName = Invoke-Command -ComputerName $ComputerName -ScriptBlock $Command
                                                $JobScriptBlock = {param ($ComputerName)
                                                                        $CommandScriptBlock = {$result = C:\Windows\System32\hostname.exe; return $result}
                                                                        $RemoteHostName = Invoke-Command -ComputerName $ComputerName -ScriptBlock $CommandScriptBlock
                                                                        return $RemoteHostName
                                                                    }
                                                $Job = Start-Job -ScriptBlock $JobScriptBlock -ArgumentList $ComputerName
                                                $Job | Wait-Job -Timeout 30 -ErrorAction SilentlyContinue
                                                IF($Job.state -eq "Running")
                                                {
                                                    WRITE-HOST "The job timer ran out for Invoke-Command, forcibly stopping and removing the job!" -ForegroundColor yellow -BackgroundColor Red
                                                    $Job | Stop-Job
                                                    $Job | Remove-Job
                                                }
                                                ELSE
                                                {
                                                    $RemoteHostName = $Job | Receive-Job 
                                                    $Job | Remove-Job
                                                }

                                                IF ($RemoteHostName -ne $null) #Attempt to obtain RemoteHostName on NSLookupRemotehostAddress using Invoke-Command
                                                {
                                                    IF ($INTERACTIVE) {
                                                                    $Message = "SUCCESS: RemoteHostName: $RemoteHostName was obtainined on the remote IP $NSLookupRemoteHostAddress using Invoke-Command"
                                                                    Write-Host $Message -BackgroundColor DarkGreen -ForegroundColor Yellow
                                                                    } #Write-Log -Path $LogFile -Message $Message -Level Info
                                                    $objVerifyComputerName.RemoteHostNameViaInvkCmd = $RemoteHostName
                                                }
                                                ELSE
                                                {
                                                    IF ($INTERACTIVE) {
                                                                    $Message = "FAILURE: RemoteHostName: Unable to obtain the hostname on the remote IP $NSLookupRemoteHostAddress using Invoke-Command"
                                                                    Write-Host $Message -BackgroundColor Red -ForegroundColor Yellow
                                                                    } #Write-Log -Path $LogFile -Message $Message -Level Warn
                                                    $objVerifyComputerName.RemoteHostNameViaInvkCmd = "FAILED"
                                                }
                                            }
                                        }
                                        ELSE
                                        {
                                            IF ($INTERACTIVE) {
                                                            $Message = "FAILURE: NSLOOKUP: Unable to resolve $ComputerName in DNS, SKIPPING: RemoteHostName Resolution"
                                                            Write-Host $Message -BackgroundColor Red -ForegroundColor Yellow
                                                            } #Write-Log -Path $LogFile -Message $Message -Level Warn
                                            $objVerifyComputerName.NSLOOKUP = "FAILED"
                                        }

                                        IF ($RemoteHostName -eq $ComputerName)
                                        {
                                            IF ($INTERACTIVE) { 
                                                            $Message = "SUCCESS: MATCH: IP in DNS Remotely Resolves to: $RemoteHostName thereby matching: $ComputerName as expected!"
                                                            Write-Host $Message -BackgroundColor DarkGreen -ForegroundColor White
                                                            } #Write-Log -Path $LogFile -Message $Message -Level Info
                                            $objVerifyComputerName.ComputerNameVerified = $True
                                        }
                                        ELSE
                                        {
                                            IF ($INTERACTIVE) { 
                                                            $Message = "FAILURE: MATCH: IP in DNS Remotely Resolves to: $RemoteHostName thereby NOT matching: $ComputerName as expected!"
                                                            Write-Host $Message -BackgroundColor Red -ForegroundColor Yellow
                                                            } #Write-Log -Path $LogFile -Message $Message -Level Warn
                                            $objVerifyComputerName.ComputerNameVerified = $False
                                        }
                                    } #END IF(!$PingOnly)

                                    $Stop = $true
                                }
                                elseif ($reply.status -eq "TimedOut")
                                {
                                    IF ($INTERACTIVE) {
                                                    Write-Host "FAILURE: PINGTEST: TimedOut: Unable to ping $ComputerName" -BackgroundColor Yellow -ForegroundColor Red
                                                    $Message = "TROUBLESHOOTING: Ensure $ComputerName is online; check machine based firewalls (i.e. HBSS AV/HIPS) and/or LAN/CAN/WAN Firewalls"
                                                    Write-Host $Message -BackgroundColor Yellow -ForegroundColor Red
                                                    } #Write-Log -Path $LogFile -Message "FAILURE: PINGTEST: TimedOut: Unable to ping $ComputerName" -Level Warn
                                                    #Write-Log -Path $LogFile -Message $Message -Level Warn
                                    $objVerifyComputerName.PINGTEST = "TimedOut"
                                }
                                elseif ($reply.status -eq "DestinationHostUnreachable")
                                {
                                    IF ($INTERACTIVE) {
                                                    Write-Host "FAILURE: PINGTEST: DestinationHostUnreachable: Unable to ping $ComputerName" -BackgroundColor Yellow -ForegroundColor Red
                                                    $Message = "TROUBLESHOOTING: Ensure $ComputerName is online; check machine based firewalls (i.e. HBSS AV/HIPS) and/or LAN/CAN/WAN Firewalls"
                                                    Write-Host $Message -BackgroundColor Yellow -ForegroundColor Red
                                                    } #Write-Log -Path $LogFile -Message "FAILURE: PINGTEST: DestinationHostUnreachable: Unable to ping $ComputerName" -Level Warn
                                                    #Write-Log -Path $LogFile -Message $Message -Level Warn
                                    $objVerifyComputerName.PINGTEST = "DestinationHostUnreachable"
                                }
                            } #END TRY
                        Catch { 
                                IF ($INTERACTIVE) {
                                                Write-Host "FAILURE: PINGTEST: Unable to ping $ComputerName" -BackgroundColor Yellow -ForegroundColor Red
                                                $Message = "TROUBLESHOOTING: Ensure $ComputerName is online; check machine based firewalls (i.e. HBSS AV/HIPS) and/or LAN/CAN/WAN Firewalls"
                                                Write-Host $Message -BackgroundColor Yellow -ForegroundColor Red
                                                } #Write-Log -Path $LogFile -Message "FAILURE: PINGTEST: Unable to ping $ComputerName" -Level Warn
                                                #Write-Log -Path $LogFile -Message $Message -Level Warn
                                $objVerifyComputerName.PINGTEST = "FAILURE"
                                $objVerifyComputerName.ComputerNameVerified = $False
                                #$reply = [ordered]@{}
                                #$reply.status = "Failure"
                                } #END CATCH
                        $AttemptCount++
                    }

                    $VeryifyComputerNameResults += $objVerifyComputerName
                } #END FOREACH ($ComputerName in $ComputerNames)
                RETURN $VeryifyComputerNameResults
            } #END PROCESS

            END
            {

            } #END END

        } #END FUNCTION Verify-ComputerName V1.45

        $GetWindowsCapabilityScriptBlock =
        {
            PARAM ($OnlyIfInstalled, $OnlyIfStaged, $OnlyIfNotPresent, $CapabilityNameFilter)
            [string] $ScriptBlockString = "Get-WindowsCapability -Online "
            IF(($OnlyIfInstalled) -and ($CapabilityNameFilter.Length -gt 0))
            {
                $ScriptBlockString += "| Where-Object {(`$_.State -eq `"Installed`") -and (`$_.Name -match (`$CapabilityNameFilter))}"
            }
            ELSEIF(($OnlyIfStaged) -and ($CapabilityNameFilter.Length -gt 0))
            {
                $ScriptBlockString += "| Where-Object {(`$_.State -eq `"Staged`") -and (`$_.Name -match (`$CapabilityNameFilter))}"
            }
            ELSEIF(($OnlyIfNotPresent) -and ($CapabilityNameFilter.Length -gt 0))
            {
                $ScriptBlockString += "| Where-Object {(`$_.State -eq `"NotPresent`") -and (`$_.Name -match (`$CapabilityNameFilter))}"
            }
            ELSEIF($OnlyIfInstalled)
            {
                $ScriptBlockString += "| Where-Object {`$_.State -eq `"Installed`"}"
            }
            ELSEIF($OnlyIfStaged)
            {
                $ScriptBlockString += "| Where-Object {`$_.State -eq `"Staged`"}"
            }
            ELSEIF($OnlyIfNotPresent)
            {
                $ScriptBlockString += "| Where-Object {`$_.State -eq `"NotPresent`"}"
            }
            ELSEIF($CapabilityNameFilter.Length -gt 0)
            {
                $ScriptBlockString += "| Where-Object {(`$_.Name -match (`$CapabilityNameFilter))}"
            }

                            
            $ScriptBlock = [Scriptblock]::Create($ScriptBlockString)
            [ARRAY]$FOD_CAPABILITIES = @(Invoke-Command -ScriptBlock $ScriptBlock)

            RETURN $FOD_CAPABILITIES
        }



        ####################################### CREATE ComputerInfoAttributes LOG DIRECTORY, IF NEEDED ##################################################################################
        IF(($RunningResults2csv) -AND (!(Test-Path $ResultsFilePath -PathType Container )))
        {
            Write-Host "WARNING: RESULTS CONTAINER `"$ResultsFilePath`" DOES NOT EXIST; ATTEMPTING TO CREATE..." -BackgroundColor Yellow -ForegroundColor Red
            TRY { 
                    md $ResultsFilePath -Force -ErrorAction Stop
                    $ResultsFolderExists = $true
                    Write-Host "SUCCESS: RESULTS CONTAINER `"$ResultsFilePath`" WAS SUCCESSFULLY CREATED!" -BackgroundColor Blue -ForegroundColor White
                }
            CATCH 
                {
                        $ResultsFolderExists = $false
                        Write-Host "FAILURE: UNABLE TO CREATE RESULTS CONTAINER, ABORTING ACTION(s)!" -BackgroundColor RED -ForegroundColor YELLOW
                }
        }ELSEIF($RunningResults2csv){Write-Host "INFO: RESULTS CONTAINER `"$ResultsFilePath`" ALREADY EXISTS; CONTINUING SCRIPT..." -BackgroundColor Blue -ForegroundColor White; $ResultsFolderExists = $true}
        ####################################### END: CREATE ComputerInfoAttributes LOG DIRECTORY, IF NEEDED ##################################################################################
        

        ####################################### CREATE THE CSV Summary Output File, IF NEEDED ##################################################################################
        IF(($RunningResults2csv) -AND ($ResultsFolderExists))
        {
            $SumOutFile = $ResultsFilePath + "WindowsCapabilitiesSummary_RunningResults_" + (Get-date -Format 'MM-dd-yyyy_HHmmss') + '.CSV' #Set the name of the CSV Summary Output File
            $objCSV = PSO_WindowsCapabilitiesSummary
            TRY {

                    $objCSV | Export-Csv -Path $SumOutFile -NoTypeInformation
                    #Clears the null values for the 2 row
                    Set-Content -Path $SumOutFile -Value (get-content -Path $SumOutFile | Select-String -Pattern ',,,,,,,,,,' -NotMatch)
                    $SumResultsLogFileExists = $true
                    Write-Host "SUCCESS: RESULTS LOG FILE `"$SumOutFile`" WAS SUCCESSFULLY CREATED!" -BackgroundColor Blue -ForegroundColor White

                }
            CATCH
                {
                    $SumResultsLogFileExists = $false
                    Write-Host "FAILURE: UNABLE TO CREATE SUMMARY RESULTS LOG FILE, ABORTING LOGGING ACTION(s)!" -BackgroundColor RED -ForegroundColor YELLOW
                } 
        }
        ####################################### END: CREATE THE CSV Summary Output File #######################################################################################


        ####################################### GET YOUR LOCAL DOMAIN NFO #######################################################################################
        $DomainName = $env:USERDOMAIN
        $Root = [ADSI]"LDAP://RootDSE"
        Try { $DomainRoot = ($Root.Get("rootDomainNamingContext")).ToUpper() }
        Catch { Write-Warning "Unable to contact Active Directory because $($Error[0]); aborting script!"; Start-Sleep -Seconds 5; Exit }
        $DomainDN = "DC=$DomainName,$DomainRoot"
        ####################################### END: GET YOUR LOCAL DOMAIN NFO ##################################################################################


        ####################################### Set $InstallationsDN path, based on domain ###########################################################################
        $InstallationsDN = "OU=Installations," + $DomainDN
        ####################################### END: Set $InstallationsDN path, based on domain ######################################################################
        
        
        ####################################### GET YOUR LOCAL SITE NAME ############################################################################################
        $sAMAccountName = "$env:ComputerName`$"
        $searcher = [adsisearcher]"(&(objectClass=Computer)(sAMAccountName=$sAMAccountName))"
        $searcher.SearchRoot = "LDAP://$($InstallationsDN)"
        $searcher.SearchScope = 'Subtree'
        $searcher.PropertiesToLoad.Add('CanonicalName') | Out-Null
        $Site = ($searcher.Findall()).properties.canonicalname.Split('/')[2].ToUpper() 
        $SitePrefix = $site.Substring(0,4)
        ####################################### END: GET YOUR LOCAL SITE NAME #######################################################################################


        ####################################### SET YOUR LOCAL baseDN PATH ############################################################################################
        $baseDN = "OU=$Site,$InstallationsDN"
        ####################################### END: SET YOUR LOCAL baseDN PATH ############################################################################################


        ####################################### Set $DomainControllersDN path, based on domain ###########################################################################
        $DomainControllersDN = "OU=Domain Controllers," + $DomainDN
        ####################################### END: Set $DomainControllersDN path, based on domain ######################################################################

        
        ####################################### GET YOUR LOCAL DC NAME(s) ############################################################################################
        $searcher = [adsisearcher]"(&(objectClass=Computer)(CN=$SitePrefix*))"
        $searcher.SearchRoot = "LDAP://$($DomainControllersDN)"
        $searcher.SearchScope = 'Subtree'
        $searcher.PropertiesToLoad.Add('CanonicalName') | Out-Null
        #($Searcher.FindAll()).properties.canonicalname.Split('/')
        $DCsADSPath = (($Searcher.FindAll()).Properties.adspath)
        [array]$DCs=$null
        FOREACH($ADSPath in $DCsADSPath)
        {
            $DCs += $ADSPath.Remove(0,10).Split(",")[0]
        }
        ####################################### END: GET YOUR LOCAL DC NAME(s) #######################################################################################

        
        ####################################### GET YOUR LOCAL DOMAIN CONTROLER THAT LOGGED YOU ON ##################################################################################
        #   $logonServer is used as -Server to the Get-ADComputer below, not sure that this is necessary(?)
        $logonServer = nltest /dsgetdc: /force
        $logonServer = $logonServer.item(0)
        $logonServer = $logonServer.TrimStart("DC: \\")
        ####################################### END: GET YOUR LOCAL DOMAIN CONTROLER THAT LOGGED YOU ON #############################################################################
        

        ####################################### USING A GUI BASED FORM, GET THE targetOU WHERE YOU WANT TO START YOUR SEARCH IN AD ##################################################################################
        IF($LocateTargetOU)
        {
            Write-Progress -Activity "Waiting on a user selected Sub-OU from AD to search. You may need to ALT-TAB to bring the form into focus." -PercentComplete (-1) 
            $targetOU = Choose-ADOUForm
        }
        ####################################### END: GET THE targetOU WHERE YOU WANT TO START YOUR SEARCH IN AD #####################################################################################################


    } #END BEGIN
    
    PROCESS 
    {   
        [array]$ComputerProperties = "CN", "OperatingSystem", "Created", "Enabled", "LastLogonTimeStamp", "IPv4Address", "CanonicalName", "Description"

        ######################################### PROVIDED YOUR targetOU COLLECTED FROM THE GUI LOGIC ISN'T NULL, PROCESS THE REPORT #########################################
        IF(($targetOU -ne "CANCELED") -AND ($targetOU.Length -gt 0))
        {
            Write-Progress -Activity "Getting target AD Computer objects based on user selected Sub-OU. This process may take a while, depending on the size of the targeted Sub-OU structure and the -SearchScope specified." -PercentComplete (-1)

            IF($ComputerNames.count -gt 0)
            {
                ### Initiate a counter for the write-progress logic below ###
                $ComputerNamesCount = 1

                [array]$ADComputerObjects = $null

                $DCsIndex = 0
                $DCsIndexCount = $DCs.Count
                FOREACH($ComputerName in $ComputerNames)
                {
                    IF($DCsIndex -gt 3) # This logic round robins the DCs with queries
                    {
                        $DCsIndex = 0
                    }
                    $DC = $DCs[$DCsIndex]
                    ### Show progress status...
                    Write-Progress -Id 0 -Activity "Getting all AD Computer Objects in $targetOU" -Status "$ComputerNamesCount of $($ComputerNames.Count)" -PercentComplete (($ComputerNamesCount / $ComputerNames.Count) * 100)
                    
                    $sAMAccountName = $NULL
                    IF($ComputerName.Contains('$'))
                    {
                        $sAMAccountName = $ComputerName
                    }
                    ELSE
                    {
                        $sAMAccountName = $ComputerName + "$"
                    }
                    
                    $ADComputerObjects += Get-ADComputer -Filter 'sAMAccountName -eq $sAMAccountName' -Properties $ComputerProperties -Server $DC -SearchBase $targetOU -SearchScope $SearchScope
                    $ComputerNamesCount++
                    $DCIndex++
                } #END FOREACH($ComputerName in $ComputerNames)
            }
            ELSEIF($ComputerListCSV.Length -gt 1)
            {
                #Write-Host "ComputerListCSVLengthGT1: $ComputerListCSV" -ForegroundColor White -BackgroundColor DarkCyan
                IF(Test-Path $ComputerListCSV.Replace("`"", "") -PathType Leaf)
                {
                    [array]$ComputersInCSV = IMPORT-CSV -LiteralPath $ComputerListCSV.Replace("`"", "")

                    IF(!((($ComputersInCSV | Where-Object {$_.ComputerName.Length -gt 0}).Count -gt 0) -or (($ComputersInCSV | Where-Object {$_.sAMAccountName.Length -gt 0}).Count -gt 0) -or (($ComputersInCSV | Where-Object {$_.AD_ComputerName.Length -gt 0}).Count -gt 0)))
                    {
                        Write-Host "No ComputerName(s), sAMAccountName(s), or AD_ComputerName(s) found in file: $ComputerListCSV, please specify another CSV file or ensure that either a ComputerName, sAMAccountName, or AD_ComputerName column exists with values in it!" -BackgroundColor red -ForegroundColor Yellow
                    }
                    ELSE
                    {
                        IF(($ComputersInCSV | Where-Object {$_.ComputerName.Length -gt 0}).Count -gt 0) 
                        {
                            [array]$ComputerNames = $ComputersInCSV.ComputerName
                        }
                        ELSEIF(($ComputersInCSV | Where-Object {$_.sAMAccountName.Length -gt 0}).Count -gt 0)
                        {
                            [array]$ComputerNames = $ComputersInCSV.sAMAccountName
                        }
                        ELSEIF(($ComputersInCSV | Where-Object {$_.AD_ComputerName.Length -gt 0}).Count -gt 0)
                        {
                            [array]$ComputerNames = $ComputersInCSV.AD_ComputerName
                        }
                        Write-Host "ComputersInCSV Count:$($ComputerNames.count)" -BackgroundColor Blue -ForegroundColor White
                    }

                    ### Initiate a counter for the write-progress logic below ###
                    $ComputerNamesCount = 1

                    [array]$ADComputerObjects = $null
                    $DCsIndex = 0
                    $DCsIndexCount = $DCs.Count

                    FOREACH($ComputerName in $ComputerNames)
                    {
                        IF($DCsIndex -gt 3)
                        {
                            $DCsIndex = 0
                        }
                        $DC = $DCs[$DCsIndex]

                            ### Show progress status...
                        Write-Progress -Id 0 -Activity "Getting all AD Computer Objects in $targetOU" -Status "$ComputerNamesCount of $($ComputerNames.Count)" -PercentComplete (($ComputerNamesCount / $ComputerNames.Count) * 100)
                    
                        $sAMAccountName = $NULL
                        IF($ComputerName.Contains('$'))
                        {
                            $sAMAccountName = $ComputerName
                        }
                        ELSE
                        {
                            $sAMAccountName = $ComputerName + "$"
                        }
                        $ADComputerObjects += Get-ADComputer -Filter 'sAMAccountName -eq $sAMAccountName' -Properties $ComputerProperties -Server $DC -SearchBase $targetOU -SearchScope $SearchScope

                        $ComputerNamesCount++
                        $DCIndex++
                    } #END FOREACH($ComputerName in $ComputerNames)
                } #END IF(Test-Path $ComputerListCSV -PathType Leaf)
            }
            ELSEIF($SearchAD)
            {
                IF(($targetOU.Length -gt 1) -and ($SearchScope.Length -gt 1))
                {
                    Write-Progress -Activity "Getting target AD Computer objects based on user selected Sub-OU. This process may take a while, depending on the size of the targeted Sub-OU structure." -PercentComplete (-1)     
                    [array]$PossibleADComputerObjects = $NULL

                    IF($SearchScope -eq "Subtree")
                    {
                        ### Collect all the Sub-OUs, starting at a the search base of $targetOU (Sub-OU processing was chosen, to elimante AD Query Time-Outs) ###
                        $ADOUs = Get-ADOrganizationalUnit -Filter * -Server $logonServer -SearchBase $targetOU -SearchScope $SearchScope | Select-Object -ExpandProperty DistinguishedName

                        ### Initiate a counter for the write-progress logic below ###
                        $ADOUsCount = 1
                        $DCsIndex = 0
                        $DCsIndexCount = $DCs.Count
                        foreach ($ADOU in $ADOUs) ### Stepping through all Sub-OUs of the $targetOU ###
                        {
                            IF($DCsIndex -gt 3)
                            {
                                $DCsIndex = 0
                            }
                            $DC = $DCs[$DCsIndex]
                            ### Show progress status...
                            Write-Progress -Id 0 -Activity "Getting all AD Computer Objects in $ADOU" -Status "$ADOUsCount of $($ADOUs.Count)" -PercentComplete (($ADOUsCount / $ADOUs.Count) * 100)
                            $PossibleADComputerObjects += Get-ADComputer -Filter * -Properties $ComputerProperties -Server $DC -SearchBase $ADOU -SearchScope OneLevel
                            $ADOUsCount++
                            $DCIndex++
                        }
                        $PossibleADComputerObjects = $PossibleADComputerObjects | Sort-Object -Property CN
                    }
                    ELSE
                    {
                        $PossibleADComputerObjects = Get-ADComputer -Filter * -Properties $ComputerProperties -Server $logonServer -SearchBase $targetOU -SearchScope $SearchScope | Sort-Object -Property CN
                    }
                }
                ELSE
                {
                    Write-Host "ERROR: You selected the -SearchAD parameter switch, but you failed to select -targetOU or -LocateTargetOU and specify a starting OU in AD where your search should begin!" -BackgroundColor Red -ForegroundColor Yellow
                }

                
                IF(($PossibleADComputerObjects.COUNT -ge 1) -and ($PossibleADComputerObjects.COUNT -le 500))
                {   
                    Write-Progress -Activity "Populating Windows Computers Selection List Form...please wait..." -PercentComplete (-1)                
                    $FormTitle = "Select Computers for Windows Capabilities Discovery"
                    $ListLabel="Please select one or more Computers from the list to include in the Windows Capabilities Discovery, and then OK to continue, or Cancel to abort: "
                    [array]$ItemsListViewColumnNames = "CN", "Description", "DistinguishedName"
                    [array]$ADComputerObjects = $null

                    #Two-Dimensional Array is in $PossibleADComputerObjects, use -ItemsList parameter in Select-ItemsFromListForm FUNCTION
                    $ADComputerObjects = Select-ItemsFromListViewForm -FormTitle $FormTitle -ListLabel $ListLabel -IncludeOkBtn -IncludeCancelBtn `
                    -IncludeSelectAllBtn -IncludeDeSelectAllBtn -ItemsListViewColumnNames $ItemsListViewColumnNames `
                    -ItemsListCollection $PossibleADComputerObjects -ReturnObjects
                }
                ELSEIF($PossibleADComputerObjects.COUNT -gt 500) 
                {   
                    Write-Host "More than 500 computers were detected, therefore individual selection is not possible, thus all COMPUTERS will have their Windows Capabilities collected during the Discovery." -ForegroundColor white -BackgroundColor Blue
                    [array]$ADComputerObjects = $null
                    $ADComputerObjects = $PossibleADComputerObjects
                }
                ELSE
                {
                    Write-Host "WARNING: There were no COMPUTERS found in the TargetOU Sected: `"$targetOU`", based on the -targetOU and -SearchScope you specified; therefore no COMPUTERS had their Windows Capabilities Collected during the Discovery!" -ForegroundColor Red -BackgroundColor Yellow
                    $Description = $NULL
                }
                
            } #END ELSEIF($SearchAD)
            ELSE
            {
                Write-Host "ERROR: No computers were specified via the supported parameters. Re-run this script passing in computers via one of three parameter options (i.e. -ComputerNames, -ComputerListCSV, or -SearchAD)!" -BackgroundColor red -ForegroundColor Yellow
            }
            
            IF($ADComputerObjects.Count -ge 1)
            {
                [array]$ADComputerObjectsResults = $null
                [array]$ComputerFODCapabilitiesResults = $null
                $ADComputerObjectsCount = 1
                $TZbias = (Get-WmiObject -Query 'Select Bias from Win32_TimeZone').bias
                FOREACH ($ADComputerObject in $ADComputerObjects)
                {
                    $ComputerName = $null; $ComputerName = $ADComputerObject.CN
                    
                    ### Show progress status...
                    Write-Progress -Id 1 -Activity "Getting all selected Computer Objects' Windows Capabilities" -Status "$ADComputerObjectsCount of $($ADComputerObjects.Count)" -PercentComplete (($ADComputerObjectsCount / $ADComputerObjects.Count) * 100)

                    $objADComputer = PSO_WindowsCapabilitiesSummary
                    $objADComputer.AD_ComputerName = $ADComputerObject.CN
                    $objADComputer.AD_OperatingSystem = $ADComputerObject.OperatingSystem
                    $objADComputer.AD_Created = $ADComputerObject.Created
                    $objADComputer.AD_Enabled = $ADComputerObject.Enabled
                    $objADComputer.AD_LastLogonTimeStamp = ([datetime]::FromFileTimeUtc($ADComputerObject.LastLogonTimeStamp)).AddMinutes($TZbias)
                    $objADComputer.AD_IPv4Address = $ADComputerObject.IPv4Address
                    $objADComputer.AD_CanonicalName = $ADComputerObject.CanonicalName
                    $objADComputer.AD_Description = $ADComputerObject.Description
                    
                    IF($ComputerName -ne $env:COMPUTERNAME)
                    {
                        $ComputerNameVerified = Verify-ComputerName -ComputerNames $ComputerName
                    }
                    IF(($ComputerName -eq $ENV:COMPUTERNAME) -OR ($ComputerNameVerified.ComputerNameVerified))
                    {
                        $objADComputer.PS1_VerifiedComputer = $true

                        ####################################### CREATE THE CSV Output File (Per Computer), IF NEEDED ##################################################################################
                        IF(($RunningResults2csv) -AND ($ResultsFolderExists))
                        {
                            $outFile = $ResultsFilePath + "$($ComputerName)_WindowsCapabilities_" + (Get-date -Format 'MM-dd-yyyy_HHmmss') + '.CSV' #Set the name of the CSV Summary Output File
                            $objCSV = PSO_WindowsCapabilities
                            TRY {

                                    $objCSV | Export-Csv -Path $outFile -NoTypeInformation
                                    #Clears the null values for the 2 row
                                    Set-Content -Path $outFile -Value (get-content -Path $outFile | Select-String -Pattern ',,,,' -NotMatch)
                                    $ResultsLogFileExists = $true
                                    Write-Host "SUCCESS: RESULTS LOG FILE `"$outFile`" WAS SUCCESSFULLY CREATED!" -BackgroundColor Blue -ForegroundColor White

                                }
                            CATCH
                                {
                                    $ResultsLogFileExists = $false
                                    Write-Host "FAILURE: UNABLE TO CREATE RESULTS LOG FILE, ABORTING LOGGING ACTION(s)!" -BackgroundColor RED -ForegroundColor YELLOW
                                } 
                        }
                        ####################################### END: CREATE THE CSV Output File (Per Computer) #######################################################################################
                        
                        $FOD_CAPABILITIES = $null
                        
                        IF($ComputerName -eq $ENV:COMPUTERNAME)
                        {
                            [ARRAY]$FOD_CAPABILITIES = Invoke-Command -ScriptBlock $GetWindowsCapabilityScriptBlock -ArgumentList $OnlyIfInstalled, $OnlyIfStaged, $OnlyIfNotPresent, $CapabilityNameFilter
                        }
                        ELSE
                        {
                            [ARRAY]$FOD_CAPABILITIES = Invoke-Command -ComputerName $ComputerName -ScriptBlock $GetWindowsCapabilityScriptBlock -ArgumentList $OnlyIfInstalled, $OnlyIfStaged, $OnlyIfNotPresent, $CapabilityNameFilter
                        }
                        $objADComputer.FOD_CAPABILITIES_Installed_COUNT = $($FOD_CAPABILITIES | WHERE-OBJECT {$_.State -eq "Installed"}).Count
                        $objADComputer.FOD_CAPABILITIES_Staged_COUNT = $($FOD_CAPABILITIES | WHERE-OBJECT {$_.State -eq "Staged"}).Count
                        $objADComputer.FOD_CAPABILITIES_NotPresent_COUNT = $($FOD_CAPABILITIES | WHERE-OBJECT {$_.State -eq "NotPresent"}).Count

                        [array]$ComputerFODCapabilities = $null
                        FOREACH ($FOD_CAPABILITY in $FOD_CAPABILITIES)
                        {
                            $objComputerFODCapability = PSO_WindowsCapabilities
                            $objComputerFODCapability.AD_ComputerName = $ComputerName
                            $objComputerFODCapability.AD_OperatingSystem = $ADComputerObject.OperatingSystem
                            $objComputerFODCapability.FOD_CAPABILITY_Name = $FOD_CAPABILITY.Name
                            $objComputerFODCapability.FOD_CAPABILITY_State = $FOD_CAPABILITY.State 
                            $objComputerFODCapability.AD_CanonicalName = $ADComputerObject.CanonicalName

                            #GENERATE THE ComputerInfoAttributes REPORT, IF -RunningResults2csv PARAMETER IS USED.
                            IF ($RunningResults2csv)
                            {
                                $objComputerFODCapability | Export-CSV -LiteralPath $outFile -Append
                            }

                            $ComputerFODCapabilities += $objComputerFODCapability
                        }

                        IF(($RunningResults2csv) -and (Test-Path $outFile -PathType Leaf) -and ($ComputerFODCapabilities.Count -gt 0))
                        {
                            $xlsFilePath = $outFile.Replace(".CSV", ".XLSX")

                            Export-CSVToXLS -CsvFileLocation $outFile -ExcelFilePath $xlsFilePath
                            Start-Sleep -Seconds 3
                        
                            IF(Test-Path $xlsFilePath -PathType Leaf)
                            {
                                Add-Type -AssemblyName Microsoft.Office.Interop.Excel
                                $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
        
                                #Instantiate an Excel COM object
                                $Excel = New-Object -ComObject Excel.Application
                                #Exit script if Excel not found (check last execution status variable)
                                If (-Not $?) {

                                    #Write a custom error 
                                    Write-Error "Excel not installed. Script execution stopped."
                                    Exit 2

                                }   #End of If (-Not $?)...
                                $Excel.Visible = $false
                                $Excel.Workbooks.Open($xlsFilePath) 2>&1>$NULLL
                                $Excel.ActiveWorkbook.ActiveSheet.Cells[1, 1].EntireRow.Font.Bold = $true;
                                $Excel.ActiveWorkbook.ActiveSheet.Cells[1, 1].EntireRow.AutoFilter() 2>&1>$NULL
                                $Excel.ActiveWorkbook.ActiveSheet.UsedRange.EntireColumn.autofit() 2>&1>$NULLL
                                $Excel.Rows.Item("2:2").Select() #Change V1_05
                                $Excel.ActiveWindow.FreezePanes = $true #Change V1_05
                                $Excel.DisplayAlerts = $false 
                                $Excel.ActiveWorkbook.saveas($xlsFilePath, $xlFixedFormat) 2>&1>$NULL
                                $Excel.Quit()
                                Remove-Variable -Name Excel  2>&1>$NULL
                                Remove-Item -Path $outFile 2>&1>$NULL
                            } #END IF(Test-Path $xlsFilePath -PathType Leaf)
                        } #END IF(($RunningResults2csv) -and (Test-Path $outFile -PathType Leaf) -and ($ADComputerObjectsResults.Count -gt 0))

                        $ComputerFODCapabilitiesResults += @($ComputerFODCapabilities)

                    }
                    ELSE
                    {
                        $objADComputer.PS1_VerifiedComputer = $false
                        $objADComputer.FOD_CAPABILITIES_Installed_COUNT = "Not Evaluated"
                        $objADComputer.FOD_CAPABILITIES_NotPresent_COUNT = "Not Evaluated"
                    }

                    
                    #GENERATE THE ComputerInfoAttributes REPORT, IF -RunningResults2csv PARAMETER IS USED.
                    IF ($RunningResults2csv)
                    {
                        $objADComputer | Export-CSV -LiteralPath $SumOutFile -Append
                    }

                    $ADComputerObjectsResults += $objADComputer

                    $ADComputerObjectsCount++

                } #END FOREACH ($ADComputerObject in $ADComputerObjects)
            } #END IF($ADComputerObjects.Count -ge 1)

            IF (!$RunningResults2csv)
            {
                RETURN $ADComputerObjectsResults, $ComputerFODCapabilitiesResults
            }
        } #END IF(($targetOU -ne "CANCELED") -AND ($targetOU.Length -gt 0))
        ELSEIF($targetOU -eq "CANCELED")
        {
            Write-Host "WARNING: You `"CANCELED`" the operation, therefore no COMPUTERS had their Windows Capabilities Collected for the report!" -ForegroundColor Red -BackgroundColor Yellow
            RETURN "CANCELED"
        }
        ELSEIF($ComputerNames[0] -eq $env:ComputerName)
        {
            [array]$ADComputerObjectsResults = $null
            [array]$ComputerFODCapabilitiesResults = $null
            $ComputerName = $ComputerNames[0]
            
            Write-Progress -Id 1 -Activity "Getting Windows Capabilities from LocalHost:$ComputerName"

            $objADComputer = PSO_WindowsCapabilitiesSummary
            $objADComputer.AD_ComputerName = $ComputerName
            $objADComputer.AD_OperatingSystem = "NotEnumerated"
            $objADComputer.AD_Created = "NotEnumerated"
            $objADComputer.AD_Enabled = "NotEnumerated"
            $objADComputer.AD_LastLogonTimeStamp = "NotEnumerated"
            $objADComputer.AD_IPv4Address = "NotEnumerated"
            $objADComputer.AD_CanonicalName = "NotEnumerated"
            $objADComputer.AD_Description = "NotEnumerated"
            $objADComputer.PS1_VerifiedComputer = $true

            ####################################### CREATE THE CSV Output File (Per Computer), IF NEEDED ##################################################################################
            IF(($RunningResults2csv) -AND ($ResultsFolderExists))
            {
                $outFile = $ResultsFilePath + "$($ComputerName)_WindowsCapabilities_" + (Get-date -Format 'MM-dd-yyyy_HHmmss') + '.CSV' #Set the name of the CSV Summary Output File
                $objCSV = PSO_WindowsCapabilities
                TRY {

                        $objCSV | Export-Csv -Path $outFile -NoTypeInformation
                        #Clears the null values for the 2 row
                        Set-Content -Path $outFile -Value (get-content -Path $outFile | Select-String -Pattern ',,,,' -NotMatch)
                        $ResultsLogFileExists = $true
                        Write-Host "SUCCESS: RESULTS LOG FILE `"$outFile`" WAS SUCCESSFULLY CREATED!" -BackgroundColor Blue -ForegroundColor White

                    }
                CATCH
                    {
                        $ResultsLogFileExists = $false
                        Write-Host "FAILURE: UNABLE TO CREATE RESULTS LOG FILE, ABORTING LOGGING ACTION(s)!" -BackgroundColor RED -ForegroundColor YELLOW
                    } 
            }
            ####################################### END: CREATE THE CSV Output File (Per Computer) #######################################################################################
            
            $FOD_CAPABILITIES = $null
            [ARRAY]$FOD_CAPABILITIES = Invoke-Command -ScriptBlock $GetWindowsCapabilityScriptBlock -ArgumentList $OnlyIfInstalled, $OnlyIfStaged, $OnlyIfNotPresent, $CapabilityNameFilter
            $objADComputer.FOD_CAPABILITIES_Installed_COUNT = $($FOD_CAPABILITIES | WHERE-OBJECT {$_.State -eq "Installed"}).Count
            $objADComputer.FOD_CAPABILITIES_Staged_COUNT = $($FOD_CAPABILITIES | WHERE-OBJECT {$_.State -eq "Staged"}).Count
            $objADComputer.FOD_CAPABILITIES_NotPresent_COUNT = $($FOD_CAPABILITIES | WHERE-OBJECT {$_.State -eq "NotPresent"}).Count

            [array]$ComputerFODCapabilities = $null
            FOREACH ($FOD_CAPABILITY in $FOD_CAPABILITIES)
            {
                $objComputerFODCapability = PSO_WindowsCapabilities
                $objComputerFODCapability.AD_ComputerName = $ComputerName
                $objComputerFODCapability.AD_OperatingSystem = "NotEnumerated"
                $objComputerFODCapability.FOD_CAPABILITY_Name = $FOD_CAPABILITY.Name
                $objComputerFODCapability.FOD_CAPABILITY_State = $FOD_CAPABILITY.State 
                $objComputerFODCapability.AD_CanonicalName = "NotEnumerated"

                #GENERATE THE ComputerInfoAttributes REPORT, IF -RunningResults2csv PARAMETER IS USED.
                IF ($RunningResults2csv)
                {
                    $objComputerFODCapability | Export-CSV -LiteralPath $outFile -Append
                }

                $ComputerFODCapabilities += $objComputerFODCapability
            }

            IF(($RunningResults2csv) -and (Test-Path $outFile -PathType Leaf) -and ($ComputerFODCapabilities.Count -gt 0))
            {
                $xlsFilePath = $outFile.Replace(".CSV", ".XLSX")

                Export-CSVToXLS -CsvFileLocation $outFile -ExcelFilePath $xlsFilePath
                Start-Sleep -Seconds 3
                        
                IF(Test-Path $xlsFilePath -PathType Leaf)
                {
                    Add-Type -AssemblyName Microsoft.Office.Interop.Excel
                    $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
        
                    #Instantiate an Excel COM object
                    $Excel = New-Object -ComObject Excel.Application
                    #Exit script if Excel not found (check last execution status variable)
                    If (-Not $?) {

                        #Write a custom error 
                        Write-Error "Excel not installed. Script execution stopped."
                        Exit 2

                    }   #End of If (-Not $?)...
                    $Excel.Visible = $false
                    $Excel.Workbooks.Open($xlsFilePath) 2>&1>$NULLL
                    $Excel.ActiveWorkbook.ActiveSheet.Cells[1, 1].EntireRow.Font.Bold = $true;
                    $Excel.ActiveWorkbook.ActiveSheet.Cells[1, 1].EntireRow.AutoFilter() 2>&1>$NULL
                    $Excel.ActiveWorkbook.ActiveSheet.UsedRange.EntireColumn.autofit() 2>&1>$NULLL
                    $Excel.Rows.Item("2:2").Select() #Change V1_05
                    $Excel.ActiveWindow.FreezePanes = $true #Change V1_05
                    $Excel.DisplayAlerts = $false 
                    $Excel.ActiveWorkbook.saveas($xlsFilePath, $xlFixedFormat) 2>&1>$NULL
                    $Excel.Quit()
                    Remove-Variable -Name Excel  2>&1>$NULL
                    Remove-Item -Path $outFile 2>&1>$NULL
                } #END IF(Test-Path $xlsFilePath -PathType Leaf)
            } #END IF(($RunningResults2csv) -and (Test-Path $outFile -PathType Leaf) -and ($ADComputerObjectsResults.Count -gt 0))

            $ComputerFODCapabilitiesResults += @($ComputerFODCapabilities)

            #GENERATE THE ComputerInfoAttributes REPORT, IF -RunningResults2csv PARAMETER IS USED.
            IF ($RunningResults2csv)
            {
                $objADComputer | Export-CSV -LiteralPath $SumOutFile -Append
            }

            $ADComputerObjectsResults += $objADComputer
            
            IF (!$RunningResults2csv)
            {
                RETURN $ADComputerObjectsResults, $ComputerFODCapabilitiesResults
            }
        }
        
    } #END PROCESS

    END
    {
        IF(($RunningResults2csv) -and (Test-Path $SumOutFile -PathType Leaf) -and ($ADComputerObjectsResults.Count -gt 0))
        {
            $xlsFilePath = $SumOutFile.Replace(".CSV", ".XLSX")

            Export-CSVToXLS -CsvFileLocation $SumOutFile -ExcelFilePath $xlsFilePath
            Start-Sleep -Seconds 3
                        
            IF(Test-Path $xlsFilePath -PathType Leaf)
            {
                Add-Type -AssemblyName Microsoft.Office.Interop.Excel
                $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
        
                #Instantiate an Excel COM object
                $Excel = New-Object -ComObject Excel.Application
                #Exit script if Excel not found (check last execution status variable)
                If (-Not $?) {

                    #Write a custom error 
                    Write-Error "Excel not installed. Script execution stopped."
                    Exit 2

                }   #End of If (-Not $?)...
                $Excel.Visible = $true
                $Excel.Workbooks.Open($xlsFilePath) 2>&1>$NULLL
                $Excel.ActiveWorkbook.ActiveSheet.Cells[1, 1].EntireRow.Font.Bold = $true;
                $Excel.ActiveWorkbook.ActiveSheet.Cells[1, 1].EntireRow.AutoFilter() 2>&1>$NULL
                $Excel.ActiveWorkbook.ActiveSheet.UsedRange.EntireColumn.autofit() 2>&1>$NULLL
                $Excel.Rows.Item("2:2").Select() #Change V1_05
                $Excel.ActiveWindow.FreezePanes = $true #Change V1_05
                $Excel.DisplayAlerts = $false 
                $Excel.ActiveWorkbook.saveas($xlsFilePath, $xlFixedFormat) 2>&1>$NULL
                Remove-Variable -Name Excel  2>&1>$NULL
                Remove-Item -Path $SumOutFile 2>&1>$NULL
            } #END IF(Test-Path $xlsFilePath -PathType Leaf)
        } #END IF(($RunningResults2csv) -and (Test-Path $SumOutFile -PathType Leaf) -and ($ADComputerObjectsResults.Count -gt 0))

    } #END END
} #END FUNCTION Get-WindowsCapabilities #VERSION 1.063

#THE CODE BELOW USES A WINDOWS FORM TO DYNAMICALLY BUILD THE PARAMETERS BEING PASSED TO THE Get-WindowsCapabilities FUNCTION
FUNCTION Get-WindowsCapabilitiesParametersForm #Version 1.06
{
    [CmdletBinding()]
    param(
        
        [Parameter(Mandatory=$false)] 
        [array]$ComputerSelectionMethod,

        [Parameter(Mandatory=$false)] 
        [array]$TargetOUSelectionMethod,

        [Parameter(Mandatory=$false)] 
        [array]$SearchScopeMethod
    )

    BEGIN
    {
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
        [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing") 

        $objForm = New-Object System.Windows.Forms.Form 

        $objForm.Text = "Get Windows Capabilities Options" 
        $objForm.Size = '800,800'
        $objForm.StartPosition = "CenterScreen"

        $objForm.KeyPreview = $True
        
        #region   OK Button
        $OKButton = New-Object System.Windows.Forms.Button
        $OKButton.Location = '300,700'
        $OKButton.Size = '75,23'
        $OKButton.Text = "OK"

        # Got rid of the Click event for OK Button, and instead just assigned its DialogResult property to OK.
        $OKButton.DialogResult = [System.Windows.Forms.DialogResult]::OK

        $objForm.Controls.Add($OKButton)

        # Setting the form's AcceptButton property causes it to automatically intercept the Enter keystroke and
        # treat it as clicking the OK button (without having to write your own KeyDown events).
        $objForm.AcceptButton = $OKButton
        #endregion 
        
        #region Cancel Button
        $CancelButton = New-Object System.Windows.Forms.Button
        $CancelButton.Location = '390,700'
        $CancelButton.Size = '75,23'
        $CancelButton.Text = "Cancel"

        # Got rid of the Click event for Cancel Button, and instead just assigned its DialogResult property to Cancel.
        $CancelButton.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

        $objForm.Controls.Add($CancelButton)

        # Setting the form's CancelButton property causes it to automatically intercept the Escape keystroke and
        # treat it as clicking the OK button (without having to write your own KeyDown events).
        $objForm.CancelButton = $CancelButton
        #endregion 

        # Add QuickQuery Button
        $QuickQueryButton = New-Object System.Windows.Forms.Button
        $QuickQueryButton.Location = New-Object System.Drawing.Size(480,700)
        $QuickQueryButton.Size = New-Object System.Drawing.Size(200,23)
        $QuickQueryButton.Text = "QUICK QUERY: Local Computer"
        
        #Add QuickQuery Button event 
        $QuickQueryButtonClickedScriptBlock = { $Script:QuickQueryButtonClicked = $True; $Script:CapabilityNameFilter = $objTextBoxCapabilityNameFilter.Text;
                                $objForm.Dispose() }
        $QuickQueryButton.Add_Click($QuickQueryButtonClickedScriptBlock) #>

        $objForm.Controls.Add($QuickQueryButton) 
                
        $objLabelComputerSelectionMethod = New-Object System.Windows.Forms.Label
        $objLabelComputerSelectionMethod.Location = '10,10'
        $objLabelComputerSelectionMethod.Size = '200,15'
        $objLabelComputerSelectionMethod.Text = "Computer Selection Method:"
        $objForm.Controls.Add($objLabelComputerSelectionMethod)

        $objLabelComputerNames = New-Object System.Windows.Forms.Label
        $objLabelComputerNames.Location = '10,80'
        $objLabelComputerNames.Size = '400,15'
        $objLabelComputerNames.Text = "Manually Enter ComputerNames (separated by commas `",`"):"
        
        $objTextBoxComputerNames = New-Object System.Windows.Forms.TextBox
        $objTextBoxComputerNames.Location = '10,100'
        $objTextBoxComputerNames.Size = '760,60'
        $objTextBoxComputerNames.Multiline = $true

        $objLabelComputerNamesCSV = New-Object System.Windows.Forms.Label
        $objLabelComputerNamesCSV.Location = '10,80'
        $objLabelComputerNamesCSV.Size = '760,15'
        $objLabelComputerNamesCSV.Text = "Enter the full file path and file name to the CSV file containing the ComputerNames for evaluation: (i.e. C:\SomePath\Computers2Check.csv)"
        
        $objTextBoxComputerNamesCSV = New-Object System.Windows.Forms.TextBox
        $objTextBoxComputerNamesCSV.Location = '10,100'
        $objTextBoxComputerNamesCSV.Size = '760,20'
        
        IF($ComputerSelectionMethod.Count -gt 0)
        {
            $listBoxComputerSelectionMethod = New-Object System.Windows.Forms.ListBox 
            $listBoxComputerSelectionMethod.Location = '10,30'
            $listBoxComputerSelectionMethod.Size = '275,10'
            $listBoxComputerSelectionMethod.Height = 50
            $listBoxComputerSelectionMethod.MultiColumn = 0
            $listBoxComputerSelectionMethod.SelectionMode = 1
            #Populate the list box with items to choose from
            foreach ($Item in $ComputerSelectionMethod ) 
            {
                [void]$listBoxComputerSelectionMethod.Items.Add($Item)
            }

            [void]$objForm.Controls.Add($listBoxComputerSelectionMethod)
            
            $SelectedIndexChanged = {

                IF($listBoxComputerSelectionMethod.SelectedItem -eq "Search AD By OU")
                {
                    $objForm.Controls.Remove($objLabelComputerNames)
                    $objForm.Controls.Remove($objTextBoxComputerNames)
                    $objForm.Controls.Remove($objLabelComputerNamesCSV)
                    $objForm.Controls.Remove($objTextBoxComputerNamesCSV)
                    $objForm.Controls.Remove($objLabelTargetOUSelectionMethod) #Change V1_03
                    $objForm.Controls.Remove($listBoxTargetOUSelectionMethod) #Change V1_03
                    $objForm.Controls.Remove($objLabelTargetOUDirectSelectionMethod) #Change V1_03
                    $objForm.Controls.Remove($objTextBoxTargetOUDirectSelectionMethod) #Change V1_03
                }
                ELSEIF($listBoxComputerSelectionMethod.SelectedItem -eq "Import ComputerNames From CSV File")
                {
                    $objForm.Controls.Remove($objLabelComputerNames)
                    $objForm.Controls.Remove($objTextBoxComputerNames)
                    $objForm.Controls.Add($objLabelComputerNamesCSV)
                    $objForm.Controls.Add($objTextBoxComputerNamesCSV)
                    $objForm.Controls.Add($objLabelTargetOUSelectionMethod) #Change V1_03
                    $objForm.Controls.Add($listBoxTargetOUSelectionMethod) #Change V1_03
                }
                ELSEIF($listBoxComputerSelectionMethod.SelectedItem -eq "Manually Enter ComputerNames")
                {
                    $objForm.Controls.Remove($objLabelComputerNamesCSV)
                    $objForm.Controls.Remove($objTextBoxComputerNamesCSV)
                    $objForm.Controls.Add($objLabelComputerNames)
                    $objForm.Controls.Add($objTextBoxComputerNames)
                    $objForm.Controls.Add($objLabelTargetOUSelectionMethod) #Change V1_03
                    $objForm.Controls.Add($listBoxTargetOUSelectionMethod) #Change V1_03
                }
            }
            $listBoxComputerSelectionMethod.Add_SelectedIndexChanged($SelectedIndexChanged)
        } #END IF($ComputerSelectionMethod.Count -gt 0)
    
        $objLabelTargetOUSelectionMethod = New-Object System.Windows.Forms.Label
        $objLabelTargetOUSelectionMethod.Location = '10,190'
        $objLabelTargetOUSelectionMethod.Size = '200,15'
        $objLabelTargetOUSelectionMethod.Text = "Target OU Selection Method:"
        #$objForm.Controls.Add($objLabelTargetOUSelectionMethod) #Change V1_03

        $objLabelTargetOUDirectSelectionMethod = New-Object System.Windows.Forms.Label
        $objLabelTargetOUDirectSelectionMethod.Location = '10,260'
        $objLabelTargetOUDirectSelectionMethod.Size = '750,15'
        $objLabelTargetOUDirectSelectionMethod.Text = "Specify Target OU: (i.e. OU=TLOU,OU=BLISS,OU=INSTALLATIONS,DC=NASW,DC=DS,DC=ARMY,DC=MIL)"
        
        $objTextBoxTargetOUDirectSelectionMethod = New-Object System.Windows.Forms.TextBox
        $objTextBoxTargetOUDirectSelectionMethod.Location = '10,280'
        $objTextBoxTargetOUDirectSelectionMethod.Size = '750,30'

        IF($TargetOUSelectionMethod.Count -gt 0)
        {
            $listBoxTargetOUSelectionMethod = New-Object System.Windows.Forms.ListBox 
            $listBoxTargetOUSelectionMethod.Location = '10,210'
            $listBoxTargetOUSelectionMethod.Size = '275,10'
            $listBoxTargetOUSelectionMethod.Height = 50
            $listBoxTargetOUSelectionMethod.MultiColumn = 0
            $listBoxTargetOUSelectionMethod.SelectionMode = 1
            #Populate the list box with items to choose from
            foreach ($Item in $TargetOUSelectionMethod ) 
            {
                [void]$listBoxTargetOUSelectionMethod.Items.Add($Item)
            }

            #$objForm.Controls.Add($listBoxTargetOUSelectionMethod) #Change V1_03

            $TargetOUSelectionMethod_SelectedIndexChanged = {
            
                IF($listBoxTargetOUSelectionMethod.SelectedItem -eq "Use Installation TLOU")  #Change V1_03
                {
                    $objForm.Controls.Remove($objLabelTargetOUDirectSelectionMethod)  #Change V1_03
                    $objForm.Controls.Remove($objTextBoxTargetOUDirectSelectionMethod)  #Change V1_03
                }
                ELSEIF($listBoxTargetOUSelectionMethod.SelectedItem -eq "Locate Target OU in AD via GUI")
                {
                    $objForm.Controls.Remove($objLabelTargetOUDirectSelectionMethod)
                    $objForm.Controls.Remove($objTextBoxTargetOUDirectSelectionMethod)
                }
                ELSEIF($listBoxTargetOUSelectionMethod.SelectedItem -eq "Specify Target OU Directly")
                {
                    $objForm.Controls.Add($objLabelTargetOUDirectSelectionMethod)
                    $objForm.Controls.Add($objTextBoxTargetOUDirectSelectionMethod)
                }
            }

            $listBoxTargetOUSelectionMethod.Add_SelectedIndexChanged($TargetOUSelectionMethod_SelectedIndexChanged)
        } #END IF($TargetOUSelectionMethod.Count -gt 0)

        $objLabelSearchScope = New-Object System.Windows.Forms.Label
        $objLabelSearchScope.Location = '10,330'
        $objLabelSearchScope.Size = '200,15'
        $objLabelSearchScope.Text = "SearchScope Method:"
        $objForm.Controls.Add($objLabelSearchScope)

        IF($SearchScopeMethod.Count -gt 0)
        {
            $listBoxSearchScope = New-Object System.Windows.Forms.ListBox 
            $listBoxSearchScope.Location = '10,350'
            $listBoxSearchScope.Size = '275,10'
            $listBoxSearchScope.Height = 50
            $listBoxSearchScope.MultiColumn = 0
            $listBoxSearchScope.SelectionMode = 1
            #Populate the list box with items to choose from
            foreach ($Item in $SearchScopeMethod ) 
            {
                [void]$listBoxSearchScope.Items.Add($Item)
            }
            [void]$objForm.Controls.Add($listBoxSearchScope) 
        }

        ##### BEGIN: CREATE REPORT ####
        $objLabelCheckBoxCreateReport = New-Object System.Windows.Forms.Label
        $objLabelCheckBoxCreateReport.Location = '400,330'
        $objLabelCheckBoxCreateReport.Size = '140,15'
        $objLabelCheckBoxCreateReport.Text = "ONLY Create Report:"
        $objForm.Controls.Add($objLabelCheckBoxCreateReport)

        $objCheckBoxCreateReport = New-Object System.Windows.Forms.CheckBox
        $objCheckBoxCreateReport.Location = '400,350'
        $objCheckBoxCreateReport.Size = '20,20'
        $objCheckBoxCreateReport.Checked = $false
        $objForm.Controls.Add($objCheckBoxCreateReport)

        $objToolTipCreateReport = New-Object System.Windows.Forms.ToolTip
        $objToolTipCreateReport.IsBalloon = $true
        $objToolTipCreateReport.SetToolTip($objLabelCheckBoxCreateReport, "Select this check box to ONLY generate an Excel report of all the Windows Capabilities collected for the selected computers. `nSelecting this option will bypass Install/Remove options of Windows Capabilities.")
        $objToolTipCreateReport.SetToolTip($objCheckBoxCreateReport, "Select this check box to ONLY generate an Excel report of all the Windows Capabilities collected for the selected computers. `nSelecting this option will bypass Install/Remove options of Windows Capabilities.")
        ##### END: CREATE REPORT ####


        ##### BEGIN: Capability Name Filter ####
        $objLabelTextBoxCapabilityNameFilter = New-Object System.Windows.Forms.Label
        $objLabelTextBoxCapabilityNameFilter.Location = '400,390'
        $objLabelTextBoxCapabilityNameFilter.Size = '120,15'
        $objLabelTextBoxCapabilityNameFilter.Text = "Capability Name Filter:"
        $objForm.Controls.Add($objLabelTextBoxCapabilityNameFilter)

        $objTextBoxCapabilityNameFilter = New-Object System.Windows.Forms.TextBox
        $objTextBoxCapabilityNameFilter.Location = '400,410'
        $objTextBoxCapabilityNameFilter.Size = '200,20'
        $objForm.Controls.Add($objTextBoxCapabilityNameFilter)

        $objToolTipCapabilityNameFilter = New-Object System.Windows.Forms.ToolTip
        $objToolTipCapabilityNameFilter.IsBalloon = $true
        $objToolTipCapabilityNameFilter.SetToolTip($objLabelTextBoxCapabilityNameFilter, "Optionally provide a key word, (i.e. `"RSAT`") in order to filter the names of the Windows Capabilities to search for.")
        $objToolTipCapabilityNameFilter.SetToolTip($objTextBoxCapabilityNameFilter, "Optionally provide a key word, (i.e. `"RSAT`") in order to filter the names of the Windows Capabilities to search for.")
        ##### END: Capability Name Filter ####
        
        ##### BEGIN: RadioButtons: ALL - OnlyIfInstalled, Staged, AND NotPresent ####
        $objRadioButtonALL = New-Object System.Windows.Forms.RadioButton
        $objRadioButtonALL.Location = '400,470'
        $objRadioButtonALL.Size = '250,20'
        $objRadioButtonALL.Text = "ALL - `"Installed`", `"Staged`", AND `"NotPresent`""  
        $objRadioButtonALL.Checked = $true

        $objRadioButtonOnlyIfInstalled = New-Object System.Windows.Forms.RadioButton
        $objRadioButtonOnlyIfInstalled.Location = '400,490'
        $objRadioButtonOnlyIfInstalled.Size = '150,20'
        $objRadioButtonOnlyIfInstalled.Text = "Only if `"Installed`""  
        $objRadioButtonOnlyIfInstalled.Checked = $false

        $objRadioButtonOnlyIfStaged = New-Object System.Windows.Forms.RadioButton
        $objRadioButtonOnlyIfStaged.Location = '400,510'
        $objRadioButtonOnlyIfStaged.Size = '150,20'
        $objRadioButtonOnlyIfStaged.Text = "Only if `"Staged`""  
        $objRadioButtonOnlyIfStaged.Checked = $false

        $objRadioButtonOnlyIfNotPresent = New-Object System.Windows.Forms.RadioButton
        $objRadioButtonOnlyIfNotPresent.Location = '400,530'
        $objRadioButtonOnlyIfNotPresent.Size = '150,20'
        $objRadioButtonOnlyIfNotPresent.Text = "Only if `"NotPresent`""  
        $objRadioButtonOnlyIfNotPresent.Checked = $false

        <# NOT NECESSARY, BUT SHOWS A GOOD EXAMPLE OF ADDING A SCRIPT BLOCK TO EXECUTE WHEN A RADIO BUTTON IS CLICKED!
        $RadioButtonClicked = {
            
            IF($objRadioButtonALL.Checked -eq $true)
            {
                $objRadioButtonOnlyIfStaged.Checked = $false
                $objRadioButtonOnlyIfNotPresent.Checked = $false
                $objRadioButtonOnlyIfInstalled.Checked = $false
                write-host "ALL was Clicked!"
            }
            ELSEIF($objRadioButtonOnlyIfInstalled.Checked -eq $true)
            {
                $objRadioButtonALL.Checked = $false
                $objRadioButtonOnlyIfStaged.Checked = $false
                $objRadioButtonOnlyIfNotPresent.Checked = $false
                write-host "OnlyIfInstalled was Clicked!"
            }
            ELSEIF($objRadioButtonOnlyIfStaged.Checked -eq $true)
            {
                $objRadioButtonALL.Checked = $false
                $objRadioButtonOnlyIfInstalled.Checked = $false
                $objRadioButtonOnlyIfNotPresent.Checked = $false
                write-host "OnlyIfStaged was Clicked!"
            }
            ELSEIF($objRadioButtonOnlyIfNotPresent.Checked -eq $true)
            {
                $objRadioButtonALL.Checked = $false
                $objRadioButtonOnlyIfInstalled.Checked = $false
                $objRadioButtonOnlyIfStaged.Checked = $false
                write-host "OnlyIfNotPresent was Clicked!"
            }
        }
        
        $objRadioButtonALL.Add_Click($RadioButtonClicked)
        $objRadioButtonOnlyIfInstalled.Add_Click($RadioButtonClicked)
        $objRadioButtonOnlyIfStaged.Add_Click($RadioButtonClicked)
        $objRadioButtonOnlyIfNotPresent.Add_Click($RadioButtonClicked)
        #>

        
        $objForm.Controls.Add($objRadioButtonALL)
        $objForm.Controls.Add($objRadioButtonOnlyIfInstalled)
        $objForm.Controls.Add($objRadioButtonOnlyIfStaged)
        $objForm.Controls.Add($objRadioButtonOnlyIfNotPresent)

        
        $objToolTipALLInstalledStagedNotPresent = New-Object System.Windows.Forms.ToolTip
        $objToolTipALLInstalledStagedNotPresent.IsBalloon = $true
        $objToolTipALLInstalledStagedNotPresent.SetToolTip($objRadioButtonALL, "Select this radio button to include ALL Capabilities states, (i.e. Installed, Staged, NotPresent)")


        $objToolTipOnlyIfInstalled = New-Object System.Windows.Forms.ToolTip
        $objToolTipOnlyIfInstalled.IsBalloon = $true
        $objToolTipOnlyIfInstalled.SetToolTip($objRadioButtonOnlyIfInstalled, "Select this radio button to include ONLY the Installed Capabilities.")

        $objToolTipOnlyIfStaged = New-Object System.Windows.Forms.ToolTip
        $objToolTipOnlyIfStaged.IsBalloon = $true
        $objToolTipOnlyIfStaged.SetToolTip($objRadioButtonOnlyIfInstalled, "Select this radio button to include ONLY the Staged Capabilities.")

        $objToolTipOnlyIfNotPresent = New-Object System.Windows.Forms.ToolTip
        $objToolTipOnlyIfNotPresent.IsBalloon = $true
        $objToolTipOnlyIfNotPresent.SetToolTip($objRadioButtonOnlyIfNotPresent, "Select this check box to include ONLY the Capabilities which are Not Present.")
        ##### END: RadioButtons: Both, OnlyIfInstalled, NotPresent ####
        
                                        
        $labelMessage = New-Object System.Windows.Forms.Label
        $labelMessage.Location = '10,600'
        $labelMessage.Size = '750,60'
         
        $objForm.Topmost = $True
        $objForm.Add_Shown({$objForm.Activate()})
        
    } #END BEGIN
            
    PROCESS
    {
        do
        {
            $MissingInformation = $false
            $itemComputerSelectionMethodSelected = $null;
            $itemTargetOUSelectionMethodSelected = $null;
            $itemSearchScopeMethodSelected = $null;
                        
            IF ($Message -ne $null)
            {
                $labelMessage.Text = $Message
                $labelMessage.BackColor = "yellow"
                $labelMessage.ForeColor = "red"
                $Message = $null
                $objForm.Controls.Add($labelMessage)
            }

            $Results = $objForm.ShowDialog()

            IF($Script:QuickQueryButtonClicked)
            {
                $itemComputerSelectionMethodSelected = "ComputerNames"
                [array]$ComputerNames = @($env:COMPUTERNAME)
                $itemTargetOUSelectionMethodSelected = "NULL"
                $targetOU = "NULL"
                $itemSearchScopeMethodSelected = "NULL"
                $GenerateReport = $objCheckBoxCreateReport.CheckState
                $CapabilityNameFilter = $Script:CapabilityNameFilter
                $OnlyIfInstalled = $objRadioButtonOnlyIfInstalled.Checked
                $OnlyIfStaged = $objRadioButtonOnlyIfStaged.Checked
                $OnlyIfNotPresent = $objRadioButtonOnlyIfNotPresent.Checked
            }
            ELSEIF($Results -eq [System.Windows.Forms.DialogResult]::OK) #USER SELECTED "OK" Button on the Form
            {   
                $itemComputerSelectionMethodSelected = $listBoxComputerSelectionMethod.SelectedItem;
                SWITCH ($listBoxComputerSelectionMethod.SelectedItem){
                    {$itemComputerSelectionMethodSelected -eq "Search AD By OU"}
                        {
                            $itemComputerSelectionMethodSelected = "SearchAD"
                            $listBoxTargetOUSelectionMethod.SetSelected( 1, $TRUE) #Change V1_03
                        }
                    {$itemComputerSelectionMethodSelected -eq "Import ComputerNames From CSV File"}
                        {
                            IF($objTextBoxComputerNamesCSV.Text.Length -gt 5)
                            {
                                $itemComputerSelectionMethodSelected = "ComputerListCSV"
                                $ComputerNamesCSV = $objTextBoxComputerNamesCSV.Text
                            }
                            ELSE
                            {
                                $MissingInformation = $true
                                $Message += "MESSAGE: You need to specify the path and the name of the CSV file containing the list of computers you wish to import for assessment.`n"
                            }
                        }
                    {$itemComputerSelectionMethodSelected -eq "Manually Enter ComputerNames"}
                        {
                            IF($objTextBoxComputerNames.Text.Length -ge 8)
                            {
                                $itemComputerSelectionMethodSelected = "ComputerNames"
                                [array]$ComputerNames = $objTextBoxComputerNames.Text.Replace(" ", "").Split(",")
                            }
                            ELSE
                            {
                                $MissingInformation = $true
                                $Message += "MESSAGE: You need to specify one or more computer names, separated by commas in the ComputerNames TextBox above.`n"
                            }
                        }
                } #END SWITCH ($listBoxComputerSelectionMethod.SelectedItem)
               
                $itemTargetOUSelectionMethodSelected = $listBoxTargetOUSelectionMethod.SelectedItem;
                SWITCH ($listBoxTargetOUSelectionMethod.SelectedItem){
                    {$itemTargetOUSelectionMethodSelected -eq "Use Installation TLOU"}
                        {
                            $itemTargetOUSelectionMethodSelected = "targetOU"

                            ####################################### GET YOUR LOCAL DOMAIN NFO #######################################################################################
                            $DomainName = $env:USERDOMAIN
                            $Root = [ADSI]"LDAP://RootDSE"
                            Try { $DomainRoot = ($Root.Get("rootDomainNamingContext")).ToUpper() }
                            Catch { Write-Warning "Unable to contact Active Directory because $($Error[0]); aborting script!"; Start-Sleep -Seconds 5; Exit }
                            $DomainDN = "DC=$DomainName,$DomainRoot"
                            ####################################### END: GET YOUR LOCAL DOMAIN NFO ##################################################################################


                            ####################################### Set $InstallationsDN path, based on domain ###########################################################################
                            $InstallationsDN = "OU=Installations," + $DomainDN
                            ####################################### END: Set $InstallationsDN path, based on domain ######################################################################
        
        
                            ####################################### GET YOUR LOCAL SITE NAME ############################################################################################
                            $sAMAccountName = "$env:ComputerName`$"
                            $searcher = [adsisearcher]"(&(objectClass=Computer)(sAMAccountName=$sAMAccountName))"
                            $searcher.SearchRoot = "LDAP://$($InstallationsDN)"
                            $searcher.SearchScope = 'Subtree'
                            $searcher.PropertiesToLoad.Add('CanonicalName') | Out-Null
                            $Site = ($searcher.Findall()).properties.canonicalname.Split('/')[2].ToUpper() 
                            $SitePrefix = $site.Substring(0,4)
                            ####################################### END: GET YOUR LOCAL SITE NAME #######################################################################################
                                                        

                            ####################################### SET YOUR LOCAL baseDN PATH ############################################################################################
                            $targetOU = "OU=$Site,$InstallationsDN"
                            ####################################### END: SET YOUR LOCAL baseDN PATH ############################################################################################
                        }
                    {$itemTargetOUSelectionMethodSelected -eq "Locate Target OU in AD via GUI"}
                        {
                            $itemTargetOUSelectionMethodSelected = "LocateTargetOU"
                        }
                    {$itemTargetOUSelectionMethodSelected -eq "Specify Target OU Directly"}
                        {
                            IF($objTextBoxTargetOUSelectionMethod.Text.Length -gt 46)
                            {
                                $itemTargetOUSelectionMethodSelected = "targetOU"
                                $targetOU = $objTextBoxTargetOUSelectionMethod.Text
                            }
                            ELSE
                            {
                                $MissingInformation = $true
                                $Message += "MESSAGE: You need to specify a valid Distinguished Name path to the TargetOU, in the TargetOU TextBox above.`n"
                            }
                        }
                } #END SWITCH ($listBoxTargetOUSelectionMethod.SelectedItem)
                
                $itemSearchScopeMethodSelected = $listBoxSearchScope.SelectedItem;

                $GenerateReport = $objCheckBoxCreateReport.CheckState 
                
                IF($objTextBoxCapabilityNameFilter.Text.Length -ge 2)
                {
                    $CapabilityNameFilter = $objTextBoxCapabilityNameFilter.Text
                }
                ELSEIF($objTextBoxCapabilityNameFilter.Text.Length -eq 1)
                {
                    $MissingInformation = $true
                    $Message += "MESSAGE: You need to specify at least 2 characters in the Capability Name Filter above OR clear the filter, correct issues and then OK to continue, or select Cancel to abort.`n"
                }
                                
                $OnlyIfInstalled = $objRadioButtonOnlyIfInstalled.Checked
                $OnlyIfStaged = $objRadioButtonOnlyIfStaged.Checked
                $OnlyIfNotPresent = $objRadioButtonOnlyIfNotPresent.Checked
                
                IF ($itemComputerSelectionMethodSelected -eq $null)
                {
                    $Message += "MESSAGE: You failed to select a `"Computer Selection Method`" above, correct issues and then OK to continue, or select Cancel to abort.`n"
                }
                IF ($itemTargetOUSelectionMethodSelected -eq $null)
                {
                    $Message += "MESSAGE: You failed to select a `"TargetOU Selection Method`" above, correct issues and then OK to continue, or select Cancel to abort.`n"
                }
                IF ($itemSearchScopeMethodSelected -eq $null)
                {
                    $Message += "MESSAGE: You failed to select a `"SearchScope Method`" above, correct issues and then OK to continue, or select Cancel to abort.`n"
                }

            } #END IF($Results -eq [System.Windows.Forms.DialogResult]::OK) #USER SELECTED "OK" Button on the Form
            ELSEIF ($Results -eq [System.Windows.Forms.DialogResult]::Cancel)  #USER SELECTED "Cancel" Button on the Form
            {
                write-host 'User canceled...terminating script!' -BackgroundColor Red -ForegroundColor Yellow
                $itemsNotSelected = "CANCELED"
                $objForm.Dispose()
                $objForm.Close()

                RETURN $itemsNotSelected, $null, $null, $null, $null, $null, $null, $null, $null, $null, $null
            }
                    
        }WHILE (($itemComputerSelectionMethodSelected -eq $null) -or ($itemTargetOUSelectionMethodSelected -eq $null) -or ($itemSearchScopeMethodSelected -eq $null) -or ($MissingInformation))
        
        $objForm.Dispose()
        $objForm.Close()

        RETURN $itemComputerSelectionMethodSelected, $ComputerNamesCSV, $ComputerNames, $itemTargetOUSelectionMethodSelected, $targetOU, $itemSearchScopeMethodSelected, $GenerateReport, $CapabilityNameFilter, $OnlyIfInstalled, $OnlyIfStaged, $OnlyIfNotPresent

    } #END PROCESS

    END
    {
        $objForm.Dispose()
        $objForm.Close()
    } #END END
} #end function Get-WindowsCapabilitiesParametersForm #Version 1.06

[array]$ComputerSelectionMethod = "Search AD By OU", "Import ComputerNames From CSV File", "Manually Enter ComputerNames"
[array]$targetOUSelectionMethod = "Use Installation TLOU", "Locate Target OU in AD via GUI", "Specify Target OU Directly"
[array]$SearchScopeMethod = "Subtree", "OneLevel", "Base"

$GetWindowsCapabilitiesParameters = Get-WindowsCapabilitiesParametersForm -ComputerSelectionMethod $ComputerSelectionMethod -TargetOUSelectionMethod $targetOUSelectionMethod -SearchScopeMethod $SearchScopeMethod

$itemComputerSelectionMethodSelected = $GetWindowsCapabilitiesParameters[0]
$ComputerNamesCSV = $GetWindowsCapabilitiesParameters[1]
$ComputerNames = $GetWindowsCapabilitiesParameters[2]
$itemTargetOUSelectionMethodSelected = $GetWindowsCapabilitiesParameters[3]
$targetOU = $GetWindowsCapabilitiesParameters[4]
$itemSearchScopeMethodSelected = $GetWindowsCapabilitiesParameters[5]
$GenerateReport = $GetWindowsCapabilitiesParameters[6]
$CapabilityNameFilter = $GetWindowsCapabilitiesParameters[7]
$OnlyIfInstalled = $GetWindowsCapabilitiesParameters[8]
$OnlyIfNotPresent  = $GetWindowsCapabilitiesParameters[9]

IF(!($itemComputerSelectionMethodSelected -eq "CANCELED"))
{
    IF($itemComputerSelectionMethodSelected -eq "SearchAD")
    {
        [string]$ScriptBlockString = "Get-WindowsCapabilities -SearchAD "
    }
    ELSEIF($itemComputerSelectionMethodSelected -eq "ComputerListCSV")
    {
        [string]$ScriptBlockString = "Get-WindowsCapabilities -ComputerListCSV `$ComputerNamesCSV "
    }
    ELSEIF($itemComputerSelectionMethodSelected -eq "ComputerNames")
    {
        [string]$ScriptBlockString = "Get-WindowsCapabilities -ComputerNames `$ComputerNames "
    }

    IF($itemTargetOUSelectionMethodSelected -eq "LocateTargetOU")
    {
        $ScriptBlockString += "-LocateTargetOU "
    }
    ELSEIF($itemTargetOUSelectionMethodSelected -eq "targetOU")
    {
        $ScriptBlockString += "-targetOU `"$targetOU`" "
    }

    IF($itemSearchScopeMethodSelected -ne "NULL")
    {
        $ScriptBlockString += "-SearchScope $itemSearchScopeMethodSelected "
    }

    IF($GenerateReport -eq "Checked")
    {
        $ScriptBlockString += "-RunningResults2csv "
    }
    
    IF($CapabilityNameFilter.Length -ge 2)
    {
        $ScriptBlockString += "-CapabilityNameFilter `$CapabilityNameFilter "
    }

    IF($OnlyIfInstalled)
    {
        $ScriptBlockString += "-OnlyIfInstalled "
    }

    IF($OnlyIfNotPresent)
    {
        $ScriptBlockString += "-OnlyIfNotPresent "
    }

    
    #Write-Host "ScriptBlockString: $ScriptBlockString"
    
    $ScriptBlock = [Scriptblock]::Create($ScriptBlockString)

    #cls

    $GetWindowsCapabilitiesResults = @(Invoke-Command -ScriptBlock $ScriptBlock)
}
ELSE {$GetWindowsCapabilitiesResults = "CANCELED"}





IF(($GetWindowsCapabilitiesResults[1].COUNT -ge 1) -and ($GetWindowsCapabilitiesResults -ne "CANCELED"))
{
    $WindowsFeatures = $GetWindowsCapabilitiesResults[1]

    $FormTitle = "Select Windows Features to Install or Remove"
    $ListLabel="Please select one or more Windows Features from the list below to INSTALL on, or REMOVE from, the computers, and then OK to continue, or Cancel to abort:`n(If you select an item that is already `"Installed`", it will be REMOVED, and if you select an item that is `"NotPresent`", it will be INSTALLED.)"
    [array]$ItemsListViewColumnNames = "AD_ComputerName", "FOD_CAPABILITY_Name", "FOD_CAPABILITY_State"

    $FeaturesToInstallRemove = @(Select-ItemsFromListViewForm -FormTitle $FormTitle -ListLabel $ListLabel `
    -IncludeOkBtn -IncludeCancelBtn -IncludeSelectAllBtn -IncludeDeSelectAllBtn -FormSizeHoriZ 800 -FormSizeVert 640 `
    -ItemsListViewColumnNames $ItemsListViewColumnNames -ItemsListCollection $WindowsFeatures -ReturnObjects)
    
    #Initialize Common Variabbles
    $WindowsUpdateAUKey = "Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
    $UseWUServerValue = "UseWUServer"
    $Disable = 0
    $Enable = 1
    $FODSource1809 = "\\blisw6syaaa7nec\IMO_Info\FOD\1809"
    $FODSource1903 = "\\blisw6syaaa7nec\IMO_Info\FOD\1903"
    $FODSource1909 = "\\blisw6syaaa7nec\IMO_Info\FOD\1909"
    $FODSource20H2 = "\\blisw6syaaa7nec\IMO_Info\FOD\20H2\"
    $LocalFODSource = "$env:ProgramData\armylocal\FOD"
    

    ############# BEGIN: Script Blocks for Reading and Setting Registry Values ################
    $TestRegKey_ScriptBlock = { param ($KeyPath) Test-Path -Path $KeyPath}
    $CreateRegKey_ScriptBlock = { param ($KeyPath) New-Item -Path $KeyPath -Force | Out-Null}
    $CreateRegValue_ScriptBlock = { param ($KeyPath, $ValueName, $Value, $PropertyType) New-ItemProperty -Path $KeyPath -Name $ValueName -Value $Value -PropertyType $PropertyType -Force | Out-Null}
    $UpdateRegValue_ScriptBlock = { param ($KeyPath, $ValueName, $Value, $PropertyType) New-ItemProperty -Path $KeyPath -Name $ValueName -Value $Value -PropertyType $PropertyType -Force | Out-Null}
    
    $TestRegValue_ScriptBlock = { param ($KeyPath, $ValueName)
        function Test-RegistryValue
        {
            param (
                    [parameter(Mandatory=$true)]
                    [ValidateNotNullOrEmpty()]$KeyPath,
                    [parameter(Mandatory=$true)]
                    [ValidateNotNullOrEmpty()]$ValueName
                    )
            try {
                    Get-ItemProperty -Path $KeyPath | Select-Object -ExpandProperty $ValueName -ErrorAction Stop | Out-Null
                    return $true
                }
            catch {
                    return $false
                }

        } #END Test-RegistryValue FUNCTION
        Test-RegistryValue -KeyPath $KeyPath -ValueName $ValueName
    } #end $TestRegValue_ScriptBlock
        
    #Method in PS6 uses Get-ItemPropertyValue
    #i.e. $AGMRegValueValue = Get-ItemPropertyValue -Path $AGMRegKey -Name $AGMRegValue
    $GetRegValue_ScriptBlock = { param ($KeyPath, $ValueName)
        function Get-RegValue([String] $KeyPath, [String] $ValueName) 
        {
            (Get-ItemProperty -LiteralPath $KeyPath -Name $ValueName).$ValueName
        }
        Get-RegValue -KeyPath $KeyPath -ValueName $ValueName
    } #This function was written for PS4 to allow greater compatibility
    
    $RestartWindowsUpdateServiceScriptBlock = {Restart-Service wuauserv}
        
    $CheckPendingReboot_ScriptBlock = {            
    # Check Pending Reboot FUNCTION for registry
        FUNCTION Check-PendingReboot {
            $CBSRebootKey = Get-ChildItem "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" -ErrorAction Ignore
            $WURebootKey = Get-Item "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" -ErrorAction Ignore
            $FileRenamePending = Get-ItemProperty "HKLM:SYSTEM\CurrentControlSet\Control\Session Manager" -Name PendingFileRenameOperations -ErrorAction SilentlyContinue  
            try 
            { 
                $util = [wmiclass]"\\.\root\ccm\clientsdk:CCM_ClientUtilities"
                $status = $util.DetermineIfRebootPending()
                if(($status -ne $null) -and $status.RebootPending){
                    $SCCMRebootPending = $true
                }
            }catch{}
          
            if (($CBSRebootKey -ne $null) -OR ($WURebootKey -ne $null) -or ($FileRenamePending -ne $null) -or ($SCCMRebootPending)) {
                $true
            }
            else {
                $false
            }
        } #end  FUNCTION Check-PendingReboot
            Check-PendingReboot
    } #END $CheckPendingReboot_ScriptBlock 

    $WindowsBuildSupported_ScriptBlock = 
    {
        # Windows 10 1809 build
        $1809Build = "17763"
        # Windows 10 1903 build
        $1903Build = "18362"
        # Windows 10 1909 build
        $1909Build = "18363"
        # Windows 10 20H2 build
        $20H2Build = "19042"
        # Get running Windows build
        $WindowsBuild =(Get-WmiObject -Class Win32_OperatingSystem).BuildNumber

        IF (($WindowsBuild -eq $1809Build) -OR ($WindowsBuild -eq $1903Build) -OR ($WindowsBuild -eq $1909Build) -OR ($WindowsBuild -eq $20H2Build)) {$true, $WindowsBuild} ELSE {$false, $WindowsBuild}
    }

    $LocalFODSourceFilesExist_ScriptBlock = {
        Param ($LocalFODSource)
        IF((Test-Path $LocalFODSource) -and (Get-ChildItem -Path $LocalFODSource -File *.cab).count -ge 1)
        {
            Write-Host "Local FOD Source Files Found on $env:COMPUTERNAME" -BackgroundColor Blue -ForegroundColor White
            RETURN $TRUE
        }
        ELSE {RETURN $FALSE}
    }

    ############# END: Script Blocks for Reading and Setting Registry Values ################
    
    
    IF(($FeaturesToInstallRemove.COUNT -ge 1) -and ($FeaturesToInstallRemove -ne "CANCELED"))
    {
        #Write-Host "I'm Here: 1" -BackgroundColor Blue -ForegroundColor White
        $NewComputer = $null
        $FeatureIndex = 1
        $FeatureCount = $FeaturesToInstallRemove.Count
        FOREACH ($Feature in $FeaturesToInstallRemove)
        {
            $CurrentComputer = $Feature.AD_ComputerName
            IF($PrevComputer -eq $null)
            {
                $PrevComputer = $CurrentComputer
                $NewComputer = $true
                $PendingReboot = $NULL
                $WindowsBuildSupported = $null
                $LocalFODSourceFilesExist = $NULL
            }
            ELSEIF($CurrentComputer -ne $PrevComputer)
            {
                ############# BEGIN: Re-Enable the use of a Windows Update Automatic Update Server (WSUS) ################
                IF(($PrevComputer -ne $env:COMPUTERNAME) -and ($WSUSDisabled)) #For REMOTE Computers
                {
                    Invoke-Command  -ComputerName $PrevComputer -ScriptBlock $UpdateRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue, $Enable, "DWord" 2> $null
                    Invoke-Command -ComputerName $PrevComputer -ScriptBlock $RestartWindowsUpdateServiceScriptBlock
                    $WSUSDisabled = $null
                }
                ELSEIF(($PrevComputer -eq $env:COMPUTERNAME) -and ($WSUSDisabled)) #For LOCAL Computer
                {
                    Invoke-Command -ScriptBlock $UpdateRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue, $Enable, "DWord" 2> $null
                    Invoke-Command -ScriptBlock $RestartWindowsUpdateServiceScriptBlock
                    $WSUSDisabled = $null
                }
                ############# END: Re-Enable the use of a Windows Update Automatic Update Server (WSUS) ################
                $PrevComputer = $CurrentComputer
                $NewComputer = $true
                $PendingReboot = $NULL
                $WindowsBuildSupported = $null
                $LocalFODSourceFilesExist = $NULL
            }

            ############# BEGIN: Check for Pending Reboot on the $CurrentComputer ###############################
            IF(($PendingReboot -eq $NULL) -and ($NewComputer) -AND ($CurrentComputer -ne $env:COMPUTERNAME)) #For REMOTE Computers
            {
                $PendingReboot = Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $CheckPendingReboot_ScriptBlock
            }
            ELSEIF(($PendingReboot -eq $NULL) -and ($NewComputer) -AND ($CurrentComputer -eq $env:COMPUTERNAME)) #For LOCAL Computer
            {
                $PendingReboot = Invoke-Command -ScriptBlock $CheckPendingReboot_ScriptBlock
            }
            ############# END: Check for Pending Reboot on the $CurrentComputer ###############################


            ############# BEGIN: Check for Windows Build Supported on the $CurrentComputer ###############################
            IF(($WindowsBuildSupported -eq $NULL) -and ($NewComputer) -AND ($CurrentComputer -ne $env:COMPUTERNAME)) #For REMOTE Computers
            {
                $WindowsBuildSupported = Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $WindowsBuildSupported_ScriptBlock
            }
            ELSEIF(($WindowsBuildSupported -eq $NULL) -and ($NewComputer) -AND ($CurrentComputer -eq $env:COMPUTERNAME)) #For LOCAL Computer
            {
                $WindowsBuildSupported = Invoke-Command -ScriptBlock $WindowsBuildSupported_ScriptBlock
            }
            ############# END: Check for Windows Build Supported on the $CurrentComputer ###############################
            

            ############# BEGIN: Check for Local FOD Source Files on the $CurrentComputer ###############################
            IF(($LocalFODSourceFilesExist -eq $NULL) -and ($NewComputer) -AND ($CurrentComputer -ne $env:COMPUTERNAME)) #For REMOTE Computers
            {
                $LocalFODSourceFilesExist = Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $LocalFODSourceFilesExist_ScriptBlock -ArgumentList $LocalFODSource
            }
            ELSEIF(($LocalFODSourceFilesExist -eq $null) -and ($NewComputer) -and ($CurrentComputer -eq $env:COMPUTERNAME)) #For LOCAL Computer
            {
                $LocalFODSourceFilesExist = Invoke-Command -ScriptBlock $LocalFODSourceFilesExist_ScriptBlock -ArgumentList $LocalFODSource
            }
            ############# END: Check for Local FOD Source Files on the $CurrentComputer ###############################
            
                        
            ############# BEGIN: CAN ONLY PERFORM FOD INSTALLS ON SYSTEMS WHICH YOU HAVE ADMIN RIGHTS TO, THERE ARE NO PENDING REBOOTS, AND THE WINDOWS BUILD SUPPORTS FOD UPDATES
            IF(($PendingReboot -ne $true -and $PendingReboot -ne $null) -and ($WindowsBuildSupported[0] -eq $true -and $WindowsBuildSupported -ne $null))
            {
                    
                ############# BEGIN: Disable the use of a Windows Update Automatic Update Server (WSUS) ################
                IF(($NewComputer) -AND ($CurrentComputer -ne $env:COMPUTERNAME)) #For REMOTE Computers
                {
                    $WindowsUpdateAUKeyExists = Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $TestRegKey_ScriptBlock -ArgumentList $WindowsUpdateAUKey 2> $null
                    IF($WindowsUpdateAUKeyExists) # Get check existence of UseWUServer value, get its value if exist
                    {
                        $UseWUServerExists = Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $TestRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue 2> $null
                        IF($UseWUServerExists)
                        {
                            $UseWUServer = Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $GetRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue 2> $null

                            IF($UseWUServer -eq 1)
                            {
                                Invoke-Command  -ComputerName $CurrentComputer -ScriptBlock $UpdateRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue, $Disable, "DWord" 2> $null
                                Invoke-Command -ComputerName $CurrentComputer -ScriptBlock $RestartWindowsUpdateServiceScriptBlock
                                $WSUSDisabled = $true            
                            }
                        }
                    }
                }
                ELSEIF(($NewComputer) -AND ($CurrentComputer -eq $env:COMPUTERNAME)) #For LOCAL Computer
                {
                    $WindowsUpdateAUKeyExists = Invoke-Command -ScriptBlock $TestRegKey_ScriptBlock -ArgumentList $WindowsUpdateAUKey 2> $null
                    IF($WindowsUpdateAUKeyExists) # Get check existence of UseWUServer value, get its value if exist
                    {
                        $UseWUServerExists = Invoke-Command -ScriptBlock $TestRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue 2> $null
                        IF($UseWUServerExists)
                        {
                            $UseWUServer = Invoke-Command -ScriptBlock $GetRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue 2> $null

                            IF($UseWUServer -eq 1)
                            {
                                Invoke-Command -ScriptBlock $UpdateRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue, $Disable, "DWord" 2> $null
                                Invoke-Command -ScriptBlock $RestartWindowsUpdateServiceScriptBlock
                                $WSUSDisabled = $true            
                            }
                        }
                    }
                }
                ############# END: Disable the use of a Windows Update Automatic Update Server (WSUS) ################


                ############# BEGIN: Install or Remove the Feature from the $CurrentComputer #########################
                IF($Feature.FOD_CAPABILITY_State -eq "NotPresent") #Install FOD Capability
                {
                    Write-Progress -Id 1 -Activity "Getting all selected Computer Objects' Windows Capabilities" -Status "Ready" -Completed
                    Write-Progress -Activity "Populating Windows Computers Selection List Form...please wait..." -Status "Ready" -Completed
                    
                    $FOD_CAPABILITY_Name = $Feature.FOD_CAPABILITY_Name.ToString()
                                        
                    #USE LOCAL FOD SOURCE FILES IF AVAILABLE
                    IF($LocalFODSourceFilesExist)
                    {
                        Write-Host "Attempting to Install $FOD_CAPABILITY_Name using FOD Files from $LocalFODSource" -BackgroundColor DarkMagenta -ForegroundColor White
                        $Command = "DISM.EXE /Online /Add-Capability /CapabilityName:$FOD_CAPABILITY_Name /Source:$LocalFODSource /LimitAccess"
                    }
                    ELSEIF($WindowsBuildSupported[1] -eq "17763")
                    {   #USE REMOTE SHARE FOD SOURCE FILES IF LOCAL FOD FILES DO NOT EXIST
                        Write-Host "Attempting to Install $FOD_CAPABILITY_Name using FOD Files from $FODSource1809" -BackgroundColor DarkMagenta -ForegroundColor White
                        $Command = "DISM.EXE /Online /Add-Capability /CapabilityName:$FOD_CAPABILITY_Name /Source:$FODSource1809 /LimitAccess"
                    }
                    ELSEIF($WindowsBuildSupported[1] -eq "18362")
                    {   #USE REMOTE SHARE FOD SOURCE FILES IF LOCAL FOD FILES DO NOT EXIST
                        Write-Host "Attempting to Install $FOD_CAPABILITY_Name using FOD Files from $FODSource1903" -BackgroundColor DarkMagenta -ForegroundColor White
                        $Command = "DISM.EXE /Online /Add-Capability /CapabilityName:$FOD_CAPABILITY_Name /Source:$FODSource1903 /LimitAccess"
                    }
                    ELSEIF($WindowsBuildSupported[1] -eq "18363")
                    {   #USE REMOTE SHARE FOD SOURCE FILES IF LOCAL FOD FILES DO NOT EXIST
                        Write-Host "Attempting to Install $FOD_CAPABILITY_Name using FOD Files from $FODSource1909" -BackgroundColor DarkMagenta -ForegroundColor White
                        $Command = "DISM.EXE /Online /Add-Capability /CapabilityName:$FOD_CAPABILITY_Name /Source:$FODSource1909 /LimitAccess"
                    }   
                    ELSEIF($WindowsBuildSupported[1] -eq "19042")
                    {   #USE REMOTE SHARE FOD SOURCE FILES IF LOCAL FOD FILES DO NOT EXIST
                        Write-Host "Attempting to Install $FOD_CAPABILITY_Name using FOD Files from $FODSource20H2" -BackgroundColor DarkMagenta -ForegroundColor White
                        $Command = "DISM.EXE /Online /Add-Capability /CapabilityName:$FOD_CAPABILITY_Name /Source:$FODSource20H2 /LimitAccess"
                    }
                    
                    Write-Host "Installing FOD: $FOD_CAPABILITY_Name on $CurrentComputer" -ForegroundColor White -BackgroundColor DarkGreen
                    & psexec64.exe \\$CurrentComputer -s -nobanner -accepteula powershell -executionpolicy bypass -command $Command *>&1 | Tee-Object -Variable 'Output' # REDIRECT STD OUT ONLY >, REDIRECT STD ERROR ONLY 2>, REDIRECTS ALL *>&1, REDIRECT STD OUT ONLY, WITH CMD --% 1>, REDIRECT STD ERROR ONLY, WITH CMD --% 2>

                    <#
                    $Output.gettype()
                    Write-Host "OutputType=$($Output.gettype())"
                    foreach ($Line in $Output)
                    {
                        Write-Host "$Line"
                        $Line.Length
                    }
                    #>

                    #USE Microsoft Windows Update Site IF REMOTE SHARE FOD SOURCE FILES DO NOT EXIST OR OTHER ERRORS WERE ENCOUNTERED
                    $SourceFilesNotFound = "The source files could not be found. "
                    $CapabilityNameNotRecognized = "A Windows capability name was not recognized."
                    IF(($CapabilityNameNotRecognized -in $Output) -or ($SourceFilesNotFound -in $Output))
                    {
                        $Command = "DISM.EXE /Online /Add-Capability /CapabilityName:$FOD_CAPABILITY_Name"
                        Write-Host "Attempting Installation of FOD: $FOD_CAPABILITY_Name on $CurrentComputer from Microsoft Windows Update site, since the local source failed." -ForegroundColor Red -BackgroundColor Yellow
                        Write-Host "Output Error Encountered for Installation of FOD: $FOD_CAPABILITY_Name on $CurrentComputer`n $Output" -ForegroundColor Red -BackgroundColor Yellow
                        & psexec64.exe \\$CurrentComputer -s -nobanner -accepteula powershell -executionpolicy bypass -command $Command
                    }
                }
                ELSEIF($Feature.FOD_CAPABILITY_State -eq "Installed") #Remove FOD Capability
                {
                    Write-Progress -Id 1 -Activity "Getting all selected Computer Objects' Windows Capabilities" -Status "Ready" -Completed
                    Write-Progress -Activity "Populating Windows Computers Selection List Form...please wait..." -Status "Ready" -Completed
                    
                    $FOD_CAPABILITY_Name = ($Feature.FOD_CAPABILITY_Name).ToString()
                    $Command = "DISM.EXE /Online /Remove-Capability /CapabilityName:$FOD_CAPABILITY_Name"
                    Write-Host "Removing FOD: $FOD_CAPABILITY_Name on $CurrentComputer" -ForegroundColor Red -BackgroundColor Yellow
                    & psexec64.exe \\$CurrentComputer -s -nobanner -accepteula powershell -executionpolicy bypass -command $Command
                }
                $NewComputer = $false
                ############# END: Install or Remove the Feature from the $CurrentComputer #########################


                #LOGIC To Re-Enable WSUS Update after Installing/Removing the last feature on the last computer.
                IF($FeatureIndex -eq $FeatureCount)
                {
                    ############# BEGIN: Re-Enable the use of a Windows Update Automatic Update Server (WSUS) ################
                    IF(($CurrentComputer -ne $env:COMPUTERNAME) -and ($WSUSDisabled)) #For REMOTE Computers
                    {
                        Invoke-Command  -ComputerName $PrevComputer -ScriptBlock $UpdateRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue, $Enable, "DWord" 2> $null
                        Invoke-Command -ComputerName $PrevComputer -ScriptBlock $RestartWindowsUpdateServiceScriptBlock            
                        $WSUSDisabled = $null
                    }
                    ELSEIF(($CurrentComputer -eq $env:COMPUTERNAME) -and ($WSUSDisabled)) #For LOCAL Computer
                    {
                        Invoke-Command -ScriptBlock $UpdateRegValue_ScriptBlock -ArgumentList $WindowsUpdateAUKey, $UseWUServerValue, $Enable, "DWord" 2> $null
                        Invoke-Command -ScriptBlock $RestartWindowsUpdateServiceScriptBlock
                        $WSUSDisabled = $null          
                    }
                    ############# END: Re-Enable the use of a Windows Update Automatic Update Server (WSUS) ################
                } #END IF($FeatureIndex -eq $FeatureCount)

            } #END IF(($PendingReboot -ne $true -and $PendingReboot -ne $null) -and ($WindowsBuildSupported -eq $true -and $WindowsBuildSupported -ne $null))
            ELSE #
            {
                IF($PendingReboot){Write-Host "$CurrentComputer is pending a reboot; reboot the computer and attempt again." -BackgroundColor Yellow -ForegroundColor Red}
                IF(!($WindowsBuildSupported)){Write-Host "$CurrentComputer has a build which is not supported by this script." -BackgroundColor Yellow -ForegroundColor Red}
            }
            ############# END: CAN ONLY PERFORM FOD INSTALLS ON SYSTEMS WHICH YOU HAVE ADMIN RIGHTS TO, THERE ARE NO PENDING REBOOTS, AND THE WINDOWS BUILD SUPPORTS FOD UPDATES

            $FeatureIndex++
        }
    }

} #END IF($GetWindowsCapabilitiesResults[1].COUNT -ge 1)

