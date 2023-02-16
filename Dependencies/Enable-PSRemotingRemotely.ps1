FUNCTION Enable-PSRemotingRemotely #Version 4
{
 <#
.SYNOPSIS
    Enables PSremoting on a single computer name or array of computer names passed in as a parameter, or on the local host if no parameters are provided. Returns $PSenabledResults
.DESCRIPTION
    Enables PSremoting on a single computer name or array of computer names passed in as a parameter, or on the local host if no parameters are provided. Returns $PSenabledResults
.PARAMETER $ComputerNames
    Single computer name, or array of computer names that will be targeted for the enablement of PSremoting
.PARAMETER $OutputFilePath
    The path location where the $PSenabledResults file will be saved on the local filesystem.
.INPUTS
    Pipe input is possible
.OUTPUTS 
    Returns $PSenabledResults to a csv file named PSenabledResults_MMddyyyy.csv, or PSenabledResults_MMddyyyy_ComputerName.csv in the $OutputFilePath
.EXAMPLES
    .\PSRemotingRemotely.ps1  
    .\PSRemotingRemotely.ps1 -ComputerNames Computer1, Computer2
    .\PSRemotingRemotely.ps1 -ComputerNames Computer1, Computer2 -LogPath "c:\temp\PSenabledResults"
    $targetComputers | PSRemotingRemotely.ps1
.NOTES
    Author:         Michael D. Sloan
    Date:           4/10/2018
    
    Changelog:
        4/10/2018     PSRemotingRemotely - Initial Release
        4/16/2018     Added multi-threaded processing     
        9/18/2018     Added LogPath Parameter and Logging
        
#>

     #Enables PSremoting on a single computer name or array of computer names passed in as a parameter, or on the local host if no parameters are provided. Returns $PSenabledResults
    [cmdletbinding()]    
    #Parameters passed into the funtion
    Param (
        [Parameter(Mandatory=$false, ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
		[array]$ComputerNames = $env:computername,
        
        [Parameter(ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true,Mandatory=$false,
        HelpMessage="Enter the path where you want the log file and CSV Results to be written.")]
		[string]$LogPath = "C:\temp\Launcher\Logs\psremoting\"
              
         <#,
        
        [Parameter(Mandatory=$false, ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$OutputFileName = "PSEnabledResults_{0}.csv" -f (Get-Date -Format "MMddyyyy"),

        [Parameter(Mandatory=$false, ValueFromPipeline=$true,ValueFromPipelineByPropertyName=$true)]
        [string]$OutputFilePath = "{0}\Desktop\" -f $env:USERPROFILE #>
          )
    
    begin 
    {
         function Write-Log #Version 2
        { 
            <# 
                .Synopsis 
                   Write-Log writes a message to a specified log file with the current time stamp. 
                .DESCRIPTION 
                   The Write-Log function is designed to add logging capability to other scripts. 
                   In addition to writing output and/or verbose you can write to a log file for 
                   later debugging. 
                .NOTES 
                   Created by: Jason Wasser @wasserja 
                   Modified: 11/24/2015 09:30:19 AM   

                   Modified by: Michael Sloan
                   Modified: 9/18/2018
		   
		   Modified by: Drew Burgess
                   Modified: 11/15/2020
 
                   Changelog: 
                    * Code simplification and clarification - thanks to @juneb_get_help 
                    * Added documentation. 
                    * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks 
                    * Revised the Force switch to work as it should - thanks to @JeffHicks 
 
                   To Do: 
                    * Add error handling if trying to create a log file in a inaccessible location. 
                    * Add ability to write $Message to $Verbose or $Error pipelines to eliminate 
                      duplicates. 
                .PARAMETER Message 
                   Message is the content that you wish to add to the log file.  
                .PARAMETER Path 
                   The path to the log file to which you would like to write. By default the function will  
                   create the path and file if it does not exist.  
                .PARAMETER Level 
                   Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational) 
                .PARAMETER NoClobber 
                   Use NoClobber if you do not wish to overwrite an existing file. 
                .EXAMPLE 
                   Write-Log -Message 'Log message'  
                   Writes the message to c:\Logs\PowerShellLog.log. 
                .EXAMPLE 
                   Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log 
                   Writes the content to the specified log file and creates the path and file specified.  
                .EXAMPLE 
                   Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
                   Writes the message to the specified log file as an error message, and writes the message to the error pipeline. 
                .LINK 
                   https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
        #>     
    
            [CmdletBinding()] 
            Param 
            ( 
                [Parameter(Mandatory=$true, 
                           ValueFromPipelineByPropertyName=$true)] 
                [ValidateNotNullOrEmpty()] 
                [Alias("LogContent")] 
                [string]$Message, 
 
                [Parameter(Mandatory=$false)] 
                [Alias('LogPath')] 
                [string]$Path='C:\Logs\PowerShellLog.log', 
         
                [Parameter(Mandatory=$false)] 
                [ValidateSet("Error","Critical","Warn","Info")] 
                [string]$Level="Info", 
         
                [Parameter(Mandatory=$false)] 
                [switch]$NoClobber 
            ) 
 
            Begin 
            { 
                # Set VerbosePreference to Continue so that verbose messages are displayed. 
                $VerbosePreference = 'Continue' 
            } 
            Process 
            { 
         
                # If the file already exists and NoClobber was specified, do not write to the log. 
                if ((Test-Path $Path) -AND $NoClobber) { 
                    Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
                    Return 
                    } 
 
                # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
                elseif (!(Test-Path $Path)) { 
                    Write-Verbose "Creating $Path." 
                    $NewLogFile = New-Item $Path -Force -ItemType File 
                    } 
 
                else { 
                    # Nothing to see here yet. 
                    } 
 
                # Format Date for our Log File 
                $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 

                function write-critical ($critical)
                {
                    write-host -ForegroundColor Magenta "CRITICAL: " $critical
                }
 
                # Write message to error, warning, or verbose pipeline and specify $LevelText 
                switch ($Level) { 
                    'Error' { 
                        Write-Error $Message 
                        $LevelText = 'ERROR:' 
                        }
                    'Critical' { 
                        Write-Critical $Message 
                        $LevelText = 'CRITICAL:' 
                        }      
                    'Warn' { 
                        Write-Warning $Message 
                        $LevelText = 'WARNING:' 
                        } 
                    'Info' { 
                        Write-Verbose $Message 
                        $LevelText = 'INFO:' 
                        } 
                    } 
         
                # Write log entry to $Path 
                "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append 
            } 
            End 
            { 
            } 
        }

        function Test-Port
        {
            param(
                    [string]$srv,
                    $port=5985,
                    $timeout=750,
                    [switch]$verbose=$false
                 )

            $ErrorActionPreference = "SilentlyContinue"
    
            $tcpclient = new-Object System.Net.Sockets.TcpClient
            $iar = $tcpclient.BeginConnect($srv,$port,$null,$null)
            $wait = $iar.AsyncWaitHandle.WaitOne($timeout,$false)
 
            if(!$wait)
                {
                    $tcpclient.Close()
                    if($verbose){Write-Host "Connection Timeout"}
                    Return $false
                }
            else
                {
                    $error.Clear()
                    $tcpclient.EndConnect($iar) | out-Null
                    if($error[0])
                        {
                            if($verbose){write-host $error[0]}
                            $failed = $true
                        }
                    $tcpclient.Close()
                }
            if($failed)
                {
                    return $false
                }
            else
                {
                    return $true
                }
        }

        $Today = get-date
        $TodayShort = Get-Date -Format "yyyy-MM-dd" 
        $LogFile = "$LogPath\PSRemotingEnabledRemotely__$TodayShort.log"

        if (!(Test-Path $LogPath)) #If LogPath doesn't exist, create the folder path
        {
            $NewLogPath = New-Item $LogPath -Force -ItemType Directory
            $Message = "********************Logging Directory Created for PSRemoting********************"
            Write-Log -Message $Message -Path $LogFile -Level Warn 

            $Message = "********************START OF NEW ENTRY********************"
            Write-Log -Message $Message -Path $LogFile -Level Info
        }
        else
        {
            $Message = "********************START OF NEW ENTRY********************"
            Write-Log -Message $Message -Path $LogFile -Level Info
        }

        $TotalComputerNames = $ComputerNames.Count
        Write-Log -Message "$TotalComputerNames targeted for PSRemoting Enabled Remotely." -Path $LogFile -Level Info

        $Message = "**********************************************************"
        Write-Log -Message $Message -Path $LogFile -Level Info

        FUNCTION CreatePSObjectForPSRemotingRemotely
        {
	        $obj = New-Object PSObject
            $obj | Add-Member -MemberType NoteProperty -Name "ComputerName" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "PSremotingEnabled" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "PSremotingProtocolVersion" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "PSVersion" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "WinRMhttpPort5985" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "WinRMhttpsPort5986" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "ClientServerComPort135" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "NetBIOSPort139" -Value ""
            $obj | Add-Member -MemberType NoteProperty -Name "SMBPort445" -Value ""

            return $obj
        }
        
        

        

    } #end begin

    process 
    {     
        #Zero out $PSenabledResults Array
        [array]$PSenabledResults = $null
                
        $Count = 1
        foreach($Computer in $ComputerNames)
        {
            Write-Progress -Id 0 -Activity "Enabling PSremoting on one or more computers" -Status "$Count of $($ComputerNames.Count)" -PercentComplete (($Count / $ComputerNames.Count) * 100)

            $objPSVersionTable = "" #Powershell Version Table, based on current $Computer in the array for processing, initilizing variable
            $objPSenabledNFO = CreatePSObjectForPSRemotingRemotely #Creating a PSObject to hold information about PS remoting status.
                        
            IF ($Computer.name){$Computer=$Computer.name} #If $CompterNames is an array of PS Computer Objects, this parses just the Computer Name portion.
            #write-host $Computer "is the current computer processing"

            $Message = "**********************************BEGIN ENTRY FOR: $Computer**********************************"
            Write-Log -Message $Message -Path $LogFile -Level Info
            Write-Log -Message "Attempting to enable PSRemoting Remotely on $Computer." -Path $LogFile -Level Info

            
            $WinRMhttpPort = Test-Port -srv $Computer -port 5985 #WinRM over HTTP
            $WinRMhttpsPort = Test-Port -srv $Computer -port 5986 #WinRM over HTTPS
            $ClientServerComPort = Test-Port -srv $Computer -port 135 #Client/Server Comms
            $NetBIOSPort = Test-Port -srv $Computer -port 139 #NetBIOS Session Service
            $SMBPort = Test-Port -srv $Computer -port 445 #SMB i.e. CIFS
            try {
                    #Invoke a PS Command on the remote computer, to test the availability of WinRM service on the computer and to get its PSVersionTable
                    $objPSVersionTable = Invoke-Command -ComputerName $Computer -ScriptBlock {$PSVersionTable} -ErrorAction stop 2>> $LogFile
                                        
                    #Success Output
                    #IF($PSBoundParameters['Verbose']) { Write-Verbose "PSremoting - Enabled on $Computer" } else { $true }
                         
                    Write-Progress -id 1 -Activity "PSremoting - Enabled on $Computer"
                    Write-Log -Message "PSremoting - Enabled on $Computer" -Path $LogFile -Level Info
                    $objPSenabledNFO.ComputerName = $Computer
                    $objPSenabledNFO.PSremotingEnabled = $true
                            
                    #Getting PSremotingProtocolVersion
                    $PSremotingProtocolVersionMajor = $objPSVersionTable.PSRemotingProtocolVersion.Major
                    $PSremotingProtocolVersionMinor = $objPSVersionTable.PSRemotingProtocolVersion.Minor
                    $PSremotingProtocolVersionBuild = $objPSVersionTable.PSRemotingProtocolVersion.Build
                    $PSremotingProtocolVersionRevision = $objPSVersionTable.PSRemotingProtocolVersion.Revision
                    $objPSenabledNFO.PSremotingProtocolVersion =  ("{0}.{1}.{2}.{3}" -f $PSremotingProtocolVersionMajor,$PSremotingProtocolVersionMinor,$PSremotingProtocolVersionBuild,$PSremotingProtocolVersionRevision)
                            
                    #Getting PSVersion
                    $PSVersionMajor = $objPSVersionTable.PSVersion.Major 
                    $PSVersionMinor = $objPSVersionTable.PSVersion.Minor
                    $PSVersionBuild = $objPSVersionTable.PSVersion.Build
                    $PSVersionRevision = $objPSVersionTable.PSVersion.Revision
                    $objPSenabledNFO.PSVersion = ("{0}.{1}.{2}.{3}" -f $PSVersionMajor,$PSVersionMinor,$PSVersionBuild,$PSVersionRevision)
                            
                    #Getting TCP Ports
                    $objPSenabledNFO.WinRMhttpPort5985 = $WinRMhttpPort
                    $objPSenabledNFO.WinRMhttpsPort5986 = $WinRMhttpsPort
                    $objPSenabledNFO.ClientServerComPort135 = $ClientServerComPort
                    $objPSenabledNFO.NetBIOSPort139 = $NetBIOSPort
                    $objPSenabledNFO.SMBPort445 = $SMBPort
                    
                    #Write to Result $PSenabledResults Set        
                    $PSenabledResults += $objPSenabledNFO
                }
            catch {
                    #Failure Output
                    #IF($PSBoundParameters['Verbose']) { Write-Verbose "PSremoting - Not Working on $Computer"; write-error $_.ToString() } else { $false }
                            
                    Write-Progress -id 1 -Activity "PSremoting - Not Working on $Computer, testing ports necessary for PSExec..."
                    Write-Log -Message "PSremoting - Not Working on $Computer, testing ports necessary for PSExec..." -Path $LogFile -Level warn
                        
                    if ($ClientServerComPort -and $NetBIOSPort -and $SMBPort) #Testing Client/Server port, NetBIOS Session Service Port, and SMB Port for availability, if ALL are available, attempt to enable PSremoting
                        {
                            Write-Progress -id 1 -Activity "TCP ports necessary for PSExec tested successfully on $Computer, attempting to enable PSremoting"
                            Write-Log -Message "TCP ports necessary for PSExec tested successfully on $Computer, attempting to enable PSremoting" -Path $LogFile -Level Info
                            $psexecResult = C:\Windows\System32\PsExec.exe \\$Computer -s powershell -executionpolicy bypass -command  "Enable-PSRemoting -Force -ErrorAction SilentlyContinue" >> $LogFile
                                
                            try {
                                    #Invoke a PS Command on the remote computer, to test the availability of WinRM service on the computer and to get its PSVersionTable
                                    $objPSVersionTable = Invoke-Command -ComputerName $Computer -ScriptBlock {$PSVersionTable} -ErrorAction stop
                                    Write-Progress -id 1 -Activity "Running psexec.exe $Computer -s Enable-PSRemoting -Force -ErrorAction SilentlyContinue on $Computer"
                                    Write-Log -Message "Running psexec.exe $Computer -s Enable-PSRemoting -Force -ErrorAction SilentlyContinue on $Computer" -Path $LogFile -Level Info
                                    $objPSenabledNFO.ComputerName = $Computer
                                    $objPSenabledNFO.PSremotingEnabled = "Attempting...SUCCEEDED"
                                                                       
                                    #Getting PSremotingProtocolVersion
                                    $PSremotingProtocolVersionMajor = $objPSVersionTable.PSRemotingProtocolVersion.Major
                                    $PSremotingProtocolVersionMinor = $objPSVersionTable.PSRemotingProtocolVersion.Minor
                                    $PSremotingProtocolVersionBuild = $objPSVersionTable.PSRemotingProtocolVersion.Build
                                    $PSremotingProtocolVersionRevision = $objPSVersionTable.PSRemotingProtocolVersion.Revision
                                    $objPSenabledNFO.PSremotingProtocolVersion =  ("{0}.{1}.{2}.{3}" -f $PSremotingProtocolVersionMajor,$PSremotingProtocolVersionMinor,$PSremotingProtocolVersionBuild,$PSremotingProtocolVersionRevision)
                        
                                    #Getting PSVersion
                                    $PSVersionMajor = $objPSVersionTable.PSVersion.Major 
                                    $PSVersionMinor = $objPSVersionTable.PSVersion.Minor
                                    $PSVersionBuild = $objPSVersionTable.PSVersion.Build
                                    $PSVersionRevision = $objPSVersionTable.PSVersion.Revision
                                    $objPSenabledNFO.PSVersion = ("{0}.{1}.{2}.{3}" -f $PSVersionMajor,$PSVersionMinor,$PSVersionBuild,$PSVersionRevision)
                                    
                                    #Re-Evaluating WinRM TCP Ports
                                    $WinRMhttpPort = Test-Port -srv $Computer -port 5985 #WinRM over HTTP
                                    $WinRMhttpsPort = Test-Port -srv $Computer -port 5986 #WinRM over HTTPS
                                    Write-Log -Message "WinRM Port Status on $Computer : HTTP Port Open: $WinRMhttpPort   HTTPS Port Open: $WinRMhttpsPort" -Path $LogFile -Level Info
                                    IF($WinRMhhtpPort -OR $WinRMhttpsPort)
                                    {
                                        Write-Log -Message "WinRM appears to have been successfully enabled on $Computer" -Path $LogFile -Level Info
                                    }
                                    ELSE
                                    {
                                        Write-Log -Message "WinRM DOESN'T appear to have been successfully enabled on $Computer" -Path $LogFile -Level Critical
                                    }

                                    #Getting TCP Ports
                                    $objPSenabledNFO.WinRMhttpPort5985 = $WinRMhttpPort
                                    $objPSenabledNFO.WinRMhttpsPort5986 = $WinRMhttpsPort
                                    $objPSenabledNFO.ClientServerComPort135 = $ClientServerComPort
                                    $objPSenabledNFO.NetBIOSPort139 = $NetBIOSPort
                                    $objPSenabledNFO.SMBPort445 = $SMBPort

                                    #Write to Result $PSenabledResults Set
                                    $PSenabledResults += $objPSenabledNFO
                                }
                            catch {
                                    Write-Progress -id 1 -Activity "Failed to enable PSremoting on $Computer"
                                    Write-Log -Message "Failed to enable PSremoting on $Computer" -Path $LogFile -Level Critical
                                    $objPSenabledNFO.ComputerName = $Computer
                                    $objPSenabledNFO.PSremotingEnabled = "Attempting...FAILED"
                                    $objPSenabledNFO.PSremotingProtocolVersion = "Not Available"
                                    $objPSenabledNFO.PSVersion = "Not Available"
                                    $objPSenabledNFO.WinRMhttpPort5985 = $WinRMhttpPort
                                    $objPSenabledNFO.WinRMhttpsPort5986 = $WinRMhttpsPort
                                    $objPSenabledNFO.ClientServerComPort135 = $ClientServerComPort
                                    $objPSenabledNFO.NetBIOSPort139 = $NetBIOSPort
                                    $objPSenabledNFO.SMBPort445 = $SMBPort
                                    $PSenabledResults += $objPSenabledNFO

                                }
                        }
                    else 
                        {
                            Write-Progress -id 1 -Activity "TPC ports necessary for PSExec failed on $Computer, aborting attempt to enable PSremoting"
                            Write-Log -Message "TPC ports necessary for PSExec failed on $Computer, aborting attempt to enable PSremoting" -Path $LogFile -Level Critical
                            $objPSenabledNFO.ComputerName = $Computer
                            $objPSenabledNFO.PSremotingEnabled = "Attempt aborted..."
                            $objPSenabledNFO.PSremotingProtocolVersion = "Not Available"
                            $objPSenabledNFO.PSVersion = "Not Available"
                            $objPSenabledNFO.WinRMhttpPort5985 = $WinRMhttpPort
                            $objPSenabledNFO.WinRMhttpsPort5986 = $WinRMhttpsPort
                            $objPSenabledNFO.ClientServerComPort135 = $WinRMhttpsPort
                            $objPSenabledNFO.NetBIOSPort139 = $NetBIOSPort
                            $objPSenabledNFO.SMBPort445 = $SMBPort
                            $PSenabledResults += $objPSenabledNFO
                        }
                }
            $Message = "**********************************`nEND ENTRY FOR: $Computer`n**********************************"
            Write-Log -Message $Message -Path $LogFile -Level Info
            $Count ++

        } #END foreach($Computer in $ComputerNames)
        #foreach ($Result in $PSenabledResults){write-host $Result}
        return $PSenabledResults
    } #END process
    end {}
} #END FUNCTION PSRemotingRemotely


<#
$Today = get-date
$TodayShort = Get-Date -Format "yyyy-MM-dd" 
$LogPath = 'C:\Logs\PSRemoting'
$outputCSV = $LogPath + "\PSRemotingEnablementResults"+"__$TodayShort.csv"

Enable-PSRemotingRemotely -ComputerNames $targetComputers -LogPath $LogPath | Export-Csv -notypeinformation -Path $outputCSV
#>






