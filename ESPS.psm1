#-----Script Info---------------------------------------------------------------------------------------------
# Name:ESPS.psm1
# Author: Einar Stenberg
# Date: 24.11.14
# Version: 20
# Job/Tasks:
#--------------------------------------------------------------------------------------------------------------


#-----Changelog------------------------------------------------------------------------------------------------
#v1.  Script created ES
#v2.  Added Function Refresh-GPO
#v3.  Added Function Uninstall-WindowsUpdate and bugfix for argumentlist in invokecommand for Refresh-GPO
#v4.  Added Function Get-ESI
#v5.  Update Get-ESI with wildcard search for displayname
#v6.  Update Get-ESI with wildcard for compure and username, added switch -XtraAdInfo to retrieve additional info from ad
#v7.  Added Function out-Excel
#v8.  Update Get-ESI to return all entries if no switches are set
#v9.  Added Function Out-Voice
#v10. Update Get-ESI to use computername instead of computer for csv header, added verbose debugging and pipeline output
#v11. Update Get-ESI with RemoteAssistance switch with grid view
#v12. Added function Remove-Software
#v13. Added function Out-HTML
#v14. Added function Start-Progress
#v15. Added function ForEach-Parallel
#v16. Fix for Get-ESI -XtraADInfo to add information as members of the objects, solves problem when more than one is returned
#v17. Added function Get-ADPasswordExpiration
#v18. Updated Get-ESI's -xtraadinfo switch to contain results from Get-ADPasswordExpiration
#v19. Updated Update-ESPS with new switches for serverside repo functions
#v20. Added function New-Password
#v21. Added function Get-Windowskey
#--------------------------------------------------------------------------------------------------------------





#-----Functions---------------------------------------------------------------------------------------------


Function Update-ESPS {

<#
.SYNOPSIS
Pulls updates for ESPS and other included modules from source
.DESCRIPTION
Requests credentials and creates a session with Azure instance to be used with session terminator cmdlets
Part of ESPS by ES
.EXAMPLE
Update-ESPS
#>

Param(
[switch]$ListModules,
$InstallModule
)

$tempfile="$([System.IO.Path]::GetTempFileName()).ps1"
Invoke-WebRequest https://wiki.ap3x.org/bin/download/Main/ESPS+Install++and+update+powershell+module+collection/ESPS.ps1 -OutFile $tempfile
If ($ListModules){& $tempfile -listmodules}
If ($InstallModule){& $tempfile -installmodule $InstallModule}
If (!$ListModules -and !$installmodule){& $tempfile}
Remove-Item $tempfile -Force
}


Function Refresh-GPO {
<#
.SYNOPSIS
Refreshes group memberships and issues a Group Policy Update to the computer
.DESCRIPTION
Refreshes group memberships and issues a Group Policy Update to the computer
Switches for automatic logoff and reboot can be added aswell as credentials
Part of ESPS by ES
.PARAMETER ComputerName
Defines target computer
.PARAMETER Logoff
Allows logoff on target machine if neccessary to complete refresh
.PARAMETER Boot
Allows reboot on target machine if neccessary to complete refresh
.PARAMETER PromptForCredentials
Displays a prompt asking for credentials to run on the remote computer
.PARAMETER Credential
Allows for passing a <PSCredential> to the remote computer for authentication
.EXAMPLE
Refresh-GPO computer.domain.local
.EXAMPLE
Refresh-GPO computer.domain.local -PromptForCredentials -Boot -Logoff
.EXAMPLE
Refresh-GPO computer.domain.local -Credentials $usercredentials
#>

Param(
[Parameter(Mandatory=$true,Position=0)]
$ComputerName,
[switch]$Logoff,
[switch]$Boot,
[switch]$PromptForCredentials,
$Credential
)


#Assigns Switch values
If ($Logoff){$logoffstring=" /logoff"}
If ($Boot){$bootstring=" /boot"}


#Runs if switch -PromptForCredentials is set
If ($PromptForCredentials -And !$Credential){
    Invoke-Command -AsJob -ComputerName $ComputerName -Credential $PromptCred -ArgumentList $logoffstring,$bootstring  -ScriptBlock {
        Param ($logoffstring, $bootstring)
        klist -lh 0 -li 0x3e7 purge
        gpupdate /force$bootstring$logoffstring
    }
}

#Runs if no switches are set
If (!$PromptForCredentials -And !$Credential){
    Invoke-Command -AsJob -ComputerName $ComputerName -ArgumentList $logoffstring,$bootstring -ScriptBlock {
        Param ($logoffstring, $bootstring)
        klist -lh 0 -li 0x3e7 purge
        gpupdate /force$bootstring$logoffstring
    }
}

#Runs if objet -Credentials is set
If ($Credential -And !$PromptForCredentials){
    Invoke-Command -AsJob -ComputerName $ComputerName -Credential $Credential -ArgumentList $logoffstring,$bootstring -ScriptBlock {
        Param ($logoffstring, $bootstring)
        klist -lh 0 -li 0x3e7 purge
        gpupdate /force$bootstring$logoffstring
    }
}


}

Function Uninstall-WindowsUpdate {
<#
.SYNOPSIS
Uninstalls Windows update with selected KBnumber from target computer
.DESCRIPTION
Uninstalls Windows update with selected KBnumber from target computer
Switches for automatic reboot can be added aswell as credentials
Part of ESPS by ES
.PARAMETER ComputerName
Defines target computer
.PARAMETER NoRestart
Machine will not restart after uninstall even if requered by update
.PARAMETER ForceRestart
Machine will restart after uninstall even if not requiered by update
.PARAMETER Log
Enables logging
.PARAMETER PromptForCredentials
Displays a prompt asking for credentials to run on the remote computer
.PARAMETER Credential
Allows for passing a <PSCredential> to the remote computer for authentication
.EXAMPLE
Uninstall-WindowsUpdate computer.domain.local -kbNumber 12345678
.EXAMPLE
Uninstall-WindowsUpdate computer.domain.local -kbNumber 12345678 -PromptForCredentials -NoRestart
#>

Param(
[Parameter(Mandatory=$true,Position=0)]
$ComputerName,
[Parameter(Mandatory=$true,Position=1)]
$KBNumber,
[switch]$NoRestart,
[switch]$ForceRestart,
[switch]$Log,
[switch]$PromptForCredentials,
$Credential
)


#Assigns Switch values
If ($ForceRestart){$RestartString=" /ForceRestart"}
If ($PromptRestart){$RestartString=" /PromptRestart"}
If ($NoRestart){$RestartString=" /NoRestart"}
If ($Log){$LogString=" /log"}

#Runs if switch -PromptForCredentials is set
If ($PromptForCredentials -And !$Credential){
    Invoke-Command -ComputerName $ComputerName -Credential $PromptCred -ArgumentList $KBNumber,$RestartString,$LogString,$ComputerName -ScriptBlock {
        Param ($KBNumber, $RestartString, $LogString, $ComputerName)
        Write-Host "$ComputerName Executing command: wusa /uninstall /KB:$KBNumber /quiet$RestartString$LogString"
        Start-Process -FilePath 'wusa.exe' -ArgumentList " /uninstall /KB:$KBNumber /quiet$RestartString$LogString" -Wait
    }
}
#Runs if no switches are set
If (!$PromptForCredentials -And !$Credential){
    Invoke-Command -ComputerName $ComputerName -ArgumentList $KBNumber,$RestartString,$LogString,$ComputerName -ScriptBlock {
        Param ($KBNumber, $RestartString, $LogString, $ComputerName)
        Write-Host "$ComputerName Executing command: wusa /uninstall /KB:$KBNumber /quiet$RestartString$LogString"
        Start-Process -FilePath 'wusa.exe' -ArgumentList " /uninstall /KB:$KBNumber /quiet$RestartString$LogString" -Wait
    }
}

#Runs if objet -Credentials is set
If ($Credential -And !$PromptForCredentials){
    Invoke-Command -ComputerName $ComputerName -Credential $Credential -ArgumentList $KBNumber,$RestartString,$LogString,$ComputerName -ScriptBlock {
        Param ($KBNumber, $RestartString, $LogString, $ComputerName)
        Write-Host "$ComputerName Executing command: wusa /uninstall /KB:$KBNumber /quiet$RestartString$LogString"
        Start-Process -FilePath 'wusa.exe' -ArgumentList " /uninstall /KB:$KBNumber /quiet$RestartString$LogString" -Wait
    }
}


}


Function Get-ESI{
<#
.SYNOPSIS
Retrives information from ESInfo's resultset
.DESCRIPTION
Retrives information from ESInfo's resultset as an object containing all the results

Note: 
To Store the location of ESInfo resultset permanently in your profile 
run the command once with the -SetESInfoFilePath switch enabled.
This will allow you to run the cmdlet in the future without specifing the path for the inputfile

Part of ESPS by ES
.PARAMETER Computername
Input computername to search for entry with that computername
.PARAMETER Username
Input Username to search for entry with that Username
.PARAMETER Computername
Input Computername to search for entry with that Computername
.PARAMETER Displayname
Input Displayname to search for entry with that Displayname
.PARAMETER ESIFile
Defines path to ESInfo resultfile
If not defined will load variable stored in profile $ESInfoPath
.PARAMETER SetESInfoFilePath
Saves path from parameter ESIFile in profile variable permanently
.PARAMETER RA
Launches RemoteAssistance for connecting to the desired computer
.EXAMPLE
Get-ESI -Username johndoe
.EXAMPLE
Get-ESI -Computername pc01
.EXAMPLE
Get-ESI -ESIFile \\server\share\folder\Domain.csv -SetESInfoFilePath
This will store the location of ESIfile in your profile permanently
.EXAMPLE
Get-ESI -RA -Username Johndoe
Will launch RemoteAssistance connection for found users or display a list if multiple matches are found
#>

Param(
[string]$ComputerName,
[string]$Username,
[string]$DisplayName,
[Parameter(Mandatory=$false,Position=0)]
#Defaults the ESInfo csv input to the variable $ESInfoPath, allowing preconfiguration of this in PS-profile
[string]$ESIFile = $Global:ESInfoPath,
#Sets the $ESInfoPath variable stored in profile to the path provided in $ESIFile input
[switch]$SetESInfoFilePath,
[switch]$XtraADInfo,
[switch]$RA
)

Write-Verbose "Using ESIFile: $ESIFile"
#Imports file specified in $ESIFile to variable $List
$List=Import-Csv $ESIFile -Delimiter ";"

#Returns based on ComputerName
If ($ComputerName -And !$Username -And !$DisplayName) {
    Write-Verbose "ComputerName used for filter"
    $result=$List | Where-Object {$_.COMPUTERNAME -like "$ComputerName"}
    }
#Returns based on UserName
If ($Username -And !$DisplayName -And !$ComputerName) {
    Write-Verbose "UserName used for filter"
    $result=$List | Where-Object {$_.USERNAME -like "$Username"}
    }
#Returns based on DisplayName
If ($DisplayName -And !$ComputerName -And !$Username) {
    Write-Verbose "DisplayName used for filter"
    $result=$List | Where-Object {$_.DISPLAYNAME -like "$DisplayName"}
    }
#Returns all info on record if no parameters are set
If (!$DisplayName -And !$ComputerName -And !$Username) {$result=$List}



#Prints extra AD information to screen
If ($XtraADInfo){
    #combines results from AD query with
    Write-Verbose "XtraADInfo enabled"
    $result = $result | ForEach-Object {
        $tempCombined = $_
        $adresult=get-aduser $_.username -Properties mail,telephonenumber -ErrorAction SilentlyContinue| Select-Object mail,telephonenumber
        $ADpasswordexpiration=Get-ADPasswordExpiration $_.username -ErrorAction SilentlyContinue #Checks if password is expired and includes it in the output
        $tempCombined | add-member -MemberType NoteProperty -name PasswordLastSet -Value $ADpasswordexpiration.PasswordLastSet
        $tempCombined | add-member -MemberType NoteProperty -name PasswordExpired -Value $ADpasswordexpiration.PasswordExpired
        $tempCombined | add-member -MemberType NoteProperty -name mail -Value $adresult.mail
        $tempCombined | add-member -MemberType NoteProperty -name telephonenumber -Value $adresult.telephonenumber
        Write-Output $tempCombined
    }
    
    Write-Output $result

}
#prints to screen
Else{
    if ($RA){
        Write-Verbose "RA Enabled"
        if ($result.count -gt 1)
        {
            Write-Verbose "RA with Gridview chosen with $result.count"
            $result=$result | Out-GridView -PassThru
            If ($result -ne $null)
            {
                Write-Verbose "RA output selected from Gridview"
                msra.exe /offerra $result.COMPUTERNAME
                Write-Output $result
            }
            Else
            {
                Write-Verbose "RA with no users selected from gridview"
                Write-Host "No result selected!"
            }

        }
        else
        {
            If ($result -ne $null)
            {
                Write-Verbose "RA with single user"
                msra.exe /offerra $result.COMPUTERNAME
                Write-Output $result
            }
            Else
            {
                Write-Verbose "RA with no users found"
                Write-Host "No result found!"
            }
        }
    }
    Else
    {
        If ($result -ne $null)
            {
                Write-Verbose "Returning result without RA"
                Write-Output $result
            }
            Else
            {
                Write-Verbose "Returning no result without RA"
                Write-Host "No result found!"
            }
    }
}




#Sets the path for esinfo permanently in your local powershell profile
If ($SetESInfoFilePath -and $ESIFile) {
    #Creates profile file if it does not exist already
    if (!(Test-Path $Profile)){New-Item -Type File -Force $Profile; Add-Content $Profile ""}
    #Removes old entry in profile
    $tempprofile=Get-Content $Profile | ForEach-Object {if ($_ -notlike "`$ESInfoPath*"){$_}}
    $tempprofile | Set-Content $Profile
    #Adds new entry in profile
    Add-Content $Profile "`$ESInfoPath = `"$ESIFile`""
    Write-Verbose "New ESIFile path set to $ESIFile"
}

}


Function Out-Excel{

<#
.SYNOPSIS
Pases information to an excel session
.DESCRIPTION
Pases information to an excel session

Part of ESPS by ES
.EXAMPLE
Get-ADUser johndoe | Out-Excel
Will show user information in excel
#>


#creates temp variable
param($Path = "$env:temp\$(Get-Date -Format yyyyMMddHHmmss).csv")
#Pipes all input to temp variable created
$input | Export-CSV -Path $Path -UseCulture -Encoding UTF8 -NoTypeInformation
#launches excel with the info from the temp variable
Invoke-Item -Path $Path

}


Function Out-Voice {
    <# 
    .SYNOPSIS 
        Used to allow PowerShell to speak to you or sends data to a WAV file for later listening.
    .DESCRIPTION 
        Used to allow PowerShell to speak to you or sends data to a WAV file for later listenin

        Part of ESPS
    .PARAMETER InputObject
        Data that will be spoken or sent to a WAV file.
    .PARAMETER Rate
        Sets the speaking rate
    .PARAMETER Volume
        Sets the output volume
    .PARAMETER ToWavFile
        Append output to a Waveform audio format file in a specified format
    .PARAMETER ListVoices
       Prints a list of all available voices
    .PARAMETER Voice
       Input name of desired voice, requires " at each side of name
    .EXAMPLE
        "This is a test" | Out-Voice
 
        Description
        -----------
        Speaks the string that was given to the function in the pipeline.
 
    .EXAMPLE
        "Today's date is $((get-date).toshortdatestring())" | Out-Voice
 
        Description
        -----------
        Says todays date
 
     .EXAMPLE
        "Today's date is $((get-date).toshortdatestring())" | Out-Voice -ToWavFile "C:\temp\test.wav"
 
        Description
        -----------
        Says todays date

    #>
 
    [cmdletbinding(
        )]
    Param (
        [parameter(ValueFromPipeline='True')]
        [string[]]$InputObject,
        [parameter()]
        [ValidateRange(-10,10)]
        [Int]$Rate,
        [parameter()]
        [ValidateRange(1,100)]
        $Volume,
        [parameter()]
        [string]$ToWavFile,
        [switch]$ListVoices,
        [string]$Voice
        )
    Begin {
        $Script:parameter = $PSBoundParameters
        Write-Verbose "Listing parameters being used"
        $PSBoundParameters.GetEnumerator() | ForEach {
            Write-Verbose "$($_)"
        }
        Write-Verbose "Loading assemblies"
        Add-Type -AssemblyName System.speech
        Write-Verbose "Creating Speech object"
        $speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
        Write-Verbose "Setting volume"
        If ($PSBoundParameters['Volume']) {
            $speak.Volume = $PSBoundParameters['Volume']
        } Else {
            Write-Verbose "No volume given, using default: 100"
            $speak.Volume = 100
        }
        Write-Verbose "Setting speech rate"
        If ($PSBoundParameters['Rate']) {
            $speak.Rate = $PSBoundParameters['Rate']
        } Else {
            Write-Verbose "No rate given, using default: -2"
            $speak.rate = -2
        }
        If ($PSBoundParameters['ToWavFile']) {
            Write-Verbose "Saving speech to wavfile: $Towavfile"
            $speak.SetOutputToWaveFile("$Towavfile")
        }
        If ($ListVoices){
            Write-Verbose "Listing voices available"
            $speak.GetInstalledVoices().VoiceInfo | Select-Object Name,Gender,Age,Culture
        }
        If ($Voice){
            Write-Verbose "Voice $Voice selected"
            $speak.SelectVoice($Voice)
        }
    }
    Process {
        If ($PSBoundParameters['ToWavFile']){
            ForEach ($line in $inputobject) {
                Write-Verbose "Speaking: $line"       
                $Speak.Speak(($line | Out-String)) | Out-Null
            }
        }
        Else{
            ForEach ($line in $inputobject) {
                Write-Verbose "Speaking: $line"       
                $Speak.SpeakAsync(($line | Out-String)) | Out-Null
            }
        }
    }
    End {
        If ($PSBoundParameters['ToWavFile']) {
            Write-Verbose "Performing cleanup"
            $speak.dispose()
        }
    }
}


Function Remove-Software{

<#
.SYNOPSIS
Uninstalls software from a machine
.DESCRIPTION
Allows you to select from all installed software on a machine and uninstall what you want

Note: Might take a while to complete!

Part of ESPS by ES
.PARAMETER Computername
Computername of machine to perform operations on
.PARAMETER GPOReinstall
Also removes gpo software installation registry key
.PARAMETER GPOReinstallOnly
Only removes gpo software installation registry key
.PARAMETER ListOnly
Only lists installed software
.EXAMPLE
Remove-Software -ComputerName pc01.domain.local
Will return a Gridview of all installed software in which you can select what you want to remove
#>


Param(
[string]$ComputerName = "localhost",
[Parameter(Mandatory=$false,Position=0)]
[switch]$ListOnly,
[switch]$GPOReinstall,
[switch]$GPOReinstallOnly
)

#Default option
If (!$Name -and !$Listonly -and !$GPOReinstallOnly)
{
    Write-verbose "No application name given; returning full list"

    $SelectedSW = 
    Invoke-Command -Computername $Computername -ScriptBlock {
        If (Test-Path ('HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) #Checks for 64bit os
        {
            Write-verbose "64bit OS detected"
            $tempREG=Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | Where-Object {$_.DisplayName -ne $null} | select DisplayName,Publisher,Version
        }
        $tempREG+=Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*  | Where-Object {$_.DisplayName -ne $null} | select DisplayName,Publisher,Version
        $tempREG | Write-Output
    } | Out-GridView -PassThru


    $SelectedSW | Foreach-object{
        write-host "Removing" $_.DisplayName "From Computer" $computername"..."
        $Result=$_.Displayname
        $Result+=Invoke-Command -Computername $Computername -ArgumentList $_.DisplayName -ScriptBlock {param($Displayname); (Get-WmiObject -Class win32_product -Filter "Name like '$DisplayName'").Uninstall() | Select PSComputerName,ReturnValue | Write-Output }
        $Result

        If ($GPOReinstall)
        {
            Invoke-command -Computername $Computername -ArgumentList $_.DisplayName -ScriptBlock {
                param($Displayname)
                If (Test-Path ('hklm:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\AppMgmt\*'))
                {
                    Write-verbose "64Bit OS Detected for GPOreinstall"
                    $tempreg=Get-Item "hklm:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\AppMgmt\*"
                }
                Else
                {
                    Write-verbose "32Bit OS Detected for GPOreinstall"
                    $tempreg=Get-Item "hklm:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\AppMgmt\*"
                }
                $tempreg | ForEach-Object {$_ | Get-ItemProperty | where {$_."Deployment Name" -like "*$Displayname*"} | Remove-Item -Recurse}
            }
            Invoke-GPUpdate -Computer $ComputerName -Force -Target computer
            Write-host "RegistryCheck for gporeinstall Successfull;" $_.DisplayName "will be reinstalled on next reboot"
        }
        "Uninstallation of " + $_.DisplayName + "completed" | Write-Output
        "Uninstallation of " + $_.DisplayName + "completed" | Out-Voice
    }#end foreach-object
}


#GPOReinstallOnly
If ($GPOReinstallOnly -and !$Listonly)
{
    Write-Verbose "GporeinstallOnly selected"
    If(!$Name)
    {
        Write-Verbose "name not entered, displaying list"
        $SelectedSW=Invoke-command -Computername $Computername -ScriptBlock {
            Write-verbose "Connected to remote computer for regisrty key search"
            If (Test-Path ('hklm:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\AppMgmt\*'))
            {
                Write-verbose "64Bit OS Detected for GPOreinstall"
                $tempreg=Get-Item "hklm:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Group Policy\AppMgmt\*" 
            }
            Else
            {
                Write-verbose "32Bit OS Detected for GPOreinstall"
                $tempreg=Get-Item "hklm:\SOFTWARE\Microsoft\Windows\CurrentVersion\Group Policy\AppMgmt\*"
            }
            $tempreg | Write-Output
        } | Out-GridView -PassThru
        
        Invoke-command -Computername $Computername -ArgumentList $SelectedSW -ScriptBlock {
            param($SelectedSW)
            write-verbose "Connected to remote computer for regisrty key removal"
            $SelectedSW | Remove-Item -Recurse
        }
        Invoke-GPUpdate -Computer $ComputerName -Force -Target computer
    }
}


#ListOnly
If ($Listonly -and !$GPOReinstallOnly -and !$Name)
{
    Write-verbose "No application name given; returning full list"
    $SelectedSW = 
    Invoke-Command -Computername $Computername -ScriptBlock {
        If (Test-Path ('HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall')) #Checks for 64bit os
        {
            Write-verbose "64bit OS detected"
            $tempREG=Get-ItemProperty -Path HKLM:\SOFTWARE\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*  | Where-Object {$_.DisplayName -ne $null} | select DisplayName,Publisher,Version
        }
        $tempREG+=Get-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\*  | Where-Object {$_.DisplayName -ne $null} | select DisplayName,Publisher,Version
        $tempREG | Write-Output
    }
    $SelectedSW | Write-Output
}


}

Function Out-HTML{

<#
.SYNOPSIS
Creates html for table from pipline input
.DESCRIPTION
Creates html for table from pipline input
Also provides a lot of customization options

Part of ESPS by ES
.PARAMETER input
piplineinput in table format
.PARAMETER Title
Defines the title of the page
.PARAMETER BackGroundColor
Defines the BackGround Color for the entire page (default lightblue)
Format: text eg. lightblue, yellow, purple
.PARAMETER TableBackGroundColor
Defines the Table BackGround Color (default white)
Format: text eg. lightblue, yellow, purple
.PARAMETER TableBorderWidth
Defines the Table Border Width in pixels (default 1)
.PARAMETER TableBorderColor
Defines the Table Border Color in pixels (default black)
Format: text eg. lightblue, yellow, purple
.PARAMETER TablePadding
Defines the Table Padding in pixels (default 3)
.PARAMETER TableOuterBorderColor
Defines the Table OuterBorder Color (default black)
Format: text eg. lightblue, yellow, purple
.PARAMETER TableOuterBorderWidth
Defines the Table OuterBorder Width in pixels (default 1)
.PARAMETER TableHeaderBorderWidth
Defines the Table Header Border Width in pixels (default 2)
.PARAMETER TableHeaderPadding
Defines the Table Header Padding in pixels (default 3)
.PARAMETER TableHeaderBorderColor
Defines the Table Header Border Color (default black)
Format: text eg. lightblue, yellow, purple
.PARAMETER TableHeaderBackGroundColor
Defines the Table Header BackGround Color (default Gainsboro)
Format: text eg. lightblue, yellow, purple
.PARAMETER Font
Defines the font type (default verdana)
.PARAMETER Fontsize
Defines the fontsize (default x-large)
Can be defined with html names or numbers
.PARAMETER DisableTimeStamp
Disables the timestamp printed on the title line
.PARAMETER RefreshInterval
Defines the auorefreshinterval for the page in seconds (default 60)
.EXAMPLE
Get-ADuser -filter * | Out-html -RefreshInterval 30 -Font Helvetica > c:\list.htm

#>

Param(
[parameter(ValueFromPipeline='True',Mandatory=$true,Position=0)]
$input,
[string]$Title,
[string]$BackGroundColor = "lightblue",
[string]$TableBackGroundColor = "white",
[int]$TableBorderWidth = "1",
[string]$TableBorderColor = "black",
[string]$TablePadding = "3",
[string]$TableOuterBorderColor = "black",
[int]$TableOuterBorderWidth = "2",
[int]$TableHeaderBorderWidth = "1",
[int]$TableHeaderPadding = "3",
[string]$TableHeaderBorderColor = "black",
[string]$TableHeaderBackGroundColor = "Gainsboro",
[string]$Font = "verdana",
[string]$Fontsize = "x-large",
[int]$RefreshInterval = "60",
[switch]$DisableTimeStamp
)

#Check if timestamp is disabled
If (!$DisableTimeStamp){$timedate = Get-Date -format yyyyMMdd_HHmm}


#Styling options
$html = "<style>"
$html = $html + "BODY{background-color:$BackGroundColor;font-size:$Fontsize;font-family:$Font;}"
$html = $html + "TABLE{border-width: $TableOuterBorderWidth`px;border-style: solid;border-color: $TableOuterBorderColor;border-collapse: collapse;}"
$html = $html + "TH{border-width: $TableHeaderBorderWidth`px;padding: $TableHeaderPadding`px;border-style: solid;border-color:$TableHeaderBorderColor;background-color:$TableHeaderBackGroundColor}"
$html = $html + "TD{border-width: $TableBorderWidth`px;padding: $TablePadding`px;border-style: solid;border-color: $TableBorderColor;background-color:$TableBackGroundColor}"
$html = $html + "</style>"
$html = $html + "<META HTTP-EQUIV=`"refresh`" CONTENT=`"$RefreshInterval`">"


#Creates and formats the html file
$html=$input | ConvertTo-HTML -head $html -body "<H2>$Title $timedate</H2>"

$html | Write-output
}


Function Start-Progress {
<#
.SYNOPSIS
Create progress indicator for scriptblock
.DESCRIPTION
Create progress indicator for scriptblock

Part of ESPS by ES
#>
  param(
    [ScriptBlock]
    $code
  )
  
  $newPowerShell = [PowerShell]::Create().AddScript($code)
  $handle = $newPowerShell.BeginInvoke()
  
  while ($handle.IsCompleted -eq $false) {
    Write-Host '.' -NoNewline
    Start-Sleep -Milliseconds 500
    Write-Host '!' -NoNewline
  }
  
  Write-Host ''
  
  $newPowerShell.EndInvoke($handle)
  
  $newPowerShell.Runspace.Close()
  $newPowerShell.Dispose()
}

Function ForEach-Parallel {
<#
.SYNOPSIS
allows running of foreach in paralell and allows for limitation of threads
.DESCRIPTION
allows running of foreach in paralell and allows for limitation of threads

Part of ESPS by ES
.PARAMETER MaxThreads
Number of paralell threads allowed
#>

    param(
        [Parameter(Mandatory=$true,position=0)]
        [System.Management.Automation.ScriptBlock] $ScriptBlock,
        [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
        [PSObject]$InputObject,
        [Parameter(Mandatory=$false)]
        [int]$MaxThreads=5
    )
    BEGIN {
        $iss = [system.management.automation.runspaces.initialsessionstate]::CreateDefault()
        $pool = [Runspacefactory]::CreateRunspacePool(1, $maxthreads, $iss, $host)
        $pool.open()
        $threads = @()
        $ScriptBlock = $ExecutionContext.InvokeCommand.NewScriptBlock("param(`$_)`r`n" + $Scriptblock.ToString())
    }
    PROCESS {
        $powershell = [powershell]::Create().addscript($scriptblock).addargument($InputObject)
        $powershell.runspacepool=$pool
        $threads+= @{
            instance = $powershell
            handle = $powershell.begininvoke()
        }
    }
    END {
        $notdone = $true
        while ($notdone) {
            $notdone = $false
            for ($i=0; $i -lt $threads.count; $i++) {
                $thread = $threads[$i]
                if ($thread) {
                    if ($thread.handle.iscompleted) {
                        $thread.instance.endinvoke($thread.handle)
                        $thread.instance.dispose()
                        $threads[$i] = $null
                    }
                    else {
                        $notdone = $true
                    }
                }
            }
        }
    }
}

Function Get-ADPasswordExpiration{
<#
.SYNOPSIS
Fetches password expiry for a user
.DESCRIPTION
Fetches password expiry for a user

Part of ESPS by ES
.PARAMETER Username
Username of user to check
.PARAMETER Nextday
allows to specifiy amount of days to check if password have expired for
.EXAMPLE
Get-ADPasswordExpiration -Username johndoe
Will return data for user johndoe
.EXAMPLE
Get-ESI -DisplayName ole* | Get-ADPasswordExpiration
Can also be combined with pipeline input from Get-esi
#>
Param(
[parameter(ValueFromPipeline='True',Mandatory=$true,ValueFromPipelineByPropertyName)]
[String]$Username,
[Int]$NextDay
)

    Process
    {
    
    $MaxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge.TotalDays 
    $PasswordBeginDate = (Get-Date).AddDays(-$MaxPasswordAge) 
    $PasswordExpriyDate = (Get-date).AddDays(-($MaxPasswordAge-$NextDay)) 
  
    $ADPasswordInfo = Get-ADUser -ErrorAction SilentlyContinue -Filter {Enabled -eq $true -and SamAccountName -eq $username}` -Properties PasswordNeverExpires,PasswordLastSet,PasswordExpired ` | Select-Object SamAccountName,@{Expression={$($_.PasswordNeverExpires -eq $false) ` -and $($_.PasswordLastSet -ge $PasswordBeginDate.Date) -and ` $($_.PasswordLastSet -le $PasswordExpriyDate.Date)};Label="ExpiredOnNext($NextDay)Day"},` PasswordLastSet,PasswordExpired 
     
    $ADPasswordInfo 
    }
}


Function New-Password{

<#
.SYNOPSIS
Creates a string for use as password and puts it in clipboard
.DESCRIPTION
Creates a string for use as password, default lenght is 8 char
Part of ESPS by ES
.EXAMPLE
new-password
.EXAMPLE
new-password 5
.EXAMPLE
new-password -length 10 -NoSymbol -Noclip 
#>


Param(
[Parameter(Position=0)]
[int]$length = 8,
[switch]$NoUcase,
[switch]$NoLcase,
[switch]$NoNumbers,
[switch]$NoSymbol,
[switch]$NoClip
)
$numbers=0..9
$Ucase=$NULL;For ($a=65;$a –le 90;$a++) {$Ucase+=,[char][byte]$a }
$Lcase=$NULL;For ($a=97;$a –le 122;$a++) {$Lcase+=,[char][byte]$a }
$symbol=".",",","-","!","?","=","#"

$allowedtypes="numbers","Ucase","Lcase","symbol"

If ($NoUcase) { $allowedtypes= $allowedtypes | Where-Object {$_ -ne "Ucase"} }
If ($NoLcase) { $allowedtypes= $allowedtypes | Where-Object {$_ -ne "Lcase"} }
If ($NoNumbers) { $allowedtypes= $allowedtypes | Where-Object {$_ -ne "Numbers"} }
If ($NoSymbol) { $allowedtypes= $allowedtypes | Where-Object {$_ -ne "symbol"} }

If (!$NoUcase) { $pass= "$pass$($Ucase | Get-Random)"}
If (!$NoLcase) {$pass= "$pass$($Lcase | Get-Random)"}
If (!$NoNumbers) {$pass= "$pass$($numbers | Get-Random)"}
If (!$NoSymbol) {$pass= "$pass$($symbol| Get-Random)"}

If ($allowedtypes){
    while ($pass.length -ne $length){
    
      
        
        $type= $allowedtypes | Get-Random
        Switch ($type){
            Ucase{$pass= "$pass$($Ucase | Get-Random)"}
            Lcase{$pass= "$pass$($Lcase | Get-Random)"}
            Numbers{$pass= "$pass$($numbers | Get-Random)"}
            symbol{$pass= "$pass$($symbol| Get-Random)"}
        
        }


    }
}
else{Write-host "ERROR: You must allow at least one charactertype"}

Write-Output $pass

If (!$Noclip){ $pass | CLIP}
}



Function Get-Windowskey{

<#
.SYNOPSIS
Gets serialkey information and displays it on screen
.DESCRIPTION
Gets serialkey information and displays it on screen, can be used on remote or local computer. Credentials can pe provided with-credential switch
Part of ESPS by ES
.EXAMPLE
get-windowskey
.EXAMPLE
get-windowskey -computername pc01.domain.local
.EXAMPLE
get-windowskey -computername pc01.domain.local -credential (get-credential)
#>


Param(
[Parameter(Position=0)]
$computername,
$credential
)


$scriptblock={
    ## function to retrieve the Windows Product Key from any PC
    ## modified to output to file with computername
    param ($targets = ".")
    $hklm = 2147483650
    $regPath = "Software\Microsoft\Windows NT\CurrentVersion"
    $regValue = "DigitalProductId"
    Foreach ($target in $targets) {
        $productKey = $null
        $win32os = $null
        $wmi = [WMIClass]"\\$target\root\default:stdRegProv"
        $data = $wmi.GetBinaryValue($hklm,$regPath,$regValue)
        $binArray = ($data.uValue)[52..66]
        $charsArray = "B","C","D","F","G","H","J","K","M","P","Q","R","T","V","W","X","Y","2","3","4","6","7","8","9"
        ## decrypt base24 encoded binary data
        For ($i = 24; $i -ge 0; $i--) {
            $k = 0
            For ($j = 14; $j -ge 0; $j--) {
                $k = $k * 256 -bxor $binArray[$j]
                $binArray[$j] = [math]::truncate($k / 24)
                $k = $k % 24
            }
            $productKey = $charsArray[$k] + $productKey
            If (($i % 5 -eq 0) -and ($i -ne 0)) {
                $productKey = "-" + $productKey
            }
        }
        $win32os = Get-WmiObject Win32_OperatingSystem -computer $target
        $obj = New-Object Object
        $obj | Add-Member Noteproperty Computer -value $env:computername
        $obj | Add-Member Noteproperty Caption -value $win32os.Caption
        $obj | Add-Member Noteproperty CSDVersion -value $win32os.CSDVersion
        $obj | Add-Member Noteproperty OSArch -value $win32os.OSArchitecture
        $obj | Add-Member Noteproperty BuildNumber -value $win32os.BuildNumber
        $obj | Add-Member Noteproperty RegisteredTo -value $win32os.RegisteredUser
        $obj | Add-Member Noteproperty ProductID -value $win32os.SerialNumber
        $obj | Add-Member Noteproperty ProductKey -value $productkey
        $obj
    }
}


If ($computername){ Invoke-Command -ComputerName $computername -ScriptBlock $scriptblock }
elseif ($computername -and $credential){ Invoke-Command -ComputerName $computername -ScriptBlock $scriptblock -Credential $credential}
else {& $scriptblock}

}
