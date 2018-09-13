#-----Script Info---------------------------------------------------------------------------------------------
# Name:UpdateESPS.ps1
# Author:ES
# Date:20.11.2014
# Version: 1
# Job/Tasks: Installs/Updates all Powershell modules from ES
#--------------------------------------------------------------------------------------------------------------

#-----Changelog------------------------------------------------------------------------------------------------
#v1. Script created ES
#v2. ......
#
#
#--------------------------------------------------------------------------------------------------------------


#-----VARIABLES-----------------------------------------------------------------------------------------------
Param(
[switch]$ListModules,
$InstallModule
)





#-----Functions---------------------------------------------------------------------------------------------

#Download modules

Function Download-module {
Param(
[Parameter(Mandatory=$true,Position=0)]
$ModuleName,
[Parameter(Mandatory=$true,Position=1)]
$FileUrl,
$Extrafiles,
$Extrafolders
)

$Mydoc=[environment]::getfolderpath("mydocuments")

#Create top level Folders
if (Test-Path $Mydoc\WindowsPowershell\modules\){}
else {mkdir $Mydoc\WindowsPowershell\modules\}

#Create folder for module
if (Test-Path $Mydoc\WindowsPowershell\modules\$ModuleName\){}
else {mkdir $Mydoc\WindowsPowershell\modules\$ModuleName\}


#Download file
Invoke-WebRequest $FileUrl -OutFile $Mydoc\WindowsPowershell\modules\$ModuleName\$ModuleName.psm1
#ExtraFolders
$Extrafolders
If ($Extrafolders){
    $Extrafolders | ForEach-Object {
        New-Item -ItemType Directory -Force $_
    }
}

#ExtraFiles
If ($Extrafiles){
    $Extrafiles | ForEach-Object {
        Invoke-WebRequest $_ -OutFile $Mydoc\WindowsPowershell\modules\$ModuleName\$($_.Split("{/}")[-1])
    }
}
}



#-----Start Script---------------------------------------------------------------------------------------------



#Available modules
$tempmodulelist=[System.IO.Path]::GetTempFileName()
Invoke-WebRequest https://wiki.ap3x.org/bin/download/Main/ESPS+Install++and+update+powershell+module+collection/Modulelist.json -OutFile $tempmodulelist
$modulelist=Get-Content -Raw $tempmodulelist | ConvertFrom-Json
Remove-Item $tempmodulelist


#Print module list
IF ($ListModules){
    
    $modulelist | ForEach-Object{
        $printmodules=@{}
        If(Test-Path $Mydoc\WindowsPowershell\modules\$($_.modulename)\){$printmodules.installed=$true}
        else{$printmodules.installed=$false}
        $printmodules.modulename=$_.modulename
        $printmodules.desc
    }
    
    #| select modulename,description | Write-Output
}
#If not printing modulelist
Else{

#Download/update main module
Download-module -ModuleName ESPS -FileUrl https://wiki.ap3x.org/bin/download/Main/ESPS+Install++and+update+powershell+module+collection/ESPS.psm1

#Module location
$Mydoc=[environment]::getfolderpath("mydocuments")

#Update modules
If (!$InstallModule){
    $modulelist | ForEach-Object{
        If(Test-Path $Mydoc\WindowsPowershell\modules\$($_.modulename)\){
            If ($_.fileurl){Download-module -ModuleName $_.modulename -FileUrl $_.fileurl -Extrafiles $_.extrafiles -Extrafolders $_.extrafolders}
        }
    }
}  
#Install new modules
If ($installmodule){
    $InstallModule=$InstallModule | ConvertFrom-Csv -Header Modulename
    $modulelist = $modulelist | Where-Object { $_.modulename -in $InstallModule.modulename }
    $modulelist | ForEach-Object{
        If ($_.fileurl){Download-module -ModuleName $_.modulename -FileUrl $_.fileurl -Extrafiles $_.extrafiles -Extrafolders $_.extrafolders}
        If ($_.installscript){ iex (New-Object Net.WebClient).DownloadString("$($_.installscript)")}
    }
}

}


