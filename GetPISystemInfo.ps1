<#
William Ehrich
2/5/2016

This script collects diagnostic information for machines hosting PI Interfaces, PI Data Archive Servers, PI AF Server, PI Coresight Servers, and PI Clients
and saves the files to a folders on the desktop called 'PISystemInfo' and 'PISystemInfo.zip'

By default, most servers are configured to not allow PowerShell scripts. To run this script, you will first need to allow scripts using:

Set-Executionpolicy Unrestricted

Then you can run the following script:
.\GetPISystemInfo.ps1

To specify a time period for the PI Message Logs and Event Logs, input the start and end times in PI time format as arguments.

For example,

.\GetPISystemInfo.ps1 *-8h *

will return the the PI Message Logs and Event Logs for the most recent 8 hours.

Please direct questions or comments regarding this script to wehrich@osisoft.com

#>

param(
        [Parameter(Mandatory=$false)][string]$inputStartTime,
        [Parameter(Mandatory=$false)][string]$inputEndTime
    )

    $defaultStartTime = "*-2h"
    $defaultEndTime = "*"

    # if there are two arguments
    if($inputEndTime){
        $startTime = $inputStartTime
        $endTime = $inputEndTime
        #write-host "start time = $startTime and end time = $endTime"
    }
    else{
        $startTime = $defaultStartTime
        $endTime = $defaultEndTime
        #write-host "start time = $startTime and end time = $endTime"
    }

$startDTM = (Get-Date)
$4HoursAgo = $startDTM.AddHours(-4)
Set-Location $env:USERPROFILE\Desktop
$loc = Get-Location

Write-Host "`nThis script collects PI System diagnostic information `nand saves the files to a folder on the desktop called 'PISystemInfo'"
Write-Host "`nCreating Output Directory: "$loc"\PISystemInfo"
New-Item -ItemType directory -Path $loc\PISystemInfo | Out-Null
Set-Location $env:USERPROFILE\Desktop\PISystemInfo
$outputDirectory = Get-Location
New-Item -ItemType directory -Path $outputDirectory\EventLogs,$outputDirectory\PILogs,$outputDirectory\PIBuffering,$outputDirectory\Windows,$outputDirectory\Script | Out-Null

if(Test-Path $env:PISERVER\adm\pidiag.exe)
{
Set-Location $env:PISERVER\adm
# start time
.\pidiag.exe -t $startTime > $outputDirectory\Script\StartTimeOutput.txt
$startTimeFirstLine = Get-Content $outputDirectory\Script\StartTimeOutput.txt -First 1
$startTimeExtractedStringTime = $StartTimeFirstLine.Substring(0,18)
#Write-Host "$startTimeExtractedStringTime= extractedStringTime"
$startTimeDatePS = [DateTime]$startTimeExtractedStringTime
#Write-Host "$startTimeDatePS= startTimeDatePS"
#end time
.\pidiag.exe -t $endTime > $outputDirectory\Script\EndTimeOutput.txt
$endTimeFirstLine = Get-Content $outputDirectory\Script\EndTimeOutput.txt -First 1
$endTimeExtractedStringTime = $endTimeFirstLine.Substring(0,18)
#Write-Host "$endTimeExtractedStringTime= extractedStringTime"
$endTimeDatePS = [DateTime]$endTimeExtractedStringTime
#Write-Host "$endTimeDatePS= startTimeDatePS"
}
elseif(Test-Path $env:PIHOME\adm\pidiag.exe)
{
Set-Location $env:PIHOME\adm
# start time
.\pidiag.exe -t $startTime > $outputDirectory\Script\StartTimeOutput.txt
$startTimeFirstLine = Get-Content $outputDirectory\Script\StartTimeOutput.txt -First 1
$startTimeExtractedStringTime = $StartTimeFirstLine.Substring(0,18)
#Write-Host "$startTimeExtractedStringTime= extractedStringTime"
$startTimeDatePS = [DateTime]$startTimeExtractedStringTime
#Write-Host "$startTimeDatePS= startTimeDatePS"
#end time
.\pidiag.exe -t $endTime > $outputDirectory\Script\EndTimeOutput.txt
$endTimeFirstLine = Get-Content $outputDirectory\Script\EndTimeOutput.txt -First 1
$endTimeExtractedStringTime = $endTimeFirstLine.Substring(0,18)
#Write-Host "$endTimeExtractedStringTime= extractedStringTime"
$endTimeDatePS = [DateTime]$endTimeExtractedStringTime
#Write-Host "$endTimeDatePS= startTimeDatePS"
}
else
{
$startTimeDatePS = Get-Date
$startTimeDatePS = $startTimeDatePS.AddHours(-2)
#Write-Host "startTimeDatePS = $startTimeDatePS"
$endTimeDatePS = Get-Date
#Write-Host "endTimeDatePS = $endTimeDatePS"
}
#Write-Host "startTimeDatePS = $startTimeDatePS"
#Write-Host "endTimeDatePS = $endTimeDatePS"

function getPIPCLog
{
Write-Host "`n(1/11): Obtaining PIPC Log..."
Copy-Item $env:PIHOME\dat\pipc.log $outputDirectory\PILogs
}

function getPIMessageLog
{
Write-Host "`n(2/11): Obtaining PI Message Log from $startTimeDatePS to $endTimeDatePS..."
If (Test-Path $env:PISERVER\adm\pigetmsg.exe)
{set-Location $env:PISERVER\adm
.\pigetmsg -st $startTime -et $endTime > $outputDirectory\PILogs\PIMessageLog.txt}
If (Test-Path $env:PIHOME\adm\pigetmsg.exe)
{set-Location $env:PIHOME\adm
.\pigetmsg -st $startTime -et $endTime > $outputDirectory\PILogs\PIMessageLog.txt}
}

function getPIAFInformation
{
If(Test-Path $env:PIHOME64\AF\AFService.exe)
{Write-Host "`n(3/11): Obtaining PI AF Event Log from $startTimeDatePS to $endTimeDatePS"
New-Item -ItemType directory -Path $outputDirectory\PIAF | Out-Null
Set-Location $env:PIHOME64\AF
& "$env:PIHOME64\AF\afdiag.exe" > $outputDirectory\PIAF\AFdiagOutput.txt
$AFEventLogEvents = Get-WinEvent -ProviderName AF | Where-Object {$_.TimeCreated -gt $startTimeDatePS -and $_.TimeCreated -lt $endTimeDatePS}
For($m = 1; $m -le $AFEventLogEvents.count; $m++)
{ Write-Progress -Activity "Gathering AF Event Log from $startTimeDatePS to $endTimeDatePS" -status "$m Events Obtained" -percentComplete ($m / $AFEventLogEvents.count*100) -ParentId 1}
$AFEventLogEvents | Export-Csv $outputDirectory\PIAF\PIAFEventLog.csv}
else
{Write-Host "`n(3/11): PI AF Server not detected..."}
}

function getPICoresightInformation
{
If(Test-Path $env:PIHOME64\Coresight)
{Write-Host "`n(4/11): Obtaining PI Coresight Event Logs from $startTimeDatePS to $endTimeDatePS..."
New-Item -ItemType directory -Path $outputDirectory\PICoresight | Out-Null
# Obtain OSIsoft-PIDataServices Log
$OSIsoftPIDataServicesLogEvents = Get-WinEvent -ProviderName OSIsoft-PIDataServices | Where-Object {$_.TimeCreated -gt $startTimeDatePS -and $_.TimeCreated -lt $endTimeDatePS}
For($n = 1; $n -le $OSIsoftPIDataServicesLogEvents.count; $n++)
{ Write-Progress -Activity "OSIsoft-PIDataServices from $startTimeDatePS to $endTimeDatePS" -status "Found Event $n" -percentComplete ($n / $OSIsoftPIDataServicesLogEvents.count*100) -ParentId 1}
$OSIsoftPIDataServicesLogEvents | Export-Csv $outputDirectory\PICoresight\OSIsoft-PIDataServices.csv
# Obtain OSIsoft-PISymbols Log
$OSIsoftPISymbolsEvents = Get-WinEvent -ProviderName OSIsoft-PISymbols | Where-Object {$_.TimeCreated -gt $startTimeDatePS -and $_.TimeCreated -lt $endTimeDatePS}
For($o = 1; $o -le $OSIsoftPISymbolsEvents.count; $o++)
{ Write-Progress -Activity "OSIsoft-PISymbols from $startTimeDatePS to $endTimeDatePS" -status "Found Event $o" -percentComplete ($o / $OSIsoftPISymbolsEvents.count*100) -ParentId 1}
$OSIsoftPISymbolsEvents | Export-Csv $outputDirectory\PICoresight\OSIsoft-PISymbols.csv
# Obtain OSIsoft-PISystemSearch Log
$OSIsoftPISystemSearchEvents = Get-WinEvent -ProviderName OSIsoft-PISystemSearch -MaxEvents 300 | Where-Object {$_.TimeCreated -gt $startTimeDatePS -and $_.TimeCreated -lt $endTimeDatePS} 
For($p = 1; $p -le $OSIsoftPISystemSearchEvents.count; $p++)
{ Write-Progress -Activity "OSIsoft-PISystemSearch from $startTimeDatePS to $endTimeDatePS" -status "Found Event $p" -percentComplete ($p / $OSIsoftPISystemSearchEvents.count*100) -ParentId 1}
$OSIsoftPISystemSearchEvents | Export-Csv $outputDirectory\PICoresight\OSIsoft-PISystemSearch.csv
# Obtain OSIsoft-Search Log
$OSIsoftSearchEvents = Get-WinEvent -ProviderName OSIsoft-Search | Where-Object {$_.TimeCreated -gt $startTimeDatePS -and $_.TimeCreated -lt $endTimeDatePS}
For($q = 1; $q -le $OSIsoftSearchEvents.count; $q++)
{ Write-Progress -Activity "OSIsoft-Search from $startTimeDatePS to $endTimeDatePS" -status "Found Event $q" -percentComplete ($q / $OSIsoftSearchEvents.count*100) -ParentId 1} 
$OSIsoftSearchEvents | Export-Csv $outputDirectory\PICoresight\OSIsoft-Search.csv
# Obtain PIWebAPI Log
$PIWebAPIEvents = Get-WinEvent -ProviderName PIWebAPI | Where-Object {$_.TimeCreated -gt $startTimeDatePS -and $_.TimeCreated -lt $endTimeDatePS}
For($r = 1; $r -le $PIWebAPIEvents.count; $r++)
{ Write-Progress -Activity "PIWebAPI from $startTimeDatePS to $endTimeDatePS" -status "Found Event $r" -percentComplete ($r / $PIWebAPIEvents.count*100) -ParentId 1}
$PIWebAPIEvents | Export-Csv $outputDirectory\PICoresight\PIWebAPI.csv}
else
{Write-Host "`n(4/11): PI Coresight not detected..."}
}

function getPIInterfacesInformation
{
If(Test-Path $env:PIHOME\Interfaces)
{Write-Host "`n(5/11): Obtaining Batch Files..."
New-Item -ItemType directory -Path $outputDirectory\PIInterfaceBatchFiles | Out-Null
Set-Location $env:PIHOME\Interfaces
get-childitem -recurse -filter *.bat | Copy-Item -Destination $outputDirectory\PIInterfaceBatchFiles}
else
{Write-Host "`n(5/11): PI Interfaces not detected..."}
}

function getPIBufferingInformation
{
Write-Host "`n(6/11): Obtaining Buffering Information..."
# If pibufss.exe is located in the 32-bit directory
If (Test-Path $env:PIHOME\bin\pibufss.exe)
{set-Location $env:PIHOME\bin
cmd /C "pibufss -cfg > %userprofile%\Desktop\PISystemInfo\PIBuffering\pibufss-cfg_output.txt"}
# If pibufss.exe is located in the 64-bit directory
If (Test-Path $env:PIHOME64\bin\pibufss.exe)
{set-Location $env:PIHOME64\bin
cmd /C "pibufss -cfg > %userprofile%\Desktop\PISystemInfo\PIBuffering\pibufss-cfg_output.txt"}
Copy-Item $env:PIHOME\dat\piclient.ini $outputDirectory\PIBuffering
# The following lines can be uncommented to obtain the results of bufreport.bat for interface nodes that have PI Buffer Subystem 4.3 and later installed
<#
set-Location $env:PIHOME64\bin
bufreport.bat > $outputDirectory\PIBuffering\output.txt
Set-Location $env:userprofile\appdata\local\temp
get-childitem -recurse -filter bufreport* | Copy-Item -Destination $outputDirectory\PIBuffering
#>
}

function getOperatingSystemInformation
{
Write-Host "`n(7/11): Obtaining Operating System Information..."
#Get-WmiObject -Class Win32_OperatingSystem -ComputerName . | Select-Object -Property BuildNumber,BuildType,OSType,Version,ServicePackMajorVersion,ServicePackMinorVersion | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\Windows\OperatingSystem.txt
#Add-Content $outputDirectory\Windows\OperatingSystem.txt "Operating System Version Table" 
#Add-Content $outputDirectory\Windows\OperatingSystem.txt "https://msdn.microsoft.com/en-us/library/windows/desktop/ms724832(v=vs.85).aspx"
Copy-Item -Path $env:SystemRoot\System32\drivers\etc\hosts -Destination $env:USERPROFILE\desktop\PISystemInfo\Windows\HostsFile.txt
cmd /C "systeminfo > %userprofile%\Desktop\PISystemInfo\Windows\OperatingSystem.txt"
}

function getApplicationLogInformation
{
Write-Host "`n(8/11): Obtaining Application Event Log from $startTimeDatePS to $endTimeDatePS..."
#Get-EventLog -After $8HoursAgo -Before $startDTM -LogName Application | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\EventLogs\ApplicationEventLog.txt
$applicationLogEvents = Get-EventLog -After $startTimeDatePS -Before $endTimeDatePS -LogName Application
For($k = 1; $k -le $applicationLogEvents.count; $k++)
{ Write-Progress -Activity "Gathering Application Log from $startTimeDatePS to $endTimeDatePS" -status "$k Events Obtained" -percentComplete ($k / $applicationLogEvents.count*100) -ParentId 1}
$applicationLogEvents | Export-Csv $outputDirectory\EventLogs\ApplicationEventLog.csv
}

function getSystemLogInformation
{
Write-Host "`n(9/11): Obtaining System Event Log from $startTimeDatePS to $endTimeDatePS..."
#Get-EventLog -After $8HoursAgo -Before $startDTM -LogName System | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\EventLogs\SystemEventLog.txt
$systemLogEvents = Get-EventLog -After $startTimeDatePS -Before $endTimeDatePS -LogName System
For($l = 1; $l -le $systemLogEvents.count; $l++)
{ Write-Progress -Activity "Gathering System Log from $startTimeDatePS to $endTimeDatePS" -status "$l Events Obtained" -percentComplete ($l / $systemLogEvents.count*100) -ParentId 1}
$systemLogEvents | Export-Csv $outputDirectory\EventLogs\SystemEventLog.csv
}

function getServicesInformation
{
Write-Host "`n(10/11): Obtaining List of Services..."
#Get-WmiObject -Class Win32_Service | select DisplayName, StartName, State | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\Windows\Services.txt
$wmiQuery = "Select Name, DisplayName, StartName, State from win32_service"
$colItems = Get-WmiObject -Query $wmiQuery
For($i = 1; $i -le $colItems.count; $i++)
{ Write-Progress -Activity "Gathering Services" -status "$i Services Found" -percentComplete ($i / $colItems.count*100) -ParentId 1}
$colItems | Select DisplayName, StartName, State, Name | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\Windows\Services.txt
}

function getInstalledProgramsInformation
{
Write-Host "`n(11/11): Obtaining List of Installed Programs (Make Take a Couple of Minutes)..."
#Get-WMIObject -Class Win32_Product | select Name, Vendor, Version, Caption, IdentifyingNumber | Sort-Object Name | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\Windows\InstalledPrograms.txt
$wmiQuery2 = "Select Name, Version from win32_product"
$colItems2 = Get-WmiObject -Query $wmiQuery2
For($j = 1; $j -le $colItems2.count; $j++)
{ Write-Progress -Activity "Gathering Installed Programs" -status "$j Programs Found" -percentComplete ($j / $colItems2.count*100) -ParentId 1}
Write-Progress -Activity "Gathering Installed Programs" -status "$j Programs Found" -percentComplete (100) -ParentId 1 -Completed
$colItems2 | Select Name, Version | Format-Table -autosize -wrap | Out-File -width 1000 -FilePath $outputDirectory\Windows\InstalledPrograms.txt
}

$modulesCompleted = 0
$numModules = 11

Write-Progress -Activity "Obtaining PIPC Log" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getPIPCLog -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining PI Message Log" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getPIMessageLog -Static
$modulesCompleted++

Write-Progress -Activity "Looking for PI AF" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getPIAFInformation -Static
$modulesCompleted++

Write-Progress -Activity "Looking for PI Coresight" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getPICoresightInformation -Static
$modulesCompleted++

Write-Progress -Activity "Looking for PI Interface" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getPIInterfacesInformation -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining Buffering Information" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getPIBufferingInformation -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining Operating System Information" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getOperatingSystemInformation -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining Application Event Log" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getApplicationLogInformation -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining System Event Log" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getSystemLogInformation -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining List of Services" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getServicesInformation -Static
$modulesCompleted++

Write-Progress -Activity "Obtaining List of Installed Programs (May Take a Couple of Minutes)..." -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1
getInstalledProgramsInformation -Static
$modulesCompleted++
Write-Progress -Activity "Obtaining List of Installed Programs" -status "($modulesCompleted/$numModules) modules complete" -percentComplete ($modulesCompleted/$numModules*100) -Id 1 -Completed

# create compressed folder
$sourceCompression = "$env:USERPROFILE\Desktop\PISystemInfo"
$destinationCompression = "$env:USERPROFILE\Desktop\PISystemInfo.zip"
Add-Type -Assembly "system.io.compression.filesystem"
[io.compression.zipfile]::CreateFromDirectory($sourceCompression,$destinationCompression)

Write-Host "`nData Gather Complete! `nThe files should be in folders on the desktop called 'PISystemInfo' and 'PISystemInfo.zip.'"

$endDTM = (Get-Date)
Write-Host "`nExecution Time: $(($endDTM-$startDTM).totalseconds) seconds"

Write-Host "Press any key to exit ..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")