###########################################################
#
#	Script Name:  
#	Version: 1.0
#	Author: kittuk
#	Date: 	08/05/2021 08:46:00
#
#	Description: 
#
###########################################################


#region---------------Start HeWinre--------------------#
Clear-Host

##################Find Start Up Vairables##########################################
$startupVariables=""
new-variable -force -name startupVariables -value ( Get-Variable |
   % { $_.Name } )
###################################################################################
function Get-ScriptDirectory {
    if ($psise) {
        Split-Path $psise.CurrentFile.FullPath
    }
    else {
        $global:PSScriptRoot
    }
}
$scriptPath = Get-ScriptDirectory

#Check PowerShell Version and Exit if less than 3
if ($PSVersionTable.PSVersion.Major -le 3){[Console]::WriteLine("Error please upgrade Powershell");Exit 1}

function Rotate-Log{ 
    #function checks to see if file in question is larger than the paramater specified if it is it will roll a log and delete the oldes log if there are more than x logs. 
    param([string]$fileName, [int64]$filesize = 1mb , [int] $logcount = 5) 
     
    $logRollStatus = $true 
    if(test-path $filename) 
    { 
        $file = Get-ChildItem $filename 
        if((($file).length) -ige $filesize) #this starts the log roll 
        { 
            $fileDir = $file.Directory 
            $fn = $file.name #this gets the name of the file we started with 
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            $filefullname = $file.fullname #this gets the fullname of the file we started with 
            #$logcount +=1 #add one to the count as the base file is one more than the count 
            for ($i = ($files.count); $i -gt 0; $i--) 
            {  
                #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of the file we are on 
                $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
                $operatingFile = $files | ?{($_.name).trim($fn) -eq $i} 
                if ($operatingfile) 
                 {$operatingFilenumber = ($files | ?{($_.name).trim($fn) -eq $i}).name.trim($fn)} 
                else 
                {$operatingFilenumber = $null} 
 
                if(($operatingFilenumber -eq $null) -and ($i -ne 1) -and ($i -lt $logcount)) 
                { 
                    $operatingFilenumber = $i
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    Write-MYLog -LogStatus "INFO" -LogMessage "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force 
                } 
                elseif($i -ge $logcount) 
                { 
                    if($operatingFilenumber -eq $null) 
                    {  
                        $operatingFilenumber = $i - 1 
                        $operatingFile = $files | ?{($_.name).trim($fn) -eq $operatingFilenumber} 
                        
                    } 
                    Write-MYLog -LogStatus "INFO" -LogMessage "deleting " ($operatingFile.FullName) 
                    remove-item ($operatingFile.FullName) -Force 
                } 
                elseif($i -eq 1) 
                { 
                    $operatingFilenumber = 1 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    Write-MYLog -LogStatus "INFO" -LogMessage "moving to $newfilename" 
                    move-item $filefullname -Destination $newfilename -Force 
                } 
                else 
                { 
                    $operatingFilenumber = $i +1  
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i-1)} 
                    Write-MYLog -LogStatus "INFO" -LogMessage "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force    
                } 
                     
            } 
 
                     
          } 
         else 
         { $logRollStatus = $false} 
    } 
    else 
    { 
        $logrollStatus = $false 
    } 
    Write-MYLog -LogStatus "INFO" -Logmessage "Rotate Logs = $logRollStatus" 
} 

 #Function to write to a custom log
function Write-MYLog {
     Param ([string]$LogStatus,[string]$Logmessage)
     #$shortDate = (Get-Date -Format (Get-culture).DateTimeFormat.ShortDatePattern) -replace "/",""
     $Global:logFile = ".\logs\WinGet-Offline.log"
     if (!(Test-Path ".\logs")){New-Item -ItemType Directory ".\Logs"}
     if (!(Test-Path $logfile)){New-Item -ItemType File $logfile }
     $logtime = (Get-Date -Format (Get-culture).DateTimeFormat.FullDateTimePattern)
     $logData = $LogStatus + "," + $pid+ "," + $LogTime + "," + $Logmessage
     Add-content $Logfile -value $LogData

    #if (!(test-path ` HKLM:\SYSTEM\CurrentControlSet\Services\Eventlog\Application\Huntsman )){new-eventlog -logname Huntsman -source Applocker -ErrorAction SilentlyContinue}
    #Write-EventLog -LogName Huntsman -EntryType $EntryType -Source Applocker -ID $eid -Message "$errmessage"
}


Rotate-Log -fileName $logfile -filesize 1mb -logcount 5

#region start
try{
#Install NuGet
if (!(Get-PackageProvider -name "NuGet" -ea 0)){Install-PackageProvider -Name "NuGet" -Force -Confirm:$false -ea 0}

# Install NtObjectManager module
if (!(Get-InstalledModule NtOBjectManager -ea 0)){Install-Module NtObjectManager -Force -Confirm:$false -ea 0}

# Install winget appx
$vclibs = Invoke-WebRequest -Uri "https://store.rg-adguard.net/api/GetFiles" -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body "type=PackageFamilyName&url=Microsoft.VCLibs.140.00_8wekyb3d8bbwe&ring=RP&lang=en-US" -UseBasicParsing | Foreach-Object Links | Where-Object outerHTML -match "Microsoft.VCLibs.140.00_.+_x64__8wekyb3d8bbwe.appx" | Foreach-Object href
$vclibsuwp = Invoke-WebRequest -Uri "https://store.rg-adguard.net/api/GetFiles" -Method "POST" -ContentType "application/x-www-form-urlencoded" -Body "type=PackageFamilyName&url=Microsoft.VCLibs.140.00.UWPDesktop_8wekyb3d8bbwe&ring=RP&lang=en-US" -UseBasicParsing | Foreach-Object Links | Where-Object outerHTML -match "Microsoft.VCLibs.140.00.UWPDesktop_.+_x64__8wekyb3d8bbwe.appx" | Foreach-Object href

Invoke-WebRequest $vclibsuwp -OutFile Microsoft.VCLibs.140.00.UWPDesktop_8wekyb3d8bbwe.appx
Invoke-WebRequest $vclibs -OutFile Microsoft.VCLibs.140.00_8wekyb3d8bbwe.appx

Add-AppxPackage -Path .\Microsoft.VCLibs.140.00.UWPDesktop_8wekyb3d8bbwe.appx
Add-AppxPackage -Path .\Microsoft.VCLibs.140.00_8wekyb3d8bbwe.appx

# install winget bundle
$url = 'https://github.com/microsoft/winget-cli/releases/latest'
$request = [System.Net.WebRequest]::Create($url)
$response = $request.GetResponse()
$realTagUrl = $response.ResponseUri.OriginalString
$version = $realTagUrl.split('/')[-1].Trim('v')
$version
#6.2.3
$fileName = "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle"
$realDownloadUrl = $realTagUrl.Replace('tag', 'download') + '/' + $fileName
$realDownloadUrl
#https://github.com/microsoft/winget-cli/releases/download/v1.0.11692/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle
#https://github.com/PowerShell/PowerShell/releases/download/v6.2.3/PowerShell-6.2.3-win-x64.zip
Invoke-WebRequest -Uri $realDownloadUrl -OutFile $env:TEMP/$fileName

#Invoke-WebRequest https://github.com/microsoft/winget-cli/releases/download/v1.0.11692/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle -OutFile Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle
Add-AppxPackage -Path $env:TEMP\$fileName #Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.appxbundle

# Create reparse point 
$installationPath = (Get-AppxPackage Microsoft.DesktopAppInstaller).InstallLocation
Set-ExecutionAlias -Path "C:\Windows\System32\winget.exe" -PackageName "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe" -EntryPoint "Microsoft.DesktopAppInstaller_8wekyb3d8bbwe!winget" -Target "$installationPath\AppInstallerCLI.exe" -AppType Desktop -Version 3
explorer.exe "shell:appsFolder\Microsoft.DesktopAppInstaller_8wekyb3d8bbwe!winget"
}
Catch{
   #display errors
   $ErrorMessage = $_.Exception.Message
   $FailedItem = $_.Exception.ItemName
   #"We failed $FailedItem. The error message was $ErrorMessage"
   Write-MYLog -LogStatus "Error" -Logmessage "----------------------------------------"
   Write-MYLog -LogStatus "Error" -Logmessage "$FailedItem :- $ErrorMessage"
   Write-MYLog -LogStatus "Error" -Logmessage "----------------------------------------"
}
Finally{
#exit AzureRM
#Set-AzureRmContext -Context ([Microsoft.Azure.Commands.Profile.Models.PSAzureContext]::new())
#Cleanup-Variables
#Start-Sleep $loopevery
#$startupVariables2
new-variable -force -name startupVariables2 -value ( Get-Variable |
   % { $_.Name } )
#function to clean up custom Variables
function Cleanup-Variables {
  Get-Variable |
    Where-Object { $startupVariables -notcontains $_.Name } #|
    % { Remove-Variable -Name "$($_.Name)" -Force -Scope "global" }
}
$myvars = Compare-Object $startupVariables $startupVariables2
foreach ($myvar in $myvars.inputobject){
#$myvar
Clear-Variable $myvar -Scope Global
}
}
#endregion start

#endregion------------End Here----------------------#
