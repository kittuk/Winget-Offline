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


#region---------------Start Here--------------------#
Clear-Host

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

#endregion------------End Here----------------------#



