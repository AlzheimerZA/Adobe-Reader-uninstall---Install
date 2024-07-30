<#
.COPYRIGHT
Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
See LICENSE in the project root for license information.
#############################################################################
#                                     			 		                    #
#   This Sample Code is provided for the purpose of illustration only       #
#   and is not intended to be used in a production environment.  THIS       #
#   SAMPLE CODE AND ANY RELATED INFORMATION ARE PROVIDED "AS IS" WITHOUT    #
#   WARRANTY OF ANY KIND, EITHER EXPRESSED OR IMPLIED, INCLUDING BUT NOT    #
#   LIMITED TO THE IMPLIED WARRANTIES OF MERCHANTABILITY AND/OR FITNESS     #
#   FOR A PARTICULAR PURPOSE.  We grant You a nonexclusive, royalty-free    #
#   right to use and modify the Sample Code and to reproduce and distribute #
#   the object code form of the Sample Code, provided that You agree:       #
#   (i) to not use Our name, logo, or trademarks to market Your software    #
#   product in which the Sample Code is embedded; (ii) to include a valid   #
#   copyright notice on Your software product in which the Sample Code is   #
#   embedded; and (iii) to indemnify, hold harmless, and defend Us and      #
#   Our suppliers from and against any claims or lawsuits, including        #
#   attorneys' fees, that arise or result from the use or distribution      #
#   of the Sample Code.                                                     #
#                                     			 		                    #
#   Author: John Guy                                                        #
#   Version 1.0         Date Last modified:      30 July 2024        	    #
#                                                                 		    #
#                                     			 		                    #
#############################################################################
#>


# Check if the script is running with administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "Please run this script as an administrator."
    Pause
    exit
}

# Function to write output to a log file
function Write-Log
{
    Param ([string]$LogString)
    $LogFile = "C:\Windows\Logs\RemoveAcrobatReader-$(get-date -f yyyy-MM-dd).log"
    $DateTime = "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
    $LogMessage = "$DateTime $LogString"
    Add-content $LogFile -value $LogMessage
}

# Get installed programs for both 32-bit and 64-bit architectures
$paths = @('HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall\','HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\')

$installedPrograms = foreach ($registryPath in $paths) {
    try {
        Get-ChildItem -LiteralPath $registryPath | Get-ItemProperty | Where-Object { $_.PSChildName -ne $null }
    } catch {
        Write-Log ("Failed to access registry path: $registryPath. Error: $_")
        return @()
    }
}

# Filter programs with Adobe Acrobat Reader in their display name, excluding Standard and Professional and version 24.002.20965
# Change Line 66 if the verison is updates
$adobeReaderEntries = $installedPrograms | Where-Object {
    $_.DisplayName -like '*Adobe Acrobat*' -and
    $_.DisplayName -notlike '*Standard*' -and
    $_.DisplayName -notlike '*Professional*' -and
    $_.DisplayVersion -notlike '*24.002.20965*'
}

if ($adobeReaderEntries.Count -eq 0) {
    Write-Log "No Adobe Acrobat Reader installations found to uninstall."
    Write-Host "No Adobe Reader to Uninstall"
    pause
    exit
}

# Try to uninstall Adobe Acrobat Reader for each matching entry
foreach ($entry in $adobeReaderEntries) {
    $productCode = $entry.PSChildName

    try {
        # Use the MSIExec command to uninstall the product
        Start-Process -FilePath "msiexec.exe" -ArgumentList "/x $productCode /qn" -Wait -PassThru

        Write-Log ("Adobe Acrobat Reader has been successfully uninstalled using product code: $productCode")
        Write-host ("Adobe Acrobat Reader has been successfully uninstalled using product code: $productCode")
    } catch {
        Write-Log ("Failed to uninstall Adobe Acrobat Reader with product code $productCode. Error: $_")
    }
}


#INSTALL THE LATEST VERSION
# Define the latest version URL (Update URL If required) Remove "#" to use URL
#$installerUrl = "https://ardownload2.adobe.com/pub/adobe/reader/win/AcrobatDC/2400220965/AcroRdrDC2400220965_en_US.exe"
$installerPath = "$env:TEMP\AcroRdrDC_Installer.exe"

# Install from file share CHANGE FILE SHARE NAME!!!
$fileSharePath = "\\FILESHARENAME\AcroRdrDC_Installer.exe"

# Create a WebClient object
$webClient = New-Object System.Net.WebClient

# Download the Adobe Reader installer
#Write-Output "Downloading Adobe Reader..."
#$webClient.DownloadFile($installerUrl, $installerPath)

# Copy the installer from the file share to the local path
Write-Output "Copying Adobe Reader installer from file share..."
Copy-Item -Path $fileSharePath -Destination $installerPath -Force

# Install Adobe Reader
Write-Output "Installing Adobe Reader..."
Start-Process -FilePath $installerPath -ArgumentList "/sAll /msi /norestart" -Wait -NoNewWindow

# Clean up
Write-Output "Cleaning up..."
Remove-Item -Path $installerPath

Write-Output "Adobe Reader update complete."