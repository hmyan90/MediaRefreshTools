# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------

<#
.SYNOPSIS
    PowerShell script for download Dynamic Update(DU) from [Microsoft Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx).

.Description
    The scripts will create six subfolders under .PARAMETER DownloadPath:
	    SSU/
	    LCU/
	    SafeOS/
	    SetupDU/
	    LP/
	    FOD/

    Then the script will download latest SSU into SSU/, latest LCU into LCU/, latest SafeOS into SafeOS/, latest SetupDU into SetupDU/.
    Script won't download Language Packs and Feature On Demands ISO, so you need to download and put them into correspondding subfolder.

.PARAMETER DownloadPath
    Specifies the path to store downloaded packages. This directory should be empty.

.PARAMETER Platform
    Specifies for which platform you need to download dynamic update.

.PARAMETER Product
    Specifies for Windows product you need to download dynamic update. Only support "Windows 10" for now.

.PARAMETER Version
    Specifies for Windows Media version you need to download dynamic update. Default value is "1809".

.PARAMETER DUReleaseMonth
    Specifies month of release of Dynamic Updates, it should be current month if you want to get latest DU. Default value is current month.

.PARAMETER ShowLinksOnly
    Specifies that you only want to see download links of DU, but won't actually download them.

.PARAMETER LogPath
    Specifies the location of log file.

.Example
    .\DownloadContents.ps1 -DownloadPath .\downloads -Version 2019-06 -DUReleaseMonth 1809

#>

#Requires -Version 5.1

Param
(
    [Parameter(Mandatory = $True, HelpMessage = "Specifies the path to store downloaded packages.")]
    [ValidateScript( { Test-Path $_ -PathType 'Container' })]
    [string]$DownloadPath,

    [Parameter(HelpMessage = "Specifies Windows platform")]
    [ValidateSet("x86", "x64", "arm64")]
    [string]$Platform = "x64",

    [Parameter(HelpMessage = "Specifies Windows product")]
    [ValidateSet("windows 10")]
    [string]$Product = "windows 10",

    [Parameter(HelpMessage = "Specifies Windows Media version")]
    [ValidateSet("1809", "1903")]
    [string]$Version = "1809",

    [Parameter(HelpMessage = "Specifies month of release of Dynamic Updates")]
    [ValidateScript( { [DateTime]::ParseExact($_, "yyyy-MM", $null) })]
    [string]$DUReleaseMonth = ("{0:d4}-{1:d2}" -f (Get-Date).Year, (Get-Date).Month),

    [Parameter(HelpMessage = "Show download links only")]
    [switch]$ShowLinksOnly,

    [Parameter(DontShow, HelpMessage = "Download files through HTTPS")]
    [switch]$ForceSSL,

    [Parameter(HelpMessage = "Specifies the location of log file")]
    [string]$LogPath = "$PSScriptRoot\download_content.log")

. "$PSScriptRoot\common\includeUtilities.ps1" | Out-Null
. "$PSScriptRoot\DownloadContents\DU.ps1" | Out-Null

[string]$global:LogPath = $LogPath


function layoutSubfolderForDownloadedPackages {
    [cmdletbinding()]
    Param()

    if ( !(Test-FolderEmpty $DownloadPath) ) {
        Out-Log "$DownloadPath is not empty. Please specify an empty directory." -level $([Constants]::LOG_ERROR)
        return $False
    }

    try {
        # Setup all subfolders
        Add-Folder $SSUPath
        Add-Folder $LCUPath
        Add-Folder $SafeOSPath
        Add-Folder $SetupDUPath
        Add-Folder $FODPath
        Add-Folder $LPPath
        return $True
    }
    catch {
        Out-Log "Failed to create subfolders. $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        Remove-Folder $SSUPath
        Remove-Folder $LCUPath
        Remove-Folder $SafeOSPath
        Remove-Folder $SetupDUPath
        Remove-Folder $FODPath
        Remove-Folder $LPPath
        return $False
    }
}

function CleanupWhenFail {
    [cmdletbinding()]
    Param()

    Remove-Folder $SSUPath
    Remove-Folder $LCUPath
    Remove-Folder $SafeOSPath
    Remove-Folder $SetupDUPath
    Remove-Folder $FODPath
    Remove-Folder $LPPath
}

function CleanUpWhenSuccess {
    [cmdletbinding()]
    param()

    Out-Log "Download Dynamic Update finish."
    return
}

function Main {
    [cmdletbinding()]
    param()

    Out-Log "Enter DownloadContents execute" -level $([Constants]::LOG_DEBUG)
    [string]$script:SSUPath = Join-Path $DownloadPath $([Constants]::SSU_DIR)
    [string]$script:LCUPath = Join-Path $DownloadPath $([Constants]::LCU_DIR)
    [string]$script:SafeOSPath = Join-Path $DownloadPath $([Constants]::SAFEOS_DIR)
    [string]$script:SetupDUPath = Join-Path $DownloadPath $([Constants]::SETUPDU_DIR)
    [string]$script:FODPath = Join-Path $DownloadPath $([Constants]::FOD_DIR)
    [string]$script:LPPath = Join-Path $DownloadPath $([Constants]::LP_DIR)

    $createSubfolder = layoutSubfolderForDownloadedPackages
    if ($createSubfolder -eq $False) { return }

    $downloadDUInstance = New-Object DownloadDU -ArgumentList $SSUPath,
    $LCUPath,
    $SafeOSPath,
    $SetupDUPath,
    $Platform,
    $DUReleaseMonth,
    $Product,
    $Version,
    $ShowLinksOnly,
    $ForceSSL

    if ( !($downloadDUInstance.DoDownload()) ) { CleanupWhenFail; return }

    CleanUpWhenSuccess
    return
}

Main
