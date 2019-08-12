# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------

<#
.SYNOPSIS
    PowerShell script for download Dynamic Update(DU) from [Microsoft®Update Catalog](https://www.catalog.update.microsoft.com/Home.aspx).

.Description
    The scripts will create six sub folders under downloadPath specified in parameter for you:
	    SSU/
	    LCU/
	    SafeOS/
	    SetupDU/
	    LP/
	    FOD/
	
And the script will download latest SSU into SSU/, latest LCU into LCU/, latest SafeOS into SafeOS/, latest SetupDU into SetupDU/.
Script won't download Language Packs and Feature On Demands ISO, so you need to download and put them into correspondding sub folder. 

.PARAMETER downloadPath
    Specifies the path to store downloaded packages. This directory should be empty.

.PARAMETER platform
    Specifies for which platform you need to download dynamic update.

.PARAMETER product
    Specifies for Windows product you need to download dynamic update.

.PARAMETER version
    Specifies for Windows version you need to download dynamic update.

.PARAMETER releaseMonth
    Specifies month of release of Dynamic Updates, it should be current month if you want to get latest DU.

.PARAMETER showLinksOnly
    Specifies you only want to see download links of DU, but won't actually download them.

.PARAMETER logPath
    Specifies the location of log file.

#>

#Requires -Version 5.1

Param
(
    [Parameter(Mandatory = $true, HelpMessage = "Specifies the path to store downloaded packages.")]
    [ValidateScript({Test-Path $_ -PathType 'Container'})]
    [string]$downloadPath,

    [Parameter(HelpMessage = "Specifies Windows platform")]
    [ValidateSet("x86", "x64", "arm64")]
    [string]$platform = "x64",

    [Parameter(HelpMessage = "Specifies Windows product")]
    [ValidateSet("windows 10")]
    [string]$product = "windows 10",

    [Parameter(HelpMessage = "Specifies Windows version")]
    [ValidateSet("1809", "1903")]
    [string]$version = "1809",

    [Parameter(HelpMessage = "Specifies month of release of Dynamic Updates")]
    [ValidateScript({[DateTime]::ParseExact($_, "yyyy-MM", $null)})]
    [string]$releaseMonth = "2019-07",

    [Parameter(HelpMessage = "Show download links only")]
    [switch]$showLinksOnly,

    [Parameter(HelpMessage = "Use https to download")]
    [switch]$forceSSL,

    [Parameter(HelpMessage = "Specifies the location of log file")]
    [string]$logPath = "$PSScriptRoot\download_content.log")

. "$PSScriptRoot\common\includeUtilities.ps1" | Out-Null
. "$PSScriptRoot\DownloadContents\DU.ps1" | Out-Null

[string]$global:logPath = $logPath


function layoutSubfolderForDownloadedPackages {
    [cmdletbinding()]
    Param()

    if ( !(Test-FolderEmpty $downloadPath) ) {
        Out-Log "$downloadPath is not empty. Please specify an empty directory." -level $([Constants]::LOG_ERROR)
        return $False
    }

    try {
        # Setup all sub directory
        Add-Folder $SSUPath
        Add-Folder $LCUPath
        Add-Folder $SafeOSPath
        Add-Folder $SetupDUPath
        Add-Folder $FODPath
        Add-Folder $LPPath
        return $true
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
    [string]$script:SSUPath = Join-Path $downloadPath $([Constants]::SSU_DIR)
    [string]$script:LCUPath = Join-Path $downloadPath $([Constants]::LCU_DIR)
    [string]$script:SafeOSPath = Join-Path $downloadPath $([Constants]::SAFEOS_DIR)
    [string]$script:SetupDUPath = Join-Path $downloadPath $([Constants]::SETUPDU_DIR)
    [string]$script:FODPath = Join-Path $downloadPath $([Constants]::FOD_DIR)
    [string]$script:LPPath = Join-Path $downloadPath $([Constants]::LP_DIR)

    $createSubFolder = layoutSubfolderForDownloadedPackages
    if ($createSubFolder -eq $false) { return }

    $downloadDUInstance = New-Object DownloadDU -ArgumentList $SSUPath,
                                                              $LCUPath,
                                                              $SafeOSPath,
                                                              $SetupDUPath,
                                                              $platform,
                                                              $releaseMonth,
                                                              $product,
                                                              $version,
                                                              $showLinksOnly,
                                                              $forceSSL
                                                              
    if ( !($downloadDUInstance.DoDownload()) ) { CleanupWhenFail; return }

    CleanUpWhenSuccess
    return
}

Main
