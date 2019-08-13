# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------

<#
.SYNOPSIS
    This script provides functionality to refresh Windows 10 Media with Dynamic Updates and adding additional Language Packs and Feature On Demands Offline.

.Description
    Before run the script, please download all dynamic update packages, Feature On Demand ISO, Language Pack ISO, and place them in directory $packagesPath before continue.
        Download latest SSU package, e.g. windows10.0-kb4499728-x64.msu, put it in $packagesPath/SSU/windows10.0-kb4493510-x64.msu
        Download latest LCU package, e.g. windows10.0-kb4497934-x64.msu, put it in $packagesPath/LCU/windows10.0-kb4497934-x64.msu
        Download latest SafeOS package, e.g Windows10.0-KB4499728-x64.cab, put it in $packagesPath/SafeOS/Windows10.0-KB4499728-x64.cab
        Download latest SetupDU package, e.g. Windows10.0-KB4499543-x64.cab, put it in $packagesPath/SetupDU/Windows10.0-KB4499728-x64.cab
        Download OEM Feature On Demand ISO, e.g. 17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso,
                and put it in $packagesPath/FOD/17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso
        Download OEM Language Pack ISO, e.g. 17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso,
                and put it in $packagesPath/LP/17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso

.PARAMETER media
    Specifies the location of media that needs to be refreshed.

.PARAMETER index
    Specifies the edition/index that need to build at last. We will only refresh this edition/index with DU contents, FoDs and LangPacks.

.PARAMETER packagesPath
    Specifies the location of all DUs, FoDs and LPs. This folder should contain below optional subdirectories:
        SSU/
        LCU/
        SafeOS/
        SetupDU/
        LP/
        FOD/

.PARAMETER target
    Specifies the location to store the refreshed media. This directory will also be used as working directory,
    so initially script will check to make sure this directory is empty and there are enough space for working.

.PARAMETER langList
    Specifies which languages need to be installed, not support Language Interface Packs (LIPs) for this version. This will add:
        1) Language Packs for Main OS
        2) Recovery Language for WinRE.
        3) Language Packs for Windows Setup. You can use .PARAMETER winSetupLang to choose add or not for Windows Setup, default is not.

.PARAMETER winSetupLang
    Specifies whether to install language packages for Windows Setup

.PARAMETER capabilityList
    Specifies which capabilities need to be installed.

.PARAMETER wimSize
    Specifies maximum size in MB for each of the split .swm files to be created. If install.wim does not exceed this size, won't split.
    Default value is 32000MB.

.PARAMETER cleanupImage
    Specifies whether you need to cleanup image components. Usually it takes about 1 hour to clean all .wim images. But this will also make the image small.

.PARAMETER logPath
    Specifies the location of log file.

.Example
    .\RefreshMedia.ps1 -media "E:\MediaRefreshTest\old media" -index 1 -packagesPath "E:\MediaRefreshTest\packages" -target "E:\MediaRefreshTest\new media"

.Example
    .\RefreshMedia.ps1 -media "E:\MediaRefreshTest\old media" -index 1 -packagesPath "E:\MediaRefreshTest\packages" -target "E:\MediaRefreshTest\new media"
                       -capabilityList "Tools.DeveloperMode.Core~~~~0.0.1.0"

.Example
    .\RefreshMedia.ps1 -media "E:\MediaRefreshTest\old media" -index 1 -packagesPath "E:\MediaRefreshTest\packages" -target "E:\MediaRefreshTest\new media"
                       -capabilityList "Language.Basic~~~fr-FR~0.0.1.0","Language.OCR~~~fr-FR~0.0.1.0" -langList "fr-fr"
#>

#Requires -Version 5.1

Param
(
    [Parameter(Mandatory = $true, HelpMessage = "Specifies the path to original media directory.")]
    [string]$media,

    [Parameter(Mandatory = $true, HelpMessage = "Specifies the Image Index that needs to be refresh.")]
    [ValidateRange(1, 11)][int]$index = 1,

    [Parameter(Mandatory = $true, HelpMessage = "Specifies downloaded packages path")]
    [string]$packagesPath,

    [Parameter(Mandatory = $true, HelpMessage = "Specifies the path to store destination media")]
    [string]$target,

    [Parameter(HelpMessage = "Specifies list of capabilities you want to add")]
    [string[]]$capabilityList,

    [Parameter(HelpMessage = "Specifies list of languages you want to add")]
    [string[]]$langList,

    [Parameter(HelpMessage = "Specifies whether to install language packages for Windows Setup")]
    [switch]$winSetupLang = $false,

    [Parameter(HelpMessage = "Specifies whether to cleanup image components")]
    [switch]$cleanupImage = $true,

    [Parameter(HelpMessage = "Specifies maximum size in MB for each of the split .swm files to be created.")]
    [ValidateRange(1, 32000)]
    [int]$wimSize = 32000,

    [Parameter(HelpMessage = "Specifies the location of log file")]
    [string]$logPath = "$PSScriptRoot\refresh_media.log"
)

[string]$global:logPath = $logPath
[string]$oldMediaPath = $media
[int]$imageIndex = $index
[string]$newMediaPath = $target

. "$PSScriptRoot\common\includeUtilities.ps1" | Out-Null
. "$PSScriptRoot\RefreshMedia\DU.ps1" | Out-Null
. "$PSScriptRoot\RefreshMedia\FoD.ps1" | Out-Null
. "$PSScriptRoot\RefreshMedia\LangPack.ps1" | Out-Null


function CheckFreeSpace {
    [cmdletbinding()]
    param()

    try {
        [int]$mediaSizeMB = [System.Math]::Floor((Get-ChildItem $oldMediaPath -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB)
        $needSizeMB = $mediaSizeMB * 3
    }
    catch {
        $needSizeMB = 15000
    }

    try {
        $currentDrive = (Get-Item $newMediaPath).PSDrive.Name + ":"
        [int]$freeSizeMB = [System.Math]::Floor((Get-WmiObject -Class Win32_Volume -Filter "DriveLetter = '$currentDrive'").FreeSpace / 1MB)
    }
    catch {
        Out-Log "Failed to get current disk free space." $([Constants]::LOG_ERROR)
        return $False
    }

    Out-log "Current disk $currentDrive free space is: $freeSizeMB MB. Needed working disk space is: $needSizeMB MB"

    if ($needSizeMB -gt $freeSizeMB) {
        Out-log "No enough disk space. Please specify another disk for Target store." -level $([Constants]::LOG_ERROR)
        return $False
    }
    else {
        return $True
    }
}


function CheckParameters {
    [cmdletbinding()]
    param()

    $origInstallWimPath = Join-Path $oldMediaPath $([Constants]::INSTALL_WIM_PATH)

    # Check essential folders exist
    if ( !(Test-FolderExist $oldMediaPath) ) {
        Out-Log "$oldMediaPath does not exist." $([Constants]::LOG_ERROR)
        return $False
    }

    if ( !(Test-FolderExist $newMediaPath) ) {
        Out-Log "$newMediaPath does not exist." $([Constants]::LOG_ERROR)
        return $False
    }
    else {
        if ( !(Test-FolderEmpty $newMediaPath) ) {
            Out-Log "$newMediaPath is not empty. Please specify an empty directory." $([Constants]::LOG_ERROR)
            return $False
        }
    }

    if ( !(Test-FolderExist $packagesPath) ) {
        Out-Log "$packagesPath does not exist." $([Constants]::LOG_ERROR)
        return $False
    }

    # Check FOD related
    $FODISODir = Join-Path $packagesPath $([Constants]::FOD_DIR)

    if ( !(Test-FolderExist $FODISODir) ) {
        Out-Log "$FODISODir does not exist." -level $([Constants]::LOG_DEBUG)
        if ( $capabilityList.count -gt 0 ) {
            Out-Log "Cannot add capability when you don't have FoD ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
    }
    else {
        $fullNames = (Get-ChildItem -Path $FODISODir).FullName
        if ($fullNames.count -eq 0) {
            Out-Log "No FoD ISO were found." -level $([Constants]::LOG_DEBUG)
            if ( $capabilityList.count -gt 0 ) {
                Out-Log "Cannot add capability when you don't have FoD ISO." -level $([Constants]::LOG_ERROR)
                return $False
            }
        }
        elseif ($fullNames.count -ge 2) {
            Out-Log "More than one FoD ISO were found in $FODISODir, please only include one ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
        else {
            # todo: need to check capability input?
            ;
        }
    }

    # Check LangPack related
    $LPISODir = Join-Path $packagesPath $([Constants]::LP_DIR)

    if ( !(Test-FolderExist $LPISODir) ) {
        Out-Log "$LPISODir does not exist." -level $([Constants]::LOG_DEBUG)
        if ( $langList.count -gt 0 ) {
            Out-Log "Cannot add Language Pack when you don't have Language Pack ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
    }
    else {
        $fullNames = (Get-ChildItem -Path $LPISODir).FullName
        if ($fullNames.count -eq 0) {
            Out-Log "No Language Pack ISO were found." -level $([Constants]::LOG_DEBUG)
            if ( $langList.count -gt 0 ) {
                Out-Log "Cannot add Language Pack when you don't have Language Pack ISO." -level $([Constants]::LOG_ERROR)
                return $False
            }
        }
        elseif ($fullNames.count -ge 2) {
            Out-Log "More than one Language Pack ISO were found in $LPISODir, please only include one ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
        else {
            $arch = Get-Architecture($origInstallWimPath)
            if ( $arch -eq $([Constants]::ARCH_UNKNOWN) ) {
                Out-Log "Image $imagePath architecture is: $arch" -level $([Constants]::LOG_ERROR)
                return $False
            }
            $isLangInputValid = (Test-LangInput $arch $fullNames $langList)
            if ( !$isLangInputValid ) { return $False }
        }
    }

    # Check Image Index
    try {
        $imageName = (Get-ImageName $origInstallWimPath $imageIndex)
        Out-Log "Image Name is: $imageName"
    }
    catch {
        Out-Log "Image $origInstallWimPath Index is error. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        return $False
    }

    return (CheckFreeSpace)
}


function GetWinREFromInstallWim {
    [cmdletbinding()]
    param([string]$installWimPath)

    Out-Log "Export winre.wim from install.wim"
    $installMountPoint = Join-Path $workingPath $([Constants]::INSTALL_MOUNT)
    $origWinREPath = Join-Path $installMountPoint "windows\system32\recovery\winre.wim"
    $dstWinREPath = Join-Path $workingPath "winre.wim"

    try {
        Mount-Image $installWimPath 1 $installMountPoint

        Copy-Files $origWinREPath $dstWinREPath
        Dismount-DiscardImage $installMountPoint

        return $dstWinREPath
    }
    catch {
        Out-Log "Failed to get winre.wim from $installWimPath. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

        try {
            Dismount-DiscardImage $installMountPoint
            Remove-File $dstWinREPath
        }
        catch {
            # Here is try best to do cleanup
            Out-Log "Failed to cleanup. $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        }

        return $null
    }

}


function CleanAndAssembleMedia {
    [cmdletbinding()]
    param([string]$installWimPath,
        [string]$winREPath,
        [string]$bootWimPath)

    # Do Cleanup Component in winre.wim, boot.wim and install.wim
    # Do Export winre.wim, boot.wim and install.wim
    # Reason why do export: Export can remove unnecessary resource files, help cleanup Wim Image a bit

    Out-Log "Cleanup and assemble media "
    $installMountPoint = Join-Path $workingPath $([Constants]::INSTALL_MOUNT)
    $winREMountPoint = Join-Path $workingPath $([Constants]::WINRE_MOUNT)
    $winPEMountPoint = Join-Path $workingPath $([Constants]::WINPE_MOUNT)

    try {
        # Cleanup Winre.wim
        Out-Log "Cleanup winre.wim " -level $([Constants]::LOG_DEBUG)
        Mount-Image $winREPath 1 $winREMountPoint
        if ($cleanupImage) {
            Restore-Image $winREMountPoint # cleanup install.wim here
        }
        Dismount-CommitImage $winREMountPoint

        $newTmpWinREPath = Join-Path $workingPath "winre2.wim"
        Export-Image $winREPath 1 $newTmpWinREPath
        Move-File $newTmpWinREPath $winREPath

        # Copy Winre.wim and cleanup install.wim
        Out-Log "Cleanup install.wim " -level $([Constants]::LOG_DEBUG)
        Mount-Image $installWimPath $imageIndex $installMountPoint
        Copy-Files $WinREPath "$installMountPoint\windows\system32\recovery\winre.wim"
        if ($cleanupImage) {
            Restore-Image $installMountPoint # cleanup install.wim here
        }
        Dismount-CommitImage $installMountPoint

        $newTmpInstallWimPath = Join-Path $newMediaPath "sources/install2.wim"
        Export-Image $installWimPath $imageIndex $newTmpInstallWimPath
        Move-File $newTmpInstallWimPath $installWimPath

        # Export boot.wim
        Out-Log "Cleanup boot.wim " -level $([Constants]::LOG_DEBUG)
        For ($index = 1; $index -le 2; $index++) {
            Mount-Image $bootWimPath $index $winPEMountPoint
            if ($cleanupImage) {
                Restore-Image $winPEMountPoint
            }
            Dismount-CommitImage $winPEMountPoint
        }

        $newTmpBootWimPath = Join-Path $newMediaPath "sources/boot2.wim"
        For ($index = 1; $index -le 2; $index++) {
            Export-Image $bootWimPath $index $newTmpBootWimPath
        }
        Move-File $newTmpBootWimPath $bootWimPath

        return $True
    }
    catch {
        Out-Log "Failed to cleanup and assemble media. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        try {
            Dismount-DiscardImage $winREMountPoint
            Dismount-DiscardImage $installMountPoint
            Dismount-DiscardImage $winPEMountPoint
        }
        catch {
            # Here is try best to do cleanup
            Out-Log "Failed to cleanup. $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        }

        return $False
    }
}


function SplitInstallWim {
    [cmdletbinding()]
    param([string]$installWimPath,
        [int]$maxInstallSize)

    try {
        $dstWimPath = $installWimPath.substring(0, $installWimPath.length - 4) + ".swm"

        if ( (Get-Item $installWimPath).length -gt "$($maxInstallSize)MB" ) {
            Out-Log "Begin to split install.wim"
            Split-Image $installWimPath $dstWimPath $maxInstallSize
            Remove-Item $installWimPath
        }
    }
    catch {
        # won't treat this as fatal error
        Out-Log "Failed to split install.wim. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
    }

}

function CleanUpWhenSuccess {
    [cmdletbinding()]
    param()

    try {
        Remove-Folder $workingPath
    }
    catch {
        $err = $_.Exception.Message
        Out-Log "Failed to cleanup $workingPath (likely bug of script if happens). Detail: $err" -level $([Constants]::LOG_ERROR)
    }
}


function CleanupWhenFail {
    [cmdletbinding()]
    param()

    try {
        Clear-Folder $newMediaPath
    }
    catch {
        $err = $_.Exception.Message
        Out-Log "Failed to cleanup $newMediaPath (likely bug of script if happens). Detail: $err" -level $([Constants]::LOG_ERROR)
    }
}


function Main {
    [cmdletbinding()]
    param()

    Out-Log "Enter RefreshMedia execute" -level $([Constants]::LOG_DEBUG)
    [string]$script:workingPath = Join-Path $newMediaPath $([Constants]::WORKING_DIR)

    $ok = (CheckParameters)
    if ($ok -eq $False) { return }

    $origInstallWimPath = Join-Path $oldMediaPath $([Constants]::INSTALL_WIM_PATH)
    [string]$script:arch = Get-Architecture($origInstallWimPath)
    if ( $arch -eq $([Constants]::ARCH_UNKNOWN) ) {
        Out-Log "Image $imagePath architecture is: $arch" -level $([Constants]::LOG_ERROR)
        return
    }
    else {
        Out-Log "Get media architecture is: $arch"
    }

    try {
        # Setup working directory
        Add-Folder $workingPath

        # Copy old Media files into new Media
        Copy-Files $oldMediaPath\* $newMediaPath
    }
    catch {
        Out-Log $_.Exception.Message -level $([Constants]::LOG_ERROR)
        CleanupWhenFail
        return
    }

    $winREPath = (GetWinREFromInstallWim $origInstallWimPath)
    if ( !$winREPath) { return }

    $bootWimPath = Join-Path $newMediaPath $([Constants]::BOOT_WIM_PATH)
    $dstInstallWimPath = Join-Path $newMediaPath $([Constants]::INSTALL_WIM_PATH)

    $patchDUInstance = New-Object PatchDU -ArgumentList $dstInstallWimPath,
    $imageIndex,
    $bootWimPath,
    $winREPath,
    $workingPath,
    $packagesPath,
    $newMediaPath

    $patchFODInstance = New-Object PatchFOD -ArgumentList $dstInstallWimPath,
    $imageIndex,
    $bootWimPath,
    $winREPath,
    $workingPath,
    $packagesPath,
    $newMediaPath,
    $capabilityList

    $patchLPInstance = New-Object PatchLP -ArgumentList $dstInstallWimPath,
    $imageIndex,
    $bootWimPath,
    $winREPath,
    $workingPath,
    $packagesPath,
    $newMediaPath,
    $langList,
    $arch,
    $winSetupLang

    if ( !($patchDUInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !($patchFODInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !($patchLPInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !(CleanAndAssembleMedia $dstInstallWimPath $winREPath $bootWimPath) ) { CleanupWhenFail; return }
    SplitInstallWim $dstInstallWimPath $wimSize

    Out-Log "Refresh Media Success!"
    CleanUpWhenSuccess
    return
}

Main
