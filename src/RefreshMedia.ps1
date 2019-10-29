# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------

<#
.SYNOPSIS
    This script provides functionality to refresh Windows 10 Media with Dynamic Updates and adding additional Language Packs and Feature On Demand Offline.

.Description
    Before run the script, please download all dynamic update packages, Feature On Demand ISO, Language Pack ISO, and place them in directory .PARAMETER PackagesPath before continue.
        Download latest SSU package, e.g. windows10.0-kb4499728-x64.msu, put it in $PackagesPath/SSU/windows10.0-kb4493510-x64.msu
        Download latest LCU package, e.g. windows10.0-kb4497934-x64.msu, put it in $PackagesPath/LCU/windows10.0-kb4497934-x64.msu
        Download latest SafeOS package, e.g Windows10.0-KB4499728-x64.cab, put it in $PackagesPath/SafeOS/Windows10.0-KB4499728-x64.cab
        Download latest SetupDU package, e.g. Windows10.0-KB4499543-x64.cab, put it in $PackagesPath/SetupDU/Windows10.0-KB4499728-x64.cab
        Download OEM Feature On Demand ISO, e.g. 17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso,
                and put it in $PackagesPath/FOD/17763.1.180914-1434.rs5_release_amd64fre_FOD-PACKAGES_OEM_PT1_amd64fre_MULTI.iso
        Download OEM Language Pack ISO, e.g. 17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso,
                and put it in $PackagesPath/LP/17763.1.180914-1434.rs5_release_CLIENTLANGPACKDVD_OEM_MULTI.iso

.PARAMETER Media
    Specifies the location of media that needs to be refreshed. You can copy RTM Media to local and specify local path,
    or mount RTM Media ISO and specify mount path.

.PARAMETER Index
    Specifies the edition/index that need to build at last. We will only refresh this edition/index with DU contents, FoDs and LPs.

.PARAMETER PackagesPath
    Specifies the location of all DUs, FoDs and LPs. This folder should contain below optional subdirectories:
        SSU/
        LCU/
        SafeOS/
        SetupDU/
        LP/
        FOD/

.PARAMETER Target
    Specifies the location to store the refreshed media. This directory will also be used as working directory,
    so script will check to make sure this directory is empty and there are enough space for working initially.

.PARAMETER LangList
    Specifies which languages need to be installed, not support Language Interface Packs (LIPs) for this version. This will add:
        1) Language Packs for Main OS
        2) Recovery Language for WinRE.
        3) Language Packs for WinPE. You can use .PARAMETER WinPELang to disable it if you don't need add language for WinPE.

.PARAMETER WinPELang
    Specifies whether to install language packages for Windows Preinstallation Environment

.PARAMETER CapabilityList
    Specifies which capabilities need to be installed.

.PARAMETER WimSize
    Specifies maximum size in MB for each of the split .swm files to be created. If install.wim does not exceed this size, won't split.
    Default value is 32000MB.

.PARAMETER CleanupImage
    Specifies whether you need to cleanup .wim image components. It can take about 1 hour or more to clean all .wim images, it will make .wim clean and smaller.

.PARAMETER LogPath
    Specifies the location of log file.

.Example
    .\RefreshMedia.ps1 -Media "E:\MediaRefreshTest\old media" -Index 1 -PackagesPath "E:\MediaRefreshTest\packages" -Target "E:\MediaRefreshTest\new media"

.Example
    .\RefreshMedia.ps1 -Media "E:\MediaRefreshTest\old media" -Index 1 -PackagesPath "E:\MediaRefreshTest\packages" -Target "E:\MediaRefreshTest\new media" `
-CapabilityList "Language.Basic~~~fr-FR~0.0.1.0","Language.OCR~~~fr-FR~0.0.1.0" -LangList "fr-fr"

.Example
    .\RefreshMedia.ps1 -Media "E:\MediaRefreshTest\old media" -Index 1 -PackagesPath "E:\MediaRefreshTest\packages" -Target "E:\MediaRefreshTest\new media" `
-CapabilityList "Language.Basic~~~fr-FR~0.0.1.0","Language.OCR~~~fr-FR~0.0.1.0" -LangList "fr-fr" -WinPELang:$false
#>

#Requires -Version 5.1

Param
(
    [Parameter(Mandatory = $true, HelpMessage = "Specifies the path to original media directory.")]
    [string]$Media,

    [Parameter(Mandatory = $true, HelpMessage = "Specifies the Image Index that needs to be refresh.")]
    [ValidateRange(1, 11)][int]$Index = 1,

    [Parameter(Mandatory = $true, HelpMessage = "Specifies downloaded packages path")]
    [string]$PackagesPath,

    [Parameter(Mandatory = $true, HelpMessage = "Specifies the path to store destination media")]
    [string]$Target,

    [Parameter(HelpMessage = "Specifies list of capabilities you want to add")]
    [string[]]$CapabilityList,

    [Parameter(HelpMessage = "Specifies list of languages you want to add")]
    [string[]]$LangList,

    [Parameter(HelpMessage = "Specifies whether to install language packages for Windows Preinstallation Environment")]
    [switch]$WinPELang = $true,

    [Parameter(HelpMessage = "Specifies whether to cleanup image components")]
    [switch]$CleanupImage = $true,

    [Parameter(HelpMessage = "Specifies maximum size in MB for each of the split .swm files to be created.")]
    [ValidateRange(1, 32000)]
    [int]$WimSize = 32000,

    [Parameter(HelpMessage = "Specifies the location of log file")]
    [string]$LogPath = "$PSScriptRoot\refresh_media.log"
)

[string]$global:LogPath = $LogPath
[string]$OldMediaPath = $Media
[int]$ImageIndex = $Index
[string]$NewMediaPath = $Target

. "$PSScriptRoot\common\includeUtilities.ps1" | Out-Null
. "$PSScriptRoot\RefreshMedia\DU.ps1" | Out-Null
. "$PSScriptRoot\RefreshMedia\FoD.ps1" | Out-Null
. "$PSScriptRoot\RefreshMedia\LangPack.ps1" | Out-Null


function CheckFreeSpace {
    [cmdletbinding()]
    param()

    try {
        [int]$mediaSizeMB = [System.Math]::Floor((Get-ChildItem $OldMediaPath -Recurse | Measure-Object -Property Length -Sum -ErrorAction Stop).Sum / 1MB)
        $needSizeMB = $mediaSizeMB * 3
    }
    catch {
        $needSizeMB = 15000
    }

    try {
        $currentDrive = (Get-Item $NewMediaPath).PSDrive.Name + ":"
        [int]$freeSizeMB = [System.Math]::Floor((Get-WmiObject -Class Win32_Volume -Filter "DriveLetter = '$currentDrive'").FreeSpace / 1MB)
    }
    catch {
        Out-Log "Failed to get current disk free space." -level $([Constants]::LOG_ERROR)
        return $False
    }

    Out-log "Current disk $currentDrive free space is: $freeSizeMB MB. Needed work disk space is: $needSizeMB MB"

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

    $origInstallWimPath = Join-Path $OldMediaPath $([Constants]::INSTALL_WIM_PATH)

    # Check essential folders exist
    if ( !(Test-FolderExist $OldMediaPath) ) {
        Out-Log "$OldMediaPath does not exist." -level $([Constants]::LOG_ERROR)
        return $False
    }

    if ( !(Test-FolderExist $NewMediaPath) ) {
        Out-Log "$NewMediaPath does not exist." -level $([Constants]::LOG_ERROR)
        return $False
    }
    else {
        if ( !(Test-FolderEmpty $NewMediaPath) ) {
            Out-Log "$NewMediaPath is not empty. Please specify an empty directory." -level $([Constants]::LOG_ERROR)
            return $False
        }
    }

    if ( !(Test-FolderExist $PackagesPath) ) {
        Out-Log "$PackagesPath does not exist." -level $([Constants]::LOG_ERROR)
        return $False
    }

    # Check <= 1 SetupDU exist
    $setupDUPath = Join-Path $PackagesPath $([Constants]::SETUPDU_DIR)
    if ( (Test-Path $setupDUPath) -and ((Get-ChildItem -Path $setupDUPath).Count -gt 1) ) {
        Out-Log "Please only place one Setup DU in $setupDUPath." -level $([Constants]::LOG_ERROR)
        return $False
    }

    # Check FOD related
    $fodISODir = Join-Path $PackagesPath $([Constants]::FOD_DIR)

    if ( !(Test-FolderExist $fodISODir) ) {
        Out-Log "$fodISODir does not exist." -level $([Constants]::LOG_DEBUG)
        if ( $CapabilityList.count -gt 0 ) {
            Out-Log "Cannot add capability when you don't have FoD ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
    }
    else {
        $fullNames = (Get-ChildItem -Path $fodISODir).FullName
        if ($fullNames.count -eq 0) {
            Out-Log "No FoD ISO were found." -level $([Constants]::LOG_DEBUG)
            if ( $CapabilityList.count -gt 0 ) {
                Out-Log "Cannot add capability when you don't have FoD ISO." -level $([Constants]::LOG_ERROR)
                return $False
            }
        }
        elseif ($fullNames.count -ge 2) {
            Out-Log "More than one FoD ISO were found in $fodISODir, please only include one ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
        else {
            # check capability input later if necessary
            ;
        }
    }

    # Check LangPack related
    $lpISODir = Join-Path $PackagesPath $([Constants]::LP_DIR)

    if ( !(Test-FolderExist $lpISODir) ) {
        Out-Log "$lpISODir does not exist." -level $([Constants]::LOG_DEBUG)
        if ( $LangList.count -gt 0 ) {
            Out-Log "Cannot add Language Pack when you don't have Language Pack ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
    }
    else {
        $fullNames = (Get-ChildItem -Path $lpISODir).FullName
        if ($fullNames.count -eq 0) {
            Out-Log "No Language Pack ISO were found." -level $([Constants]::LOG_DEBUG)
            if ( $LangList.count -gt 0 ) {
                Out-Log "Cannot add Language Pack when you don't have Language Pack ISO." -level $([Constants]::LOG_ERROR)
                return $False
            }
        }
        elseif ($fullNames.count -ge 2) {
            Out-Log "More than one Language Pack ISO were found in $lpISODir, please only include one ISO." -level $([Constants]::LOG_ERROR)
            return $False
        }
        else {
            $arch = Get-Architecture($origInstallWimPath)
            if ( $arch -eq $([Constants]::ARCH_UNKNOWN) ) {
                Out-Log "Image $origInstallWimPath architecture is: $arch" -level $([Constants]::LOG_ERROR)
                return $False
            }
            $isLangInputValid = (Test-LangInput $arch $fullNames $LangList)
            if ( !$isLangInputValid ) { return $False }
        }
    }

    # Check Image Index
    try {
        $imageName = (Get-ImageName $origInstallWimPath $ImageIndex)
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
    $installMountPoint = Join-Path $WorkingPath $([Constants]::INSTALL_MOUNT)
    $origWinREPath = Join-Path $installMountPoint "windows\system32\recovery\winre.wim"
    $dstWinREPath = Join-Path $WorkingPath "winre.wim"

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

        return $Null
    }

}


function CleanAndAssembleMedia {
    [cmdletbinding()]
    param([string]$installWimPath,
        [string]$winREPath,
        [string]$bootWimPath)

    <#
    .SYNOPSIS
        Cleanup Component in winre.wim, boot.wim and install.wim
        Export winre.wim, insert this new winre.wim back to install.wim
        Export boot.wim and install.wim
    .DESCRIPTION
        Reason why do Export: Export can remove unnecessary resource files, help cleanup Wim Image a bit
    #>

    Out-Log "Cleanup and assemble media, it can take one hour or more to cleanup"
    $installMountPoint = Join-Path $WorkingPath $([Constants]::INSTALL_MOUNT)
    $winREMountPoint = Join-Path $WorkingPath $([Constants]::WINRE_MOUNT)
    $winPEMountPoint = Join-Path $WorkingPath $([Constants]::WINPE_MOUNT)

    try {
        # Cleanup Winre.wim
        Out-Log "Cleanup winre.wim " -level $([Constants]::LOG_DEBUG)
        Mount-Image $winREPath 1 $winREMountPoint
        if ($CleanupImage) {
            Restore-Image $winREMountPoint # cleanup winre.wim here
        }
        Dismount-CommitImage $winREMountPoint

        $newTmpWinREPath = Join-Path $WorkingPath "winre2.wim"
        Export-Image $winREPath 1 $newTmpWinREPath
        Move-File $newTmpWinREPath $winREPath

        # Copy Winre.wim and cleanup install.wim
        Out-Log "Cleanup install.wim " -level $([Constants]::LOG_DEBUG)
        Mount-Image $installWimPath $ImageIndex $installMountPoint
        Copy-Files $WinREPath "$installMountPoint\windows\system32\recovery\winre.wim"
        if ($CleanupImage) {
            Restore-Image $installMountPoint # cleanup install.wim here
        }
        Dismount-CommitImage $installMountPoint

        $newTmpInstallWimPath = Join-Path $NewMediaPath "sources/install2.wim"
        Export-Image $installWimPath $ImageIndex $newTmpInstallWimPath
        Move-File $newTmpInstallWimPath $installWimPath

        # Export boot.wim
        Out-Log "Cleanup boot.wim " -level $([Constants]::LOG_DEBUG)
        $imageNumber = Get-ImageTotalEdition $bootWimPath
        For ($index = 1; $index -le $imageNumber; $index++) {
            Mount-Image $bootWimPath $index $winPEMountPoint
            if ($CleanupImage) {
                Restore-Image $winPEMountPoint # cleanup boot.wim here
            }
            Dismount-CommitImage $winPEMountPoint
        }

        $newTmpBootWimPath = Join-Path $NewMediaPath "sources/boot2.wim"
        For ($index = 1; $index -le $imageNumber; $index++) {
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

function RemoveFilesReadOnlyAttr {
    [cmdletbinding()]
    param([string]$path)

    Get-ChildItem -Path $path -Recurse | Where-Object { -not $_.PSIsContainer -and $_.IsReadOnly } |
    ForEach-Object {
        $_.IsReadOnly = $false
    }
}

function CleanUpWhenSuccess {
    [cmdletbinding()]
    param()

    try {
        Remove-Folder $WorkingPath
    }
    catch {
        $err = $_.Exception.Message
        Out-Log "Failed to cleanup $WorkingPath (likely bug of script if happens). Detail: $err" -level $([Constants]::LOG_ERROR)
    }
}


function CleanupWhenFail {
    [cmdletbinding()]
    param()

    try {
        Clear-Folder $NewMediaPath
    }
    catch {
        $err = $_.Exception.Message
        Out-Log "Failed to cleanup $NewMediaPath (likely bug of script if happens). Detail: $err" -level $([Constants]::LOG_ERROR)
    }
}


function Main {
    [cmdletbinding()]
    param()

    Out-Log "Enter RefreshMedia execute" -level $([Constants]::LOG_DEBUG)
    [string]$script:WorkingPath = Join-Path $NewMediaPath $([Constants]::WORKING_DIR)

    $ok = (CheckParameters)
    if ($ok -eq $False) { return }

    $origInstallWimPath = Join-Path $OldMediaPath $([Constants]::INSTALL_WIM_PATH)
    [string]$script:arch = Get-Architecture($origInstallWimPath)
    if ( $arch -eq $([Constants]::ARCH_UNKNOWN) ) {
        Out-Log "Image $origInstallWimPath architecture is: $arch" -level $([Constants]::LOG_ERROR)
        return
    }
    else {
        Out-Log "Get media architecture is: $arch"
    }

    $bootWimPath = Join-Path $NewMediaPath $([Constants]::BOOT_WIM_PATH)
    $dstInstallWimPath = Join-Path $NewMediaPath $([Constants]::INSTALL_WIM_PATH)

    try {
        # Setup working directory
        Add-Folder $WorkingPath

        # Copy old Media files into new Media
        Out-Log "Copy media from $OldMediaPath to $NewMediaPath, this might take a while if copy from ISO."
        Copy-Files $OldMediaPath\* $NewMediaPath

        # Check and remove media files read-only attribute
        RemoveFilesReadOnlyAttr $NewMediaPath
    }
    catch {
        Out-Log $_.Exception.Message -level $([Constants]::LOG_ERROR)
        CleanupWhenFail
        return
    }

    $winREPath = (GetWinREFromInstallWim $dstInstallWimPath)
    if ( !$winREPath) { return }

    $patchSSUInstance = New-Object PatchSSU -ArgumentList $dstInstallWimPath,
    $ImageIndex,
    $bootWimPath,
    $winREPath,
    $WorkingPath,
    $PackagesPath,
    $NewMediaPath

    $patchLPInstance = New-Object PatchLP -ArgumentList $dstInstallWimPath,
    $ImageIndex,
    $bootWimPath,
    $winREPath,
    $WorkingPath,
    $PackagesPath,
    $NewMediaPath,
    $LangList,
    $arch,
    $winPELang

    $patchFODInstance = New-Object PatchFOD -ArgumentList $dstInstallWimPath,
    $ImageIndex,
    $bootWimPath,
    $winREPath,
    $WorkingPath,
    $PackagesPath,
    $NewMediaPath,
    $CapabilityList

    $patchDUExcludeSSUInstance = New-Object PatchDUExcludeSSU -ArgumentList $dstInstallWimPath,
    $ImageIndex,
    $bootWimPath,
    $winREPath,
    $WorkingPath,
    $PackagesPath,
    $NewMediaPath

    if ( !($patchSSUInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !($patchLPInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !($patchFODInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !($patchDUExcludeSSUInstance.DoPatch())) { CleanupWhenFail; return }
    if ( !(CleanAndAssembleMedia $dstInstallWimPath $winREPath $bootWimPath) ) { CleanupWhenFail; return }
    SplitInstallWim $dstInstallWimPath $wimSize

    Out-Log "Refresh Media Success!"
    CleanUpWhenSuccess
    return
}

Main
