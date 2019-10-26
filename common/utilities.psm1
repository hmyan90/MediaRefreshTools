# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------

Set-StrictMode -Version Latest;

Enum DUType {
    Unknown = 0
    SSU = 1
    LCU = 2
    SafeOS = 3
    SetupDU = 4
}


class Constants {

    # wim relative path
    static [string]$INSTALL_WIM_PATH = "sources\install.wim"
    static [string]$BOOT_WIM_PATH = "sources\boot.wim"

    # mount point
    static [string]$INSTALL_MOUNT = "Installmount"
    static [string]$WINPE_MOUNT = "WinPEmount"
    static [string]$WINRE_MOUNT = "WinREmount"

    # environment name, for debug/show
    static [string]$MAIN_OS = "Main OS"
    static [string]$WINPE = "WinPE"
    static [string]$WINRE = "WinRE"

    # Some path define
    static [string]$SSU_DIR = "SSU"
    static [string]$LCU_DIR = "LCU"
    static [string]$SAFEOS_DIR = "SafeOS"
    static [string]$SETUPDU_DIR = "SetupDU"
    static [string]$FOD_DIR = "FOD"
    static [string]$LP_DIR = "LP"
    static [string]$WORKING_DIR = "working"

    # Image architecture
    static [string]$ARCH_X86 = "x86"
    static [string]$ARCH_ARM = "arm"
    static [string]$ARCH_X64 = "x64"
    static [string]$ARCH_ARM64 = "arm64"
    static [string]$ARCH_UNKNOWN = "Unknown"

    # Log Level
    static [string]$LOG_DEBUG = "debug"
    static [string]$LOG_INFO = "info"
    static [string]$LOG_WARNING = "warning"
    static [string]$LOG_ERROR = "error"

    # install.wim split
    static [string]$INSTALL_SWM_FILE = "install.swm"
}


function Test-FolderExist {
    [cmdletbinding()]
    param([string]$folder)

    return (Test-Path($folder) -PathType Container)
}


function Test-FolderEmpty {
    [cmdletbinding()]
    param([string]$folder)

    if ( !(Test-FolderExist $folder) ) {
        # for safety reason
        return $false
    }

    $info = (Get-ChildItem $folder | Measure-Object)
    return ($info.count -eq 0)
}


function Remove-Folder {
    [cmdletbinding()]
    param([string]$folder)

    if ( !(Test-FolderExist $folder) ) { return }

    Remove-Item -Path $folder -Recurse -Force -ErrorAction stop | Out-Null
}


function Clear-Folder {
    [cmdletbinding()]
    param([string]$folder)

    Get-ChildItem -Path $folder -Include * | remove-Item -recurse | Out-Null
}


function Remove-File {
    [cmdletbinding()]
    param([string]$filePath)

    if ( !(Test-Path $filePath) ) { return }

    Remove-Item -Path $filePath -ErrorAction stop | Out-Null
}


function Move-File {
    [cmdletbinding()]
    param([string]$srcFile,
        [string]$dstFile)

    Move-Item -Path $srcFile -Destination $dstFile -Force -ErrorAction stop | Out-Null
}

function Copy-Files {
    [cmdletbinding()]
    param([string]$srcPath,
        [string]$dstPath)

    Copy-Item -Path $srcPath -Destination $dstPath -Force -Recurse -ErrorAction stop | Out-Null
}


function Add-Folder {
    [cmdletbinding()]
    param([string]$path)

    New-Item -ItemType directory -Path $path -ErrorAction Stop | Out-Null
}


function Mount-Image {
    [cmdletbinding()]
    param([string]$wimPath,
        [int]$index,
        [string]$mountPoint)

    if ( !(Test-Path $mountPoint) ) {
        New-Item -ItemType directory -Path $mountPoint -ErrorAction stop | Out-Null
    }

    Mount-WindowsImage -ImagePath $wimPath -Index $index -Path $mountPoint -ErrorAction stop | Out-Null
}


function Dismount-DiscardImage {
    [cmdletbinding()]
    param([string]$mountPoint)

    if ( Test-FolderEmpty $mountPoint ) {
        return
    }

    Dismount-WindowsImage -Path $mountPoint -Discard -ErrorAction stop | Out-Null
}


function Dismount-CommitImage {
    [cmdletbinding()]
    param([string]$mountPoint)

    Dismount-WindowsImage -Path $mountPoint -Save -ErrorAction stop | Out-Null
}


function Install-Package {
    [cmdletbinding()]
    param([string]$imagePath,
        [string]$packagePath)

    Add-WindowsPackage -Path $imagePath -PackagePath $packagePath -ErrorAction stop | Out-Null
}


function Export-Image {

    # We can also optimize an image by exporting to a new image file with Export-WindowsImage.
    # When modify an image, DISM stores additional resource files that increase the overall size of the image.
    # Exporting the image will remove unnecessary resource files.

    [cmdletbinding()]
    param([string]$srcPath,
        [int]$index,
        [string]$dstFile)

    Export-WindowsImage -SourceImagePath $srcPath -SourceIndex $index -DestinationImagePath $dstFile -ErrorAction stop | Out-Null
}


function Restore-Image {
    [cmdletbinding()]
    param([string]$imagePath)

    dism /image:$imagePath /cleanup-image /StartComponentCleanup | Out-Null
}


function Get-Architecture {
    [cmdletbinding()]
    param([string]$imagePath)

    # Hard code mapping here, reference in source code: ntexapi_h_x.w

    try {
        $info = Get-WindowsImage -ImagePath $imagePath -Index 1 -ErrorAction stop
        if ($info.Architecture -eq 0) {
            $arch = [Constants]::ARCH_X86
        }
        elseif ($info.Architecture -eq 5) {
            $arch = [Constants]::ARCH_ARM
        }
        elseif ($info.Architecture -eq 9) {
            $arch = [Constants]::ARCH_X64
        }
        elseif ($info.Architecture -eq 12) {
            $arch = [Constants]::ARCH_ARM64
        }
        else {
            $arch = [Constants]::ARCH_UNKNOWN
        }
    }
    catch {
        $arch = [Constants]::ARCH_UNKNOWN
    }

    return $arch
}


function Mount-ISO {
    [cmdletbinding()]
    param([string]$isoPath)

    $info = (Mount-DiskImage -ImagePath $isoPath -ErrorAction stop | Get-Volume)
    return $info.DriveLetter
}


function Dismount-ISO {
    [cmdletbinding()]
    param([string]$isoPath)

    Dismount-DiskImage -ImagePath $isoPath -ErrorAction stop | Out-Null
}


function Add-Capability {
    [cmdletbinding()]
    param([string]$capabilityName,
        [string]$imagePath,
        [string]$packagePath)

    Add-WindowsCapability -Name $capabilityName -Path $imagePath -Source $packagePath -ErrorAction stop | Out-Null
}


function Split-Image {
    [cmdletbinding()]
    param([string]$imagePath,
        [string]$dstPath,
        [int]$maxInstallSize)

    Split-WindowsImage -ImagePath $imagePath -SplitImagePath $dstPath -FileSize $maxInstallSize -CheckIntegrity -ErrorAction stop | Out-Null

}


function Get-ImageName {
    [cmdletbinding()]
    param([string]$imagePath,
        [int]$index)

    $info = Get-WindowsImage -ImagePath $imagePath -Index $index -ErrorAction stop
    return $info.ImageName
}


function Get-ImageTotalEdition {
    [cmdletbinding()]
    param([string]$imagePath)

    $info = Get-WindowsImage -ImagePath $imagePath -ErrorAction stop
    return $info.Count
}


# All status messages are always written to the diagnostic log.
#
# Level:
#   "debug"   Indicates information for debug, show on screen only in debug mode.
#   "info"    Indicates ordinary operational or status information.
#   "warning" Indicates something important, but not indicative of operational
#             failure. For example, the user tried to delete a VM that didn't
#             exist.
#   "error"   Indicates a terminal failure, from which the tools cannot or
#             should not attempt to recover. For example, we called an
#             externally owned cmdlet and it threw an unexpected exception.
#

function Write-Message {
    [CmdletBinding(PositionalBinding = $false)]

    param ([validateset("debug", "info", "warning", "error")][string]$level = [Constants]::LOG_INFO,
        [parameter(Position = 0)][string]$message)


    if ($level -eq $([Constants]::LOG_INFO)) {
        Write-Host $message
    }
    elseif ($level -eq $([Constants]::LOG_WARNING)) {
        Write-Warning $message
    }
    elseif ($level -eq $([Constants]::LOG_ERROR)) {
        # Write-Error $message; # I don't want to print out stack since upper stack is Out-Log, which is useless for trace.
        Write-Host "ERROR: $message" -foregroundcolor red
    }
    else {
        Write-Debug $message
    }
}


function Out-Log {
    [cmdletbinding()]
    param(
        [Parameter(Position = 0, Mandatory = $true, ValueFromPipeline = $true, ValueFromPipelineByPropertyName = $true)][String[]]$messages,
        [Parameter(Position = 1)][String]$logfile = $global:logPath,
        [Parameter(Position = 2)][validateset("debug", "info", "warning", "error")][string]$level = [Constants]::LOG_INFO
    );

    BEGIN {
        if (!(Test-Path $logfile)) {
            New-Item $logfile -Force | Out-Null
        }
    }
    PROCESS {
        foreach ($msg in $messages) {
            Write-Message $msg -level $level
            Add-Content $logfile "$(Get-Date -Format:"yyyy-MM-dd HH:mm:ss"), $level     $msg"
        }
    }
    END { }
}


class PatchMedia {

    [string]$installWimPath
    [int]$wimIndex
    [string]$bootWimPath
    [string]$winREPath
    [string]$workingPath
    [string]$packagesPath
    [string]$newMediaPath

    PatchMedia ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath) {
        $this.installWimPath = $installWimPath
        $this.wimIndex = $wimIndex
        $this.bootWimPath = $bootWimPath
        $this.winREPath = $winREPath
        $this.workingPath = $workingPath
        $this.packagesPath = $packagesPath
        $this.newMediaPath = $newMediaPath
    }

    [bool]Initialize() {
        return $true
    }

    [bool]PatchWinPE() {
        return $true
    }

    [bool]PatchWinRE() {
        return $true
    }

    [bool]PatchMainOS() {
        return $true
    }

    [bool]PatchMediaBinaries() {
        return $true
    }

    [void]Cleanup() {
        return
    }

    [bool]TestNeedPatch() {
        # Test whether need to patch according to the existance of DU or ISO file.
        return $true
    }

    [bool]DoPatch() {
        if ( ($this.TestNeedPatch()) ) {
            if ( !($this.Initialize()) ) { return $False }
            if ( !($this.PatchWinPE()) ) { return $False }
            if ( !($this.PatchWinRE()) ) { return $False }
            if ( !($this.PatchMainOS()) ) { return $False }
            if ( !($this.PatchMediaBinaries()) ) { return $False }
            $this.Cleanup()
        }

        return $true
    }
}


class DownloadContents {

    # Properties
    [string]$platform
    [string]$product
    [string]$version
    [switch]$showLinksOnly

    DownloadContents ([string]$platform, [string]$product, [string]$version, [switch]$showLinksOnly) {
        $this.platform = $platform
        $this.product = $product
        $this.version = $version
        $this.showLinksOnly = $showLinksOnly
    }

    [bool]Initialize() {
        return $true
    }

    [bool]Download() {
        return $true
    }

    [void]Cleanup() {
        return
    }

    [bool]DoDownload() {

        if (!($this.Initialize())) { return $False }
        if (!($this.Download())) { return $False }
        $this.Cleanup()

        return $true
    }
}
