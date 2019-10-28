# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------

class PatchDU: PatchMedia {

    static $DUInfoMapping = @{
        [DUType]::SSU     = @{name = "SSU" };
        [DUType]::LCU     = @{name = "LCU" };
        [DUType]::SafeOS  = @{name = "SafeOS DU" };
        [DUType]::SetupDU = @{name = "SetupDU" };
    }

    PatchDU ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath, $packagesPath, $newMediaPath) {
    }

    static InstallPackages($imagePath, $path, $envName, [DUType]$duType) {
        if (Test-Path $path) {
            Get-ChildItem -Path $path\* -Include *.msu, *cab |

                Foreach-Object {
                    Out-Log ("Install $( [PatchDU]::DUInfoMapping.$duType.name ) $( $_.Name ) to $envName")
                    Install-Package $imagePath ($_.FullName)
                }
        }
    }

    static [bool]PatchDUs([string]$wimPath, [int]$index, [string]$mountPoint, [string]$envName, [string]$ssuPath, [string]$safeOSPath, [string]$lcuPath) {

        try {
            Out-Log "Mount Image $wimPath $index to $mountPoint" -level $([Constants]::LOG_DEBUG)

            Mount-Image $wimPath $index $mountPoint

            if ( $ssuPath ) {
                [PatchDU]::InstallPackages($mountPoint, $ssuPath, $envName, [DUType]::SSU)
            }

            if ( $safeOSPath ) {
                [PatchDU]::InstallPackages($mountPoint, $safeOSPath, $envName, [DUType]::SafeOS)
            }

            if ( $lcuPath ) {
                [PatchDU]::InstallPackages($mountPoint, $lcuPath, $envName, [DUType]::LCU)
            }

            Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
            Dismount-CommitImage $mountPoint
            return $True
        }
        catch {
            Out-Log "Failed to patch SSU/SafeOS to $envName. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

            try {
                Dismount-DiscardImage $mountPoint
            }
            catch {
                # Here is try best to do cleanup
                Out-Log "Failed to cleanup. $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
            }
            return $False
        }
    }
}


class PatchSSU: PatchDU {

    <#
    .SYNOPSIS
        Patch SSU to WinPE, WinRE and Main OS
    #>

    [string]$SSUPath

    PatchSSU ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.SSUPath = Join-Path $packagesPath $([Constants]::SSU_DIR)
    }

    [bool]TestNeedPatch() {
        $SSUNotExist = ((!(Test-Path $this.SSUPath)) -or (Test-FolderEmpty $this.SSUPath))

        if ( $SSUNotExist ) {
            Out-Log "No need to patch SSU since script cannot find any related packages in $( $this.PackagesPath )" -level $([Constants]::LOG_WARNING)
            return $False
        }
        return $True
    }

    [bool]PatchWinPE() {
        $imageNumber = Get-ImageTotalEdition $this.BootWimPath

        For ($index = 1; $index -le $imageNumber; $index++) {
            $mountPoint = Join-Path $this.WorkingPath $([Constants]::WINPE_MOUNT)
            if ( [PatchDU]::PatchDUs($this.BootWimPath, $index, $mountPoint, "$( [Constants]::WINPE )[$index]", $this.SSUPath, $Null, $Null) -eq $False ) {
                return $False
            }
        }
        return $True
    }

    [bool]PatchWinRE() {
        $mountPoint = Join-Path $this.WorkingPath $([Constants]::WINRE_MOUNT)
        return ([PatchDU]::PatchDUs($this.WinREPath, 1, $mountPoint, [Constants]::WINRE, $this.SSUPath, $Null, $Null))
    }

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.WorkingPath $([Constants]::INSTALL_MOUNT)
        return ([PatchDU]::PatchDUs($this.InstallWimPath, $this.WimIndex, $mountPoint, [Constants]::MAIN_OS, $this.SSUPath, $Null, $Null))
    }
}


class PatchDUExcludeSSU: PatchDU {

    <#
    .SYNOPSIS
        Patch SafeOS, Setup DU, LCU
    .DESCRIPTION
        Patch SafeOS to WinRE
        Patch Setup DU for Media
        Patch LCU to WinPE, Main OS
    #>

    [string]$LCUPath
    [string]$SafeOSPath
    [string]$SetupDUPath

    PatchDUExcludeSSU ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.LCUPath = Join-Path $packagesPath $([Constants]::LCU_DIR)
        $this.SafeOSPath = Join-Path $packagesPath $([Constants]::SAFEOS_DIR)
        $this.SetupDUPath = Join-Path $packagesPath $([Constants]::SETUPDU_DIR)
    }

    [bool]TestNeedPatch() {
        $SafeOSNotExist = ((!(Test-Path $this.SafeOSPath)) -or (Test-FolderEmpty $this.SafeOSPath))
        $SetupDUNotExist = ((!(Test-Path $this.SetupDUPath)) -or (Test-FolderEmpty $this.SetupDUPath))
        $LCUNotExist = ((!(Test-Path $this.LCUPath)) -or (Test-FolderEmpty $this.LCUPath))

        if ( ($SafeOSNotExist -and $SetupDUNotExist -and $LCUNotExist) ) {
            Out-Log "No need to patch SafeOS and Setup DU and LCU since script cannot find any related packages in $( $this.PackagesPath )" -level $([Constants]::LOG_WARNING)
            return $False
        }
        return $True
    }

    [bool]PatchWinPE() {
        $imageNumber = Get-ImageTotalEdition $this.BootWimPath

        For ($index = 1; $index -le $imageNumber; $index++) {
            $mountPoint = Join-Path $this.WorkingPath $([Constants]::WINPE_MOUNT)
            if ( [PatchDU]::PatchDUs($this.BootWimPath, $index, $mountPoint, "$( [Constants]::WINPE )[$index]", $null, $Null, $this.LCUPath) -eq $False) {
                return $False
            }
        }
        return $True
    }

    [bool]PatchWinRE() {
        $mountPoint = Join-Path $this.WorkingPath $([Constants]::WINRE_MOUNT)
        return ([PatchDU]::PatchDUs($this.WinREPath, 1, $mountPoint, [Constants]::WINRE, $Null, $this.SafeOSPath, $Null))
    }

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.WorkingPath $([Constants]::INSTALL_MOUNT)
        return ([PatchDU]::PatchDUs($this.InstallWimPath, $this.WimIndex, $mountPoint, [Constants]::MAIN_OS, $Null, $Null, $this.LCUPath))
    }

    static [string]GetLatestSetupDU([string]$setupDUPath) {

        if ( !(Test-Path $setupDUPath) ) { return $null }

        $fileNames = (Get-ChildItem -Path $setupDUPath).Name
        if ( $fileNames.count -eq 1 ) {
            return (Join-Path $setupDUPath $fileNames)
        }
        elseif ( $fileNames.count -eq 0 ) {
            return $Null
        }
        else {
            # Should not reach here since we already checked the number
            return $Null
        }
    }

    [bool]PatchMediaBinaries() {
        $latestSetupDUPath = [PatchDUExcludeSSU]::GetLatestSetupDU($this.SetupDUPath)
        $dstPath = Join-Path $this.NewMediaPath "sources"

        if ( $latestSetupDUPath ) {
            Out-Log "Install SetupDU $( Split-Path $latestSetupDUPath -leaf ) to new media"
            try {
                cmd.exe /c $env:SystemRoot\System32\expand.exe $latestSetupDUPath -F:* $dstPath | Out-Null
            }
            catch {
                Out-Log "Failed to patch Setup DU. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
                return $False
            }
        }
        else {
            Out-Log "Didn't find any related Setup DU." -level $([Constants]::LOG_WARNING)
        }

        return $True
    }
}