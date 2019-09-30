# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------

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

    static InstallPackages($imagePath, $Path, $envName, [DUType]$DUType) {
        if (Test-Path $Path) {
            Get-ChildItem -Path $Path\* -Include *.msu, *cab |

                Foreach-Object {
                    Out-Log ("Install $( [PatchDU]::DUInfoMapping.$DUType.name ) $( $_.Name ) to $envName")
                    Install-Package $imagePath ($_.FullName)
                }
        }
    }

    static [bool]PatchDUs([string]$wimPath, [int]$index, [string]$mountPoint, [string]$envName, [string]$SSUPath, [string]$SafeOSPath, [string]$LCUPath) {

        try {
            Out-Log "Mount Image $wimPath $index to $mountPoint" -level $([Constants]::LOG_DEBUG)

            Mount-Image $wimPath $index $mountPoint

            if ( $SSUPath ) {
                [PatchDU]::InstallPackages($mountPoint, $SSUPath, $envName, [DUType]::SSU)
            }

            if ( $SafeOSPath ) {
                [PatchDU]::InstallPackages($mountPoint, $SafeOSPath, $envName, [DUType]::SafeOS)
            }

            if ( $LCUPath ) {
                [PatchDU]::InstallPackages($mountPoint, $LCUPath, $envName, [DUType]::LCU)
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


class PatchLCU: PatchDU {

    [string]$LCUPath

    PatchLCU ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.LCUPath = Join-Path $packagesPath $([Constants]::LCU_DIR)
    }

    [bool]TestNeedPatch() {
        $LCUNotExist = ((!(Test-Path $this.LCUPath)) -or (Test-FolderEmpty $this.LCUPath))

        if ( $LCUNotExist ) {
            Out-Log "No need to patch LCU DU since script cannot find any related packages in $( $this.packagesPath )" -level $([Constants]::LOG_WARNING)
            return $False
        }
        return $True
    }

    [bool]PatchWinPE() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WINPE_MOUNT)
        return ([PatchDU]::PatchDUs($this.bootWimPath, 1, $mountPoint, [Constants]::WINPE, $null, $null, $this.LCUPath))
    }

    [bool]PatchWinSetup() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WIN_SETUP_MOUNT)
        return ([PatchDU]::PatchDUs($this.bootWimPath, 2, $mountPoint, [Constants]::WIN_SETUP, $null, $null, $this.LCUPath))
    }

    [bool]PatchWinRE() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WINRE_MOUNT)
        return ([PatchDU]::PatchDUs($this.winREPath, 1, $mountPoint, [Constants]::WINRE, $null, $null, $this.LCUPath))
    }

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::INSTALL_MOUNT)
        return ([PatchDU]::PatchDUs($this.installWimPath, $this.wimIndex, $mountPoint, [Constants]::MAIN_OS, $null, $null, $this.LCUPath))
    }
}


class PatchDUExcludeLCU: PatchDU {

    [string]$SSUPath
    [string]$SafeOSPath
    [string]$SetupDUPath

    PatchDUExcludeLCU ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.SSUPath = Join-Path $packagesPath $([Constants]::SSU_DIR)
        $this.SafeOSPath = Join-Path $packagesPath $([Constants]::SAFEOS_DIR)
        $this.SetupDUPath = Join-Path $packagesPath $([Constants]::SETUPDU_DIR)
    }

    [bool]TestNeedPatch() {
        $SSUNotExist = ((!(Test-Path $this.SSUPath)) -or (Test-FolderEmpty $this.SSUPath))
        $SafeOSNotExist = ((!(Test-Path $this.SafeOSPath)) -or (Test-FolderEmpty $this.SafeOSPath))
        $SetupDUNotExist = ((!(Test-Path $this.SetupDUPath)) -or (Test-FolderEmpty $this.SetupDUPath))

        if ( ($SSUNotExist -and $SafeOSNotExist -and $SetupDUNotExist) ) {
            Out-Log "No need to patch SSU or SafeOS or Setup DU since script cannot find any related packages in $( $this.packagesPath )" -level $([Constants]::LOG_WARNING)
            return $False
        }
        return $True
    }

    [bool]PatchWinPE() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WINPE_MOUNT)
        return ([PatchDU]::PatchDUs($this.bootWimPath, 1, $mountPoint, [Constants]::WINPE, $this.SSUPath, $this.SafeOSPath, $null))
    }

    [bool]PatchWinSetup() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WIN_SETUP_MOUNT)
        return ([PatchDU]::PatchDUs($this.bootWimPath, 2, $mountPoint, [Constants]::WIN_SETUP, $this.SSUPath, $this.SafeOSPath, $null))
    }

    [bool]PatchWinRE() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WINRE_MOUNT)
        return ([PatchDU]::PatchDUs($this.winREPath, 1, $mountPoint, [Constants]::WINRE, $this.SSUPath, $this.SafeOSPath, $null))
    }

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::INSTALL_MOUNT)
        return ([PatchDU]::PatchDUs($this.installWimPath, $this.wimIndex, $mountPoint, [Constants]::MAIN_OS, $this.SSUPath, $null, $null))
    }

    static [string]GetLatestSetupDU([string]$setupDUPath) {

        if ( !(Test-Path $setupDUPath) ) { return $null }

        $fileNames = (Get-ChildItem -Path $setupDUPath).Name
        if ( $fileNames.count -eq 1 ) {
            return (Join-Path $setupDUPath $fileNames)
        }
        elseif ( $fileNames.count -eq 0 ) {
            return $null
        }
        else {
            # Should not reach here since we already checked the number
            return $null
        }
    }

    [bool]PatchSetupBinaries() {
        $latestSetupDUPath = [PatchDUExcludeLCU]::GetLatestSetupDU($this.SetupDUPath)
        $dstPath = Join-Path $this.newMediaPath "sources"

        if ( $latestSetupDUPath ) {
            Out-Log "Install SetupDU $( Split-Path $latestSetupDUPath -leaf ) to new media"
            try {
                cmd.exe /c $env:SystemRoot\System32\expand.exe $latestSetupDUPath -F:* $dstPath | Out-Null
            }
            catch {
                Out-Log "Failed to patch Setup DU. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
                return $false
            }
        }
        else {
            Out-Log "Didn't find any related Setup DU." -level $([Constants]::LOG_WARNING)
        }

        return $true
    }
}