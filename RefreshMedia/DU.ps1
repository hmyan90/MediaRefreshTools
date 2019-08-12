# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------


class PatchDU: PatchMedia {

    [string]$SSUPath    # for MainOS and WinRE
    [string]$SafeOSPath # for WinRE only
    [string]$LCUPath    # for
    [string]$SetupDUPath

    PatchDU ([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.SSUPath = Join-Path $packagesPath $([Constants]::SSU_DIR)
        $this.SafeOSPath = Join-Path $packagesPath $([Constants]::SAFEOS_DIR)
        $this.LCUPath = Join-Path $packagesPath $([Constants]::LCU_DIR)
        $this.SetupDUPath = Join-Path $packagesPath $([Constants]::SETUPDU_DIR)
    }

    [bool]TestNeedPatch() {
        $SSUNotExist = ((!(Test-Path $this.SSUPath)) -or (Test-FolderEmpty $this.SSUPath))
        $SafeOSNotExist = ((!(Test-Path $this.SafeOSPath)) -or (Test-FolderEmpty $this.SafeOSPath))
        $LCUNotExist = ((!(Test-Path $this.LCUPath)) -or (Test-FolderEmpty $this.LCUPath))
        $SSUNotExist = ((!(Test-Path $this.SetupDUPath)) -or (!([PatchDU]::GetLatestSetupDU($this.SetupDUPath))))

        if ( ($SSUNotExist -and $SafeOSNotExist -and $LCUNotExist -and $SSUNotExist) ) {
            Out-Log "No need to patch DU since script cannot find any related packages in $( $this.packagesPath )" -level $([Constants]::LOG_WARNING)
            return $False
        }
        return $True
    }

    static InstallPackages($imagePath, $Path, $envName, $packageType) {
        if (Test-Path $Path) {
            Get-ChildItem -Path $Path\* -Include *.msu, *cab |

            Foreach-Object {
                Out-Log ("Install $packageType $( $_.Name ) to $envName")
                Install-Package $imagePath ($_.FullName)
            }
        }
    }

    [bool]PatchSsuLcuSafeOS([string]$wimPath, [int]$index, [string]$mountPoint, [string]$envName) {

        try {
            Out-Log "Mount Image $wimPath $index to $mountPoint" -level $([Constants]::LOG_DEBUG)

            Mount-Image $wimPath $index $mountPoint

            [PatchDU]::InstallPackages($mountPoint, $this.SSUPath, $envName, "SSU")
            [PatchDU]::InstallPackages($mountPoint, $this.SafeOSPath, $envName, "SafeOS DU")
            [PatchDU]::InstallPackages($mountPoint, $this.LCUPath, $envName, "LCU")

            Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
            Dismount-CommitImage $mountPoint
            return $True
        }
        catch {
            Out-Log "Failed to patch SSU/LCU/SafeOS to $envName. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

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

    [bool]PatchWinPE() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WINPE_MOUNT)
        return ($this.PatchSsuLcuSafeOS($this.bootWimPath, 1, $mountPoint, [Constants]::WINPE))
    }

    [bool]PatchWinSetup() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WIN_SETUP_MOUNT)
        return ($this.PatchSsuLcuSafeOS($this.bootWimPath, 2, $mountPoint, [Constants]::WIN_SETUP))
    }

    [bool]PatchWinRE() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::WINRE_MOUNT)
        return ($this.PatchSsuLcuSafeOS($this.winREPath, 1, $mountPoint, [Constants]::WINRE))
    }

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::INSTALL_MOUNT)
        return ($this.PatchSsuLcuSafeOS($this.installWimPath, $this.wimIndex, $mountPoint, [Constants]::MAIN_OS))
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
            # Setup DU is named as "Windows10.0-KB4490480-x64.cab"
            # Need to get latest setup DU according to KB name (todo check ?? )
            try {
                $latestFile = ($fileNames | Sort-Object { $_ -match "Windows10.0-KB([0-9]+)-(.+).cab" | % { [int]($matches[1]) } } -Descending)[0]
                return (Join-Path $setupDUPath $latestFile)
            }
            catch {
                return $null
            }
        }
    }

    [bool]PatchSetupBinaries() {
        $latestSetupDUPath = [PatchDU]::GetLatestSetupDU($this.SetupDUPath)
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