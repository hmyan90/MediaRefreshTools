# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------


class PatchFOD: PatchMedia {

    [string]$FodIsoPath
    [string[]]$CapabilityList
    [string]$FODPath

    PatchFOD([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath, [string[]]$capabilityList): base($installWimPath, $wimIndex, $bootWimPath,
        $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.FodIsoPath = [PatchFOD]::GetFodIsoPath($packagesPath)
        $this.CapabilityList = $capabilityList
    }

    [bool]TestNeedPatch() {
        return ( $this.CapabilityList.Count -gt 0)
    }

    [string]static GetFodIsoPath($packagesPath) {
        $fodIsoDir = Join-Path $packagesPath $([Constants]::FOD_DIR)
        return (Get-ChildItem $fodIsoDir).FullName
    }

    [bool]Initialize() {
        try {
            $driveLetter = Mount-ISO $this.FodIsoPath
        }
        catch {
            Out-Log "Failed to mount $( $this.FodIsoPath ). Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
            return $False
        }

        $this.FODPath = $driveLetter + ":\"
        return $True
    }

    [bool]PatchMainOS() {

        $mountPoint = Join-Path $this.WorkingPath $([Constants]::INSTALL_MOUNT)

        try {
            Out-Log "Mount Image $( $this.InstallWimPath ) $( $this.WimIndex ) to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.InstallWimPath $this.WimIndex $mountPoint

            foreach ($capabilityName in $this.CapabilityList) {
                Out-Log "Install FoD capability $capabilityName to $([Constants]::MAIN_OS)"
                Add-Capability $capabilityName $mountPoint $this.FODPath
            }

            Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
            Dismount-CommitImage $mountPoint
            return $True
        }
        catch {
            Out-Log "Failed to add all capabilities. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

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

    [void]CleanUp() {
        try {
            Dismount-ISO $this.FodIsoPath
        }
        catch {
            Out-Log "Failed to dismount $( $this.FodIsoPath ). Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_WARNING)
        }
    }
}