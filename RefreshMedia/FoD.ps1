# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------


class PatchFOD: PatchMedia {

    [string]$FODISOPath
    [string[]]$capabilityList
    [string]$FODPath

    PatchFOD([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath, [string[]]$capabilityList): base($installWimPath, $wimIndex, $bootWimPath,
        $winREPath, $workingPath, $packagesPath, $newMediaPath) {
        $this.FODISOPath = [PatchFOD]::GetFODISOPath($packagesPath)
        $this.capabilityList = $capabilityList
    }

    [bool]TestNeedPatch() {
        return ( $this.capabilityList.Count -gt 0)
    }

    [string]static GetFODISOPath($packagesPath) {
        $FODISODir = Join-Path $packagesPath $([Constants]::FOD_DIR)
        return (Get-ChildItem $FODISODir).FullName
    }

    [bool]Initialize() {
        try {
            $driveLetter = Mount-ISO $this.FODISOPath
        }
        catch {
            Out-Log "Failed to mount $( $this.FODISOPath ). Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
            return $False
        }

        $this.FODPath = $driveLetter + ":\"
        return $True
    }


    [bool]PatchMainOS() {

        $mountPoint = Join-Path $this.workingPath $([Constants]::INSTALL_MOUNT)

        try {
            Out-Log "Mount Image $( $this.installWimPath ) $( $this.wimIndex ) to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.installWimPath $this.wimIndex $mountPoint

            foreach ($capabilityName in $this.capabilityList) {
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
            Dismount-ISO $this.FODISOPath
        }
        catch {
            Out-Log "Failed to dismount $( $this.FODISOPath ). Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_WARNING)
        }
    }

}