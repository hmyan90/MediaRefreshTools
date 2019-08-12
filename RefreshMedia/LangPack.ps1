# --------------------------------------------------------------
#  Copyright © Microsoft Corporation.  All Rights Reserved.
# ---------------------------------------------------------------


function Test-LangInput {
    [cmdletbinding()]
    param([string]$arch,
        [string]$LPISOPath,
        [string[]]$langList)

    try {
        $driveLetter = Mount-ISO $LPISOPath
    }
    catch {
        Out-Log "Failed to mount ISO $LPISOPath. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        return $False
    }

    $isValidInput = $False
    if ( $langList.Count -eq 0 ) {
        $isValidInput = $True
    }
    else {
        if ( !(Test-Path -Path ($driveLetter + ":\" + $arch)) ) {
            Out-Log "Language Pack not supported on this architecture" -level $([Constants]::LOG_ERROR)
        }
        else {
            $OSLangPackPath = $driveLetter + ":\" + $arch + "\langpacks\"
            $notFoundLang = [System.Collections.ArrayList]@()

            Foreach ($lang in $langList) {
                $LPFile = Join-Path $OSLangPackPath ("Microsoft-Windows-Client-Language-Pack_" + $arch + "_" + $lang + ".cab")
                if ( !(Test-Path -Path $LPFile) ) {
                    $notFoundLang.add($lang) | Out-NULL
                }
            }

            if ( $notFoundLang.Count -gt 0 ) {
                Out-Log "$notFoundLang not supported in Language Pack, please retry." -level $([Constants]::LOG_ERROR)
            }
            else {
                $isValidInput = $True
            }
        }
    }

    try {
        Dismount-ISO $LPISOPath
    }
    catch {
        Out-Log "Failed to dismount ISO $LPISOPath. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        return $False
    }

    return $isValidInput
}


class PatchLP: PatchMedia {

    [string]$LPISOPath
    [string[]]$langList
    [string]$arch
    [string]$mainOSLangPackPath
    [string]$WinPEOCPath

    PatchLP([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath, [string[]]$langList, [string]$arch): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath,
        $packagesPath, $newMediaPath) {
        $this.LPISOPath = [PatchLP]::GetLPISOPath($packagesPath)
        $this.langList = $langList
        $this.arch = $arch
    }

    [bool]TestNeedPatch() {
        return ( $this.langList.Count -gt 0)
    }

    [string]static GetLPISOPath($packagesPath) {
        $LPISODir = Join-Path $packagesPath $([Constants]::LP_DIR)
        return (Get-ChildItem $LPISODir).FullName
    }

    [bool]Initialize() {
        try {
            $driveLetter = Mount-ISO $this.LPISOPath
        }
        catch {
            Out-Log "Failed to mount $( $this.LPISOPath ). Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
            return $False
        }

        try {
            $this.mainOSLangPackPath = Join-Path $driveLetter":" -ChildPath $this.arch | Join-Path -ChildPath "langpacks"
            $this.WinPEOCPath = Join-Path $driveLetter":" -ChildPath "Windows Preinstallation Environment" | Join-Path -ChildPath $this.arch | Join-Path -ChildPath "WinPE_OCs"
        }
        catch {
            Out-Log "Path doesn't exist. Detail: $( $_.Exception.Message ) " -level $([Constants]::LOG_ERROR)
            return $False
        }

        return $True
    }

    <#
    # Not sure about LangPack request of boot.wim, comment for now.

    AddLPToWinSetup($mountPoint) {

        $envName = [Constants]::WIN_SETUP
        Foreach ($lang in $this.langList) {
            $winPEOCLangPath = Join-Path $this.WinPEOCPath $lang

            # Install lp.cab cab
            $lpPath = (Join-Path $winPEOCLangPath "lp.cab")
            Out-Log "Install LangPack $lpPath to $envName"
            Install-Package $mountPoint $lpPath

            # Install WinPE-Setup_$lang.cab
            $LPFile = "WinPE-Setup_$lang.cab"
            $LPPath = Join-Path $winPEOCLangPath $LPFile
            if ( Test-Path -Path $LPPath ) {
                Out-Log "Install LangPack $LPPath to $envName"
                Install-Package $mountPoint $LPPath
            }
            else {
                Out-Log "$LPPath not found (possibly a bug)." -level $([Constants]::LOG_WARNING)
            }

            # Install WinPE-Setup-Client_$lang.cab
            $LPFile = "WinPE-Setup-Client_$lang.cab"
            $LPPath = Join-Path $winPEOCLangPath $LPFile
            if ( Test-Path -Path $LPPath ) {
                Out-Log "Install LangPack $LPPath to $envName"
                Install-Package $mountPoint $LPPath
            }
            else {
                Out-Log "$LPPath not found (possibly a bug)." -level $([Constants]::LOG_WARNING)
            }
        }
    }

    AddTTSToWinSetup($mountPoint) {

        $envName = [Constants]::WIN_SETUP
        $needInstallSTTS = $False
        $defaultSTTSPath = (Join-Path $this.WinPEOCPath "WinPE-Speech-TTS.cab")
        if ( !(Test-Path -Path $defaultSTTSPath) ) {
            Out-Log "$defaultSTTSPath not exists, stop install all Speech-TTS for $envName" -level $([Constants]::LOG_WARNING)
        }
        else {
            Foreach ($lang in $this.langList) {

                $STTSFile = "WinPE-Speech-TTS-$lang.cab"
                $STTSPath = Join-Path $this.WinPEOCPath $STTSFile
                if ( (Test-Path -Path $STTSPath) ) {
                    if ($needInstallSTTS -eq $False) {
                        $needInstallSTTS = $True
                        Out-Log "Install LangPack $defaultSTTSPath to $envName"
                        Install-Package $mountPoint $defaultSTTSPath
                    }
                    Out-Log "Install LangPack $STTSPath to $envName"
                    Install-Package $mountPoint $STTSPath
                }
            }
        }
    }

    AddFontSupportToWinSetup($mountPoint) {

        $envName = [Constants]::WIN_SETUP
        Foreach ($lang in $this.langList) {

            $fontSupportFile = "WinPE-FontSupport-$lang.cab"
            $fontSupportPath = Join-Path $this.WinPEOCPath $fontSupportFile
            if ( (Test-Path -Path $fontSupportPath) ) {
                Out-Log "Install LangPack $fontSupportPath to $envName"
                Install-Package $mountPoint $fontSupportPath
            }
        }
    }

    [bool]PatchWinSetup() {
        # Follow the documentation: https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/add-multilingual-support-to-windows-setup

        $mountPoint = Join-Path $this.workingPath $([Constants]::WIN_SETUP_MOUNT)
        $envName = [Constants]::WIN_SETUP

        try {
            Out-Log "Mount Image $( $this.bootWimPath ) 2 to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.bootWimPath 2 $mountPoint

            $this.AddFontSupportToWinSetup($mountPoint)
            $this.AddTTSToWinSetup($mountPoint)
            $this.AddLPToWinSetup($mountPoint)

            Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
            Dismount-CommitImage $mountPoint
            return $True
        }
        catch {
            Out-Log "Failed to add all language packages to $envName. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

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
    #>


    [bool]PatchWinRE() {
        # Follow the documentation: https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/add-language-packs-to-windows

        $mountPoint = Join-Path $this.workingPath $([Constants]::WINRE_MOUNT)
        $envName = [Constants]::WINRE

        try {
            Out-Log "Mount Image $( $this.winREPath ) 1 to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.winREPath 1 $mountPoint
            $winREInstalledOC = Get-WindowsPackage -Path $mountPoint

            Foreach ($lang in $this.langList) {

                $winPEOCLangPath = Join-Path $this.WinPEOCPath $lang
                $cabs = Get-ChildItem $winPEOCLangPath -name

                # Install lp.cab cab
                $lpPath = Join-Path $winPEOCLangPath "lp.cab"
                Out-Log "Install LangPack $lpPath to $envName"
                Install-Package $mountPoint $lpPath

                # Install OC cab
                Foreach ($package in $winREInstalledOC) {

                    if ( ($package.PackageState -eq "Installed") `
                            -and ($package.PackageName.startsWith("WinPE-")) `
                            -and ($package.ReleaseType -eq "FeaturePack") ) {

                        $index = $package.PackageName.IndexOf("-Package")
                        if ($index -ge 0) {
                            $OCCab = $package.PackageName.Substring(0, $index) + "_" + $lang + ".cab"

                            if ($cabs.Contains($OCCab)) {
                                $OCCabPath = Join-Path $winPEOCLangPath $OCCab
                                Out-Log "Install LangPack $OCCabPath to $envName"
                                Install-Package $mountPoint $OCCabPath
                            }
                        }
                    }
                }
            }

            Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
            Dismount-CommitImage $mountPoint
            return $True
        }
        catch {
            Out-Log "Failed to add all language packages to $envName. Detail: $(  $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

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

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.workingPath $([Constants]::INSTALL_MOUNT)
        $envName = [Constants]::MAIN_OS

        try {
            Out-Log "Mount Image $( $this.installWimPath ) $( $this.wimIndex ) to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.installWimPath $this.wimIndex $mountPoint

            Foreach ($lang in $this.langList) {
                $LPFile = "Microsoft-Windows-Client-Language-Pack_" + $this.arch + "_" + $lang + ".cab"
                $LPPath = Join-Path $this.mainOSLangPackPath $LPFile
                if ( Test-Path -Path $LPPath ) {
                    Out-Log "Install LangPack $LPPath to $envName"
                    Install-Package $mountPoint $LPPath
                }
                else {
                    # should not happen because we already checked language input.
                    Out-Log "$LPPath not found (possibly a bug)." -level $([Constants]::LOG_WARNING)
                }
            }

            Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
            Dismount-CommitImage $mountPoint
            return $True;
        }
        catch {
            Out-Log "Failed to add all language packages to $envName. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)

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
            Dismount-ISO $this.LPISOPath
        }
        catch {
            Out-Log "Failed to dismount $( $this.LPISOPath ). Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_WARNING)
        }
    }
}