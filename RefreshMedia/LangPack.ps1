# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------


function Test-LangInput {
    [cmdletbinding()]
    param([string]$arch,
        [string]$lpISOPath,
        [string[]]$langList)

    try {
        $driveLetter = Mount-ISO $lpISOPath
    }
    catch {
        Out-Log "Failed to mount ISO $lpISOPath. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
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
            $osLangPackPath = $driveLetter + ":\" + $arch + "\langpacks\"
            $notFoundLang = [System.Collections.ArrayList]@()

            Foreach ($lang in $langList) {
                $lpFile = Join-Path $osLangPackPath ("Microsoft-Windows-Client-Language-Pack_" + $arch + "_" + $lang + ".cab")
                if ( !(Test-Path -Path $lpFile) ) {
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
        Dismount-ISO $lpISOPath
    }
    catch {
        Out-Log "Failed to dismount ISO $lpISOPath. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_ERROR)
        return $False
    }

    return $isValidInput
}


class PatchLP: PatchMedia {

    <#
    .SYNOPSIS
        Add language to WinPE, WinRE, Main OS
    .DESCRIPTION
        For details about adding lanugage to Main OS and WinRE, see: https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/add-language-packs-to-windows
        For details about adding language to WinPE, see: https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/add-multilingual-support-to-windows-setup
    #>

    [string]$LPISOPath
    [string[]]$LangList
    [string]$Arch
    [switch]$WinPELang
    [string]$MainOSLangPackPath
    [string]$WinPEOCPath

    PatchLP([string]$installWimPath, [int]$wimIndex, [string]$bootWimPath, [string]$winREPath, [string]$workingPath, [string]$packagesPath,
        [string]$newMediaPath, [string[]]$langList, [string]$arch, [switch]$winPELang): base($installWimPath, $wimIndex, $bootWimPath, $winREPath, $workingPath,
        $packagesPath, $newMediaPath) {
        $this.LPISOPath = [PatchLP]::GetLPISOPath($packagesPath)
        $this.LangList = $langList
        $this.Arch = $arch
        $this.WinPELang = $winPELang
    }

    [bool]TestNeedPatch() {
        return ( $this.LangList.Count -gt 0)
    }

    [string]static GetLPISOPath($packagesPath) {
        $lpISODir = Join-Path $packagesPath $([Constants]::LP_DIR)
        return (Get-ChildItem $lpISODir).FullName
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
            $this.MainOSLangPackPath = Join-Path $driveLetter":" -ChildPath $this.Arch | Join-Path -ChildPath "langpacks"
            $this.WinPEOCPath = Join-Path $driveLetter":" -ChildPath "Windows Preinstallation Environment" | Join-Path -ChildPath $this.Arch | Join-Path -ChildPath "WinPE_OCs"
        }
        catch {
            Out-Log "Path doesn't exist. Detail: $( $_.Exception.Message ) " -level $([Constants]::LOG_ERROR)
            return $False
        }

        return $True
    }

    AddTTSToWinPE($mountPoint, $envName) {

        <#
        .SYNOPSIS
            add TTS package for WinPE environment
        .DESCRIPTION
            This function will try to find available speech TTS Package in Language Pack and install them
            This function will raise exception when failed, upper function need to handle this exception
        #>

        $needInstallSpeechTTS = $False
        $defaultSpeechTTSPath = (Join-Path $this.WinPEOCPath "WinPE-Speech-TTS.cab")
        if ( !(Test-Path -Path $defaultSpeechTTSPath) ) {
            Out-Log "$defaultSpeechTTSPath not exists, stop install all Speech-TTS for $envName" -level $([Constants]::LOG_WARNING)
        }
        else {
            Foreach ($lang in $this.LangList) {

                $speechTTSFile = "WinPE-Speech-TTS-$lang.cab"
                $speechTTSPath = Join-Path $this.WinPEOCPath $speechTTSFile
                if ( (Test-Path -Path $speechTTSPath) ) {
                    if ($needInstallSpeechTTS -eq $False) {
                        $needInstallSpeechTTS = $True
                        Out-Log "Install LangPack $defaultSpeechTTSPath to $envName"
                        Install-Package $mountPoint $defaultSpeechTTSPath
                    }
                    Out-Log "Install LangPack $speechTTSPath to $envName"
                    Install-Package $mountPoint $speechTTSPath
                }
            }
        }
    }

    AddFontSupportToWinPE($mountPoint, $envName) {
        <#
        .SYNOPSIS
            add Font Support package for WinPE environment
        .DESCRIPTION
            This function will try to find available Font Support in Language Pack and install them
            This function will raise exception when failed, upper function need to handle this exception
        #>

        Foreach ($lang in $this.LangList) {

            $fontSupportFile = "WinPE-FontSupport-$lang.cab"
            $fontSupportPath = Join-Path $this.WinPEOCPath $fontSupportFile
            if ( (Test-Path -Path $fontSupportPath) ) {
                Out-Log "Install LangPack $fontSupportPath to $envName"
                Install-Package $mountPoint $fontSupportPath
            }
        }
    }

    GenLangIni($mountPoint, $envName) {

        try {
            $iniFilePath = Join-Path $mountPoint "sources\lang.ini"
            if ( Test-Path $iniFilePath ) {
                Out-Log "Update lang.ini for $envName" -level $([Constants]::LOG_DEBUG)
                dism /image:$mountPoint /Gen-LangINI /distribution:$mountPoint | Out-Null
            }
        }
        catch {
            Out-Log "Failed to update lang.ini for $envName" -level $([Constants]::LOG_WARNING)
        }
    }

    AddOcLP($mountPoint, $envName) {
        <#
        .SYNOPSIS
            add language packs for WinPE Optional Components
        .DESCRIPTION
            This function will loop for all installed Optional Componment in environment, add language packs for those Optional Componment
            This function will raise exception when failed, upper function need to handle this exception
        #>

        $winPEInstalledOC = Get-WindowsPackage -Path $mountPoint

        Foreach ($lang in $this.LangList) {

            $winPEOCLangPath = Join-Path $this.WinPEOCPath $lang
            $cabs = Get-ChildItem $winPEOCLangPath -name

            # Install lp.cab cab
            $lpPath = Join-Path $winPEOCLangPath "lp.cab"
            Out-Log "Install LangPack $lpPath to $envName"
            Install-Package $mountPoint $lpPath

            # Install OC cab
            Foreach ($package in $winPEInstalledOC) {

                if ( ($package.PackageState -eq "Installed") `
                        -and ($package.PackageName.startsWith("WinPE-")) `
                        -and ($package.ReleaseType -eq "FeaturePack") ) {

                    $index = $package.PackageName.IndexOf("-Package")
                    if ($index -ge 0) {
                        $ocCab = $package.PackageName.Substring(0, $index) + "_" + $lang + ".cab"

                        if ($cabs.Contains($ocCab)) {
                            $ocCabPath = Join-Path $winPEOCLangPath $ocCab
                            Out-Log "Install LangPack $ocCabPath to $envName"
                            Install-Package $mountPoint $ocCabPath
                        }
                    }
                }
            }
        }
    }

    [bool]PatchWinPE() {

        if (-not $this.WinPELang) {
            return $true
        }

        $mountPoint = Join-Path $this.WorkingPath $([Constants]::WINPE_MOUNT)
        $imageNumber = Get-ImageTotalEdition $this.BootWimPath

        For ($index = 1; $index -le $imageNumber; $index++) {
            $envName = "$( [Constants]::WINPE )[$index]"
            try {
                Out-Log "Mount Image $( $this.BootWimPath ) Index $index to $mountPoint" -level $([Constants]::LOG_DEBUG)
                Mount-Image $this.BootWimPath $index $mountPoint

                $this.AddFontSupportToWinPE($mountPoint, $envName)
                $this.AddTTSToWinPE($mountPoint, $envName)
                $this.AddOcLP($mountPoint, $envName)
                $this.GenLangIni($mountPoint, $envName)

                Out-Log "Dismount Image $mountPoint and commit changes" -level $([Constants]::LOG_DEBUG)
                Dismount-CommitImage $mountPoint
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
        return $True
    }

    [bool]PatchWinRE() {
        $mountPoint = Join-Path $this.WorkingPath $([Constants]::WINRE_MOUNT)
        $envName = [Constants]::WINRE

        try {
            Out-Log "Mount Image $( $this.WinREPath ) 1 to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.WinREPath 1 $mountPoint

            $this.AddOcLP($mountPoint, $envName)

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

    [bool]PatchMainOS() {
        $mountPoint = Join-Path $this.WorkingPath $([Constants]::INSTALL_MOUNT)
        $envName = [Constants]::MAIN_OS

        try {
            Out-Log "Mount Image $( $this.InstallWimPath ) $( $this.WimIndex ) to $mountPoint" -level $([Constants]::LOG_DEBUG)
            Mount-Image $this.InstallWimPath $this.WimIndex $mountPoint

            Foreach ($lang in $this.LangList) {
                $LPFile = "Microsoft-Windows-Client-Language-Pack_" + $this.Arch + "_" + $lang + ".cab"
                $LPPath = Join-Path $this.MainOSLangPackPath $LPFile
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