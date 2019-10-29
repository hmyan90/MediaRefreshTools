# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------


class DownloadDU: DownloadContents {

    [string]$SearchPageTemplate = 'https://www.catalog.update.microsoft.com/Search.aspx?q={0}'
    [string]$DownloadURL = 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx'
    [string]$DetailedInfoURL = 'https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={0}'
    [int]$MaxTryMonth = 12
    [string]$SSUPath
    [string]$LCUPath
    [string]$SafeOSPath
    [string]$SetupDUPath
    [string]$DUReleaseMonth
    [switch]$ForceSSL
    $DUInfoMapping

    DownloadDU([string]$ssuPath, [string]$lcuPath, [string]$safeOSPath, [string]$setupDUPath,
        [string]$platform, [string]$duReleaseMonth, [string]$product, [string]$version,
        [switch]$showLinksOnly, [switch]$forceSSL): base($platform, $product, $version, $showLinksOnly) {
        $this.SSUPath = $ssuPath
        $this.LCUPath = $lcuPath
        $this.SafeOSPath = $safeOSPath
        $this.SetupDUPath = $setupDUPath
        $this.DUReleaseMonth = $duReleaseMonth
        $this.ForceSSL = $forceSSL
        $this.DUInfoMapping = @{
            [DUType]::SSU     = @{title = "Servicing Stack Update"; path = $ssuPath };
            [DUType]::LCU     = @{title = "Cumulative Update"; path = $lcuPath };
            [DUType]::SafeOS  = @{title = "Dynamic Update"; path = $safeOSPath; product = "Safe OS Dynamic Update" };
            [DUType]::SetupDU = @{title = "Dynamic Update"; path = $setupDUPath; description = "SetupUpdate" };
        }
    }

    [bool]DownloadFileFromURL([string]$url, [string]$destinationPath, [string]$fileName) {

        $filePath = Join-Path -Path $destinationPath -ChildPath $fileName

        try {
            # try use Start-BitsTransfer first, because it is faster for downloading large files
            if ((Get-Command Start-BitsTransfer -ErrorAction Ignore)) {
                Start-BitsTransfer -Source $url -Destination $filePath -ErrorAction stop
            }
            else {
                Write-Progress -Activity "Downloading $fileName" -Id 1
                (New-Object Net.WebClient).DownloadFile($url, $filePath)
                Write-Progress -Activity "Downloading $fileName" -Id 1 -Completed
            }
        }
        catch {
            Out-Log "Failed to download from $url" -level $([Constants]::LOG_ERROR)
            Out-Log "Exception detail: $( $_.Exception.Message )" -level $([Constants]::LOG_DEBUG)
            return $False
        }

        Out-Log "Download $fileName successfully"
        return $True
    }

    [string]RewriteURL([string]$url) {

        if ($this.ForceSSL) {
            if ($url -match '^https?:\/\/(.+)$') {
                return 'https://{0}' -f $Matches[1]
            }
            else {
                return ""
            }
        }
        else {
            return $url
        }
    }

    [string[]]GetKBDownloadLinksByGUID([string]$guid, [DUType]$duType) {

        $post = @{ size = 0; updateID = $guid; uidInfo = $guid } | ConvertTo-Json -Compress
        $body = @{ updateIDs = "[$post]" }

        try {
            $links = Invoke-WebRequest -Uri $this.DownloadURL -Method Post -Body $body -ErrorAction stop |
                Select-Object -ExpandProperty Content |
                Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" |
                Select-Object -Unique
        }
        catch {
            Out-Log "Failed to get download link for $duType" -level $([Constants]::LOG_ERROR)
            Out-Log "Exception detail: $( $_.Exception.Message )" -level $([Constants]::LOG_DEBUG)
            return @()
        }

        if (-not $links) {
            Out-Log "No download link available for $duType" -level $([Constants]::LOG_ERROR)
            return @()
        }

        $resLinks = @()
        foreach ($link in $links) {
            $tmp = $this.RewriteURL($link.matches.value)
            if ($tmp) {
                $resLinks += $tmp
            }
        }
        return $resLinks
    }

    [int]FindTableColumnIndex($columns, [string]$pattern) {
        $counter = 0
        foreach ($column in $columns) {
            if ($column.InnerHTML -like $pattern) {
                break
            }
            $counter++
        }
        return $counter
    }

    [string]GetLatestGUID($url, [DUType]$duType, [string]$curYearMonth) {

        try {
            Out-Log "Begin query $url." -level $([Constants]::LOG_DEBUG)
            $kbCatalogPage = Invoke-WebRequest -Uri $url -ErrorAction stop

            # Find the main table which contains all updates entry.
            $rows = $kbCatalogPage.ParsedHtml.getElementById('ctl00_catalogBody_updateMatches').getElementsByTagName('tr')
        }
        catch {
            Out-Log "Ignored exception detail: $( $_.Exception.Message )" -level $([Constants]::LOG_DEBUG)
            Out-Log "No $duType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $headerRow = $null
        $dataRows = @()
        foreach ($row in $rows) {
            if ($row.id -eq 'headerRow') {
                if (-not $headerRow) {
                    $headerRow = $row
                }
            }
            else {
                $dataRows += $row
            }
        }

        if ($headerRow -and $dataRows) {
            $columns = $headerRow.getElementsByTagName('td')
        }
        else {
            if (-not $headerRow) {
                Out-Log "No headRow for $duType" -level $([Constants]::LOG_DEBUG)
            }

            if (-not $dataRows) {
                Out-Log "No dataRows for $duType" -level $([Constants]::LOG_DEBUG)
            }

            Out-Log "No $duType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $dateColumnIndex = $this.FindTableColumnIndex($columns, '*<SPAN>Last Updated</SPAN>*') # date column
        $titleColumnIndex = $this.FindTableColumnIndex($columns, '*<SPAN>Title</SPAN>*') # title column
        $productColumnIndex = $this.FindTableColumnIndex($columns, '*<SPAN>Products</SPAN>*') # product column

        if (($dateColumnIndex -eq $columns.count) -or ($titleColumnIndex -eq $columns.count) ) {
            Out-Log "Indexes(date:title:product) = $dateColumnIndex, $titleColumnIndex, $productColumnIndex for $duType" -level $([Constants]::LOG_DEBUG)
            Out-Log "No $duType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $releaseDate = New-Object -TypeName DateTime -ArgumentList @(1, 1, 1)
        $guid = $null

        # we will look for a row with the most recent release date.
        foreach ($row in $dataRows) {
            try {
                # filter Products
                $duProduct = $this.DUInfoMapping.$duType.product
                if ($duProduct -and ($row.getElementsByTagName('td')[$productColumnIndex].innerHTML.Trim() -notlike "*$duProduct*")) {
                    continue
                }

                # goToDetails contains update's guid which we will use to request an update download page and get Detail page
                if ($row.getElementsByTagName('td')[$titleColumnIndex].innerHTML -match 'goToDetails\("(.+)"\);') {
                    $curGUID = $Matches[1]

                    # filter Description
                    $detailURL = $this.DetailedInfoURL -f $curGUID
                    $detailPage = Invoke-WebRequest -Uri $detailURL -ErrorAction stop

                    $duDescription = $this.DUInfoMapping.$duType.description
                    if ($duDescription -and ($detailPage.ParsedHtml.getElementById('ScopedViewHandler_desc').innerHTML.Trim() -notlike "*$duDescription*")) {
                        continue
                    }

                    $curDate = [DateTime]::ParseExact($row.getElementsByTagName('td')[$dateColumnIndex].innerHTML.Trim(), 'd', $null)
                    if ($releaseDate -lt $curDate) {
                        # We assume that MS never publishes several versions of an update on the same day.
                        $releaseDate = $curDate
                        $guid = $curGUID
                    }
                }
            }
            catch {
                Out-Log "Parse HTML failed. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_DEBUG)
            }
        }

        Out-Log "GetLatestGUID = $guid for $duType" -level $([Constants]::LOG_DEBUG)

        if (!$guid) {
            Out-Log "No $duType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
        }

        return $guid
    }

    [string]GetDUSearchShowString([DUType]$duType, [string]$curYearMonth) {
        $duTitle = $this.DUInfoMapping.$duType.title
        $duProduct = $this.DUInfoMapping.$duType.product
        $duDescription = $this.DUInfoMapping.$duType.description

        $res = "(Date: $curYearMonth)"
        if ($duTitle) {
            $res += "(Title: $duTitle)"
        }

        if ($duProduct) {
            $res += "(Product: $duProduct)"
        }

        if ($duDescription) {
            $res += "(Description: $duDescription)"
        }
        return $res
    }

    [bool]DownloadByTypeAndDate([DUType]$duType, [string]$curYearMonth) {

        $duTitle = $this.DUInfoMapping.$duType.title

        $searchCriteria = "$curYearMonth $duTitle $($this.Product) version $($this.Version) $($this.Platform)-based"
        Out-Log ("Begin searching for latest $duType, Criteria = " + $this.GetDUSearchShowString($duType, $curYearMonth))

        $url = $this.SearchPageTemplate -f $searchCriteria

        $guid = $this.GetLatestGUID($url, $duType, $curYearMonth)
        if (!$guid) { return $False }

        $downloadLinks = $this.GetKBDownloadLinksByGUID($guid, $duType)
        if (-not $downloadLinks) { return $False }
        Out-Log "Get download links: $downloadLinks" -level $([Constants]::LOG_DEBUG)

        if ($this.ShowLinksOnly) {
            Out-Log "$downloadLinks"
            return $True
        }

        foreach ($url in $downloadLinks) {
            if ($url -match '.+/(.+)$') {
                $this.DownloadFileFromURL($url, $this.DUInfoMapping.$duType.path, $Matches[1])
            }
        }

        return $True
    }

    [bool]DownloadByType([DUType]$duType) {

        if ( ($duType -ne [DUType]::SSU) -and ($duType -ne [DUType]::LCU) -and ($duType -ne [DUType]::SafeOS) -and ($duType -ne [DUType]::SetupDU)) {
            return $False
        }

        $curDate = [DateTime]::ParseExact($this.DUReleaseMonth, "yyyy-MM", $null)

        For ($i = 0; $i -le $this.MaxTryMonth; $i++) {
            $curYearMonth = ("{0:d4}-{1:d2}" -f $curDate.Year, $curDate.Month)
            if ($this.DownloadByTypeAndDate($duType, $curYearMonth) -eq $True) {
                return $True
            }
            $curDate = $curDate.date.AddMonths(-1)
        }
        return $False
    }

    [bool]Download() {

        [enum]::GetNames([DUType]) |

            ForEach-Object {
                if ($_ -ne [DUType]::Unknown) {
                    $this.DownloadByType($_)
                }
            }

        return $True
    }
}