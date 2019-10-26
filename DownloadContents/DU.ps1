# -------------------------------------------------------------------------------------------
# Copyright (c) Microsoft Corporation. All rights reserved.
# Licensed under the MIT License. See License.txt in the project root for license information.
# -------------------------------------------------------------------------------------------


class DownloadDU: DownloadContents {

    [string]$searchPageTemplate = 'https://www.catalog.update.microsoft.com/Search.aspx?q={0}'
    [string]$downloadURL = 'http://www.catalog.update.microsoft.com/DownloadDialog.aspx'
    [string]$detailedInfoURL = 'https://www.catalog.update.microsoft.com/ScopedViewInline.aspx?updateid={0}'
    [int]$maxTryMonth = 12
    [string]$SSUPath
    [string]$LCUPath
    [string]$SafeOSPath
    [string]$SetupDUPath
    [string]$platform
    [string]$duReleaseMonth
    [string]$product
    [string]$version
    [switch]$showLinksOnly
    [switch]$forceSSL
    $DUInfoMapping

    DownloadDU([string]$SSUPath, [string]$LCUPath, [string]$SafeOSPath, [string]$SetupDUPath,
        [string]$platform, [string]$duReleaseMonth, [string]$product, [string]$version,
        [switch]$showLinksOnly, [switch]$forceSSL): base($platform, $product, $version, $showLinksOnly) {
        $this.SSUPath = $SSUPath
        $this.LCUPath = $LCUPath
        $this.SafeOSPath = $SafeOSPath
        $this.SetupDUPath = $SetupDUPath
        $this.duReleaseMonth = $duReleaseMonth
        $this.forceSSL = $forceSSL
        $this.DUInfoMapping = @{
            [DUType]::SSU     = @{title = "Servicing Stack Update"; path = $SSUPath };
            [DUType]::LCU     = @{title = "Cumulative Update"; path = $LCUPath };
            [DUType]::SafeOS  = @{title = "Dynamic Update"; path = $SafeOSPath; product = "Safe OS Dynamic Update" };
            [DUType]::SetupDU = @{title = "Dynamic Update"; path = $SetupDUPath; description = "SetupUpdate" };
        }
    }

    [bool]DownloadFileFromUrl([string]$url, [string]$destinationPath, [string]$fileName) {

        $filePath = Join-Path -Path $destinationPath -ChildPath $fileName

        try {
            if ((Get-Command Start-BitsTransfer -ErrorAction Ignore)) {
                Start-BitsTransfer -Source $url -Destination $filePath -ErrorAction stop
            }
            else {
                # Invoke-WebRequest is crazy slow for large downloads
                Write-Progress -Activity "Downloading $fileName" -Id 1
                (New-Object Net.WebClient).DownloadFile($url, $filePath)
                Write-Progress -Activity "Downloading $fileName" -Id 1 -Completed
            }
        }
        catch {
            Out-Log "Failed to download from $url" -level $([Constants]::LOG_ERROR)
            return $false
        }

        Out-Log "Download $fileName successfully"
        return $true
    }

    [string]RewriteURL([string]$url) {

        if ($this.forceSSL) {
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

    [string[]]GetKBDownloadLinksByGUID([string]$guid, [DUType]$DUType) {

        $post = @{ size = 0; updateID = $guid; uidInfo = $guid } | ConvertTo-Json -Compress
        $body = @{ updateIDs = "[$post]" }

        try {
            $links = Invoke-WebRequest -Uri $this.downloadURL -Method Post -Body $body -ErrorAction stop |
                Select-Object -ExpandProperty Content |
                Select-String -AllMatches -Pattern "(http[s]?\://download\.windowsupdate\.com\/[^\'\""]*)" |
                Select-Object -Unique
        }
        catch {
            Out-Log "Failed to get download link for $DUType" -level $([Constants]::LOG_ERROR)
            return @()
        }

        if (-not $links) {
            Out-Log "No download link available for $DUType" -level $([Constants]::LOG_ERROR)
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

    [int]FindTableColumnIndex($columns, [string]$Pattern) {
        $counter = 0
        foreach ($column in $columns) {
            if ($column.InnerHTML -like $Pattern) {
                break
            }
            $counter++
        }
        return $counter
    }

    [string]GetLatestGUID($url, [DUType]$DUType, [string]$curYearMonth) {

        try {
            Out-Log "Begin query $url." -level $([Constants]::LOG_DEBUG)
            $KBCatalogPage = Invoke-WebRequest -Uri $url -ErrorAction stop

            # Find the main table which contains all updates entry.
            $rows = $KBCatalogPage.ParsedHtml.getElementById('ctl00_catalogBody_updateMatches').getElementsByTagName('tr')
        }
        catch {
            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
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
                Out-Log "No headRow for $DUType" -level $([Constants]::LOG_DEBUG)
            }

            if (-not $dataRows) {
                Out-Log "No dataRows for $DUType" -level $([Constants]::LOG_DEBUG)
            }

            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $dateColumnIndex = $this.FindTableColumnIndex($columns, '*<SPAN>Last Updated</SPAN>*') # date column
        $titleColumnIndex = $this.FindTableColumnIndex($columns, '*<SPAN>Title</SPAN>*') # title column
        $productColumnIndex = $this.FindTableColumnIndex($columns, '*<SPAN>Products</SPAN>*') # product column

        if (($dateColumnIndex -eq $columns.count) -or ($titleColumnIndex -eq $columns.count) ) {
            Out-Log "Indexes(date:title:product) = $dateColumnIndex, $titleColumnIndex, $productColumnIndex for $DUType" -level $([Constants]::LOG_DEBUG)
            Out-Log "No $DUType found for $curYearMonth" -level $([Constants]::LOG_WARNING)
            return $null
        }

        $releaseDate = New-Object -TypeName DateTime -ArgumentList @(1, 1, 1)
        $GUID = $null

        # we will look for a row with the most recent release date.
        foreach ($row in $dataRows) {
            try {
                # filter Products
                $DUProduct = $this.DUInfoMapping.$DUType.product
                if ($DUProduct -and ($row.getElementsByTagName('td')[$productColumnIndex].innerHTML.Trim() -notlike "*$DUProduct*")) {
                    continue
                }

                # goToDetails contains update's GUID which we will use to request an update download page and get Detail page
                if ($row.getElementsByTagName('td')[$titleColumnIndex].innerHTML -match 'goToDetails\("(.+)"\);') {
                    $curGuid = $matches[1]

                    # filter Description
                    $detailURL = $this.detailedInfoURL -f $curGuid
                    $detailPage = Invoke-WebRequest -Uri $detailURL -ErrorAction stop

                    $DUdescription = $this.DUInfoMapping.$DUType.description
                    if ($DUdescription -and ($detailPage.ParsedHtml.getElementById('ScopedViewHandler_desc').innerHTML.Trim() -notlike "*$DUdescription*")) {
                        continue
                    }

                    $curDate = [datetime]::ParseExact($row.getElementsByTagName('td')[$dateColumnIndex].innerHTML.Trim(), 'd', $null)
                    if ($releaseDate -lt $curDate) {
                        # We assume that MS never publishes several versions of an update on the same day.
                        $releaseDate = $curDate
                        $GUID = $curGuid
                    }
                }
            }
            catch {
                Out-Log "Parse HTML failed. Detail: $( $_.Exception.Message )" -level $([Constants]::LOG_DEBUG)
            }
        }

        Out-Log "getLatestGUID = $GUID for $DUType" -level $([Constants]::LOG_DEBUG)
        return $GUID
    }

    [string]GetDUSearchShowString([DUType]$DUType, [string]$curYearMonth) {
        $DUTitle = $this.DUInfoMapping.$DUType.title
        $DUProduct = $this.DUInfoMapping.$DUType.product
        $DUDescription = $this.DUInfoMapping.$DUType.description

        $res = "(Date: $curYearMonth)"
        if ($DUTitle) {
            $res += "(Title: $DUTitle)"
        }

        if ($DUProduct) {
            $res += "(Product: $DUProduct)"
        }

        if ($DUDescription) {
            $res += "(Description: $DUDescription)"
        }
        return $res
    }

    [bool]DownloadByTypeAndDate([DUType]$DUType, [string]$curYearMonth) {

        $DUTitle = $this.DUInfoMapping.$DUType.title

        $searchCriteria = "$curYearMonth $DUTitle $($this.product) version $($this.version) $($this.platform)-based"
        Out-Log ("Begin searching for latest $DUType, Criteria = " + $this.GetDUSearchShowString($DUType, $curYearMonth))

        $url = $this.searchPageTemplate -f $searchCriteria

        $GUID = $this.GetLatestGUID($url, $DUType, $curYearMonth)
        if (!$GUID) { return $false }

        $downloadLinks = $this.GetKBDownloadLinksByGUID($GUID, $DUType)
        if (-not $downloadLinks) { return $false }
        Out-Log "Get download links: $downloadLinks" -level $([Constants]::LOG_DEBUG)

        if ($this.showLinksOnly) {
            Out-Log "$downloadLinks"
            return $true
        }

        foreach ($url in $downloadLinks) {
            if ($url -match '.+/(.+)$') {
                $this.DownloadFileFromUrl($url, $this.DUInfoMapping.$DUType.path, $Matches[1])
            }
        }

        return $true
    }

    [bool]DownloadByType([DUType]$DUType) {

        if ( ($DUType -ne [DUType]::SSU) -and ($DUType -ne [DUType]::LCU) -and ($DUType -ne [DUType]::SafeOS) -and ($DUType -ne [DUType]::SetupDU)) {
            return $false
        }

        $curDate = [DateTime]::ParseExact($this.duReleaseMonth, "yyyy-MM", $null)

        For ($i = 0; $i -le $this.maxTryMonth; $i++) {
            $curYearMonth = ("{0:d4}-{1:d2}" -f $curDate.Year, $curDate.Month)
            if ($this.DownloadByTypeAndDate($DUType, $curYearMonth) -eq $true) {
                return $true
            }
            $curDate = $curDate.date.AddMonths(-1)
        }
        return $false
    }

    [bool]Download() {

        [enum]::GetNames([DUType]) |

            ForEach-Object {
                if ($_ -ne [DUType]::Unknown) {
                    $this.DownloadByType($_)
                }
            }

        return $true
    }
}