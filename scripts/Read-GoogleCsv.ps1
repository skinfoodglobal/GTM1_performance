$ErrorActionPreference = "Stop"
. (Join-Path $PSScriptRoot "Get-DataFolderFile.ps1")

function Get-GoogleCsvDelimiter {
    param([string[]]$Lines)

    foreach ($line in $Lines) {
        if ($line -match "`t") { return "`t" }
        if ($line -match ',') { return ',' }
    }
    return "`t"
}

function Get-GoogleCsvEncoding {
    param([string]$Path)

    $stream = [System.IO.File]::OpenRead($Path)
    try {
        $b0 = $stream.ReadByte()
        $b1 = $stream.ReadByte()
        if ($b0 -eq 0xFF -and $b1 -eq 0xFE) { return [System.Text.Encoding]::Unicode }
        if ($b0 -eq 0xFE -and $b1 -eq 0xFF) { return [System.Text.Encoding]::BigEndianUnicode }
        if ($b0 -eq 0xEF -and $b1 -eq 0xBB) { return New-Object System.Text.UTF8Encoding $true }
    }
    finally {
        $stream.Dispose()
    }

    return New-Object System.Text.UTF8Encoding $false
}

function Merge-GoogleParsedDatasets {
    param($ParsedList)

    $columnList = New-Object System.Collections.Generic.List[string]
    $columnSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::Ordinal)
    $rowList = New-Object System.Collections.Generic.List[object]
    $titles = New-Object System.Collections.Generic.List[string]
    $ranges = New-Object System.Collections.Generic.List[string]

    foreach ($parsed in @($ParsedList)) {
        if (-not $parsed) { continue }
        $title = [string]$parsed.ReportTitle
        $range = [string]$parsed.ReportRange
        if ($title -and -not $titles.Contains($title)) { $titles.Add($title) }
        if ($range -and -not $ranges.Contains($range)) { $ranges.Add($range) }

        foreach ($col in @($parsed.Headers)) {
            $name = [string]$col
            if ([string]::IsNullOrWhiteSpace($name)) { continue }
            if ($columnSet.Add($name)) { $columnList.Add($name) }
        }

        foreach ($row in @($parsed.Rows)) {
            if ($row) { [void]$rowList.Add($row) }
        }
    }

    return [PSCustomObject]@{
        Headers     = $columnList.ToArray()
        Rows        = $rowList.ToArray()
        ReportTitle = if ($titles.Count) { $titles[0] } else { '' }
        ReportRange = ($ranges -join ' / ')
    }
}

function Get-EmptyGoogleCsvResult {
    param($Lines)

    return [PSCustomObject]@{
        Headers     = @()
        Rows        = @()
        ReportTitle = if ($Lines -and @($Lines).Count -gt 0) { ([string]$Lines[0]).Trim() } else { '' }
        ReportRange = if ($Lines -and @($Lines).Count -gt 1) { ([string]$Lines[1]).Trim() } else { '' }
    }
}

function Read-GoogleCsvRows {
    param(
        [string]$Path,
        [ValidateSet('keyword', 'demandgen')]
        [string]$DatasetKind = 'keyword'
    )

    $tempCopy = Join-Path $env:TEMP ("google_csv_" + [guid]::NewGuid().ToString("N") + ".csv")
    try {
        Copy-Item -LiteralPath $Path -Destination $tempCopy -Force
        $encoding = Get-GoogleCsvEncoding -Path $tempCopy
        $rawLines = [System.IO.File]::ReadAllLines($tempCopy, $encoding)
    }
    finally {
        if (Test-Path -LiteralPath $tempCopy) {
            Remove-Item -LiteralPath $tempCopy -Force -ErrorAction SilentlyContinue
        }
    }

    if (-not $rawLines -or $rawLines.Count -lt 3) {
        return Get-EmptyGoogleCsvResult -Lines $rawLines
    }

    $dataLines = @($rawLines | Select-Object -Skip 2 | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($dataLines.Count -lt 1) {
        return Get-EmptyGoogleCsvResult -Lines $rawLines
    }

    $delimiter = Get-GoogleCsvDelimiter -Lines $dataLines
    $headerLine = $dataLines[0]
    $headers = @($headerLine.Split($delimiter) | ForEach-Object { ([string]$_).Trim().TrimStart([char]0xFEFF) })

    $dateHeader = if ($headers.Count -gt 0) { $headers[0] } else { '' }
    $primaryHeader = switch ($DatasetKind) {
        'demandgen' {
            if ($headers.Count -gt 2) { $headers[2] } else { '' }
        }
        default {
            if ($headers.Count -gt 1) { $headers[1] } else { '' }
        }
    }

    $rows = New-Object System.Collections.Generic.List[object]
    for ($i = 1; $i -lt $dataLines.Count; $i++) {
        $parts = @($dataLines[$i].Split($delimiter))
        if ($parts.Count -eq 0) { continue }

        $record = [ordered]@{}
        for ($c = 0; $c -lt $headers.Count; $c++) {
            $header = $headers[$c]
            if ([string]::IsNullOrWhiteSpace($header)) { continue }
            $value = if ($c -lt $parts.Count) { ([string]$parts[$c]).Trim().Trim('"') } else { '' }
            $record[$header] = $value
        }

        $dateVal = if ($dateHeader) { [string]$record[$dateHeader] } else { '' }
        $primaryVal = if ($primaryHeader) { [string]$record[$primaryHeader] } else { '' }
        if ([string]::IsNullOrWhiteSpace($dateVal) -or [string]::IsNullOrWhiteSpace($primaryVal)) { continue }

        [void]$rows.Add($record)
    }

    return New-Object PSObject -Property @{
        Headers     = $headers
        Rows        = $rows.ToArray()
        ReportTitle = if ($rawLines.Count -gt 0) { ([string]$rawLines[0]).Trim() } else { '' }
        ReportRange = if ($rawLines.Count -gt 1) { ([string]$rawLines[1]).Trim() } else { '' }
    }
}

function Export-GoogleDatasetPayload {
    param(
        $Parsed,
        [string]$SourceName
    )

    $columns = @()
    if ($Parsed.Headers) {
        $columns = @($Parsed.Headers | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    }
    $rows = @()
    if ($null -ne $Parsed.Rows) {
        $rows = @($Parsed.Rows | Where-Object { $_ })
    }
    return [ordered]@{
        sourceFile  = $SourceName
        reportTitle = [string]$Parsed.ReportTitle
        reportRange = [string]$Parsed.ReportRange
        columns     = $columns
        rows        = $rows
    }
}

function Export-GoogleDataJson {
    param(
        [string]$RootDir,
        [string]$OutputPath
    )

    $keywordSources = @(Get-UpdateSourceFiles -RootDir $RootDir -Keyword "Google_Daily_RAW_keyword_cursor" -Extensions @('.csv'))
    $demandGenSources = @(Get-UpdateSourceFiles -RootDir $RootDir -Keyword "Google_Daily_RAW_demandgen_cursor" -Extensions @('.csv'))

    $emptyParsed = [PSCustomObject]@{ Headers = @(); Rows = @(); ReportTitle = ''; ReportRange = '' }
    $keywordParsedList = @()
    $keywordSourceName = ''

    if ($keywordSources.Count -eq 0) {
        Write-Warning "UPDATE 폴더에 Google_Daily_RAW_keyword_cursor 키워드 csv 파일이 없습니다. 빈 데이터(0)로 처리합니다."
        $keywordParsedList = @($emptyParsed)
    }
    else {
        $keywordSourceName = ($keywordSources | ForEach-Object { $_.Name }) -join ', '
        foreach ($source in $keywordSources) {
            Write-Host "Google keyword source=$($source.Name)"
            $keywordParsedList += , (Read-GoogleCsvRows -Path $source.Path -DatasetKind 'keyword')
        }
        $mergedKeyword = Merge-GoogleParsedDatasets -ParsedList $keywordParsedList
        if (@($mergedKeyword.Rows).Count -eq 0) {
            Write-Warning "Google 검색어 CSV에 데이터 행이 없습니다. 빈 데이터(0)로 처리합니다: $keywordSourceName"
        }
        $keywordParsedList = @($mergedKeyword)
    }

    $keywordDataset = Export-GoogleDatasetPayload -Parsed $keywordParsedList[0] -SourceName $keywordSourceName

    $demandGenDataset = $null
    if ($demandGenSources.Count -eq 0) {
        Write-Warning "UPDATE 폴더에 Google_Daily_RAW_demandgen_cursor csv 파일이 없습니다."
    }
    else {
        $demandGenParsedList = @()
        foreach ($source in $demandGenSources) {
            Write-Host "Google demandGen source=$($source.Name)"
            $demandGenParsedList += , (Read-GoogleCsvRows -Path $source.Path -DatasetKind 'demandgen')
        }
        $mergedDemandGen = Merge-GoogleParsedDatasets -ParsedList $demandGenParsedList
        if (@($mergedDemandGen.Rows).Count -gt 0) {
            $demandGenDataset = Export-GoogleDatasetPayload -Parsed $mergedDemandGen -SourceName (($demandGenSources | ForEach-Object { $_.Name }) -join ', ')
        }
        else {
            Write-Warning "Google 디맨드젠 CSV 파일에서 읽을 수 있는 데이터 행이 없습니다."
        }
    }

    $payload = [ordered]@{
        sourceFiles = [ordered]@{
            keyword   = $keywordSourceName
            demandGen = if ($demandGenDataset) { $demandGenDataset.sourceFile } else { $null }
        }
        sourceFile  = $keywordDataset.sourceFile
        reportTitle = $keywordDataset.reportTitle
        reportRange = $keywordDataset.reportRange
        updatedAt   = (Get-Date).ToString("yyyy-MM-ddTHH:mm:ss")
        columns     = $keywordDataset.columns
        rows        = $keywordDataset.rows
        demandGen   = if ($demandGenDataset) { $demandGenDataset } else { $null }
    }

    if ($null -eq $payload.rows) { $payload.rows = @() }
    if ($null -eq $payload.columns) { $payload.columns = @() }

    $json = $payload | ConvertTo-Json -Depth 8 -Compress
    $json = $json -replace '"rows":\s*""', '"rows":[]'
    $json = $json -replace '"columns":\s*""', '"columns":[]'
    $utf8NoBom = New-Object System.Text.UTF8Encoding $false
    [System.IO.File]::WriteAllText($OutputPath, $json, $utf8NoBom)

    $jsPath = [System.IO.Path]::ChangeExtension($OutputPath, ".js")
    [System.IO.File]::WriteAllText($jsPath, "window.GOOGLE_DATA = $json;", $utf8NoBom)

    return $payload
}
