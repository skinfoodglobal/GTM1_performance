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
        $rawLines = [System.IO.File]::ReadAllLines($tempCopy, [System.Text.Encoding]::UTF8)
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

    $keywordSource = Get-UpdateSourceFile -RootDir $RootDir -Keyword "Google_Daily_RAW_keyword_cursor" -Extensions @('.csv')
    $demandGenSource = Get-UpdateSourceFile -RootDir $RootDir -Keyword "Google_Daily_RAW_demandgen_cursor" -Extensions @('.csv')

    $emptyParsed = [PSCustomObject]@{ Headers = @(); Rows = @(); ReportTitle = ''; ReportRange = '' }

    if (-not $keywordSource) {
        Write-Warning "UPDATE 폴더에 Google_Daily_RAW_keyword_cursor 키워드 csv 파일이 없습니다. 빈 데이터(0)로 처리합니다."
        $keywordParsed = $emptyParsed
        $keywordSourceName = ''
    }
    else {
        $keywordParsed = Read-GoogleCsvRows -Path $keywordSource.Path -DatasetKind 'keyword'
        $keywordSourceName = $keywordSource.Name
        if (@($keywordParsed.Rows).Count -eq 0) {
            Write-Warning "Google 검색어 CSV에 데이터 행이 없습니다. 빈 데이터(0)로 처리합니다: $keywordSourceName"
        }
    }

    $keywordDataset = Export-GoogleDatasetPayload -Parsed $keywordParsed -SourceName $keywordSourceName

    $demandGenDataset = $null
    if ($demandGenSource) {
        $demandGenParsed = Read-GoogleCsvRows -Path $demandGenSource.Path -DatasetKind 'demandgen'
        if (@($demandGenParsed.Rows).Count -gt 0) {
            $demandGenDataset = Export-GoogleDatasetPayload -Parsed $demandGenParsed -SourceName $demandGenSource.Name
        }
        else {
            Write-Warning "Google 디맨드젠 CSV 파일에서 읽을 수 있는 데이터 행이 없습니다: $($demandGenSource.Name)"
        }
    }
    else {
        Write-Warning "UPDATE 폴더에 Google_Daily_RAW_demandgen_cursor csv 파일이 없습니다."
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
