function Get-UpdateSourceFileStamp {
    param([string]$Name)

    if ([string]$Name -match '(20\d{6})') {
        return $Matches[1]
    }
    return '00000000'
}

function Get-UpdateSourceFiles {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootDir,

        [Parameter(Mandatory = $true)]
        [string]$Keyword,

        [string[]]$Extensions = @('.xlsx', '.csv')
    )

    $updateDir = Join-Path $RootDir "UPDATE"
    if (-not (Test-Path -LiteralPath $updateDir)) { return @() }

    $extSet = New-Object 'System.Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach ($ext in $Extensions) {
        $normalized = if ($ext.StartsWith('.')) { $ext } else { ".$ext" }
        [void]$extSet.Add($normalized)
    }

    $keywordLower = $Keyword.ToLowerInvariant()
    $files = @(
        Get-ChildItem -Path $updateDir -File -ErrorAction SilentlyContinue |
            Where-Object {
                $_ -and
                $_.Name -notmatch '^~\$' -and
                $extSet.Contains($_.Extension) -and
                $_.Name.ToLowerInvariant().Contains($keywordLower)
            } |
            Sort-Object @{ Expression = { Get-UpdateSourceFileStamp $_.Name }; Descending = $true },
                        @{ Expression = 'LastWriteTime'; Descending = $true },
                        @{ Expression = 'Name'; Descending = $true }
    )

    if ($files.Count -eq 0) { return @() }

    return @(
        $files | ForEach-Object {
            [PSCustomObject]@{
                Name = $_.Name
                Path = $_.FullName
                Keyword = $Keyword
            }
        }
    )
}

function Get-UpdateSourceFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootDir,

        [Parameter(Mandatory = $true)]
        [string]$Keyword,

        [string[]]$Extensions = @('.xlsx', '.csv')
    )

    $files = @(Get-UpdateSourceFiles -RootDir $RootDir -Keyword $Keyword -Extensions $Extensions)
    if ($files.Count -eq 0) { return $null }

    if ($files.Count -gt 1) {
        Write-Warning "UPDATE 폴더에서 '$Keyword' 키워드 파일이 $($files.Count)개입니다. 최신 파일을 사용합니다: $($files[0].Name)"
    }

    return $files[0]
}
