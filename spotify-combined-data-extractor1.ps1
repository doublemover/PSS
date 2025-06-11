<#
.SYNOPSIS
    Combined Spotify Data Extractor for Extended + Account History.

.DESCRIPTION
    Processes one or two Spotify data ZIP files, enforces one Extended + one Account ZIP,
    outputs combined music stats and separate account data files.

.NOTES
    Requires PowerShell 7
#>

param (
    [Parameter(Mandatory = $true)]
    [string]$ZipFile1,

    [Parameter(Mandatory = $false)]
    [string]$ZipFile2,

    [Parameter(Mandatory = $false)]
    [string]$OutputFolder
)

# Validate PowerShell version
if ($PSVersionTable.PSVersion.Major -lt 7) {
    Write-Warning "PowerShell 7+ is required, exiting."
    exit
}

# Style
$PSStyle.Progress.MaxWidth = 65
$psstyle.Progress.Style = "$($PSStyle.Foreground.FromRgb(29, 185, 84))"

# --- Fancy Line ---
function Write-Host-FancyLine {
    if ($fancy -eq $false) {
        return
    }
    $bars = 64
    for ($i = 0; $i -le $bars; $i++) {
        $color = ((255/($bars*0.5))*($i*0.5))
        if ($i -gt ($bars/2)) {
            $color = 255 - ((255/($bars*0.5))*($i*0.5))
        }
        Write-Host "$($PSStyle.Foreground.FromRgb($color, $color, $color))=" -NoNewline
    }
    Write-Host ""
}

$fancyGap = "        "
$startTime = Get-Date
$processedCount = 0

Write-Host-FancyLine
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Spotify Data Extractor" -NoNewline
Write-Host " v0.2" -NoNewline -ForegroundColor White 
Write-Host "              $($PSStyle.Foreground.FromRgb(44, 44, 44))by" -NoNewline
Write-Host "$($PSStyle.Foreground.FromRgb(152, 185, 20)) @doublemover" -NoNewline
Write-Host " " -BackgroundColor Black
Write-Host-FancyLine

# Initialize global vars
$processedFiles = @()
$failedFiles = @()
$processedTimestamps = [System.Collections.Generic.HashSet[string]]::new()

# Create output folder
if (-not $OutputFolder) {
    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $OutputFolder = Join-Path $PSScriptRoot "SpotifyData_$timestamp"
}
New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
$errorLogPath = Join-Path $OutputFolder "error.log"

# Prepare temp extraction folders
$tempRoot = Join-Path $env:TEMP "SpotifyExtractorTemp"
$tempExtract1 = Join-Path $tempRoot "Extract1"
$tempExtract2 = Join-Path $tempRoot "Extract2"
New-Item -ItemType Directory -Path $tempExtract1 -Force | Out-Null
New-Item -ItemType Directory -Path $tempExtract2 -Force | Out-Null

function Write-ErrorLog {
    param ($Message)
    $Message | Add-Content -Path $errorLogPath
}

function Test-ZipFileIsValidSpotifyData {
    param ($ZipPath, $TempPath)

    if (-not (Test-Path $ZipPath -PathType Leaf)) { return "Invalid" }
    if (-not ($ZipPath.ToLower().EndsWith(".zip"))) { return "Invalid" }

    # Clean temp path first
    if (Test-Path $TempPath) {
        Remove-Item -Path $TempPath -Recurse -Force -ErrorAction SilentlyContinue
    }
    New-Item -ItemType Directory -Path $TempPath -Force | Out-Null

    # Extract to temp
    try {
        Expand-Archive -Path $ZipPath -DestinationPath $TempPath -Force
    } catch {
        return "Invalid"
    }

    # Now check extracted folders
    $folders = Get-ChildItem -Path $TempPath -Directory | Select-Object -ExpandProperty Name

    if ($folders -contains "Spotify Extended Streaming History") {
        return "ExtendedStreaming"
    } elseif ($folders -contains "Spotify Account Data") {
        return "AccountData"
    } else {
        return "Invalid"
    }
}

function Expand-ZipArchive {
    param ($ZipPath, $DestinationPath)

    try {
        Write-Host "$fancyGap Extracting " -NoNewLine -ForegroundColor Cyan
        Write-Host "$ZipPath" -ForegroundColor DarkCyan
        Expand-Archive -Path $ZipPath -DestinationPath $DestinationPath -Force
        Write-Host "$fancyGap Extraction complete." -ForegroundColor Cyan
        return $true
    } catch {
        $msg = "Failed to extract $ZipPath : $($_.Exception.Message)"
        Write-Error $msg
        Write-ErrorLog $msg
        $failedFiles += $ZipPath
        return $false
    }
}

function Get-SpotifyFolderStat {
    param ($FolderPath, [bool]$IsExtended)

    $trackDict = @{}
    $localProcessedFiles = @()
    $localFailedFiles = @()

    $jsonFiles = Get-ChildItem -Path $FolderPath -Recurse -Include "Streaming_History*.json","StreamingHistory*.json" | Sort-Object Name

    foreach ($file in $jsonFiles) {
        try {
            # Write-Host "$fancyGap Processing $($file.Name) ..."
            $jsonContent = Get-Content -Path $file.FullName -Raw
            $entries = $jsonContent | ConvertFrom-Json -ErrorAction Stop

            $playCount = 0
            foreach ($entry in $entries) {
                $playCount++

                $trackName = $entry.master_metadata_track_name
                $artistName = $entry.master_metadata_album_artist_name
                $albumName = $entry.master_metadata_album_album_name
                $ts = $entry.ts
                $msPlayed = $entry.ms_played

                if (-not $trackName -or -not $artistName -or -not $albumName -or -not $ts -or -not $msPlayed) {
                    continue
                }

                $timestampKey = "$artistName|$albumName|$trackName|$ts"

                if (-not $IsExtended -and $processedTimestamps.Contains($timestampKey)) {
                    continue
                }

                if ($IsExtended) {
                    $processedTimestamps.Add($timestampKey) | Out-Null
                }

                $trackKey = "$artistName|$albumName|$trackName"

                if (-not $trackDict.ContainsKey($trackKey)) {
                    $trackDict[$trackKey] = [ordered]@{
                        track_name  = $trackName
                        artist_name = $artistName
                        album_name  = $albumName
                        ms_played   = 0
                        plays       = 0
                        timestamps  = @()
                    }
                }

                $trackEntry = $trackDict[$trackKey]
                $trackEntry.ms_played += $msPlayed
                $trackEntry.plays += 1

                $timestampDetails = @{
                    ts = $ts
                    reason_start = $entry.reason_start
                    reason_end = $entry.reason_end
                }
                if ($entry.shuffle) { $timestampDetails.shuffle = $entry.shuffle }
                if ($entry.skipped) { $timestampDetails.skipped = $entry.skipped }
                if ($entry.incognito_mode) { $timestampDetails.incognito_mode = $entry.incognito_mode }
                if ($entry.offline) {
                    $timestampDetails.offline = $entry.offline
                    if ($entry.offline_timestamp) {
                        $timestampDetails.offline_timestamp = $entry.offline_timestamp
                    }
                }

                $trackEntry.timestamps += $timestampDetails
            }

            Write-Host -NoNewline "`r$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))[$($PSStyle.Foreground.FromRgb(255, 255, 255))$processedCount$($PSStyle.Foreground.FromRgb(29, 185, 84))] Processed $($PSStyle.Foreground.FromRgb(255, 255, 255))$playCount$($PSStyle.Foreground.FromRgb(29, 185, 84)) plays from $($file.Name).                                             "
            $processedCount += $playCount
            $localProcessedFiles += $file.FullName

        } catch {
            $msg = "Failed to process $($file.FullName): $($_.Exception.Message)"
            Write-Error $msg
            Write-ErrorLog $msg
            $localFailedFiles += $file.FullName
        }
    }
    Write-Host ""
    return @{
        trackDict = $trackDict
        processedFiles = $localProcessedFiles
        failedFiles = $localFailedFiles
    }
}

function Export-SpotifyAccountData {
    param ($FolderPath)

    $lib = Get-ChildItem -Path $FolderPath -Recurse -Filter "YourLibrary.json" | Select-Object -First 1
    if ($lib) {
        Copy-Item -Path $lib.FullName -Destination (Join-Path $OutputFolder "library.json") -Force
    }

    $plFiles = Get-ChildItem -Path $FolderPath -Recurse -Filter "Playlist*.json"
    $plCombined = @()
    foreach ($pl in $plFiles) {
        $plData = Get-Content -Path $pl.FullName -Raw | ConvertFrom-Json
        $plCombined += $plData
    }
    if ($plCombined.Count -gt 0) {
        $plCombined | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $OutputFolder "playlists.json") -Encoding UTF8
    }

    $searchFiles = Get-ChildItem -Path $FolderPath -Recurse -Filter "Search-Queries*.json"
    $searchCombined = @()
    foreach ($sf in $searchFiles) {
        $sfData = Get-Content -Path $sf.FullName -Raw | ConvertFrom-Json
        $searchCombined += $sfData
    }
    $finalSearch = @()
    $prevQuery = ""
    foreach ($s in $searchCombined | Sort-Object timestamp) {
        $q = $s.searchQuery
        if ($q -notmatch "^$prevQuery.*$") {
            $finalSearch += $s
            $prevQuery = $q
        }
    }
    if ($finalSearch.Count -gt 0) {
        $finalSearch | ConvertTo-Json -Depth 5 | Set-Content -Path (Join-Path $OutputFolder "search_history.json") -Encoding UTF8
    }

    $wrappedFiles = Get-ChildItem -Path $FolderPath -Recurse -Filter "Wrapped*.json"
    $wrappedOutput = @{}
    foreach ($wf in $wrappedFiles) {
        $wfData = Get-Content -Path $wf.FullName -Raw | ConvertFrom-Json
        if ($wf.Name -match "(\d{4})") {
            $year = $Matches[1]
            $wrappedOutput[$year] = $wfData
        }
    }
    if ($wrappedOutput.Count -gt 0) {
        $wrappedOutput | ConvertTo-Json -Depth 10 | Set-Content -Path (Join-Path $OutputFolder "wrapped.json") -Encoding UTF8
    }
}

# === MAIN PROCESSING ===

$zip1TempCheckPath = Join-Path $tempRoot "Zip1Check"
$zip2TempCheckPath = Join-Path $tempRoot "Zip2Check"

$zip1Type = Test-ZipFileIsValidSpotifyData -ZipPath $ZipFile1 -TempPath $zip1TempCheckPath
$zip2Type = if ($ZipFile2) { Test-ZipFileIsValidSpotifyData -ZipPath $ZipFile2 -TempPath $zip2TempCheckPath } else { "None" }

if ($zip1Type -eq "Invalid" -or ($ZipFile2 -and $zip2Type -eq "Invalid")) {
    throw "Invalid ZIP file(s) provided. Must be Extended Streaming or Account Data ZIP."
}

if ($ZipFile2) {
    $combo = @($zip1Type, $zip2Type)
    if (($combo -contains "ExtendedStreaming") -and ($combo -contains "AccountData")) {
        # OK
    } else {
        throw "When two ZIPs provided, one must be Extended Streaming and one must be Account Data."
    }
}

$combinedTrackDict = @{}

$extendedZip = if ($zip1Type -eq "ExtendedStreaming") { $ZipFile1 } else { $ZipFile2 }
if (Expand-ZipArchive -ZipPath $extendedZip -DestinationPath $tempExtract1) {
    $statsExt = Get-SpotifyFolderStat -FolderPath $tempExtract1 -IsExtended $true
    $combinedTrackDict = $statsExt.trackDict
    $processedFiles += $statsExt.processedFiles
    $failedFiles += $statsExt.failedFiles
}

Write-Host-FancyLine

if ($ZipFile2) {
    $accountZip = if ($zip1Type -eq "AccountData") { $ZipFile1 } else { $ZipFile2 }
    if (Expand-ZipArchive -ZipPath $accountZip -DestinationPath $tempExtract2) {
        $statsAcc = Get-SpotifyFolderStat -FolderPath $tempExtract2 -IsExtended $false
        foreach ($key in $statsAcc.trackDict.Keys) {
            if (-not $combinedTrackDict.ContainsKey($key)) {
                $combinedTrackDict[$key] = $statsAcc.trackDict[$key]
            } else {
                $combinedTrackDict[$key].ms_played += $statsAcc.trackDict[$key].ms_played
                $combinedTrackDict[$key].plays += $statsAcc.trackDict[$key].plays
                $combinedTrackDict[$key].timestamps += $statsAcc.trackDict[$key].timestamps
            }
        }
        $processedFiles += $statsAcc.processedFiles
        $failedFiles += $statsAcc.failedFiles
        Export-SpotifyAccountData -FolderPath $tempExtract2
    }
}

$organizedStats = @{}
foreach ($trackKey in $combinedTrackDict.Keys) {
    $t = $combinedTrackDict[$trackKey]
    $artist = $t.artist_name
    $album = $t.album_name
    $track = $t.track_name

    if (-not $organizedStats.ContainsKey($artist)) {
        $organizedStats[$artist] = @{
            plays = 0
            msplayed = 0
            albums = @{}
        }
    }
    $organizedStats[$artist].plays += $t.plays
    $organizedStats[$artist].msplayed += $t.ms_played

    if (-not $organizedStats[$artist].albums.ContainsKey($album)) {
        $organizedStats[$artist].albums[$album] = @{}
    }
    $organizedStats[$artist].albums[$album][$track] = @{
        ms_played = $t.ms_played
        plays = $t.plays
        timestamps = $t.timestamps
    }
}

$sortedArtists = $organizedStats.GetEnumerator() | Sort-Object { $_.Value.msplayed } -Descending

$finalOut = [ordered]@{}
foreach ($artist in $sortedArtists) {
    $finalOut[$artist.Key] = $artist.Value
}

$musicStatsPath = Join-Path $OutputFolder "music_stats.json"
$finalOut | ConvertTo-Json -Depth 12 | Set-Content -Path $musicStatsPath -Encoding UTF8

Remove-Item -Path $tempRoot -Recurse -Force -ErrorAction SilentlyContinue

Write-Host-FancyLine
Write-Host "$fancyGap Files processed successfully: $($PSStyle.Foreground.FromRgb(255, 255, 255))$($processedFiles.Count)" -ForegroundColor Cyan
Write-Host "$fancyGap Files failed: $($PSStyle.Foreground.FromRgb(255, 255, 255))$($failedFiles.Count)"  -ForegroundColor Cyan
if ($failedFiles.Count -gt 0) {
    Write-Host "$fancyGap Failed files written to error.log"  -ForegroundColor Cyan
}
Write-Host "$fancyGap Music stats written to: $($PSStyle.Foreground.FromRgb(255, 255, 255))SpotifyData_$timestamp"  -ForegroundColor Cyan
Write-Host "$fancyGap Processed in $(((Get-Date) - $startTime).TotalSeconds.ToString("0.0")) seconds" -ForegroundColor DarkCyan
Write-Host-FancyLine
