<#
.SYNOPSIS
Spotify Data Extractor
By @doublemover

.DESCRIPTION
Extracts and analyzes your Spotify data export.
Accepts either a ZIP file or an already-extracted folder.
PS7 Required, has a shitty PS5 Fallback
Optionally outputs CSV for excel, graphing etc

.OUTPUT
Generates:
- account_info.json
- library.json
- playlists.json
- inferences.json
- search_history.json
- wrapped.json
- music_stats.json
- [Optional] top_artists_overall.csv, top_tracks_overall.csv, listening_by_year.csv, listening_by_month.csv

.PARAMETER InputPath
Path to ZIP file OR extracted folder.

.PARAMETER GenerateCSV
If specified, generates additional CSV summary files.

.EXAMPLE
pwsh -File .\spotify-data-extractor-ps7-enhanced.ps1 -InputPath my_spotify_data.zip -GenerateCSV

pwsh -File .\spotify-data-extractor-ps7-enhanced.ps1 -InputPath "C:\ExtractedSpotifyData" -GenerateCSV

#>

param(
  [Parameter(Mandatory=$true)]
  [string]$InputPath,

  [switch]$GenerateCSV
)

$PSStyle.Progress.MaxWidth = 65
$psstyle.Progress.Style = "$($PSStyle.Foreground.FromRgb(29, 185, 84))"


# --- Fancy Line ---
function Write-Host-FancyLine {
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

Write-Host-FancyLine
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Spotify Data Extractor" -NoNewline
Write-Host " v0.1" -NoNewline -ForegroundColor White 
Write-Host "$fancyGap      $($PSStyle.Foreground.FromRgb(44, 44, 44))by" -NoNewline
Write-Host "$($PSStyle.Foreground.FromRgb(152, 185, 20)) @doublemover" -NoNewline
Write-Host " " -BackgroundColor Black
Write-Host-FancyLine

# --- PowerShell version check ---
$IsPS7OrHigher = $PSVersionTable.PSVersion.Major -ge 7

if (-not $IsPS7OrHigher) {
    Write-Warning "You are running PowerShell $($PSVersionTable.PSVersion). This script is optimized for PowerShell 7+."
    Write-Warning "For best performance, upgrade PowerShell: winget install --id Microsoft.Powershell --source winget"
    Write-Warning "Running fallback compatibility path..."
} else {

}


if (-not $IsPS7OrHigher) {
# --- PS5 Fallback Path ---
# Only generate basic music_stats.json (no CSV, no progress bars)

# Determine input type
if ((Test-Path $InputPath -PathType Leaf) -and ($InputPath.ToLower().EndsWith(".zip"))) {
    Write-Host "$fancyGap Input is ZIP file. Extracting..." -ForegroundColor DarkCyan
    $extractDir = Join-Path $env:TEMP ("spotifydata_" + [guid]::NewGuid())
    Expand-Archive -Path $InputPath -DestinationPath $extractDir -Force
    $basePath = Join-Path $extractDir "Spotify Account Data"
    $cleanupExtract = $true
}
elseif (Test-Path $InputPath -PathType Container) {
    Write-Host "$fancyGap Input is extracted folder." -ForegroundColor DarkCyan
    $basePath = Join-Path $InputPath "Spotify Account Data"
    if (-not (Test-Path $basePath)) {
        throw "Spotify Account Data folder not found inside $InputPath"
    }
    $cleanupExtract = $false
}
else {
    throw "Invalid InputPath: must be a .zip file or a folder."
}

# --- Determine username for output folder ---
$userdataPath = Join-Path $basePath "Userdata.json"
$username = "SpotifyData_" + (Get-Date -Format "yyyyMMdd_HHmmss")

if (Test-Path $userdataPath) {
    try {
        $userdata = Get-Content $userdataPath -Raw | ConvertFrom-Json
        if ($userdata.username) {
            $username = $userdata.username
        } elseif ($userdata.email) {
            $username = $userdata.email.Split('@')[0]
        }
    } catch {
        Write-Warning "Could not parse Userdata.json to extract username. Using timestamp."
    }
} else {
    Write-Warning "Userdata.json not found. Using timestamp."
}

Write-Host "$fancyGap Extraction complete. Building stats for " -ForegroundColor Green
Write-Host "$username..." -NoNewline -ForegroundColor White

$outputFolder = Join-Path (Get-Location) ($username + "_spotifyData")
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

Write-Host "$fancyGap Output: " -ForegroundColor Cyan
Write-Host "$outputFolder..." -NoNewline -ForegroundColor White

$startTime = Get-Date

function Load-Json($path) {
  if ($null -ne $path -and (Test-Path $path)) {
      $txt = Get-Content $path -Raw
      if ($txt.Trim().Length -gt 0) {
          try {
              $data = $txt | ConvertFrom-Json
              if ($data -is [System.Collections.IEnumerable] -or $data -is [PSCustomObject]) {
                  return $data
              }
          } catch {
              return $null
          }
      }
  }
  return $null
}

# Load streaming history
$historyFiles = Get-ChildItem -Path $basePath -Filter "StreamingHistory_music_*.json"
$allEntries = @()
foreach ($hf in $historyFiles) {
    $entries = Load-Json $hf.FullName
    if ($entries) {
        $allEntries += $entries
    }
}

# Build basic music_stats.json
$artistDict = @{}
$trackDict = @{}
foreach ($entry in $allEntries) {
    $artist = $entry.artistName
    $track  = $entry.trackName
    $ms     = $entry.msPlayed

    if (-not $artistDict.ContainsKey($artist)) { $artistDict[$artist] = @{ msPlayed = 0; playCount = 0 } }
    $artistDict[$artist].msPlayed += $ms
    $artistDict[$artist].playCount += 1

    $tkey = "$artist - $track"
    if (-not $trackDict.ContainsKey($tkey)) { $trackDict[$tkey] = @{ artist = $artist; track = $track; msPlayed = 0; playCount = 0 } }
    $trackDict[$tkey].msPlayed += $ms
    $trackDict[$tkey].playCount += 1
}

$musicStats = @{
    totalArtists = $artistDict.Count
    totalTracks  = $trackDict.Count
    totalMsPlayed = 0
    totalPlayCount = 0
}

foreach ($v in $trackDict.Values) {
    $musicStats.totalMsPlayed += $v.msPlayed
    $musicStats.totalPlayCount += $v.playCount
}

$musicStats | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "music_stats.json") -Encoding UTF8

Write-Host "$fancyGap Done! Generated: music_stats.json"

if ($cleanupExtract) {

$generatedFiles = Get-ChildItem -Path $outputFolder -File
$generatedCount = $generatedFiles.Count

Write-Host-FancyLine
Write-Host "DONE! " -ForegroundColor Green
Write-Host ("$fancyGap Output " + $generatedCount + " files to:") -ForegroundColor White
Write-Host $outputFolder -ForegroundColor White
Write-Host ("$fancyGap Processed " + $totalHours + " hours, " + $totalMinutes + " minutes in " + ((Get-Date) - $startTime).TotalSeconds.ToString("0.0") + " seconds.") -ForegroundColor White
Write-Host-FancyLine

    Write-Host "Cleaning up temporary extracted files..."
    Remove-Item $extractDir -Recurse -Force
}
} else {
# Determine input type
if ((Test-Path $InputPath -PathType Leaf) -and ($InputPath.ToLower().EndsWith(".zip"))) {
    Write-Host "$fancyGap Input is zip. " -NoNewline -ForegroundColor Cyan
    Write-Host "Extracting. " -NoNewline -ForegroundColor DarkCyan
    $extractDir = Join-Path $env:TEMP ("spotifydata_" + [guid]::NewGuid())
    Expand-Archive -Path $InputPath -DestinationPath $extractDir -Force
    $basePath = Join-Path $extractDir "Spotify Account Data"
    $cleanupExtract = $true
}
elseif (Test-Path $InputPath -PathType Container) {
    Write-Host "$fancyGap Input is folder." -NoNewline -ForegroundColor Cyan
    $basePath = Join-Path $InputPath "Spotify Account Data"
    if (-not (Test-Path $basePath)) {
        throw "Spotify Account Data folder not found inside $InputPath"
    }
    $cleanupExtract = $false
}
else {
    throw "Invalid InputPath: must be a .zip file or a folder."
}

# --- Determine username for output folder ---
$userdataPath = Join-Path $basePath "Userdata.json"
$username = "SpotifyData_" + (Get-Date -Format "yyyyMMdd_HHmmss")

if (Test-Path $userdataPath) {
    try {
        $userdata = Get-Content $userdataPath -Raw | ConvertFrom-Json
        if ($userdata.username) {
            $username = $userdata.username
        } elseif ($userdata.email) {
            $username = $userdata.email.Split('@')[0]
        }
    } catch {
        Write-Host ""
        Write-Warning "Could not parse Userdata.json to extract username. Using timestamp."
    }
} else {
    Write-Host ""
    Write-Warning "Userdata.json not found. Using timestamp."
}

# Create output folder
$outputFolder = Join-Path (Get-Location) ($username + "_spotifyData")
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

Write-Host "Complete." -ForegroundColor Blue
Write-Host "$fancyGap Building stats for " -NoNewLine -ForegroundColor Magenta
Write-Host "$username" -ForegroundColor White

$startTime = Get-Date
$outputFolder = Join-Path (Get-Location) ($username + "_spotifyData")
if (-not (Test-Path $outputFolder)) {
    New-Item -ItemType Directory -Path $outputFolder | Out-Null
}

Write-Host-FancyLine

$startTime = Get-Date

function Load-Json($path) {
  if ($null -ne $path -and (Test-Path $path)) {
      $txt = Get-Content $path -Raw
      if ($txt.Trim().Length -gt 0) {
          try {
              $data = $txt | ConvertFrom-Json
              if ($data -is [System.Collections.IEnumerable] -or $data -is [PSCustomObject]) {
                  return $data
              }
          } catch {
              return $null
          }
      }
  }
  return $null
}

# --- Account Info ---
$accountFiles = @(
  "Identity.json", "Userdata.json", "UserAddress.json",
  "Identifiers.json", "Payments.json", "Follow.json", "S4XProfile.json"
)
$accountInfo = @{}
foreach ($fname in $accountFiles) {
  $fpath = Join-Path $basePath $fname
  $data = Load-Json $fpath
  if ($data) { $accountInfo[$fname -replace '\.json$',''] = $data }
}
$accountInfo | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "account_info.json") -Encoding UTF8

# --- Library, Playlists, Inferences, Search, Wrapped ---
$library = Load-Json (Join-Path $basePath "YourLibrary.json")
$library | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "library.json") -Encoding UTF8

$playlistFiles = Get-ChildItem -Path $basePath -Filter "Playlist*.json"
$playlists = [System.Collections.Generic.List[object]]::new()
foreach ($pf in $playlistFiles) {
    $data = Load-Json $pf.FullName
    if ($data) { $playlists.Add($data) }
}
$playlists | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "playlists.json") -Encoding UTF8

$inferences = Load-Json (Join-Path $basePath "Inferences.json")
$inferences | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "inferences.json") -Encoding UTF8

$searches = Load-Json (Join-Path $basePath "SearchQueries.json")
$searches | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "search_history.json") -Encoding UTF8

$wrappedFiles = Get-ChildItem -Path $basePath -Filter "Wrapped*.json"
$wrappedByYear = @{}
foreach ($wf in $wrappedFiles) {
  $yr = ($wf.Name -replace '\D','')
  $data = Load-Json $wf.FullName
  if ($data) {
      $wrappedByYear["$yr"] = @{}
      foreach ($k in $data.PSObject.Properties.Name) {
          $wrappedByYear["$yr"][$k] = $data.$k
      }
  }
}
$wrappedByYear | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "wrapped.json") -Encoding UTF8

# --- Streaming History (Sequential Load with Progress) ---
$historyFiles = Get-ChildItem -Path $basePath -Filter "StreamingHistory_music_*.json"

$allEntries = [System.Collections.Generic.List[object]]::new()
$total = $historyFiles.Count
$counter = 0

foreach ($hf in $historyFiles) {
    $counter++
    Write-Progress -Activity "Loading Streaming History" -Status "$counter / $total ($($hf.Name))" -PercentComplete (($counter / $total) * 100)

    $entries = Get-Content $hf.FullName -Raw | ConvertFrom-Json
    if ($entries) {
        foreach ($e in $entries) { $allEntries.Add($e) }
    }
}
Write-Progress -Activity "Loading Streaming History" -Completed

# Enrich each play with year/month
foreach ($entry in $allEntries) {
  if ($entry.endTime) {
      $dt = [datetime]::ParseExact($entry.endTime, "yyyy-MM-dd HH:mm", $null)
      $entry | Add-Member -NotePropertyName "year" -NotePropertyValue $dt.Year
      $entry | Add-Member -NotePropertyName "month" -NotePropertyValue $dt.Month
      $entry | Add-Member -NotePropertyName "yearMonth" -NotePropertyValue ("{0}-{1}" -f $dt.Year, $dt.Month)
  }
}

# --- Organize for stats ---
$musicStats = @{
  totalArtists   = 0
  totalTracks    = 0
  totalMsPlayed  = 0
  totalPlayCount = 0
  topArtistsByMonth = @{}
  topTracksByMonth = @{}
  yearlyBreakdown = @{}
}

# Group by yearMonth
$entriesByMonth = @{}
foreach ($entry in $allEntries) {
  $ym = $entry.yearMonth
  if ($null -ne $ym) {
      if (-not $entriesByMonth.ContainsKey($ym)) {
          $entriesByMonth[$ym] = [System.Collections.Generic.List[object]]::new()
      }
      $entriesByMonth[$ym].Add($entry)
  }
}

# --- Compute topArtistsByMonth and topTracksByMonth with progress ---
$monthKeys = $entriesByMonth.Keys
$totalMonths = $monthKeys.Count
$counterMonths = 0

foreach ($ym in $monthKeys) {
    $counterMonths++
    Write-Progress -Activity "Building Top Artists/Tracks By Month" -Status "$counterMonths / $totalMonths ($ym)" -PercentComplete (($counterMonths / $totalMonths) * 100)

    $artistMap = @{}
    $trackMap = @{}
    foreach ($entry in $entriesByMonth[$ym]) {
        $artist = $entry.artistName
        $track = $entry.trackName
        $ms = $entry.msPlayed

        if (-not $artistMap.ContainsKey($artist)) { $artistMap[$artist] = @{ msPlayed = 0; playCount = 0 } }
        $artistMap[$artist].msPlayed += $ms
        $artistMap[$artist].playCount += 1

        if (-not $trackMap.ContainsKey($track)) { $trackMap[$track] = @{ msPlayed = 0; playCount = 0 } }
        $trackMap[$track].msPlayed += $ms
        $trackMap[$track].playCount += 1
    }

    $topA = $artistMap.GetEnumerator() | Sort-Object { $_.Value.msPlayed } -Descending | Select-Object -First 10 | ForEach-Object {
        @{ artist = $_.Key; msPlayed = $_.Value.msPlayed; playCount = $_.Value.playCount }
    }
    $musicStats.topArtistsByMonth[$ym] = $topA

    $topT = $trackMap.GetEnumerator() | Sort-Object { $_.Value.msPlayed } -Descending | Select-Object -First 10 | ForEach-Object {
        @{ track = $_.Key; msPlayed = $_.Value.msPlayed; playCount = $_.Value.playCount }
    }
    $musicStats.topTracksByMonth[$ym] = $topT
}
Write-Progress -Activity "Building Top Artists/Tracks By Month" -Completed

# --- Build yearlyBreakdown with progress ---
$yearKeys = ($allEntries | Select-Object -ExpandProperty year | Sort-Object -Unique)
$totalYears = $yearKeys.Count
$counterYears = 0

foreach ($year in $yearKeys) {
    $counterYears++
    Write-Progress -Activity "Building Yearly Breakdown" -Status "$counterYears / $totalYears ($year)" -PercentComplete (($counterYears / $totalYears) * 100)

    $entriesThisYear = $allEntries | Where-Object { $_.year -eq $year }

    if (-not $musicStats.yearlyBreakdown.ContainsKey("$year")) {
        $musicStats.yearlyBreakdown["$year"] = @{ artists = @{} }
    }
    $artistBlock = $musicStats.yearlyBreakdown["$year"].artists

    foreach ($entry in $entriesThisYear) {
        $artist = $entry.artistName
        $track  = $entry.trackName
        $ms     = $entry.msPlayed
        $endTime = $entry.endTime

        if (-not $artistBlock.ContainsKey($artist)) { $artistBlock[$artist] = @{ tracks = @{}; msPlayed = 0; playCount = 0 } }
        $artistBlock[$artist].msPlayed += $ms
        $artistBlock[$artist].playCount += 1

        $trackBlock = $artistBlock[$artist].tracks
        if (-not $trackBlock.ContainsKey($track)) { $trackBlock[$track] = @{ msPlayed = 0; playCount = 0; endTimes = [System.Collections.Generic.List[string]]::new() } }
        $trackBlock[$track].msPlayed += $ms
        $trackBlock[$track].playCount += 1
        $trackBlock[$track].endTimes.Add($endTime)
    }
}
Write-Progress -Activity "Building Yearly Breakdown" -Completed

# --- Totals ---
$artistDict = @{}
$trackDict = @{}
foreach ($entry in $allEntries) {
  $artist = $entry.artistName
  $track  = $entry.trackName
  $ms     = $entry.msPlayed

  if (-not $artistDict.ContainsKey($artist)) { $artistDict[$artist] = @{ msPlayed = 0; playCount = 0 } }
  $artistDict[$artist].msPlayed += $ms
  $artistDict[$artist].playCount += 1

  $tkey = "$artist - $track"
  if (-not $trackDict.ContainsKey($tkey)) { $trackDict[$tkey] = @{ artist = $artist; track = $track; msPlayed = 0; playCount = 0 } }
  $trackDict[$tkey].msPlayed += $ms
  $trackDict[$tkey].playCount += 1
}

$musicStats.totalArtists = $artistDict.Count
$musicStats.totalTracks  = $trackDict.Count

$musicStats.totalMsPlayed = 0
$musicStats.totalPlayCount = 0
foreach ($v in $trackDict.Values) {
    $musicStats.totalMsPlayed += $v.msPlayed
    $musicStats.totalPlayCount += $v.playCount
}

$musicStats | ConvertTo-Json -Depth 8 | Set-Content (Join-Path $outputFolder "music_stats.json") -Encoding UTF8


if ($GenerateCSV) {
# --- Optional CSV Exports ---
if ($GenerateCSV) {
    Write-Host "Generating CSV summary files..."

    $topOverallArtists = $artistDict.GetEnumerator() | Sort-Object { $_.Value.msPlayed } -Descending | Select-Object -First 50 | ForEach-Object {
        [PSCustomObject]@{
            Artist    = $_.Key
            MsPlayed  = $_.Value.msPlayed
            PlayCount = $_.Value.playCount
        }
    }
    $topOverallArtists | Export-Csv -Path (Join-Path $outputFolder "top_artists_overall.csv") -NoTypeInformation -Encoding UTF8

    $topOverallTracks = $trackDict.GetEnumerator() | Sort-Object { $_.Value.msPlayed } -Descending | Select-Object -First 50 | ForEach-Object {
        [PSCustomObject]@{
            Track     = $_.Value.track
            Artist    = $_.Value.artist
            MsPlayed  = $_.Value.msPlayed
            PlayCount = $_.Value.playCount
        }
    }
    $topOverallTracks | Export-Csv -Path (Join-Path $outputFolder "top_tracks_overall.csv") -NoTypeInformation -Encoding UTF8

    $listeningByYear = $musicStats.yearlyBreakdown.Keys | ForEach-Object {
        [PSCustomObject]@{
            Year      = $_
            MsPlayed  = ($musicStats.yearlyBreakdown[$_].artists.Values | ForEach-Object { $_.msPlayed } | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
            PlayCount = ($musicStats.yearlyBreakdown[$_].artists.Values | ForEach-Object { $_.playCount } | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
        }
    }
    $listeningByYear | Export-Csv -Path (Join-Path $outputFolder "listening_by_year.csv") -NoTypeInformation -Encoding UTF8

    $listeningByMonth = $musicStats.topArtistsByMonth.Keys | ForEach-Object {
        [PSCustomObject]@{
            YearMonth = $_
            MsPlayed  = ($musicStats.topArtistsByMonth[$_] | ForEach-Object { $_.msPlayed } | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
            PlayCount = ($musicStats.topArtistsByMonth[$_] | ForEach-Object { $_.playCount } | Measure-Object -Sum | Select-Object -ExpandProperty Sum)
        }
    }
    $listeningByMonth | Export-Csv -Path (Join-Path $outputFolder "listening_by_month.csv") -NoTypeInformation -Encoding UTF8
    }
}

# --- Runtime summary ---
$endTime = Get-Date
$durationSeconds = ($endTime - $startTime).TotalSeconds

# Find first and last endTime
$validEndTimes = $allEntries | Where-Object { $_.endTime } | Sort-Object endTime
if ($validEndTimes.Count -gt 0) {
    $firstEndTime = [datetime]::ParseExact($validEndTimes[0].endTime, "yyyy-MM-dd HH:mm", $null)
    $lastEndTime  = [datetime]::ParseExact($validEndTimes[-1].endTime, "yyyy-MM-dd HH:mm", $null)
    $span = $lastEndTime - $firstEndTime

    # Build human readable span
    $totalDays = [math]::Floor($span.TotalDays)
    $totalWeeks = [math]::Floor($totalDays / 7)
    $totalMonths = [math]::Floor($totalDays / 30.44)
    $totalYears = [math]::Floor($totalDays / 365.25)

    # Total listening time
    $totalHours = [math]::Floor($musicStats.totalMsPlayed / 1000 / 60 / 60)
    $totalMinutes = [math]::Floor(($musicStats.totalMsPlayed / 1000 / 60) % 60)

    Write-Host "                             SUMMARY                             " -ForegroundColor White
    Write-Host-FancyLine

    Write-Host "$fancyGap Counted " -NoNewline -ForegroundColor Blue
    Write-Host ("{0}" -f $musicStats.totalPlayCount) -NoNewline -ForegroundColor White 
    Write-Host " played tracks from " -NoNewline -ForegroundColor Blue
    Write-Host ("{0}" -f $musicStats.totalArtists) -NoNewline -ForegroundColor White
    Write-Host " artists" -ForegroundColor Blue
    Write-Host "$fancyGap Total listening time: " -NoNewline -ForegroundColor DarkCyan
    if ($totalHours -gt 0) {
        Write-Host ("{0}" -f $totalHours) -NoNewline -ForegroundColor White
        Write-Host " hours " -NoNewline -ForegroundColor DarkCyan
    }
    Write-Host ("{0}" -f $totalMinutes) -NoNewline -ForegroundColor White
    Write-Host " minutes" -ForegroundColor DarkCyan
    Write-Host "$fancyGap Timespan: " -NoNewline  -ForegroundColor DarkMagenta
    if ($totalYears -gt 0) {
        Write-Host ("{0}" -f $totalYears) -NoNewline -ForegroundColor White
        Write-Host " years " -NoNewline -ForegroundColor DarkMagenta
    }
    if ($totalMonths -gt 0) {
        Write-Host ("{0}" -f $totalMonths) -NoNewline -ForegroundColor White 
        Write-Host " months " -NoNewline -ForegroundColor DarkMagenta
    }
    if ($totalWeeks -gt 0) {
        Write-Host ("{0}" -f $totalWeeks) -NoNewline -ForegroundColor White
        Write-Host " weeks " -NoNewline -ForegroundColor DarkMagenta
    }
    Write-Host ("{0}" -f $totalDays) -NoNewline -ForegroundColor White
    Write-Host " days" -ForegroundColor DarkMagenta

    Write-Host-FancyLine
} else {
    Write-Host-FancyLine
    Write-Host "                             SUMMARY                             " -ForegroundColor White
    Write-Host-FancyLine
    Write-Host "$fancyGap No valid endTime fields found in streaming history." -ForegroundColor Red
    Write-Host ("$fancyGap Processing completed in {0:N1} seconds" -f $durationSeconds) -ForegroundColor White
    Write-Host-FancyLine
}

    if ($cleanupExtract) {
        $generatedFiles = Get-ChildItem -Path $outputFolder -File
        $generatedCount = $generatedFiles.Count

        Write-Host "$fancyGap Output " -NoNewLine -ForegroundColor DarkCyan
        Write-Host $generatedCount -NoNewline -ForegroundColor White
        Write-Host " files to: " -NoNewLine -ForegroundColor DarkCyan
        Write-Host ("{0}_spotifyData" -f $username) -ForegroundColor Cyan
        Write-Host "$fancyGap Processed in " -NoNewline -ForegroundColor Magenta
        Write-Host ((Get-Date) - $startTime).TotalSeconds.ToString("0.0") -NoNewline -ForegroundColor White
        Write-Host " seconds." -ForegroundColor Magenta
        Write-Host-FancyLine

        Write-Host "$fancyGap Cleaning up temporary files..." -NoNewline -ForegroundColor DarkMagenta
        Remove-Item $extractDir -Recurse -Force
        Write-Host " $($PSStyle.Foreground.FromRgb(29, 185, 84))Done!"
        Write-Host-FancyLine
    }
}
