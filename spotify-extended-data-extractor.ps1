<#
.SYNOPSIS
Spotify Extended Extractor
By @doublemover

.DESCRIPTION
Extracts and analyzes your Spotify extended play data. Accepts a ZIP file.
Requires PS7

.OUTPUT
Generates:
- spotify_tracks_output.json
If there are values with missing URI:
- spotify_tracks_missing_uri.json

.PARAMETER ZipFile
Path to ZIP file

.PARAMETER Fast
Goes faster

.EXAMPLE
pwsh -File .\spotify-extended-data-extractor.ps1 .\my_extended_spotify_data.zip
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ZipFile,

    [switch]$Fast
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
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Spotify Extended Extractor" -NoNewline
Write-Host " v0.1" -NoNewline -ForegroundColor White 
Write-Host "          $($PSStyle.Foreground.FromRgb(44, 44, 44))by" -NoNewline
Write-Host "$($PSStyle.Foreground.FromRgb(152, 185, 20)) @doublemover" -NoNewline
Write-Host " " -BackgroundColor Black
Write-Host-FancyLine

# Resolve path
$ResolvedZipPath = Resolve-Path -Path $ZipFile -ErrorAction Stop
if (-not (Test-Path $ResolvedZipPath)) {
    Write-Error "Zip file not found: $ResolvedZipPath"
    exit 1
}

# Output paths
$FolderName = "SpotifyData_" + (Get-Date -Format "yyyyMMdd_HHmmss")
$BaseOutputPath = "$PSScriptRoot\SpotifyData_" + $FolderName
New-Item -ItemType Directory -Path $BaseOutputPath -Force | Out-Null

$ExtractPath = Join-Path $BaseOutputPath "Extracted"
New-Item -ItemType Directory -Path $ExtractPath -Force | Out-Null

# Extract ZIP (robust fallback)
Write-Host "$fancyGap Extracting ZIP file..." -NoNewline -ForegroundColor Cyan
try {
    [System.IO.Compression.ZipFile]::ExtractToDirectory($ResolvedZipPath, $ExtractPath)
} catch {
    Write-Host "." -ForegroundColor DarkCyan
    Expand-Archive -Path $ResolvedZipPath -DestinationPath $ExtractPath -Force
}

# Gather JSON files and sort them
$JsonFiles = Get-ChildItem -Path $ExtractPath -Recurse -Filter "Streaming_History_Audio*.json" | Sort-Object Name

# Init data structures
$TrackDict = @{}
$MissingUriList = @()
$SkippedRecords = 0
$TotalTimestamps = 0
$PlaysProcessed = 0

# Process files
foreach ($JsonFile in $JsonFiles) {
    $shortName = $JsonFile.Name -replace "Streaming_History_Audio_", ""
    $shortName = $shortName -replace ".json", ""

    Write-Host -NoNewline "`r$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Processing $($shortName) ... $($PSStyle.Foreground.FromRgb(255, 255, 255))$PlaysProcessed $($PSStyle.Foreground.FromRgb(29, 185, 84))plays                  "

    try {
        $JsonData = Get-Content -Raw -Path $JsonFile.FullName | ConvertFrom-Json
    } catch {
        Write-Warning "$fancyGap Failed to parse JSON in file $($JsonFile.FullName). Skipping."
        continue
    }

    foreach ($Entry in $JsonData) {
        try {
            $SpotifyUri = $Entry.spotify_track_uri
            $Timestamp = $Entry.ts
            $TrackName = $Entry.master_metadata_track_name
            $AlbumName = $Entry.master_metadata_album_album_name
            $AlbumArtistName = $Entry.master_metadata_album_artist_name

            if (-not $Timestamp -or -not $TrackName) {
                $SkippedRecords++
                continue
            }

            if ([string]::IsNullOrWhiteSpace($SpotifyUri)) {
                $MissingUriList += $Entry
                continue
            }

            # Init object for this track name if needed
            if (-not $TrackDict.ContainsKey($TrackName)) {
                $TrackDict[$TrackName] = @{
                    album_name        = @()
                    album_artist_name = @()
                    total_ms_played   = 0
                    timestamps        = @{}
                }
            }

            # Accumulate album_name
            if ($AlbumName -and -not ($TrackDict[$TrackName].album_name -contains $AlbumName)) {
                $TrackDict[$TrackName].album_name += $AlbumName
            }

            # Accumulate album_artist_name
            if ($AlbumArtistName -and -not ($TrackDict[$TrackName].album_artist_name -contains $AlbumArtistName)) {
                $TrackDict[$TrackName].album_artist_name += $AlbumArtistName
            }

            # Accumulate ms_played
            $TrackDict[$TrackName].total_ms_played += ($Entry.ms_played | ForEach-Object { $_ } )

            # Build timestamp entry with conditional fields
            $TimestampEntry = @{
                reason_start = $Entry.reason_start
                reason_end   = $Entry.reason_end
            }

            if ($Entry.shuffle)           { $TimestampEntry.shuffle = $Entry.shuffle }
            if ($Entry.skipped)           { $TimestampEntry.skipped = $Entry.skipped }
            if ($Entry.offline)           { $TimestampEntry.offline = $Entry.offline }
            if ($Entry.incognito_mode)    { $TimestampEntry.incognito_mode = $Entry.incognito_mode }
            if ($Entry.offline_timestamp) { $TimestampEntry.offline_timestamp = $Entry.offline_timestamp }

            # Add/update timestamp
            $TrackDict[$TrackName].timestamps["$Timestamp"] = $TimestampEntry

            $TotalTimestamps++
            $PlaysProcessed++

            if ($PlaysProcessed % 10 -eq 0) {
                if (-not $Fast) {
                    Write-Host -NoNewline "`r$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Processing $($shortName) ... $($PSStyle.Foreground.FromRgb(255, 255, 255))$PlaysProcessed $($PSStyle.Foreground.FromRgb(29, 185, 84))plays"
                }
            }
        } catch {
            $SkippedRecords++
            continue
        }
    }
}

# Save output
$OutputFileGood = Join-Path $BaseOutputPath "spotify_tracks_output.json"
$OutputFileMissing = Join-Path $BaseOutputPath "spotify_tracks_missing_uri.json"

$TrackDict | ConvertTo-Json -Depth 12 | Out-File -Encoding utf8 -FilePath $OutputFileGood
if ($MissingUriList.length > 0) {
    $MissingUriList | ConvertTo-Json -Depth 10 | Out-File -Encoding utf8 -FilePath $OutputFileMissing
}
# Summary
Write-Host ""
Write-Host-FancyLine
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Unique track names: $($PSStyle.Foreground.FromRgb(255, 255, 255))$($TrackDict.Count)"
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Total timestamps processed: $($PSStyle.Foreground.FromRgb(255, 255, 255))$TotalTimestamps"
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Records with missing URI: $($PSStyle.Foreground.FromRgb(255, 255, 255))$($MissingUriList.Count)"
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Malformed/skipped records: $($PSStyle.Foreground.FromRgb(255, 255, 255))$SkippedRecords"
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Output written to: $($PSStyle.Foreground.FromRgb(255, 255, 255))$FolderName"
Write-Host "$fancyGap $($PSStyle.Foreground.FromRgb(29, 185, 84))Processed in $($PSStyle.Foreground.FromRgb(255, 255, 255))$(((Get-Date) - $startTime).TotalSeconds.ToString("0.0")) $($PSStyle.Foreground.FromRgb(29, 185, 84))seconds"
if ($MissingUriList.length > 0) {
    Write-Host "$fancyGap Missing URI records written to: $OutputFileMissing"
}
Write-Host-FancyLine
Write-Host "$fancyGap Cleaning up temporary files..." -NoNewline -ForegroundColor DarkMagenta
Remove-Item $ExtractPath -Recurse -Force
Write-Host " Done." -ForegroundColor Cyan
Write-Host-FancyLine
