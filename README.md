# 2xMVR PSS
Powershell Scripts for Various Bullshit

---

<details open>
<summary><h1>Spotify Data Extractor</h1></summary>

This script extracts and combines your full Spotify listening history and account data into json files for personal analysis. It is intended for users who want complete access to their listening data, offline and under their control.

<details>
<summary><h2>Obtaining Your Data</h2></summary>

### Extended Streaming History

1. Visit: [https://www.spotify.com/account/privacy/](https://www.spotify.com/account/privacy/)
2. Under **Download your data**, request **Extended streaming history**.
3. Wait for the email from Spotify with your download link.
4. Download the ZIP file, typically named `my_extended_spotify_data.zip`.

### Account Data

1. On the same page: [https://www.spotify.com/account/privacy/](https://www.spotify.com/account/privacy/)
2. Under **Download your data**, request **Account data**.
3. Wait for the email from Spotify with your download link.
4. Download the ZIP file, typically named `my_spotify_data.zip`.

</details>

## Usage

This script requires **PowerShell 7** or later.

To install PowerShell 7, run this from cmd or earlier versions of PowerShell:

```
winget install --id Microsoft.Powershell --source winget
```

Type `PowerShell 7` into the start menu specifically to launch it.

### Running the Script

Place this script and your ZIP file(s) in the same folder, or provide full paths to the ZIP(s)

```powershell
.\spotify-data-extractor.ps1 .\my_extended_spotify_data.zip .\my_spotify_data.zip
```

If two ZIPs are provided, one must be an Extended Streaming History ZIP and the other must be an Account Data ZIP.

## Output

The script creates an output folder named:

```
SpotifyData_YYYYMMDD_HHMMSS\
```

This folder will contain:

- `music_stats.json` — Complete play history, organized by artist → album → track, sorted by play count, including:
  - Total `plays` and `msplayed` per artist.
  - Per-track stats including:
    - `ms_played`
    - `plays`
    - Timestamps of plays, each containing:
      - `ts`, `reason_start`, `reason_end`
      - Only if applicable: `shuffle`, `skipped`, `incognito_mode`, `offline`, `offline_timestamp`
- `library.json` — Your saved tracks/albums, from Account Data.
- `playlists.json` — Your playlists, from Account Data.
- `search_history.json` — Cleaned list of searches, combining related incremental searches into final search terms.
- `wrapped.json` — Merged Spotify Wrapped data, organized by year.
- `error.log` — If any files failed to process.

## Notes

- The script de-duplicates play data: plays present in both Extended Streaming History and Account Data will not be double-counted.
- The `music_stats.json` output is structured for ease of further processing or visualization.
- All processing is done locally — **no data is sent to any external service**.
- The script is compatible with the standard ZIP format and folder structure as provided directly by Spotify.
- It has been tested on data exports from accounts with over 300,000 plays
- It generally finishes in less than a minute

<details>
<summary><h2>Troubleshooting</h2></summary>

- If you receive `Invalid ZIP file(s)` errors:
  - Verify that the ZIP files were downloaded directly from Spotify and have not been modified or extracted/re-compressed.
  - Ensure that:
    - The Extended Streaming History ZIP contains a `Spotify Extended Streaming History` folder.
    - The Account Data ZIP contains a `Spotify Account Data` folder.

- If `music_stats.json` appears empty:
  - Ensure that the Extended Streaming History ZIP contains one or more `Streaming_History*.json` files.
  - If only using Account Data, note that it typically contains fewer plays and may not provide complete history.

- If `search_history.json` looks incorrect:
  - The script attempts to collapse incremental search sequences. In some cases, Spotify's raw search logs may not perfectly reflect user intent.
 
- If it's running for a really long time and you are sure you do not have a slow computer:
  - This should not happen but if it does please file an issue or contact me on discord
</details>

## Support

If you encounter issues or have questions about this script:

- [Open an issue](https://github.com/doublemover/PSS/issues), **or**
- Contact `@2xMVR` on Discord for assistance.
- X/Bsky `@doublemover`
</details>

---

If one of these scripts was useful:

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/E1E71G7Y0T)
