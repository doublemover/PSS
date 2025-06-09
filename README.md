# 2xMVR PSS
Powershell Scripts for Various Bullshit

---

<details open>
  <summary><h2><a href="https://github.com/doublemover/PSS/blob/main/spotify-data-extractor.ps1">Spotify Data Extractor</a></h2></summary>

  - Requires Powershell 7 
  - Usage: ` PS C:\Users\you\PowerShell-Scripts> .\spotify-data-extractor1.ps1 -InputPath my_spotify_data.zip`
  - Consumes a spotify data zip or extracted folder then outputs:
    - `account_info.json` 
    - `music_stats.json`
      - collective statistics of all played tracks
      - total artists/ms/playcount/tracks
      - top artists by month w/ playcount & total playtime
      - yearly breakdown organized by artist
    - `library.json`
      - tracks/albums/artists saved to your library
    - `playlists.json`
      - playlist stats
    - `wrapped.json`
      - an attempt to shove all fields from any wrapped into one file
    - `inferences.json`
    - `search_history.json`
    - [Optional, `-GenerateCSV`] `top_artists_overall.csv`, `top_tracks_overall.csv`, `listening_by_year.csv`, `listening_by_month.csv`
</details>

<details open>
  <summary><h2><a href="https://github.com/doublemover/PSS/blob/main/spotify-extended-data-extractor.ps1">Spotify Extended Extractor</a></h2></summary>

  - Requires Powershell 7
  - Usage: ` PS C:\Users\you\PowerShell-Scripts> .\spotify-extended-data-extractor.ps1.ps1 -ZipFile my_spotify_data.zip`
  - Consumes a spotify extended streaming data zip then outputs:
    - `spotify_tracks_output.json` 
    - If there are values with missing URI:
      - `spotify_tracks_missing_uri.json`
</details>

---

If one of these was useful:

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/E1E71G7Y0T)
