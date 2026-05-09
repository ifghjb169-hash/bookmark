# Bookmark Export Tool

Windows desktop app for moving Chrome bookmarks into CSV files or attempting to append them to Google Sheets.

## Run

Double-click `start_bookmark_tool.bat`, or run:

```powershell
py -3 bookmark_tool.py
```

## Download

Release builds are published from GitHub Actions when a version tag such as `v1.0.0` is pushed:

https://github.com/secure-artifacts/bookmark/releases

Download the `BookmarkExportTool-<version>.exe` asset from the latest release. Release assets are produced by the workflow in `.github/workflows/release.yml`; do not use manually uploaded files as official builds.

## Features

- Convert one or more Chrome bookmark HTML exports into separate CSV files.
- Drag `.html` or `.htm` bookmark files directly into the app window to add them.
- Scan Chrome user profile folders such as `C:\Users\tenp\AppData\Local\Google\Chrome\User Data\Profile 3`.
- Select which scanned Chrome profiles to export, then export one CSV per selected profile or merge selected profiles into one CSV with a `profile` column.
- Save output files to `Desktop\bookmarks` by default, or choose any other output folder from the top bar.
- UI language dropdown: Chinese, English, Vietnamese, Burmese, Hindi, Filipino, Portuguese, Spanish.
- Google Sheets tab accepts a Sheet link and an OAuth client JSON file, remembers both paths, can open the sheet in your browser, lets you choose profiles, then writes each selected Chrome profile into its own sheet tab.

## Google Sheets note

Google Sheets write operations require authorization. Create a Google Cloud OAuth client for a Desktop app, download its JSON file, select it in the app, and sign in when the browser opens. The app caches the resulting token locally with Windows DPAPI protection for the current Windows user.

Setup checklist:

1. In Google Cloud, enable the Google Sheets API for your project.
2. Create an OAuth client ID with application type `Desktop app`.
3. Download the client JSON file.
4. Make sure the Google account you sign in with has edit access to the target spreadsheet.
5. Paste the spreadsheet link in the app, select the OAuth client JSON, then click upload.

## Verify a release

Official release assets are built by GitHub Actions from a `v*` tag and include a GitHub artifact attestation.

With GitHub CLI installed, verify the provenance:

```powershell
gh attestation verify .\BookmarkExportTool-v1.0.4.exe --repo secure-artifacts/bookmark
```

You can also compare the downloaded file with the `.sha256` checksum published next to the EXE:

```powershell
Get-FileHash .\BookmarkExportTool-v1.0.4.exe -Algorithm SHA256
```
