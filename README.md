# Get-WindowsServerReleases

![PowerShell](https://img.shields.io/badge/PowerShell-5+-blue)

`Get-WindowsServerReleases.ps1` is a PowerShell script designed to retrieve and parse the **Windows Server release history** directly from Microsoft's official documentation page. It supports recent versions of Windows Server (2016 and newer) and uses local caching to minimize redundant web scraping.

## Install
`Install-Script -Name Get-WindowsServerReleases`

---

Screenshot:

<img width="882" height="331" alt="image" src="https://github.com/user-attachments/assets/64300a73-070b-4555-b605-0ff17632901f" />

Remarks:
 *  Yes - I know there is a [public API](https://learn.microsoft.com/en-us/graph/api/resources/windowsupdates-product?view=graph-rest-beta) with this data, and more, but you need a tenant and graph API auth in order to use it.
 *  This script relies on consistent use of table id's and titles in Microsoft HTML-structure. Seems good :)
 *  The script will need to be patched when the next OS comes in 2027 og 2028 something, and when 2016 is not longer updated. Easy fix with validate set in parameters.

---

## Features

-  Web scraping from [Microsoft Docs](https://learn.microsoft.com/en-us/windows/release-health/windows-server-release-info)
-  Caching (with optional force rebuild)
-  CSV export per OS version
-  Console logging (verbose optional)
-  Compatible with Windows PowerShell and PowerShell Core

---

## Parameters

| Parameter             | Type      | Description |
|-----------------------|-----------|-------------|
| `-WindowsServerVersion` | `string`  | (Optional) Filter output for a specific version: `Server 2016`, `Server 2019`, `Server 2022`, `Server 2025`. |
| `-PathLocalStore`       | `string`  | Path to store local CSV cache. Defaults to script location. |
| `-ForceRebuild`         | `switch`  | Forces a fresh web scrape and overwrites cached data regardless of age. |
| `-VerboseLogging`       | `switch`  | Outputs progress and activity to the console. |
| `-ShowCache`            | `switch`  | Displays existing cached data without making a web request. |

---

## Script Logic

1. **Startup & Parameter Handling**
   - Validates version input and sets path defaults.
   - If `-ShowCache` is used, it simply loads cached CSVs and exits.

2. **Web Scraping**
   - Downloads the release history page using `Invoke-WebRequest`.
   - Parses HTML to extract tables by Windows Server version.

3. **Caching**
   - Cached files are stored per version as `*.csv` (semicolon-separated).
   - Automatically reuses cache if it’s less than 24 hours old, unless `-ForceRebuild` is used.

4. **Output**
   - Returns a PowerShell object list representing the release information.
   - Supports formatting (`Format-Table`, `Out-GridView`, etc.).

---

## More examples

```powershell
# Show all server versions in a grid view
.\Get-WindowsServerReleases.ps1 | Out-GridView

# Get releases for Server 2022 with verbose logging
.\Get-WindowsServerReleases.ps1 -WindowsServerVersion 'Server 2022' -VerboseLogging | Format-Table

# Rebuild and get info for Server 2019
.\Get-WindowsServerReleases.ps1 -WindowsServerVersion 'Server 2019' -ForceRebuild

# Display cached data without fetching new content
.\Get-WindowsServerReleases.ps1 -ShowCache
```
