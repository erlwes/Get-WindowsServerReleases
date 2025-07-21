![PowerShell](https://img.shields.io/badge/PowerShell-5+-blue)

# Get-WindowsServerReleases
`Get-WindowsServerReleases.ps1` is a PowerShell script designed to retrieve and parse the **Windows Server release history** directly from Microsoft's official documentation page. It supports recent versions of Windows Server (2016 and newer) and uses local caching to minimize redundant web scraping.

## 📦 Installation

Install the module from PowerShell Gallery:

```powershell
Install-Script -Name Get-WindowsServerReleases
```

---

## 🚀 Quickstart

```powershell
Get-WindowsServerReleases.ps1 -WindowsServerVersion 'Server 2025' | Format-Table
```

<img width="882" height="331" alt="image" src="https://github.com/user-attachments/assets/64300a73-070b-4555-b605-0ff17632901f" />


---

## 📌 Features

-  Web scraping from [Microsoft Docs](https://learn.microsoft.com/en-us/windows/release-health/windows-server-release-info)
-  Caches results in offline CSV-files and reuses results intra day (re-downloads if older than 1 day)
-  Compatible with Windows PowerShell and PowerShell Core
-  Does not require a tentant and graph auth, like the official API

---

## 💻 Parameters

| Parameter             | Type      | Description |
|-----------------------|-----------|-------------|
| `-WindowsServerVersion` | `string`  | (Optional) Filter output for a specific version: `Server 2016`, `Server 2019`, `Server 2022`, `Server 2025`. |
| `-PathLocalStore`       | `string`  | Path to store local CSV cache. Defaults to script location. |
| `-ForceRebuild`         | `switch`  | Forces a fresh web scrape and overwrites cached data regardless of age. |
| `-VerboseLogging`       | `switch`  | Outputs progress and activity to the console. |
| `-ShowCache`            | `switch`  | Displays existing cached data without making a web request. |

---

## 📋 Examples

### `All supported server versions to GridView`
```powershell
Get-WindowsServerReleases.ps1 | Out-GridView
```

### `Releases for Server 2022 with verbose logging`
```powershell
Get-WindowsServerReleases.ps1 -WindowsServerVersion 'Server 2022' -VerboseLogging | Format-Table
```

### `Force re-download/build of local cache`
```powershell
Get-WindowsServerReleases.ps1 -ForceRebuild
```

### `Display cached data without checking age or online content`
```powershell
Get-WindowsServerReleases.ps1 -ShowCache
```
