<#PSScriptInfo
    .VERSION 1.0.0
    .GUID 857dbda9-4724-4c09-8969-8e2a576657ef
    .AUTHOR Erlend Westervik
    .COMPANYNAME
    .COPYRIGHT
    .TAGS Windows, Server, Update, Patch, Upgrade, Operating System, OS, Win, Release, Version, Servicing, LTSC, AC, Lifecycle, Build, OOB
    .LICENSEURI
    .PROJECTURI https://github.com/erlwes/Get-WindowsServerReleases
    .ICONURI
    .EXTERNALMODULEDEPENDENCIES 
    .REQUIREDSCRIPTS
    .EXTERNALSCRIPTDEPENDENCIES
    .RELEASENOTES
        Version: 1.0.0 - Original published version

#>

<#
.SYNOPSIS
    Get Windows server release history

.DESCRIPTION
    Explain the steps

.NOTES
    1. Windows Server 2012 R2 or older OS are end-of-life (no longer extended support). Release history for these OS are no longer published, and is therefore not supported in this script. The data would be static.
    2. There is a "public" API with this information and more, but it requires a M365 tenant to access + auth
        - https://learn.microsoft.com/en-us/graph/api/resources/windowsupdates-product?view=graph-rest-beta

.PARAMETER PathLocalStore
    Specify path to save local cache of Windows server releases. Defaults to module/script location.

.PARAMETER ForceRebuild
    Re-creates the local CSV-files, regardless of how recent they are. By default, it uses the local cache if its less than one day old.

.PARAMETER VerboseLogging
    Show activity in console

.EXAMPLE
    .\Get-WindowsServerReleases.ps1 | Out-GridView

.EXAMPLE
    .\Get-WindowsServerReleases.ps1 -WindowsServerVersion 'Server 2025' -VerboseLogging | Format-Table

.EXAMPLE
    .\Get-WindowsServerReleases.ps1 -WindowsServerVersion 'Server 2016' | Format-Table

.EXAMPLE
    .\Get-WindowsServerReleases.ps1 -WindowsServerVersion -ForceRebuild

#>

Param(    
    [ValidateSet('Server 2025', 'Server 2022', 'Server 2019', 'Server 2016')]
    [String]$WindowsServerVersion,

    [String]$PathLocalStore = $PSScriptRoot,

    [Switch]$VerboseLogging,

    [Switch]$ForceRebuild,

    [Switch]$ShowCache
)

# DECLARATIONS
$Time = (Get-Date)

# FUNCTIONS
function ParseHtml($string) {
    $unicode = [System.Text.Encoding]::Unicode.GetBytes($string)
    $html = New-Object -Com 'HTMLFile'
    if ($html.PSObject.Methods.Name -Contains 'IHTMLDocument2_Write') {
        $html.IHTMLDocument2_Write($unicode)
    } 
    else {
        $html.write($Unicode)
    }
    $html.Close()
    $html
}

 Function Write-Log {
    param(
        [ValidateSet(0, 1, 2, 3, 4)]
        [int]$Level,

        [Parameter(Mandatory=$true)]
        [string]$Message
    )
    $Message = $Message.Replace("`r",'').Replace("`n",' ')
    switch ($Level) {
        0 { $Status = 'Info'    ;$FGColor = 'White'   }
        1 { $Status = 'Success' ;$FGColor = 'Green'   }
        2 { $Status = 'Warning' ;$FGColor = 'Yellow'  }
        3 { $Status = 'Error'   ;$FGColor = 'Red'     }
        4 { $Status = 'Console' ;$FGColor = 'Gray'    }
        Default { $Status = ''  ;$FGColor = 'Black'   }
    }
    if ($VerboseLogging) {
        Write-Host "$((Get-Date).ToString()) " -ForegroundColor 'DarkGray' -NoNewline
        Write-Host "$Status" -ForegroundColor $FGColor -NoNewline

        if ($level -eq 4) {
            Write-Host ("`t " + $Message) -ForegroundColor 'Cyan'
        }
        else {
            Write-Host ("`t " + $Message) -ForegroundColor 'White'
        }
    }
    if ($Level -eq 3) {
        $LogErrors += $Message
    }
}

Write-Log -Level 0 -Message "Start"
Write-Log -Level 0 -Message "Local cache - Using '$PathLocalStore' as local store"

if ($ShowCache) {
    Write-Log -Level 0 -Message "Local cache - Parameter -ShowCache used. Looking for existing CSV-cacge in this path: '$PathLocalStore'."
    $Cache = @()
    if ($WindowsServerVersion) {
        $CSVFiles = Get-Item -Path "$PathLocalStore\Windows $WindowsServerVersion (OS build*.csv"    
    }
    else {
        $CSVFiles = Get-Item -Path "$PathLocalStore\Windows Server * (OS build*.csv"    
    }
    
    if ($CSVFiles.count -ge 1) {
        Write-Log -Level 1 -Message "Local cache - $($CSVFiles.count) CSV-files found."
        Foreach ($File in $CSVFiles) {
            $Cache += Import-Csv -Path $File.FullName -Delimiter ';' -Encoding utf8
        }
        $Cache
    }
    else {
        Write-Log -Level 0 -Message "Local cache - Not found  (while using -ShowCache). Can not continue."
    }
    Write-Log -Level 0 -Message "End"
    Break
}
else {
    # WEB REQUEST
    $WR = Invoke-WebRequest -Uri 'https://learn.microsoft.com/en-us/windows/release-health/windows-server-release-info'
    if ($WR.StatusCode -eq 200) {
        Write-Log -Level 1 -Message 'Invoke-WebRequest - Status 200. Content received.'
    }
    else {
        Write-Log -Level 3 -Message "Invoke-WebRequest - Status $($WR.StatusCode)"
    }

    # HIGH LEVEL PARSING - TABLES + HEADERS
    if ($host.version.Major -gt 5) {
        Write-Log -Level 0 -Message "Parse HTML - PowerShell Core detected ($($host.version.Major).$($host.version.Minor)). Parsing HTML using 'ParseHTLM function'"
        $Document = ParseHtml $WR.Content
        $Tables = $Document.getElementsByTagName('table') | Where-Object {$_.id -match 'historyTable'}
        $ServerOSTitles = $Document.getElementsByTagName('strong') | Where-Object {$_.innerText -match 'Windows Server'} | Select-Object -ExpandProperty innerText -Unique
    }
    else {
        Write-Log -Level 0 -Message "Parse HTML - Windows PowerShell detected ($($host.version.Major).$($host.version.Minor)). Using built in 'parsedHtml'"
        $Tables = $WR.ParsedHtml.getElementsByTagName('table') | Where-Object {$_.id -match 'historyTable'}
        $ServerOSTitles = $WR.ParsedHtml.getElementsByTagName('strong') |  Where-Object {$_.IHTMLElement_innerText -match 'Windows Server'} | Select-Object -ExpandProperty IHTMLElement_innerText -Unique 
    }
    Write-Log -Level 0 -Message "Parse HTML - Found $($Tables.count) relevant tables"
    Write-Log -Level 0 -Message "Parse HTML - Found $($ServerOSTitles.count) relevant headers"

    # SANITY CHECKS ON RESULTS
    if ($Tables.count -lt 1 -or $ServerOSTitles.count -lt 1) {
        Write-Log -Level 2 -Message 'Parse HTML - Missing headers and/or titles. Can not continue.'
        Break
    }
    if ($Tables.count -ne $ServerOSTitles.count) {
        Write-Log -Level 2 -Message 'Parse HTML - In-equal count of headers vs. tables. Will not continue.'
        Break
    }

    # MATCH SERVER VERSION TO TABLES AND DETERMINE CACHE LOCATIONS
    $CSVFiles = @()
    if ($WindowsServerVersion -eq 'Server 2016') {
        $Tables = $Tables | Where-Object {$_.id -eq 'historyTable_3'}
        $CSVFiles += "$PathLocalStore\$($ServerOSTitles | Where-Object {$_ -match $WindowsServerVersion}).csv"
    }
    elseif ($WindowsServerVersion -eq 'Server 2019') {
        $Tables = $Tables | Where-Object {$_.id -eq 'historyTable_2'}
        $CSVFiles += "$PathLocalStore\$($ServerOSTitles | Where-Object {$_ -match $WindowsServerVersion}).csv"
    }
    elseif ($WindowsServerVersion -eq 'Server 2022') {
        $Tables = $Tables | Where-Object {$_.id -eq 'historyTable_1'}
        $CSVFiles += "$PathLocalStore\$($ServerOSTitles | Where-Object {$_ -match $WindowsServerVersion}).csv"
    }
    elseif ($WindowsServerVersion -eq 'Server 2025') {
        $Tables = $Tables | Where-Object {$_.id -eq 'historyTable_0'}
        $CSVFiles += "$PathLocalStore\$($ServerOSTitles | Where-Object {$_ -match $WindowsServerVersion}).csv"
    }
    else {
        $ServerOSTitles | Where-Object {$_ -match 'Windows Server'} | % {
            $CSVFiles += "$PathLocalStore\$_.csv"
        }
    }

    # DETERMINE IF CACHE EXIST AND CAN BE RE-USED
    if (!$ForceRebuild) {
        $ServerOSTitles | Where-Object {$_ -match $WindowsServerVersion} | ForEach-Object {
            $ServerOSTitle = $_
            $CSVFile = "$PathLocalStore\$ServerOSTitle.csv"
            if (Test-Path $CSVFile) {
                Write-Log -Level 1 -Message "Check cache - $ServerOSTitle - File exist ('$CSVFile')"
                [datetime]$FileDate = Get-Item -Path $CSVFile | Select-Object -ExpandProperty LastWriteTime
                $TimeDiff = ($Time - $FileDate)
                if ($TimeDiff.TotalHours -ge 24) {
                    Write-Log -Level 2 -Message "Check cache date - $ServerOSTitle - Older than 24h ($($TimeDiff.TotalHours)). Setting '-ForceRebuild' switch."
                    $ForceRebuild = $true
                }
                else {
                    Write-Log -Level 1 -Message "Check cache date - $ServerOSTitle - More recent than 24h ($($TimeDiff.TotalHours))."
                }
            }
            else {
                Write-Log -Level 2 -Message "Check cache - $ServerOSTitle - No file ('$CSVFile'). Setting '-ForceRebuild' switch."
                $ForceRebuild = $true
            }
            Clear-Variable ServerOSTitle, CSVFile
        }
    }

    # BUILD/RE-BUILD CACHE
    $Result = @()
    if ($ForceRebuild) {
        foreach ($table in $Tables) {

            # Get the headers (cells in row 0)
            $Headers = ($Table.Rows[0].Cells | Select-Object -ExpandProperty innerText).Trim()

            # Find matching Windows Server version for the table contents
            $Index = [int]($Table.id -replace '^.+_')
            $ServerTitle = $ServerOSTitles[$Index]

            Write-Log -Level 0 -Message "Convert HTML to PSObject - $ServerTitle - Begin"
            $M = Measure-Command {        
                $Object = [System.Collections.Generic.List[object]]::new()

                # Loop through rows (skip header)

                foreach ($Tablerow in ($Table.Rows | Select-Object -Skip 1)) {
                    $Cells = ($Tablerow.Cells | Select-Object -ExpandProperty innerText).Trim()

                    # Build hashtable
                    $RowHash = @{}
                    $RowHash['Windows server version'] = $ServerTitle
                    for ($i = 0; $i -lt $Headers.Count; $i++) {
                        if ($i -lt $Cells.Count) {
                            $RowHash[$Headers[$i]] = $Cells[$i]
                        }
                    }

                    # Convert to PSObject and add to list
                    $Object.Add([PSCustomObject]$RowHash)
                }
            }
            Write-Log -Level 0 -Message "Convert HTML to PSObject - $ServerTitle - Done in $($M.TotalSeconds) seconds ($($Object.Count) releases)"

            #Sanity checks before saving and poteltially overwriting last cache
            if ($Object.Build -notmatch '\d') {
                Write-Log -Level 3 -Message "Convert HTML to PSObject - $ServerTitle - The produced PSObject is not valid. Build-property is not present or has invalid values. Aborting script."
                Break
            }

            # Export to CSV (save cache)
            try {
                $Object | Export-Csv "$PathLocalStore\$ServerTitle.csv" -Delimiter ';' -Encoding utf8 -Confirm:$false -Force -ErrorAction Stop
                $Result += $Object
                Write-Log -Level 1 -Message "Export-Csv - Saved '$PathLocalStore\$ServerTitle.csv'"
            }
            catch {
                Write-Log -Level 3 -Message "Export-Csv - Failed to save '$PathLocalStore\$ServerTitle.csv'. Error: $($_.Exception.Message)"
            }
            
            # Clear variables for next table in loop
            Clear-Variable headers, index, serverTitle, m
        }
    }
    else {
        # If all needed caches are newer than 24h, just import the cache and present the data.
        $CSVFiles | ForEach-Object {
            $Result += Import-Csv -Path $_ -Delimiter ';' -Encoding utf8
        }
    }
    Write-Log -Level 0 -Message "End"
    $Result
}
