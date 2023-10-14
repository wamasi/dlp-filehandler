<#
.Synopsis
   Script to run yt-dlp, mkvmerge, subtitle edit, filebot and a python script for downloading and processing videos
.EXAMPLE
   Runs the script using the mySiteHere as a manual run with the defined config using login, cookies, archive file, mkvmerge, and sends out a discord message
   D:\_DL\dlp-script.ps1 -sn mySiteHere -l -c -mk -a -sd
.EXAMPLE
   Runs the script using the mySiteHere as a daily run with the defined config using login, cookies, no archive file, and filebot/plex
   D:\_DL\dlp-script.ps1 -sn mySiteHere -d -l -c -f
.NOTES
   See https://github.com/wamasi/dlp-filehandler for more details
   Script was designed to be ran via powershell by being called through console on a cronjob or task scheduler. copying and pasting into powershell console will not work.
#>
[CmdletBinding()]
param(
    [Parameter(ParameterSetName = 'Help', Mandatory = $True)]
    [Alias('H')]
    [switch]$help,
    [Parameter(ParameterSetName = 'NewConfig', Mandatory = $True)]
    [Alias('NC')]
    [switch]$newConfig,
    [Alias('SU')]
    [Parameter(ParameterSetName = 'SupportFiles', Mandatory = $True)]
    [switch]$supportFiles,
    [Alias('SN')]
    [ValidateScript({ if (Test-Path -Path "$PSScriptRoot\config.xml") {
                if (([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName -contains $_ ) {
                    $true
                }
                else {
                    $validSites = (([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName) -join "`r`n"
                    throw "The following Sites are valid:`r`n$validSites"
                }
            }
            else {
                throw "No valid config.xml found in $PSScriptRoot. Run ($PSScriptRoot\dlp-script.ps1 -nc) for a new config file."
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $True)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $True)]
    [string]$site,
    [Alias('OD')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$overrideBatch,
    [Alias('D')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$daily,
    [Alias('L')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$login,
    [Alias('C')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$cookies,
    [Alias('A')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$archive,
    [Alias('AT')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$archiveTemp,
    [Alias('SE')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$subtitleEdit,
    [Alias('MK')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$mkvMerge,
    [Alias('F')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$filebot,
    [Alias('SD')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$sendDiscord,
    [Alias('AL')]
    [ValidateScript({
            $langValues = 'ar', 'de', 'en', 'es', 'es-la', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und'
            if ($_ -in $langValues) {
                $true
            }
            else {
                throw "Value '{0}' is invalid. The following languages are valid:`r`n{1}" -f $_, $($langValues -join "`r`n")
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$audioLang,
    [Alias('SL')]
    [ValidateScript({
            $langValues = 'ar', 'de', 'en', 'es', 'es-la', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und'
            if ($_ -in $langValues ) {
                $true
            }
            else {

                throw "Value '{0}' is invalid. The following languages are valid:`r`n{1}" -f $_, $($langValues -join "`r`n")
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$subtitleLang,
    [Alias('T')]
    [Parameter(ParameterSetName = 'Test', Mandatory = $true)]
    [switch]$testScript,
    [Alias('DS')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [switch]$debugScript
)
# Timer for script
$scriptStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
# Setting styling to remove error characters and width
$psStyle.OutputRendering = 'Host'
$width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($width, 9999)
mode con cols=9999
# Output current time in different formats
function Get-DateTime {
    param (
        [int]$dateType
    )
    switch ($dateType) {
        1 { $datetime = Get-Date -Format 'yy-MM-dd' }
        2 { $datetime = Get-Date -Format 'MMddHHmmssfff' }
        3 { $datetime = ($(Get-Date).ToUniversalTime()).ToString('yyyy-MM-ddTHH:mm:ss.fffZ') }
        Default { $datetime = Get-Date -Format 'yy-MM-dd HH-mm-ss' }
    }
    return $datetime
}
# Pass through expressions to format them from logging
function Invoke-ExpressionConsole {
    param (
        [Parameter(Mandatory = $true)]
        [Alias('SCMFN')]
        [string]$scmFunctionName,
        [Parameter(Mandatory = $true)]
        [Alias('SCMFP')]
        [string]$scmFunctionParams
    )
    $iecArguement = "$scmFunctionParams *>&1"
    $iecObject = Invoke-Expression $iecArguement
    $iecObject -split "`n" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | ForEach-Object {
        # Write output then logic to grab filebot output for variables
        Write-Output "[$scmFunctionName] $(Get-DateTime) - $_"
        if ($scmFunctionName -eq 'Filebot') {
            # Use Select-String with regex to match from, to, and title strings
            $fromPattern = '\[MOVE\] from\s\[(.*?)\]\sto'
            $toPattern = '\[MOVE\].*?to \[(.*?)\]$'
            $titlePattern = "--set title=(.*?)' executed"
            if ($_ -match $fromPattern) {
                $fromVideoVar = ($_ | Select-String -Pattern $fromPattern | ForEach-Object { $_.Matches.Groups[1].Value }).Trim()
                Write-Output "[$scmFunctionName] $(Get-DateTime) - From: $fromVideoVar."
            }
            if ($_ -match $toPattern) {
                $toVideoVar = ($_ | Select-String -Pattern $toPattern | ForEach-Object { $_.Matches.Groups[1].Value }).Trim()
                Write-Output "[$scmFunctionName] $(Get-DateTime) - To: $toVideoVar."
            }
            if ($_ -match $titlePattern) {
                $titleVideoVar = ($_ | Select-String -Pattern $titlePattern | ForEach-Object { $_.Matches.Groups[1].Value }).Trim()
                Write-Output "[$scmFunctionName] $(Get-DateTime) - Title: $titleVideoVar."
            }
            # Once all 3 are matched update associated record and clear the variables for the next video(s)
            if ($fromVideoVar -and $toVideoVar -and $titleVideoVar) {
                Set-VideoStatus -searchKey '_vsDestPathVideo' -searchValue $fromVideoVar -nodeKey '_vsFinalPathVideo' -nodeValue $toVideoVar
                Set-VideoStatus -searchKey '_vsDestPathVideo' -searchValue $fromVideoVar -nodeKey '_vsFilebotTitle' -nodeValue $titleVideoVar
                Remove-Variable -Name fromVideoVar, toVideoVar, titleVideoVar
            }
        }
    }
}
# Get diskspace for a give filepath or Drive
function Get-DiskSpace {
    param (
        $drive
    )
    $r = ((Get-PSDrive -PSProvider 'FileSystem' | Where-Object { $_.Root -eq ([System.IO.Path]::GetPathRoot($drive)) }) | Select-Object Name, Root, Free, Used)
    $rootName = $r.Name
    $free = $r.Free
    $used = $r.Used
    $total = (Get-Volume -Partition (Get-Partition -DriveLetter $rootName)).Size
    $freeO = $free
    $usedO = $used
    $totalO = $total
    $units = @('B', 'KB', 'MB', 'GB', 'TB')
    $freeIndex = 0
    $usedIndex = 0
    $totalIndex = 0
    while ($free -ge 1KB -and $freeIndex -lt $units.Count - 1) {
        $free /= 1KB
        $freeIndex++
    }
    while ($used -ge 1KB -and $usedIndex -lt $units.Count - 1) {
        $used /= 1KB
        $usedIndex++
    }
    while ($total -ge 1KB -and $totalIndex -lt $units.Count - 1) {
        $total /= 1KB
        $totalIndex++
    }
    $freeSpace = '{0:N2}' -f $free
    $usedSpace = '{0:N2}' -f $used
    $totalSpace = '{0:N2}' -f $total
    $FreeUnit = $units[$freeIndex]
    $usedUnit = $units[$usedIndex]
    $totalUnit = $units[$totalIndex]
    $freeSpaceFormatted = "$($freeSpace)$($FreeUnit)"
    $usedSpaceFormatted = "$($usedSpace)$($usedUnit)"
    $totalSpaceFormatted = "$($totalSpace)$($totalUnit)"
    $percentageFree = [math]::Round((($freeO / $totalO) * 100), 2).ToString('F2')
    $percentageUsed = [math]::Round((($usedO / $totalO) * 100), 2).ToString('F2')
    $diskSpacePSObj = [PSCustomObject]@{
        rootName            = $rootName
        rootDrive           = $r.Root
        freeSpace           = $freeSpace
        usedSpace           = $usedSpace
        totalSpace          = $totalSpace
        freeSpaceFormatted  = $freeSpaceFormatted
        usedSpaceFormatted  = $usedSpaceFormatted
        totalSpaceFormatted = $totalSpaceFormatted
        percentageFree      = $percentageFree
        percentageUsed      = $percentageUsed
        freeUnit            = $FreeUnit
        usedUnit            = $usedUnit
        totalUnit           = $totalUnit
    }
    return $diskSpacePSObj
}

# Get filesize formatted
function Get-Filesize {
    param (
        $filePath
    )
    $Size = (Get-Item -Path $filePath).Length

    $units = @('B', 'KB', 'MB', 'GB', 'TB')
    $index = 0
    while ($Size -ge 1KB -and $index -lt $units.Count - 1) {
        $Size /= 1KB
        $index++
    }
    $filesize = [math]::Round($Size, 2).ToString('F2')
    $unit = $units[$index]
    $filesizeFormatted = "$($filesize)$($unit)"
    return $filesize, $filesizeFormatted, $unit, $Size
}
# Test if file is available to interact with
function Test-Lock {
    Param(
        [parameter(Mandatory = $true)]
        $testLockFilename,
        [switch]$literal
    )
    if ($literal) {
        $testLockInitial = Resolve-Path -LiteralPath $testLockFilename
    }
    else {
        $testLockInitial = Resolve-Path $testLockFilename
    }
    $testLockFile = Get-Item -Path ($testLockInitial) -Force
    if ($testLockFile -is [IO.FileInfo]) {
        trap {
            Write-Output "[FileLockCheck] $(Get-DateTime) - $testLockFile File locked. Waiting."
            return $true
            continue
        }
        $testLockStream = New-Object system.IO.StreamReader $testLockFile
        if ($testLockStream) { $testLockStream.Close() }
    }
    Write-Output "[FileLockCheck] $(Get-DateTime) - $testLockFile File unlocked. Continuing."
    return $false
}

function New-Folder {
    param (
        [Parameter(Mandatory = $true)]
        [string] $newFolderFullPath
    )
    if (!(Test-Path -Path $newFolderFullPath)) {
        New-Item -Path $newFolderFullPath -ItemType Directory -Force -Verbose
        Write-Output "$newFolderFullPath missing. Creating."
    }
    else {
        Write-Output "$newFolderFullPath already exists."
    }
}
function New-SuppFile {
    param (
        [Parameter(Mandatory = $true)]
        [string] $newSupportFiles
    )
    if (!(Test-Path -Path $newSupportFiles -PathType Leaf)) {
        New-Item -Path $newSupportFiles -ItemType File | Out-Null
        Write-Output "$newSupportFiles file missing. Creating."
    }
    else {
        Write-Output "$newSupportFiles file already exists."
    }
}
function New-Config {
    param (
        [Parameter(Mandatory = $true)]
        [string] $newConfigs
    )
    Write-Output "Creating $newConfigs"
    New-Item -Path $newConfigs -ItemType File -Force
    if ($newConfigs -match 'vrv') {
        $vrvConfig | Set-Content -Path $newConfigs
        Write-Output "$newConfigs created with VRV values."
    }
    elseif ($newConfigs -match 'crunchyroll') {
        $crunchyrollConfig | Set-Content -Path $newConfigs
        Write-Output "$newConfigs created with Crunchyroll values."
    }
    elseif ($newConfigs -match 'funimation') {
        $funimationConfig | Set-Content -Path $newConfigs
        Write-Output "$newConfigs created with Funimation values."
    }
    elseif ($newConfigs -match 'hidive') {
        $hidiveConfig | Set-Content -Path $newConfigs
        Write-Output "$newConfigs created with Hidive values."
    }
    elseif ($newConfigs -match 'paramountplus') {
        $paramountPlusConfig | Set-Content -Path $newConfigs
        Write-Output "$newConfigs created with ParamountPlus values."
    }
    else {
        $defaultConfig | Set-Content -Path $newConfigs
        Write-Output "$newConfigs created with default values."
    }
}
# Delete Tmp/Src/Home folder logic
function Remove-Folders {
    param (
        [parameter(Mandatory = $true)]
        [string]$removeFolder,
        [parameter(Mandatory = $false)]
        [string]$removeFolderMatch,
        [parameter(Mandatory = $true)]
        [string]$removeFolderBaseMatch
    )
    if ($removeFolder -eq $siteTemp) {
        if (($removeFolder -match $removeFolderBaseMatch) -and (Test-Path -Path $removeFolder)) {
            Write-Output "Force deleting $removeFolder folders/files."
            Remove-Item -Path $removeFolder -Recurse -Force -Verbose
        }
        else {
            Write-Output "SiteTemp($removeFolder) already deleted."
        }
    }
    else {
        if (!(Test-Path -Path $removeFolder)) {
            Write-Output "Folder($removeFolder) already deleted."
        }
        elseif ((Test-Path -Path $removeFolder) -and (Get-ChildItem -Path $removeFolder -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "Folder($removeFolder) is empty. Deleting folder."
            & $DeleteRecursion -deleteRecursionPath $removeFolder
        }
        else {
            Write-Output "Folder($removeFolder) contains files. Manual attention needed."
        }
    }
}
# Removing Site log files
function Remove-Logfiles {
    # Log cleanup
    $filledLogsLimit = (Get-Date).AddDays(-$filledLogs)
    $emptyLogsLimit = (Get-Date).AddDays(-$emptyLogs)
    if (!(Test-Path -Path $logFolderBase)) {
        Write-Output "$logFolderBase is missing. Skipping log cleanup."
    }
    else {
        Write-Output "$logFolderBase found. Starting Filledlog($filledLogs days) cleanup."
        $filledLogFiles = Get-ChildItem -Path $logFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -match '.*-T.*' -and $_.FullName -ne $logFile -and $_.CreationTime -lt $filledLogsLimit }
        if ($filledLogFiles -and ($filledLogFiles | Measure-Object).Count -gt 0) {
            foreach ($f in $filledLogFiles ) {
                $removeLog = $f.FullName
                $removeLog | Remove-Item -Recurse -Force -Verbose
            }
        }
        else {
            Write-Output "No filled logs to remove in $logFolderBase"
        }
        Write-Output "$logFolderBase found. Starting emptylog($emptyLogs days) cleanup."
        $emptyLogFiles = Get-ChildItem -Path $logFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -notmatch '.*-T.*' -and $_.FullName -ne $logFile -and $_.CreationTime -lt $emptyLogsLimit }
        if ($emptyLogFiles -and ($emptyLogFiles | Measure-Object).Count -gt 0) {
            foreach ($e in $emptyLogFiles ) {
                $removeLog = $e.FullName
                $removeLog | Remove-Item -Recurse -Force -Verbose
            }
        }
        else {
            Write-Output "No empty logs to remove in $logFolderBase"
        }
        & $DeleteRecursion -deleteRecursionPath $logFolderBase
    }
}

# Recursively deletes folders
$DeleteRecursion = {
    param(
        $deleteRecursionPath
    )
    foreach ($DeleteRecursionDirectory in Get-ChildItem -LiteralPath $deleteRecursionPath -Directory -Force) {
        & $DeleteRecursion -deleteRecursionPath $DeleteRecursionDirectory.FullName
    }
    $drCurrentChildren = Get-ChildItem -LiteralPath $deleteRecursionPath -Force
    $deleteRecursionEmpty = $drCurrentChildren -eq $null
    if ($deleteRecursionEmpty) {
        Write-Output "Force deleting '${deleteRecursionPath}' folders/files if empty."
        Remove-Item -LiteralPath $deleteRecursionPath -Force -Verbose
    }
}
# format folder/filename to get a clean name to use in filebot
function Format-Filename {
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputStr
    )
    # Part 1: Replace underscores with spaces
    # Replace a single underscore between letters with a space
    # Replace a single letter surrounded by underscores with a space
    # Remove the space before a single letter not followed by another letter and not 'i' or 'I'
    # Replace one or more consecutive underscores with a single space
    $InputStr = $InputStr -replace '(?<=\p{L})_(?=\p{L})', ' ' -replace '(?<=_)\b\p{L}\b(?=_)', ' ' -replace '(?<!_) (?!(?i:i))\b(?!\p{L})', ' ' -replace '_+', ' '
    $InputStr = $InputStr -replace '(?<=\p{L})-(?=\p{L})', ' ' -replace '(?<=-)\b\p{L}\b(?=-)', ' ' -replace '(?<!-) (?!(?i:i))\b(?!\p{L})', ' '
    #$subpattern = '(?<=\.)[^.]+(?=\.)'
    # $subtitleMatch = [regex]::Match($InputStr, $subpattern)
    # $subtitleString = ($subtitleMatch.Value).ToLower() -replace ' ', '-'
    # $InputStr = $InputStr -replace $subpattern, $subtitleString
    # Part 2: Remove leading space from single character not 'i' or 'I'
    # Replace ' space + single character (not 'i' or 'I' and not '-') + space ' with ' character + space '
    $OutputStr = ($InputStr -replace '\s(?<=[\s])([^\diIaAoO\-\s]) ', '$1 ').Trim()
    return $OutputStr
}
# sanitizng string from non-english characters with '?'
function Format-CleanString {
    param (
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [string]$InputString
    )
    $pattern = '[^\u0000-\u007F]'
    $sanitizedString = $InputString -replace $pattern, '?'
    return $sanitizedString
}

function Remove-Spaces {
    param (
        [Parameter(Mandatory = $true)]
        [string] $removeSpacesFile
    )
    (Get-Content $removeSpacesFile) | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } | Set-Content -Path $removeSpacesFile
    $removeSpacesContent = [System.IO.File]::ReadAllText($removeSpacesFile)
    $removeSpacesContent = $removeSpacesContent.Trim()
    [System.IO.File]::WriteAllText($removeSpacesFile, $removeSpacesContent)
}
# Setup exting script
function Exit-Script {
    $scriptStopWatch.Stop()
    # Cleanup folders
    Invoke-ExpressionConsole -SCMFN 'Cleanup' -SCMFP "Remove-Folders -removeFolder `"$($siteTemp)`" -removeFolderBaseMatch `"$($siteTempBaseMatch)`""
    Invoke-ExpressionConsole -SCMFN 'Cleanup' -SCMFP "Remove-Folders -removeFolder `"$($siteSrc)`" -removeFolderBaseMatch `"$($siteSrcBaseMatch)`""
    Invoke-ExpressionConsole -SCMFN 'Cleanup' -SCMFP "Remove-Folders -removeFolder `"$($siteHome)`" -removeFolderBaseMatch `"$($siteHomeBaseMatch)`""
    if ($overrideDriveList.count -gt 0) {
        foreach ($orDriveList in $overrideDriveList) {
            $orDriveListBaseMatch = ($orDriveList._vsDestPathBase).Replace('\', '\\')
            Invoke-ExpressionConsole -SCMFN 'Cleanup' -SCMFP "Remove-Folders -removeFolder `"$($orDriveList._vsDestPath)`" -removeFolderBaseMatch `"$($orDriveListBaseMatch)`""
        }
    }
    # Cleanup Log Files
    Invoke-ExpressionConsole -SCMFN 'Cleanup' -SCMFP 'Remove-Logfiles'
    $totalRuntime = $($scriptStopWatch.Elapsed.ToString('dd\:hh\:mm\:ss'))
    Write-Output "[END] $(Get-DateTime) - Script completed. Total Elapsed Time: $totalRuntime"
    Stop-Transcript
    ((Get-Content -Path $logFile | Select-Object -Skip 5) | Select-Object -SkipLast 4) | Set-Content -Path $logFile
    Remove-Spaces -removeSpacesFile $logFile
    $logTemp = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime-Temp.log"
    New-Item -Path $logTemp -ItemType File | Out-Null
    $asciiLogo | Out-File -FilePath $logTemp -Width 9999
    if ($vsvTotCount -gt 0) {
        $vsCompletedFilesTable | Out-File -FilePath $logTemp -Width 9999 -Append
    }
    if ($debugScript) {
        Get-Content $logFile -ReadCount 5000 | ForEach-Object {
            $_ | Add-Content -Path $logTemp
        }
    }
    else {
        Get-Content $logFile -ReadCount 5000 | ForEach-Object {
            $_ | Select-String -Pattern '^\[.*\].*has already been recorded in the archive|^\[.*\].*skipping.*' -NotMatch | `
                Select-String -Pattern '^\[download\]\s+(\d+\.\d+)?%.*' -NotMatch | `
                Select-String -Pattern '^\[.*\].*Downloading.*\(.*\).*' -NotMatch | `
                Select-String -Pattern '^\[debug\].*Loaded.*' -NotMatch | `
                Select-String -Pattern '^\[debug\].*Skipping writing playlist thumbnail.*' -NotMatch | `
                Select-String -Pattern '^\[download\].*Finished downloading playlist:.*' -NotMatch | `
                Select-String -Pattern '^\[download\].*Downloading playlist.*' -NotMatch | `
                Select-String -Pattern '^\[generic\].*Downloading webpage.*' -NotMatch | `
                Select-String -Pattern '^\[generic\].*Extracting URL.*' -NotMatch | `
                Select-String -Pattern '^\[redirect\].*redirect to.*' -NotMatch | `
                Select-String -Pattern '^\[DL\:.*\].*' -NotMatch | `
                Select-String -Pattern '^\[.*\] Sleeping.*seconds.*' -NotMatch | `
                Select-String -Pattern '^\[.*\].*Retrieving signed policy.*' -NotMatch | `
                Select-String -Pattern '^WARNING\: The download speed shown is only of one thread. This is a known issue.*' -NotMatch | `
                Add-Content -Path $logTemp
        }
        $modifiedLines = @()
        $lines = Get-Content $logTemp
        for ($i = 0; $i -lt $lines.Length; $i++) {
            if ($lines[$i] -match '^ERROR\: \[.*\].*FutureContent') {
                $modifiedLines += $lines[$i] -replace '^Error\: '
                $i += 2
            }
            else {
                $modifiedLines += $lines[$i]
            }
        }
        $modifiedLines | Out-File $logTemp -Force
    }
    Remove-Item -Path $logFile
    if ($vsvTotCount -gt 0) {
        $newLogFile = "$dateTime-T$vsvTotCount-E$vsvErrorCount.log"
        Rename-Item -Path $logTemp -NewName $newLogFile
    }
    elseif ($testScript) {
        $newLogFile = "$dateTime-DEBUG.log"
        Rename-Item -Path $logTemp -NewName $newLogFile
    }
    else {
        Rename-Item -Path $logTemp -NewName $logFile
    }
    return "[DLP-Script] $(Get-DateTime) - $siteNameO - Ran for $($totalRuntime) - Downloaded $vsvTotCount with $vsvErrorCount errors"
}
# Function to make API requests to Sonarr
function Invoke-SonarrApi {
    param (
        $url
    )
    $method = 'GET'
    $headers = @{ 'X-Api-Key' = $sonarrToken }
    $fullUrl = "$($sonarrHost)/api/v3/$($url)"
    if ($body) {
        $headers['Content-Type'] = 'application/json'
        Invoke-RestMethod -Uri $fullUrl -Method $method -Headers $headers -Body $body
    }
    else {
        Invoke-RestMethod -Uri $fullUrl -Method $method -Headers $headers
    }
}
# Update $vsCompletedFilesList values with write out for logs
function Set-VideoStatus {
    param (
        [parameter(Mandatory = $true)]
        [string]$searchKey,
        [parameter(Mandatory = $true)]
        [string]$searchValue,
        [parameter(Mandatory = $true)]
        [string]$nodeKey,
        [parameter(Mandatory = $true)]
        $nodeValue
    )
    $vsCompletedFilesList | Where-Object { $_.$SearchKey -eq $SearchValue } | ForEach-Object {
        $_.$nodeKey = $nodeValue
        Write-Output "[UpdateVSList] - $(Get-DateTime) - '$nodeKey' - $nodeValue for $searchValue"
    }
}

# Looks at language code in filename to match a pair of vars
function Get-SubtitleLanguage {
    param ($subFiles)
    switch -Regex ($subFiles) {
        '\.ar.*\.' { $stLang = 'ar'; $stTrackName = 'Arabic Sub'; $stLangFull = 'Arabic' }
        '\.de.*\.' { $stLang = 'de'; $stTrackName = 'Deutsch Sub'; $stLangFull = 'Deutsch' }
        '\.en.*\.|\.en-US\.' { $stLang = 'en'; $stTrackName = 'English Sub'; $stLangFull = 'English' }
        '\.es-la\.' { $stLang = 'es'; $stTrackName = 'Spanish(Latin America) Sub'; $stLangFull = 'Spanish(Latin America)' }
        '\.es-es\.|\.es.' { $stLang = 'es-es'; $stTrackName = 'Spanish(Spain) Sub'; $stLangFull = 'Spanish(Spain)' }
        '\.fr.*\.' { $stLang = 'fr'; $stTrackName = 'French Sub'; $stLangFull = 'French' }
        '\.it.*\.' { $stLang = 'it'; $stTrackName = 'Italian Sub'; $stLangFull = 'Italian' }
        '\.ja.*\.' { $stLang = 'ja'; $stTrackName = 'Japanese Sub'; $stLangFull = 'Japanese' }
        '\.pt-br\.' { $stLang = 'pt-br'; $stTrackName = 'Português(Brasil) Sub'; $stLangFull = 'Português(Brasil)' }
        '\.pt-pt\.|\.pt\.' { $stLang = 'pt-pt'; $stTrackName = 'Português(Portugal) Sub'; $stLangFull = 'Português(Portugal)' }
        '\.ru.*\.' { $stLang = 'ru'; $stTrackName = 'Russian Video'; $stLangFull = 'Russian' }
        default { $stLang = 'und'; $stTrackName = 'und sub'; $stLangFull = 'Unknown' }
    }
    $return = @($stLang, $stTrackName, $stLangFull)
    return $return
}

# Sending To Discord for new file notifications
function Invoke-Discord {
    param (
        $site,
        $series,
        $episode,
        $siteIcon,
        $color,
        $subtitle,
        $episodeSize,
        $quality,
        $videoCodec,
        $audioCodec,
        $duration,
        $release,
        $episodeURL,
        $episodeDate,
        $siteFooterIcon,
        $siteFooterText,
        $DiscordSiteUrl
    )
    $embedDebug = @()
    $fieldObjects = @()
    # Start of discord Message
    $site = $(Format-CleanString -InputString $site)
    $series = $(Format-CleanString -InputString $series)
    $episode = $(Format-CleanString -InputString $episode)
    $title = "**$site**"
    $description = "**$series**`n> $episode"
    $color = $color
    $timestamp = Get-DateTime 3
    $thumbnailObject = [PSCustomObject]@{
        url = $siteIcon
    }
    # outputting empty footer for consistant message width
    $footerObject = [PSCustomObject]@{
        icon_url = $siteFooterIcon
        text     = "$siteFooterText"
    }
    # Field Objects
    # Release Field
    $fieldTitleRelease = 'Release'
    $fieldValueRelease = "> $(Format-CleanString -InputString $release)"
    $fieldInlineRelease = 'false'
    $fObjRelease = [PSCustomObject]@{
        name   = $fieldTitleRelease
        value  = $fieldValueRelease
        inline = $fieldInlineRelease
    }
    $fieldObjects += $fObjRelease
    # Episode URL Field
    $fieldTitleEpisodeURL = 'Episode URL'
    $fieldValueEpisodeURL = "> $(Format-CleanString -InputString $episodeURL)"
    $fieldInlineEpisodeURL = 'false'
    $fObjEpisodeURL = [PSCustomObject]@{
        name   = $fieldTitleEpisodeURL
        value  = $fieldValueEpisodeURL
        inline = $fieldInlineEpisodeURL
    }
    $fieldObjects += $fObjEpisodeURL
    # Episode Release Field
    $fieldTitleepisodeDate = 'Release Date'
    $fieldValueepisodeDate = "> $(Format-CleanString -InputString $episodeDate)"
    $fieldInlineepisodeDate = 'false'
    $fObjepisodeDate = [PSCustomObject]@{
        name   = $fieldTitleepisodeDate
        value  = $fieldValueepisodeDate
        inline = $fieldInlineepisodeDate
    }
    $fieldObjects += $fObjepisodeDate
    # Duration Field
    $fieldTitleDuration = 'Duration'
    $fieldValueDuration = $duration
    $fieldInlineDuration = 'true'
    $fObjDuration = [PSCustomObject]@{
        name   = $fieldTitleDuration
        value  = $fieldValueDuration
        inline = $fieldInlineDuration
    }
    $fieldObjects += $fObjDuration
    # Quality Field
    $fieldTitleQuality = 'Quality'
    $fieldValueQuality = $quality
    $fieldInlineQuality = 'true'
    $fObjQuality = [PSCustomObject]@{
        name   = $fieldTitleQuality
        value  = $fieldValueQuality
        inline = $fieldInlineQuality
    }
    $fieldObjects += $fObjQuality
    # Size Field
    $fieldTitleSize = 'Size'
    $fieldValueSize = $episodeSize
    $fieldInlineSize = 'true'
    $fObjSize = [PSCustomObject]@{
        name   = $fieldTitleSize
        value  = $fieldValueSize
        inline = $fieldInlineSize
    }
    $fieldObjects += $fObjSize
    # Video Field
    $fieldTitleVideoCodec = 'Video Codecs'
    $fieldValueVideoCodec = "$videoCodec"
    $fieldInlineVideoCodec = 'true'
    $fObjVideoCodec = [PSCustomObject]@{
        name   = $fieldTitleVideoCodec
        value  = $fieldValueVideoCodec
        inline = $fieldInlineVideoCodec
    }
    $fieldObjects += $fObjVideoCodec
    # Audio Field
    $fieldTitleAudioCodec = 'Audio Codecs'
    $fieldValueAudioCodec = $audioCodec
    $fieldInlineAudioCodec = 'true'
    $fObjAudioCodec = [PSCustomObject]@{
        name   = $fieldTitleAudioCodec
        value  = $fieldValueAudioCodec
        inline = $fieldInlineAudioCodec
    }
    $fieldObjects += $fObjAudioCodec
    # Subtitle Field
    $fieldValueSub = @()
    $fieldTitleSub = 'Subtitle Language'
    foreach ($sub in $subtitle) {
        $sLang = Get-SubtitleLanguage $sub
        $subResult = $($sLang[2])
        $fieldValueSub += $subResult
    }
    $fieldValueSub = (($fieldValueSub -join ', ') -split ', ' | Select-Object -Unique | Sort-Object) -join ', '
    $fieldInlineSub = 'true'
    $fObjSub = [PSCustomObject]@{
        name   = $fieldTitleSub
        value  = $fieldValueSub
        inline = $fieldInlineSub
    }
    $fieldObjects += $fObjSub
    $embedDebug += "Color = $color", "Title = $site", "Series = $($series)", "Episode = $episode", "Thumbnail = $siteIcon", "Duration = $duration", "Qaulity = $quality", "Size = $fieldValueSize", `
        "Video Codecs = $videoCodec", "Audio Codecs = $audioCodec", "Subtitles = $fieldValueSub", "Release = $($release)", "Timestamp = $($timestamp)"
    # Embed object
    [System.Collections.ArrayList]$embedArray = @()
    $embedObject = [PSCustomObject]@{
        color       = $color
        title       = $title
        description = $description
        thumbnail   = $thumbnailObject
        fields      = $fieldObjects
        footer      = $footerObject
        timestamp   = $timestamp
    }
    $payload = [PSCustomObject]@{
        embeds = $embedArray
    }
    $embedArray.Add($embedObject) | Out-Null
    $payloadJson = $payload | ConvertTo-Json -Depth 4
    Invoke-ExpressionConsole -scmFunctionName 'Discord' -scmFunctionParams "write-output `"$($embedDebug -join "`n")`""
    $p = Invoke-WebRequest -Uri $DiscordSiteUrl -Body $payloadJson -Method Post -ContentType 'application/json' | Select-Object -ExpandProperty Headers
    $discordLimit = $p.'x-ratelimit-limit'
    $discordRemainingLimit = $p.'x-ratelimit-remaining'
    $discordResetAfter = [float]::Parse($p.'x-rateLimit-reset-after')
    $discordResetMilliseconds = [int]($discordResetAfter * 1000)
    Write-Output "[Discord] $(Get-DateTime) - Rate limit: $discordRemainingLimit/$discordLimit remaining. Resetting after $discordResetAfter seconds"
    if ($discordRemainingLimit -le 2) {
        Write-Output "[Discord] $(Get-DateTime) - Sleeping for $discordResetMilliseconds"
        Start-Sleep -Milliseconds $discordResetMilliseconds
    }
}
# Run MKVMerge process
function Invoke-MKVMerge {
    param (
        [parameter(Mandatory = $true)]
        [string]$mkvVidInput,
        [parameter(Mandatory = $true)]
        [string]$mkvVidBaseName,
        [parameter(Mandatory = $true)]
        [array]$mkvVidSubtitle,
        [parameter(Mandatory = $true)]
        [string]$mkvVidTempOutput
    )
    Write-Output "[MKVMerge] $(Get-DateTime) - Starting MKVMerge with:"
    Write-Output "[MKVMerge] $(Get-DateTime) - $mkvVidInput"
    Write-Output "[MKVMerge] $(Get-DateTime) - $mkvVidSubtitle"
    Write-Output "[MKVMerge] $(Get-DateTime) - Default Video = $videoLang/$videoTrackName - Default Audio Language = $audioLang/$audioTrackName - Default Subtitle = $subtitleLang/$subtitleTrackName."
    While ($True) {
        if ((Test-Lock $mkvVidInput) -eq $True) {
            continue
        }
        else {
            $sblist = ''
            $mkvCMD = ''
            $mkvVidSubtitle | ForEach-Object {
                $sp = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
                $StLang = Get-SubtitleLanguage -subFiles $sp
                # foreach sub/sublang add to mkv command to run in mkvmerge
                # if sublang defined then that will set track to default.
                if ($stLang[0] -match $subtitleLang) {
                    $subLangCode = $stLang[0]
                    $subTrackName = $stLang[1]
                    $mkvCMD += "--language 0:`"$subLangCode`" --track-name 0:`"$subTrackName`" ( `"$sp`" ) "
                }
                else {
                    $subLangCode = $stLang[0]
                    $subTrackName = $stLang[1]
                    $mkvCMD += "--language 0:`"$subLangCode`" --track-name 0:`"$subTrackName`" --default-track-flag 0:no ( `"$sp`" ) "
                }
            }
            $mkvCMD = $mkvCMD.TrimEnd()
            if ($subFontDir -ne 'None') {
                Write-Output "[MKVMerge] $(Get-DateTime) - MKV subtitle params: `"$mkvCMD`""
                Write-Output "[MKVMerge] $(Get-DateTime) - Combining $sblist and $mkvVidInput files with $subFontDir."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$mkvVidTempOutput`" --language 0:`"$videoLang`" --track-name 0:`"$videoTrackName`" --language 1:`"$audioLang`" --track-name 1:`"$audioTrackName`" ( `"$mkvVidInput`" ) $mkvCMD --attach-file `"$subFontDir`" --attachment-mime-type application/x-truetype-font"
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-DateTime) - Merging as-is. No Font specified for $sblist and $mkvVidInput files with $subFontDir."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$mkvVidTempOutput`" --language 0:`"$videoLang`" --track-name 0:`"$videoTrackName`" --language 1:`"$audioLang`" --track-name 1:`"$audioTrackName`" ( `"$mkvVidInput`" ) $mkvCMD"
            }
        }
        Start-Sleep -Seconds 1
    }
    While (!(Test-Path -Path $mkvVidTempOutput -ErrorAction SilentlyContinue)) {
        Start-Sleep 1.5
    }
    While ($True) {
        if (((Test-Lock $mkvVidInput) -eq $True) -and ((Test-Lock $mkvVidTempOutput) -eq $True)) {
            continue
        }
        else {
            Write-Output "[MKVMerge] $(Get-DateTime) - Removing $mkvVidInput file."
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$mkvVidInput`" -Verbose"
            $mkvVidSubtitle | ForEach-Object {
                $sp = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
                While ($True) {
                    if ((Test-Lock $sp) -eq $True) {
                        continue
                    }
                    else {
                        Write-Output "[MKVMerge] $(Get-DateTime) - Removing $sp file."
                        Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$sp`" -Verbose"
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $mkvVidTempOutput) -eq $True) {
            continue
        }
        else {
            Write-Output "[MKVMerge] $(Get-DateTime) - Renaming $mkvVidTempOutput to $mkvVidInput."
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Rename-Item -Path `"$mkvVidTempOutput`" -NewName `"$mkvVidInput`" -Verbose"
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $mkvVidInput) -eq $True) {
            continue
        }
        else {
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvpropedit `"$mkvVidInput`" --edit track:s1 --set flag-default=1"
            break
        }
        Start-Sleep -Seconds 1
    }
    Set-VideoStatus -searchKey '_vsEpisodeRaw' -searchValue $mkvVidBaseName -nodeKey '_vsMKVCompleted' -nodeValue $($true)
    $videoOverrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsEpisodePath -eq $mkvVidInput } | Select-Object _vsEpisodePath, _vsDestPathDirectory -Unique
    Write-Output "[FileMoving] $(Get-DateTime) - Moving $($videoOverrideDriveList._vsEpisodePath) to $($videoOverrideDriveList._vsDestPathDirectory)."
    if (!(Test-Path -Path $videoOverrideDriveList._vsDestPathDirectory)) {
        Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "New-Folder -newFolderFullPath `"$($videoOverrideDriveList._vsDestPathDirectory)`" -Verbose"
    }
    Write-Output "[FileMoving] $(Get-DateTime) - Moving $($videoOverrideDriveList._vsEpisodePath) to $($videoOverrideDriveList._vsDestPathDirectory)."
    Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($videoOverrideDriveList._vsEpisodePath)`" -Destination `"$($videoOverrideDriveList._vsDestPathDirectory)`" -Force -Verbose"
    if (!(Test-Path -Path $videoOverrideDriveList._vsEpisodePath)) {
        Write-Output "[FileMoving] $(Get-DateTime) - Move completed for $($videoOverrideDriveList._vsEpisodePath)."
        Set-VideoStatus -searchKey '_vsEpisodePath' -searchValue $videoOverrideDriveList._vsEpisodePath -nodeKey '_vsMoveCompleted' -nodeValue $($true)
    }
}
# Function to process video files through FileBot
function Invoke-Filebot {
    param (
        [parameter(Mandatory = $true)]
        [string]$filebotPath,
        [string]$filebotContentType
    )
    Write-Output "[Filebot] $(Get-DateTime) - Looking for files to rename and move to final folder."
    $filebotVideoList = $vsCompletedFilesList | Where-Object { $_._vsDestPath -eq $filebotPath } | Select-Object _vsDestPath, _vsDestPathVideo, _vsEpisodeRaw, _vsEpisodeSubtitle, _vsOverridePath
    $filebotEndParams = '--conflict skip -non-strict --apply date tags clean --log info'
    if ($daily) {
        $filebotEndParams = '--filter "age < 5" ' + $filebotEndParams
    }
    foreach ($filebotFiles in $filebotVideoList) {
        $filebotVidInput = $filebotFiles._vsDestPathVideo
        $filebotSubInput = $filebotFiles._vsEpisodeSubtitle
        $filebotVidBaseName = $filebotFiles._vsEpisodeRaw
        $filebotOverrideDrive = $filebotFiles._vsOverridePath
        if ($siteParentFolder.trim() -ne '' -or $siteSubFolder.trim() -ne '') {
            $FilebotRootFolder = $filebotOverrideDrive + $siteParentFolder
            $filebotBaseFolder = Join-Path -Path $FilebotRootFolder -ChildPath $siteSubFolder
            $filebotParams = Join-Path -Path $filebotBaseFolder -ChildPath $filebotStructure
            $filebotSubParams = $filebotParams + "{'.'+lang.ISO2}"
            Write-Output "[Filebot] $(Get-DateTime) - Files found($filebotVidInput). Renaming video and moving files to final folder. Using path($filebotStructure)."
            Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename -r `"$filebotVidInput`" --db `"$filebotDB`" --format `"$filebotParams`" $filebotEndParams"
            if (!($mkvMerge)) {
                $filebotSubInput | ForEach-Object {
                    $filebotSubParams = $filebotStructure
                    $osp = $_ | Where-Object { $_.key -eq 'overrideSubPath' } | Select-Object -ExpandProperty value
                    $StLang = Get-SubtitleLanguage -subFiles $osp
                    $filebotSubParams = $filebotSubParams + "{'.$($StLang[0])'}"
                    Write-Output "[Filebot] $(Get-DateTime) - Files found($osp). Renaming subtitle and moving files to final folder. Using path($filebotSubParams)."
                    Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename -r `"$osp`" --db `"$filebotDB`" --format `"$filebotSubParams`" $filebotEndParams"
                }
            }
        }
        else {
            Write-Output "[Filebot] $(Get-DateTime) - Files found($filebotVidInput). ParentFolder or Subfolder path not specified. Renaming files in place  using path($filebotStructure)."
            $filebotSubInput | ForEach-Object {
                Write-Output "[Filebot] $(Get-DateTime) - Files found($filebotVidInput). ParentFolder or Subfolder path not specified. Renaming files in place."
                Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename -r `"$filebotVidInput`" --db `"$filebotDB`" --format `"$filebotParams`" $filebotEndParams"
            }
            if (!($mkvMerge)) {
                $filebotSubInput | ForEach-Object {
                    $filebotSubParams = $filebotStructure
                    $osp = $_ | Where-Object { $_.key -eq 'overrideSubPath' } | Select-Object -ExpandProperty value
                    $StLang = Get-SubtitleLanguage -subFiles $osp
                    $filebotSubParams = $filebotSubParams + "{'.$($StLang[0])'}"
                    $filebotSubParams
                    Write-Output "[Filebot] $(Get-DateTime) - Files found($osp). Renaming subtitle($($StLang[0])) and renaming files in place using path($filebotSubParams)."
                    Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename -r `"$osp`" --db `"$filebotDB`" --format `"$filebotSubParams`" $filebotEndParams"
                }
            }
        }
        if (!(Test-Path -Path $filebotVidInput)) {
            Write-Output "[Filebot] $(Get-DateTime) - Setting file($filebotVidInput) as completed."
            Set-VideoStatus -searchKey '_vsEpisodeRaw' -searchValue $filebotVidBaseName -nodeKey '_vsFBCompleted' -nodeValue $($true)
        }
        else {
            Write-Output "[Filebot] $(Get-DateTime) - Failed to match. Setting file($filebotVidInput) as errored."
            Set-VideoStatus -searchKey '_vsEpisodeRaw' -searchValue $filebotVidBaseName -nodeKey '_vsErrored' -nodeValue $($true)
        }
    }
    $vsvFBCount = ($vsCompletedFilesList | Where-Object { $_._vsFBCompleted -eq $true } | Measure-Object).Count
    if ($vsvFBCount -eq $vsvTotCount ) {
        Write-Output "[Filebot]$(Get-DateTime) - Filebot($vsvFBCount) = ($vsvTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup."
        Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -script fn:cleaner `"$siteHome`" --log info"
    }
    else {
        Write-Output "[Filebot] $(Get-DateTime) - Filebot($vsvFBCount) and Total Video($vsvTotCount) count mismatch. Manual check required."
    }
    if ($vsvFBCount -ne $vsvTotCount) {
        Write-Output "[Filebot] $(Get-DateTime) - [Cleanup] - File needs processing."
    }
}


# Update Subtitle files with font name
function Update-SubtitleStyle {
    param (
        $SubtitleFilePath,
        $subFontName
    )
    # Top
    $subtitleContentTags = Get-Content -Path $SubtitleFilePath -Raw
    $substring = $subtitleContentTags.Substring(0, $subtitleContentTags.IndexOf('[V4+ Styles]'))
    $result = @()
    $string = $substring -split "`r`n"
    $stl = Get-SubtitleLanguage $SubtitleFilePath
    foreach ($i in $string) {
        switch -Regex ( $i) {
            '(?<=^Title:).*' { $i = "Title: $($stl[2])" ; $result += $i }
            '(?<=^Original Translation:).*' { $i = 'Original Translation:' ; $result += $i }
            '(?<=^Original Editing:).*' { $i = 'Original Editing:' ; $result += $i }
            '(?<=^Original Timing:).*' { $i = 'Original Timing:' ; $result += $i }
            '(?<=^Script Updated By:).*' { $i = 'Script Updated By:' ; $result += $i }
            '(?<=^Update Details:).*' { $i = 'Update Details:' ; $result += $i }
            '(?<=^Original Script:).*' { $i = 'Original Script:'; $result += $i }
            '(?<=^ScaledBorderAndShadow:).*' { $i = 'ScaledBorderAndShadow: yes' ; $result += $i }
            '(?<=^WrapStyle:).*' { $i = 'WrapStyle: 0'; $result += $i }
            '(?<=^Collisions:).*' { $i = 'Collisions: Normal'; $result += $i }
            '(?<=^PlayResX:).*' { $i = 'PlayResX: 640'; $result += $i }
            '(?<=^PlayResY:).*' { $i = 'PlayResY: 360'; $result += $i }
            '(?<=^YCbCr Matrix:).*' { $i = 'YCbCr Matrix: TV.709'; $result += $i }
            Default { if ($i -notmatch '^\s*$') { $result += $i } }
        }
    }
    $subtitlePatterns = [ordered]@{
        '^Title.*'                 = "Title: $($stl[2])"
        '^ScaledBorderAndShadow.*' = 'ScaledBorderAndShadow: yes'
        '^Collisions.*'            = 'Collisions: Normal'
        '^WrapStyle.*'             = 'WrapStyle: 0'
        '^PlayResX.*'              = 'PlayResX: 640'
        '^PlayResY.*'              = 'PlayResY: 360'
        '^YCbCr Matrix.*'          = 'YCbCr Matrix: TV.709'
    }
    $c = 0
    # Extract existing values from original script lines
    $existingValues = @()
    foreach ($line in $result) {
        foreach ($pattern in $subtitlePatterns.Keys) {
            if ($line -match $pattern) {
                $existingValues += $line
            }
        }
    }
    # Add new lines from the patterns and associated values if not already present
    foreach ($pattern in $subtitlePatterns.Keys) {
        $patternMatched = $false
        foreach ($value in $existingValues) {
            if ($value -match $pattern) {
                Write-Output "pattern found: $pattern - $value"
                $patternMatched = $true
                $c++
                break
            }
        }
        if ( $c -ne 0) {
            if (-not $patternMatched) {
                $result += $subtitlePatterns[$pattern]
            }
        }
    }
    $subtitleContent = Get-Content $SubtitleFilePath
    $startLine = $subtitleContent | Select-String -Pattern '^Format:.*' | Select-Object -First 1 | ForEach-Object { $_.LineNumber }
    $endLine = $subtitleContent | Select-String -Pattern '^Style:.*' | Select-Object -Last 1 | ForEach-Object { $_.LineNumber }
    Write-Output "[SubtitleRegex] $(Get-DateTime) - Found lines $startLine to $endLine."
    # Middle - format header
    $formatHeader = $subtitleContent | Select-String -Pattern '^Format:.*' | Select-Object -First 1
    [string]$formatStyleBlock = ''
    $formatStyleBlock += ($formatHeader -Replace ('^Format: ', '') -replace (', ', ','))
    # Middle - style rows
    $subtitleContent | Select-String -Pattern '^Style:.*' | ForEach-Object {
        $formatStyleBlock += "`n" + ($_ -Replace ('^Style: ', '') -replace (', ', ','))
    }
    $styleBlockCSV = ConvertFrom-Csv -InputObject $formatStyleBlock -Delimiter ','
    $styleBlockCSV | ForEach-Object {
        $_.Fontname = $subFontName
        if ($siteName -eq 'hidive') {
            $_.Fontsize = 24
        }
        $_.Outline = if ($_.Outline -lt 1.5) {
            1.5
        }
        else {
            $_.Outline
        }
        $_.Shadow = 0
    }
    Write-Output "[SubtitleRegex] $(Get-DateTime) - Start of initial updated lines:"
    ($styleBlockCSV | Format-Table | Out-String) -split "`n" | Where-Object { $_.Trim() -ne '' } | ForEach-Object {
        Write-Output "[SubtitleRegex] $(Get-DateTime) - $_"
    }
    Write-Output "[SubtitleRegex] $(Get-DateTime) - End of initial updated lines."
    $styleBlockToCSV = $styleBlockCSV | ConvertTo-Csv -Delimiter ',' -UseQuotes Never
    $Headrow = $styleBlockToCSV | Select-Object -First 1
    $Header = 'Format: ' + $Headrow
    $rows = $styleBlockToCSV | Select-Object -Skip 1 | ForEach-Object {
        'Style: ' + $_
    }
    $newStrings = @()
    $newStrings = @($Header, $rows)
    Write-Output "`n[SubtitleRegex] $(Get-DateTime) - Start of final updated line block to file:"
    ($newStrings | Format-Table | Out-String) -split "`n" | Where-Object { $_.Trim() -ne '' } | ForEach-Object {
        Write-Output "[SubtitleRegex] $(Get-DateTime) - $_"
    }
    Write-Output "[SubtitleRegex] $(Get-DateTime) - End of final updated line block to file."
    $styles = "`n[V4+ Styles]`n$($newStrings[0])`n"
    foreach ($n in $newStrings[1]) {
        $styles += $n + "`n"
    }
    $styles = $styles
    # bottom
    $events = $subtitleContent | Select-Object -Skip $endLine
    $top = $result | Where-Object { $_ -ne $null -and $_.Trim() -ne '' }
    $styles = $styles | Where-Object { $_ -ne $null -and $_.Trim() -ne '' }
    $events = $events | Where-Object { $_ -ne $null -and $_.Trim() -ne '' }
    ($top + $styles + $events) | Set-Content $SubtitleFilePath
    Write-Output "[SubtitleRegex] $(Get-DateTime) - Finished updating subtitle: $SubtitleFilePath."
}

# Setting up arraylist for MKV and Filebot lists
class VideoStatus {
    [string]$_vsSite
    [string]$_vsSeries
    [string]$_vsEpisode
    [string]$_vsEpisodeDuration
    [string]$_vsEpisodeSize
    [string]$_vsEpisodeRes
    [string]$_vsEpisodeVideoCodec
    [string]$_vsEpisodeAudioCodec
    [string]$_vsSeriesDirectory
    [string]$_vsEpisodeRaw
    [string]$_vsEpisodeTemp
    [string]$_vsEpisodePath
    [string]$_vsOverridePath
    [string]$_vsDestPathBase
    [string]$_vsDestPath
    [string]$_vsDestPathDirectory
    [string]$_vsDestPathVideo
    [string]$_vsFinalPathVideo
    [string]$_vsFilebotTitle
    [string]$_vsEpisodeUrl
    [string]$_vsEpisodeDate
    [array]$_vsEpisodeSubtitle
    [bool]$_vsSECompleted
    [bool]$_vsMKVCompleted
    [bool]$_vsMoveCompleted
    [bool]$_vsFBCompleted
    [bool]$_vsErrored

    VideoStatus([string]$vsSite, [string]$vsSeries, [string]$vsEpisode, [string]$vsEpisodeDuration, [string]$vsEpisodeSize, [string]$vsEpisodeRes, [string]$vsEpisodeVideoCodec, [string]$vsEpisodeAudioCodec, [string]$vsSeriesDirectory, [string]$vsEpisodeRaw, [string]$vsEpisodeTemp, [string]$vsEpisodePath, `
            [string]$vsOverridePath, [string]$vsDestPathBase, [string]$vsDestPath, [string]$vsDestPathDirectory, [string]$vsDestPathVideo, [string]$vsFinalPathVideo, [string]$vsFilebotTitle, [string]$vsEpisodeUrl, [string]$vsEpisodeDate, [array]$vsEpisodeSubtitle, [bool]$vsSECompleted, [bool]$vsMKVCompleted, `
            [bool]$vsMoveCompleted, [bool]$vsFBCompleted, [bool]$vsErrored) {
        $this._vsSite = $vsSite
        $this._vsSeries = $vsSeries
        $this._vsEpisode = $vsEpisode
        $this._vsEpisodeDuration = $vsEpisodeDuration
        $this._vsEpisodeSize = $vsEpisodeSize
        $this._vsEpisodeRes = $vsEpisodeRes
        $this._vsEpisodeVideoCodec = $vsEpisodeVideoCodec
        $this._vsEpisodeAudioCodec = $vsEpisodeAudioCodec
        $this._vsSeriesDirectory = $vsSeriesDirectory
        $this._vsEpisodeRaw = $vsEpisodeRaw
        $this._vsEpisodeTemp = $vsEpisodeTemp
        $this._vsEpisodePath = $vsEpisodePath
        $this._vsOverridePath = $vsOverridePath
        $this._vsDestPathBase = $vsDestPathBase
        $this._vsDestPath = $vsDestPath
        $this._vsDestPathDirectory = $vsDestPathDirectory
        $this._vsDestPathVideo = $vsDestPathVideo
        $this._vsFinalPathVideo = $vsFinalPathVideo
        $this._vsFilebotTitle = $vsFilebotTitle
        $this._vsEpisodeUrl = $vsEpisodeUrl
        $this._vsEpisodeDate = $vsEpisodeDate
        $this._vsEpisodeSubtitle = $vsEpisodeSubtitle
        $this._vsSECompleted = $vsSECompleted
        $this._vsMKVCompleted = $vsMKVCompleted
        $this._vsMoveCompleted = $vsMoveCompleted
        $this._vsFBCompleted = $vsFBCompleted
        $this._vsErrored = $vsErrored
    }
}

[System.Collections.ArrayList]$vsCompletedFilesList = @()

# Start of script Variable setup
$scriptDirectory = $PSScriptRoot
$configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
#$sharedFolder = Join-Path -Path $scriptDirectory -ChildPath 'shared'
$fontFolder = Join-Path -Path $scriptDirectory -ChildPath 'fonts'
$xmlConfig = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <Directory>
        <!-- Folder to store backup of site config/archive/bat/cookie files
            ex: <backup location="E:\Backup" />
        -->
        <backup location="" />
        <!-- Temp folder used by YT-DLP
            ex: <temp location="D:\tmp" />
        -->
        <temp location="" />
        <!-- Staging folder for post-process after download
            ex: <src location="D:\src" />
        -->
        <src location="" />
        <!-- Default final staging after post-processing. Can be overridden per series/show under OverrideSeries section
            ex: <dest location="J:\tmp" />
        -->
        <dest location="" />
        <!-- Location of YT-DLP FFMPEG version
            ex: <ffmpeg location="D:\Common\ffmpeg\bin" />
        -->
        <ffmpeg location="" />
    </Directory>
    <Logs>
        <!-- How long to keep folders that did or didn't result in a download -->
        <keeplog emptylogskeepdays="0" filledlogskeepdays="7" />
    </Logs>
    <Plex>
        <!-- PLEX local IP and token used to update plex library after success processing -->
        <plexcred plexUrl="" plexToken="" />
    </Plex>
    <Filebot>
        <!-- Uses dest drive with site parentfolder/subfolder and argument to run filebot command. may want to run for 1 episode to see what the series folder name comes out to.
            ex: <fbfolder fbArgument="{n}\{'Season '+s00}\{n} - {s00e00} - {t}" />
        -->
        <fbfolder fbArgument="{n}\{'Season '+s00}\{n} - {s00e00} - {t}" />
    </Filebot>
    <OverrideSeries>
        <!-- Used to move files to different tmp folder than whats defined in defauilt above.
            <override orSeriesName="MySeriesThatIsInADifferentDriveThanDefaultTemp" orSrcdrive="I:\" />
        -->
        <override orSeriesName="" orSrcdrive="" />
        <override orSeriesName="" orSrcdrive="" />
        <override orSeriesName="" orSrcdrive="" />
    </OverrideSeries>
    <OverrideEposides>
        <override overrideType="" filenamePattern="" fileReplaceText="" />
        <override overrideType="" filenamePattern="" fileReplaceText="" />
        <override overrideType="" filenamePattern="" fileReplaceText="" />
    </OverrideEposides>
    <Discord>
        <!-- Discord webhook url and default icon pointed to static web image />
        <hook
            url="" />
        <icon
            Default=""
            Color="8359053" />
    </Discord>
    <credentials>
        <!-- Where you store the Site name, username/password, plexlibraryid, folder in library, and a custom font used to embed into video/sub
            <site sitename="MySiteHere">
                <username>MyUserName</username>
                <password>MyPassword</password>
                <plexlibraryid>4</plexlibraryid>
                <parentfolder>Video</parentfolder>
                <subfolder>A</subfolder>
                <font>Marker SD.ttf</font>
                <icon
                    url=""
                    color="" />
                <icon
                    url=""
                    color="" />
            </site>
        -->
        <site sitename="">
            <username></username>
            <password></password>
            <plexplexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
            <icon
                url=""
                color="" />
            <icon
                url=""
                color="" />
        </site>
        <site sitename="">
            <username></username>
            <password></password>
            <plexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
            <icon
                url=""
                color="" />
            <icon
                url=""
                color="" />
        </site>
        <site sitename="">
            <username></username>
            <password></password>
            <plexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
            <icon
                url=""
                color="" />
            <icon
                url=""
                color="" />
        </site>
    </credentials>
</configuration>
'@
$defaultConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,episode" "[!?$%^@:.#+-]" " "
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
--sub-format 'ass/srt/vtt'
--sub-format 'vtt/srt'
--convert-subs 'ass'
--write-subs
--embed-chapters
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/%(series).110s - S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$vrvConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,episode" "[!?$%^@:.#+-]" " "
--trim-filenames 248
--add-metadata
--sub-langs "en-US"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-chapters
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv[format_id*=-ja-JP][format_id!*=hardsub][height>=1080]+ba[format_id*=-ja-JP][format_id!*=hardsub] / b[format_id*=-ja-JP][format_id!*=hardsub][height>=1080] / b*[format_id*=-ja-JP][format_id!*=hardsub]'
-o '%(series).110s/%(series).110s - S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$crunchyrollConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,episode" "[!?$%^@:.#+-]" " "
--trim-filenames 248
--add-metadata
--sub-langs "en-US"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-chapters
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--match-filter "season !~='(?i)\b(?:dubs?)\b' & title !~='(?i)\b(?:dubs?)\b'"
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv[height>=1080]+ba[height>=1080] / bv+ba / b*'
-o '%(series).110s/%(series).110s - S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$funimationConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,episode" "[!?$%^@:.#+-]" " "
--trim-filenames 248
--add-metadata
--sub-langs 'en.*'
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-chapters
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
--no-check-certificate
-N 32
--downloader aria2c
--extractor-args 'funimation:language=japanese'
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / b'
-o '%(series).110s/%(series).110s - S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$hidiveConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,episode" "[!?$%^@:.#+-]" " "
--trim-filenames 248
--add-metadata
--sub-langs "english-subs"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-chapters
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/%(series).110s - S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$paramountPlusConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,episode" "[!?$%^@:.#+-]" " "
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-chapters
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/%(series).110s - S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$asciiLogo = @'
___________________________________________________________________________________________________________________________________________________
                                     ____  _     ____       _____ _ _      _   _                 _ _
                                    |  _ \| |   |  _ \     |  ___(_) | ___| | | | __ _ _ __   __| | | ___ _ __
                                    | | | | |   | |_) |____| |_  | | |/ _ \ |_| |/ _` | '_ \ / _` | |/ _ \ '__|
                                    | |_| | |___|  __/_____|  _| | | |  __/  _  | (_| | | | | (_| | |  __/ |
                                    |____/|_____|_|        |_|   |_|_|\___|_| |_|\__,_|_| |_|\__,_|_|\___|_|
___________________________________________________________________________________________________________________________________________________
'@

if (!(Test-Path -Path "$scriptDirectory\config.xml" -PathType Leaf)) {
    New-Item -Path "$scriptDirectory\config.xml" -ItemType File -Force
    Write-Output "No config xml found at $ScriptDirectory. Created new config. Exiting script."
    exit
}
if ($help) {
    Show-Markdown -Path "$scriptDirectory\README.md" -UseBrowser
    exit
}
if ($newConfig) {
    if (!(Test-Path -Path $configPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-Content -Path $configPath))) {
        New-Item -Path $configPath -ItemType File -Force
        $xmlConfig | Set-Content -Path $configPath
        Write-Output "$configPath File Created successfully."
    }
    else {
        Write-Output "$configPath File Exists."
    }
    exit
}
if ($supportFiles) {
    Invoke-ExpressionConsole -SCMFN 'SupportFolders' -SCMFP "New-Folder -newFolderFullPath `"$fontFolder`""
    $configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
    [xml]$configFile = Get-Content -Path $configPath
    $sitenameXML = $configFile.configuration.credentials.site | Where-Object { $_.siteName.trim() -ne '' } | Select-Object 'siteName' -ExpandProperty siteName
    $sitenameXML | ForEach-Object {
        $sn = $_.siteName
        $scc = Join-Path -Path $scriptDirectory -ChildPath 'sites'
        $scf = Join-Path -Path $scc -ChildPath $sn
        $scdf = "$($scf.TrimEnd('\'))_D"
        $siteSupportFolders = $scc, $scf, $scdf
        foreach ($F in $siteSupportFolders) {
            Invoke-ExpressionConsole -SCMFN 'SiteFolders' -SCMFP "New-Folder -newFolderFullPath `"$F`""
        }
        $sadf = Join-Path -Path $scdf -ChildPath "$($sn)_D_A"
        $sbdf = Join-Path -Path $scdf -ChildPath "$($sn)_D_B"
        $sbdc = Join-Path -Path $scdf -ChildPath "$($sn)_D_C"
        $saf = Join-Path -Path $scf -ChildPath "$($sn)_A"
        $sbf = Join-Path -Path $scf -ChildPath "$($sn)_B"
        $sbc = Join-Path -Path $scf -ChildPath "$($sn)_C"
        $siteSuppFiles = $sadf, $sbdf, $sbdc, $saf, $sbf, $sbc
        foreach ($S in $siteSuppFiles) {
            Invoke-ExpressionConsole -SCMFN 'SiteFiles' -SCMFP "New-SuppFile -newSupportFiles `"$S`""
        }
        $scfDC = Join-Path -Path $scdf -ChildPath 'yt-dlp.conf'
        $scfC = Join-Path -Path $scf -ChildPath 'yt-dlp.conf'
        $siteConfigFiles = $scfDC , $scfC
        foreach ($cf in $siteConfigFiles) {
            Invoke-ExpressionConsole -SCMFN 'SiteConfigs' -SCMFP "New-Config -newConfigs `"$cf`""
            Remove-Spaces -removeSpacesFile $cf
        }
    }
    exit
}
if ($site) {
    $date = Get-DateTime 1
    $dateTime = Get-DateTime
    $time = Get-DateTime 2
    $site = $site.ToLower()
    # Reading from XML
    $configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
    [xml]$configFile = Get-Content -Path $configPath
    $ffmpeg = $configFile.configuration.Directory.ffmpeg.location
    [int]$emptyLogs = $configFile.configuration.Logs.keeplog.emptylogskeepdays
    [int]$filledLogs = $configFile.configuration.Logs.keeplog.filledlogskeepdays
    $sonarrHost = $configFile.configuration.Sonarr.sonarrcred.sonarrUrl
    $sonarrToken = $configFile.configuration.Sonarr.sonarrcred.sonarrToken
    $plexHost = $configFile.configuration.Plex.plexcred.plexUrl
    $plexToken = $configFile.configuration.Plex.plexcred.plexToken
    # Drives
    $backupDrive = $configFile.configuration.Directory.backup.location
    $tempDrive = $configFile.configuration.Directory.temp.location
    $srcDrive = $configFile.configuration.Directory.src.location
    $destDrive = $configFile.configuration.Directory.dest.location
    $srcBackup = Join-Path -Path $backupDrive -ChildPath 'yt-dlp'
    $srcDriveSharedFonts = Join-Path -Path $srcBackup -ChildPath 'fonts'
    # Site specific params
    $siteParams = $configFile.configuration.credentials.site | Where-Object { $_.siteName -ne '' -or $_.sitename -ne $null } #| Select-Object 'siteName', 'username', 'password', 'plexlibraryid', 'parentfolder', 'subfolder', 'font', 'fbtype', 'useSonarr', 'usePlex', 'icon'
    $siteNameParams = $siteParams | Where-Object { $_.siteName.ToLower() -eq $site } | Select-Object * -First 1
    $siteNameCount = 0
    $SiteNameList = @()
    foreach ($sp in $siteParams) {
        $SiteNameID = $sp.sitename
        $siteNameCount = $siteNameCount + 1
        $siteFolderId = $sp.sitename.Substring(0, 1) + $siteNameCount
        $SiteNameList += , ($siteFolderId, $SiteNameID)
    }
    $siteFolderIdName = $SiteNameList | Where-Object { $_ -eq $site }
    $siteNameO = $siteNameParams.siteName
    $siteName = $siteNameO.ToLower()
    $siteNameRaw = $siteNameParams.siteName
    $siteFolderDirectory = Join-Path -Path $scriptDirectory -ChildPath 'sites'
    $siteUser = $siteNameParams.username
    $sitePass = $siteNameParams.password
    $SiteContentType = $siteNameParams.fbtype
    $siteLibraryID = $siteNameParams.plexlibraryid
    $siteParentFolder = $siteNameParams.parentfolder
    $siteSubFolder = $siteNameParams.subfolder
    $subFontExtension = $siteNameParams.font
    $useSonarr = $siteNameParams.useSonarr
    $usePlex = $siteNameParams.usePlex
    if ($daily) {
        $siteType = "$($siteName)_D"
        $siteFolder = Join-Path -Path $siteFolderDirectory -ChildPath $siteType
        $siteConfig = Join-Path -Path $siteFolder -ChildPath 'yt-dlp.conf'
        $logFolderBase = Join-Path -Path $siteFolder -ChildPath 'log'
        $logFolderBaseDate = Join-Path -Path $logFolderBase -ChildPath $date
        $logFile = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime.log"
    }
    else {
        $siteType = $siteName
        $siteFolder = Join-Path -Path $siteFolderDirectory -ChildPath $siteType
        $siteConfig = Join-Path -Path $siteFolder -ChildPath 'yt-dlp.conf'
        $logFolderBase = Join-Path -Path $siteFolder -ChildPath 'log'
        $logFolderBaseDate = Join-Path -Path $logFolderBase -ChildPath $date
        $logFile = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime.log"
    }
    Start-Transcript -Path $logFile -UseMinimalHeader
    Write-Output "[Setup] $(Get-DateTime) - $siteNameRaw"
    $cookieFile = Join-Path -Path $siteFolder -ChildPath "$($siteType)_C"
    if ($overrideBatch) {
        $batFile = $overrideBatch
    }
    else {
        $batFile = Join-Path -Path $siteFolder -ChildPath "$($siteType)_B"
    }
    $archiveFile = if ($archiveTemp) {
        Join-Path -Path $siteFolder -ChildPath "$($siteType)_A_Temp"
    }
    else {
        Join-Path -Path $siteFolder -ChildPath "$($siteType)_A"
    }
    $siteConfigBackup = Join-Path -Path (Join-Path -Path $srcBackup -ChildPath 'sites') -ChildPath $siteType
    $filebotArgument = $configFile.configuration.Filebot.fbfolder | Where-Object { $_.type -eq $SiteContentType } | Select-Object fbArgumentdb, fbArgumentStructure
    foreach ($fb in $filebotArgument) {
        $filebotDB = $fb.fbArgumentdb
        $filebotStructure = $fb.fbArgumentStructure
    }
    $overrideSeriesList = $configFile.configuration.OverrideSeries.override | Where-Object { $_.orSeriesId -ne '' -and $_.orSrcdrive -ne '' }
    $overrideEpisodeList = $configFile.configuration.OverrideEposides.override | Where-Object { $_.filenamePattern -ne '' -and $_.fileReplaceText -ne '' }
    $discordHookUrl = $configFile.configuration.Discord.hook.url
    $discordHookErrorURL = $configFile.configuration.Discord.hook.error
    $discordSiteIcon = $siteNameParams.icon.url
    $discordSiteColor = $siteNameParams.icon.color
    $discordIconDefault = $configFile.configuration.Discord.icon.Default
    $discordColorDefault = $configFile.configuration.Discord.icon.Color
    $discordFooterIconDefault = $configFile.configuration.Discord.icon.footerIcon
    if ($($discordSiteIcon.Trim()).length -eq 0) {
        $discordSiteIcon = $discordIconDefault
    }
    if ($($discordSiteColor.Trim()).length -eq 0) {
        $discordSiteColor = $discordColorDefault
    }
    $siteDefaultPath = Join-Path -Path (Split-Path $destDrive -Qualifier) -ChildPath ($siteParentFolder + '\' + $siteSubFolder)
    # Video track title inherits from Audio language code
    if ($audioLang -eq '' -or $null -eq $audioLang) {
        $audioLang = 'und'
    }
    if ($subtitleLang -eq '' -or $null -eq $subtitleLang) {
        $subtitleLang = 'und'
    }
    switch ($audioLang) {
        ar { $videoLang = $audioLang; $audioTrackName = 'Arabic Audio'; $videoTrackName = 'Arabic Video' }
        de { $videoLang = $audioLang; $audioTrackName = 'Deutsch Audio'; $videoTrackName = 'Deutsch Video' }
        en { $videoLang = $audioLang; $audioTrackName = 'English Audio'; $videoTrackName = 'English Video' }
        es-la { $videoLang = $audioLang; $audioTrackName = 'Spanish(Latin America) Audio'; $videoTrackName = 'Spanish(Latin America) Video' }
        es { $videoLang = $audioLang; $audioTrackName = 'Spanish(Spain) Audio'; $videoTrackName = 'Spanish(Spain) Video' }
        fr { $videoLang = $audioLang; $audioTrackName = 'French Audio'; $videoTrackName = 'French Video' }
        it { $videoLang = $audioLang; $audioTrackName = 'Italian Audio'; $videoTrackName = 'Italian Video' }
        ja { $videoLang = $audioLang; $audioTrackName = 'Japanese Audio'; $videoTrackName = 'Japanese Video' }
        pt-br { $videoLang = $audioLang; $audioTrackName = 'Português (Brasil) Audio'; $videoTrackName = 'Português (Brasil) Video' }
        pt-pt { $videoLang = $audioLang; $audioTrackName = 'Português (Portugal) Audio'; $videoTrackName = 'Português (Portugal) Video' }
        ru { $videoLang = $audioLang; $audioTrackName = 'Russian Audio'; $videoTrackName = 'Russian Video' }
        und { $audioLang = 'und'; $videoLang = 'und'; $audioTrackName = 'und audio'; $videoTrackName = 'und Video' }
    }
    switch ($subtitleLang) {
        ar { $subtitleTrackName = 'Arabic Sub' }
        de { $subtitleTrackName = 'Deutsch Sub' }
        en { $subtitleTrackName = 'English Sub' }
        es-la { $subtitleTrackName = 'Spanish(Latin America) Sub' }
        es { $subtitleTrackName = 'Spanish(Spain) Sub' }
        fr { $subtitleTrackName = 'French Sub' }
        it { $subtitleTrackName = 'Italian Sub' }
        ja { $subtitleTrackName = 'Japanese Sub' }
        pt-br { $subtitleTrackName = 'Português (Brasil) Sub' }
        pt-pt { $subtitleTrackName = 'Português (Portugal) Sub' }
        ru { $subtitleTrackName = 'Russian Video' }
        und { $subtitleTrackName = 'und sub' }
    }
    # End reading from XML
    if ($subFontExtension.Trim() -ne '') {
        $subFontDir = Join-Path -Path $fontFolder -ChildPath $subFontExtension
        if (Test-Path -Path $subFontDir) {
            $subFontName = [System.Io.Path]::GetFileNameWithoutExtension($subFontExtension)
            Write-Output "[Setup] $(Get-DateTime) - $subFontExtension set for $siteName."
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - $subFontExtension specified in $configFile is missing from $fontFolder. Exiting."
            Exit-Script
        }
    }
    else {
        $subFontDir = 'None'
        $subFontName = 'None'
        $subFontExtension = 'None'
        Write-Output "[Setup] $(Get-DateTime) - $subFontExtension - No font set for $siteName."
    }
    $dlpParams = 'yt-dlp'
    $dlpArray = @()
    if ($daily) {
        $siteTempBase = Join-Path -Path $tempDrive -ChildPath $siteFolderIdName[0]
        $siteTempBaseMatch = $siteTempBase.Replace('\', '\\')
        $siteTemp = Join-Path -Path $siteTempBase -ChildPath $time
        $siteSrcBase = Join-Path -Path $srcDrive -ChildPath $siteFolderIdName[0]
        $siteSrcBaseMatch = $siteSrcBase.Replace('\', '\\')
        $siteSrc = Join-Path -Path $siteSrcBase -ChildPath $time
        $siteHomeBase = Join-Path -Path (Join-Path -Path $destDrive -ChildPath "_$siteSubFolder") -ChildPath $siteFolderIdName[0]
        $siteHomeBaseMatch = $siteHomeBase.Replace('\', '\\')
        $siteHome = Join-Path -Path $siteHomeBase -ChildPath $time
    }
    else {
        $siteTempBase = Join-Path -Path $tempDrive -ChildPath "$($siteFolderIdName[0])M"
        $siteTempBaseMatch = $siteTempBase.Replace('\', '\\')
        $siteTemp = Join-Path -Path $siteTempBase -ChildPath $time
        $siteSrcBase = Join-Path -Path $srcDrive -ChildPath "$($siteFolderIdName[0])M"
        $siteSrcBaseMatch = $siteSrcBase.Replace('\', '\\')
        $siteSrc = Join-Path -Path $siteSrcBase -ChildPath $time
        $siteHomeBase = Join-Path -Path (Join-Path -Path $destDrive -ChildPath '_M') -ChildPath $siteFolderIdName[0]
        $siteHomeBaseMatch = $siteHomeBase.Replace('\', '\\')
        $siteHome = Join-Path -Path $siteHomeBase -ChildPath $time
    }
    if ($srcDrive -eq $tempDrive) {
        Write-Output "[Setup] $(Get-DateTime) - Src($srcDrive) and Temp($tempDrive) Directories cannot be the same"
        Exit-Script
    }
    if ((Test-Path -Path $siteConfig)) {
        Write-Output "[Setup] $(Get-DateTime) - $siteConfig file found. Continuing."
        $dlpParams = $dlpParams + " --config-location $siteConfig -P temp:$siteTemp -P home:$siteSrc"
        $dlpArray += '--config-location', "$siteConfig", '-P', "temp:$siteTemp", '-P', "home:$siteSrc"
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - $siteConfig does not exist. Exiting."
        Exit-Script
    }
    if ($login) {
        if ($siteUser -and $sitePass) {
            Write-Output "[Setup] $(Get-DateTime) - Login is true and SiteUser/Password is filled. Continuing."
            $dlpParams = $dlpParams + " -u $siteUser -p $sitePass"
            $dlpArray += '-u', "$siteUser", '-p', "$sitePass"
            if ($cookies) {
                if ((Test-Path -Path $cookieFile)) {
                    Write-Output "[Setup] $(Get-DateTime) - Cookies is true and $cookieFile file found. Continuing."
                    $dlpParams = $dlpParams + " --cookies $cookieFile"
                    $dlpArray += '--cookies', "$cookieFile"
                }
                else {
                    Write-Output "[Setup] $(Get-DateTime) - $cookieFile does not exist. Exiting."
                    Exit-Script
                }
            }
            else {
                $cookieFile = 'None'
                Write-Output "[Setup] $(Get-DateTime) - Login is true and Cookies is false. Continuing."
            }
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - Login is true and Username/Password is Empty. Exiting."
            Exit-Script
        }
    }
    else {
        if ((Test-Path -Path $cookieFile)) {
            Write-Output "[Setup] $(Get-DateTime) - $cookieFile file found. Continuing."
            $dlpParams = $dlpParams + " --cookies $cookieFile"
            $dlpArray += '--cookies', "$cookieFile"
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - $cookieFile does not exist. Exiting."
            Exit-Script
        }
    }
    if ($ffmpeg) {
        Write-Output "[Setup] $(Get-DateTime) - $ffmpeg file found. Continuing."
        $dlpParams = $dlpParams + " --ffmpeg-location $ffmpeg"
        $dlpArray += '--ffmpeg-location', "$ffmpeg"
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - FFMPEG: $ffmpeg missing. Exiting."
        Exit-Script
    }
    if ((Test-Path -Path $batFile)) {
        Write-Output "[Setup] $(Get-DateTime) - $batFile file found. Continuing."
        if (![String]::IsNullOrWhiteSpace((Get-Content -Path $batFile))) {
            Write-Output "[Setup] $(Get-DateTime) - $batFile not empty. Continuing."
            $dlpParams = $dlpParams + " -a $batFile"
            $dlpArray += '-a', "$batFile"
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - $batFile is empty. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - BAT: $batFile missing. Exiting."
        Exit-Script
    }
    if ($archive -or $archiveTemp) {
        if ((Test-Path -Path $archiveFile)) {
            Write-Output "[Setup] $(Get-DateTime) - $archiveFile file found. Continuing."
            $dlpParams = $dlpParams + " --download-archive $archiveFile"
            $dlpArray += '--download-archive', "$archiveFile"
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - Archive file missing. Creating."
            New-Item $archiveFile -ItemType File
        }
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - Using --no-download-archive. Continuing."
        $dlpParams = $dlpParams + ' --no-download-archive'
        $dlpArray += '--no-download-archive'
    }
    $videoType = Select-String -Path $siteConfig -Pattern '--remux-video.*' | Select-Object -First 1
    if ($videoType) {
        $vidType = '*.' + ($videoType -split ' ')[1]
        $vidType = $vidType.Replace("'", '').Replace('"', '')
        if ($vidType -eq '*.mkv') {
            Write-Output "[Setup] $(Get-DateTime) - Using $vidType. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - VidType(mkv) is missing. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - --remux-video parameter is missing. Exiting."
        Exit-Script
    }
    if (Select-String -Path $siteConfig '--write-subs' -SimpleMatch -Quiet) {
        Write-Output "[Setup] $(Get-DateTime) - SubtitleEdit or MKVMerge is true and --write-subs is in config. Continuing."
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - SubtitleEdit is true and --write-subs is not in config. Exiting."
        Exit-Script
    }
    $subtitleType = Select-String -Path $siteConfig -Pattern '--convert-subs.*' | Select-Object -First 1
    if ($subtitleType) {
        $subType = '*.' + ($subtitleType -split ' ')[1]
        $subType = $subType.Replace("'", '').Replace('"', '')
        if ($subType -eq '*.ass') {
            Write-Output "[Setup] $(Get-DateTime) - Using $subType. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-DateTime) - Subtype(ass) is missing. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - --convert-subs parameter is missing. Exiting."
        Exit-Script
    }
    $writeSub = Select-String -Path $siteConfig -Pattern '--write-subs.*' | Select-Object -First 1
    if ($writeSub) {
        Write-Output "[Setup] $(Get-DateTime) - --write-subs is in config. Continuing."
    }
    else {
        Write-Output "[Setup] $(Get-DateTime) - SubSubWrite is missing. Exiting."
        Exit-Script
    }
    $debugVars = [ordered]@{Site = $siteName; IsDaily = $daily; UseLogin = $login; UseCookies = $cookies; UseArchive = $archive; UseArchiveTemp = $archiveTemp; SubtitleEdit = $subtitleEdit; `
            MKVMerge = $mkvMerge; VideoTrackName = $videoTrackName; AudioLang = $audioLang; AudioTrackName = $audioTrackName; SubtitleLang = $subtitleLang; SubtitleTrackName = $subtitleTrackName; Filebot = $filebot; `
            filebotDB = $filebotDB; filebotStructure = $filebotStructure; SiteNameRaw = $siteNameRaw; SiteType = $siteType; SiteUser = $siteUser; SitePass = $sitePass; SiteFolderIdName = $siteFolderIdName[0]; `
            SiteFolder = $siteFolder; SiteParentFolder = $siteParentFolder; SiteSubFolder = $siteSubFolder; SiteLibraryId = $siteLibraryID; SiteTemp = $siteTemp; SiteSrcBase = $siteSrcBase; SiteSrc = $siteSrc; `
            SiteHomeBase = $siteHomeBase; SiteHome = $siteHome; SiteDefaultPath = $siteDefaultPath; SiteConfig = $siteConfig; CookieFile = $cookieFile; Archive = $archiveFile; Bat = $batFile; Ffmpeg = $ffmpeg; `
            SubFontName = $subFontName; SubFontExtension = $subFontExtension; SubFontDir = $subFontDir; SubType = $subType; VidType = $vidType; Backup = $srcBackup; BackupShared = $srcBackupDriveShared; `
            BackupFont = $srcDriveSharedFonts; SiteConfigBackup = $siteConfigBackup; PlexHost = $plexHost; PlexToken = $plexToken; ConfigPath = $configPath; `
            ScriptDirectory = $scriptDirectory; dlpParams = $dlpParams
    }
    if ($testScript) {
        Write-Output "[START] $dateTime - $siteNameRaw - TEST Run"
        $debugVars
        $overrideSeriesList
        $overrideEpisodeList
        Write-Output 'dlpArray:'
        $dlpArray
        if ($sendDiscord) {
            $f = Get-DiskSpace $scriptDirectory
            $discordTestEpisodeSize = (Get-Filesize $configPath)[1]
            $fieldEpisodeDate = ([DateTime]::ParseExact('20230609', 'yyyyMMdd', $null)).ToString('ddd yyyy-MM-dd')
            $fieldFooterText = "$($f.rootName): $($f.freeSpaceFormatted)/$($f.totalSpaceFormatted) - $($f.percentageFree)% Free/$($f.percentageUsed)% Used"
            $discordTestSiteName = 'Test Site Name'; $discordTestSeries = 'Test Series'; $discordTestEpisode = 'Test Episode'; $discordTestIcon = $discordIconDefault; $discordTestColor = $discordColorDefault; `
                $discordTestsubtitle = @('test1.en-US.ass', 'test2.es-es.ass', 'test3.ja.ass'); $discordTestEpisodeSize = $discordTestEpisodeSize ; $discordTestQuality = 'Test Quality'; $discordTestVideoCodec = 'Test Video Codec'; `
                $discordTestAudioCodec = 'Test Audio Codec'; $discordTestDuration = 'Test Duration'; $discordTestRelease = 'Test Release'; $discordTestEpsidodeURL = 'https://www.google.com' ; $discordTestFooterIcon = $discordFooterIconDefault
            Invoke-Discord -site $discordTestSiteName -series $discordTestSeries -episode $discordTestEpisode -siteIcon $discordTestIcon -color $discordTestColor -subtitle $discordTestsubtitle `
                -episodeSize $discordTestEpisodeSize -quality $discordTestQuality -videoCodec $discordTestVideoCodec -audioCodec $discordTestAudioCodec -duration $discordTestDuration `
                -release $discordTestRelease -episodeURL $discordTestEpsidodeURL -episodeDate $fieldEpisodeDate -siteFooterIcon $discordTestFooterIcon -siteFooterText $fieldFooterText -DiscordSiteUrl $discordHookUrl
        }
        Exit-Script
    }
    else {
        if (!$debugScript) {
            $debugVarsRemove = 'SiteUser' , 'SitePass', 'PlexToken', 'dlpParams'
            foreach ($dbv in $debugVarsRemove) {
                $debugVars.Remove($dbv)
            }
        }
        if ($daily) {
            Write-Output "[START] $dateTime - $siteNameRaw - Daily Run"
        }
        elseif ($debugScript) {
            Write-Output "[START] $dateTime - $siteNameRaw - DEBUG Run"
        }
        else {
            Write-Output "[START] $dateTime - $siteNameRaw - Manual Run"
        }
        Write-Output '[START] Debug Vars:'
        $debugVars
        Write-Output '[START] Drive/Series and Episode Overrides:'
        $overrideSeriesList | Sort-Object orSrcdrive, orSeriesName | Format-Table
        $overrideEpisodeList | Sort-Object filenamePattern, fileReplaceText | Format-Table
        # Create folders
        $createFolders = $tempDrive, $srcDrive, $backupDrive, $srcBackup, $siteConfigBackup, $srcDriveSharedFonts, $destDrive, $siteTemp, $siteSrc, $siteHome
        foreach ($cf in $createFolders) {
            Invoke-ExpressionConsole -SCMFN 'START' -SCMFP "New-Folder -newFolderFullPath `"$cf`""
        }
        # Log cleanup
        Invoke-ExpressionConsole -SCMFN 'Cleanup' -SCMFP 'Remove-Logfiles'
    }
    # yt-dlp
    & yt-dlp.exe $dlpArray *>&1 | Out-Host
    # Post-processing
    $totalDownloaded = (Get-ChildItem -Path $siteSrc -Recurse -Force -File -Include "$vidType" | Select-Object * | Measure-Object).Count
    if ($totalDownloaded -gt 0) {
        $OSeries = $overrideEpisodeList | Where-Object { $_.overrideType -eq 'Series' }
        $OSeason = $overrideEpisodeList | Where-Object { $_.overrideType -eq 'Season' }
        $subfolders = Get-ChildItem -Path $siteSrc -Directory
        foreach ($subfolder in $subfolders) {
            $subFolderName = $subfolder.Name
            $subFolderFullname = $subfolder.FullName
            Write-Output "[Processing] $(Get-DateTime) - Processing files in subfolder: $subFolderName"
            $newfoldername = ''
            $files = Get-ChildItem -Path $subFolderFullname -File -Recurse
            foreach ($f in $files) {
                $oldfullPath = $f.FullName
                Write-Output "[Processing] $(Get-DateTime) - Checking Filename: `"$oldfullPath`""
                if ($f.Extension -eq '.mkv') {
                    $formattedFilename = Format-Filename -InputStr $f.BaseName
                }
                else {
                    $subBaseName = $f.BaseName
                    $lastPeriodIndex = $subBaseName.LastIndexOf('.')
                    $formattedFilename = (Format-Filename -InputStr ($subBaseName.Substring(0, $lastPeriodIndex))) + '.' + $subBaseName.Substring($lastPeriodIndex + 1)
                }
                Write-Output "[Processing] $(Get-DateTime) - Formatted name: $formattedFilename"
                if ($OSeries.count -gt 0) {
                    foreach ($e in $OSeries) {
                        if ($formattedFilename -match $($e.filenamePattern)) {
                            $pattern = $e.filenamePattern
                            $replacementText = $e.fileReplaceText
                            $renamedFilename = ($formattedFilename -replace $pattern, "$replacementText`$2$replacementText`$4")
                            $newfoldername = (Get-Culture).TextInfo.ToTitleCase($replacementText)
                            Write-Output "[Processing] $(Get-DateTime) - Override Series Pattern: `"$($pattern)`""
                            Write-Output "[Processing] $(Get-DateTime) - Override Series CHANGED: `"$renamedFilename`""
                            break
                        }
                        else {
                            $renamedFilename = $formattedFilename
                        }
                    }
                }
                else {
                    $renamedFilename = $formattedFilename
                }
                if ($OSeason.count -gt 0) {
                    foreach ($o in $OSeason) {
                        if ($renamedFilename -match $($o.filenamePattern)) {
                            $pattern = $o.filenamePattern
                            $replacementText = $o.fileReplaceText
                            $newBaseName = ($renamedFilename -replace $pattern, "`$1$replacementText`$3") + $f.Extension
                            Write-Output "[Processing] $(Get-DateTime) - Override Season Pattern: `"$($pattern)`""
                            Write-Output "[Processing] $(Get-DateTime) - Override Season CHANGED: `"$newBaseName`""
                            break
                        }
                        else {
                            $newBaseName = $renamedFilename + $f.Extension
                        }
                    }
                }
                else {
                    $newBaseName = (Get-Culture).TextInfo.ToTitleCase($renamedFilename) + $f.Extension
                }
                Write-Output "[Processing] $(Get-DateTime) - Writing $oldfullPath to $newBaseName"
                Rename-Item $oldfullPath -NewName $newBaseName -Verbose
            }
            if ($newfoldername) {
                Rename-Item $subFolderFullname -NewName $newfoldername
            }
            else {
                $newfoldername = Format-Filename -InputStr $subFolderName
                Rename-Item $subFolderFullname -NewName $newfoldername
            }
            Write-Output "[Processing] $(Get-DateTime) - End of Subfolder: $newfoldername"
        }
        
        Get-ChildItem -Path $siteSrc -Recurse -Include "$vidType" | Sort-Object LastWriteTime | Select-Object -Unique | Get-Unique | ForEach-Object {
            [System.Collections.ArrayList]$vsEpisodeSubtitle = @{}
            $vsSite = $siteNameRaw
            $vsEpisodeRaw = $_.BaseName
            $vsSeries = (Split-Path (Split-Path $_ -Parent) -Leaf)
            $vsEpisode = $vsEpisodeRaw
            Write-Output "[Processing] $(Get-DateTime) - Video Found: $vsEpisode"
            # Origin directory
            $vsSeriesDirectoryO = $_.DirectoryName
            # New directory
            $vsSeriesDirectory = Join-Path $siteSrc -ChildPath $vsSeries
            $vsEpisodeTemp = Join-Path -Path $vsSeriesDirectory -ChildPath "$($vsEpisode).temp$($_.Extension)"
            $vsEpisodeSize = (Get-Filesize $_.FullName)[1]
            Invoke-ExpressionConsole -SCMFN 'Processing' -SCMFP "New-Item -Path `"$vsSeriesDirectory`" -ItemType Directory"
            # Checking if override exists for Series name in config
            $vsOverridePath = $overrideSeriesList | Where-Object { $_.orSeriesName.ToLower() -eq $vsSeries.ToLower() } | Select-Object -ExpandProperty orSrcdrive
            # Origin directory
            $vsEpisodePathO = $_.FullName
            # New directory
            $vsEpisodePath = Join-Path -Path $vsSeriesDirectory -ChildPath "$($vsEpisode)$($_.Extension)"
            # Run FFprobe to get video information resolution, video and audio codecs against origin filepath
            $ffprobeResVideoCodec = & ffprobe.exe -v error -select_streams v:0 -show_entries stream='height,codec_name' -of csv=p=0 "$vsEpisodePathO"
            $codec, $height = $ffprobeResVideoCodec.Split(',')
            $vsEpisodeRes = "${height}p"
            $vsEpisodeVideoCodec = $codec.ToUpper()
            $ffprobeAudioCodec = & ffprobe.exe -v error -select_streams a:0 -show_entries stream=codec_name -of csv=p=0 "$vsEpisodePathO"
            $vsEpisodeAudioCodec = $($ffprobeAudioCodec.Trim()).ToUpper()
            # Run FFprobe to get video duration against origin filepath
            $ffprobeRuntime = & ffprobe.exe -v error -show_entries format=duration -of default=noprint_wrappers=1:nokey=1 "$vsEpisodePathO"
            $duration = [TimeSpan]::FromSeconds([double]::Parse($ffprobeRuntime ))
            $vsEpisodeDuration = "$([math]::Floor($duration.TotalHours)):$($duration.Minutes.ToString('00')):$($duration.Seconds.ToString('00'))"
            $vsEpisodeDateO = ffprobe.exe -v quiet -show_entries format_tags='DATE' -of default=nw=1:nk=1 "$vsEpisodePathO"
            $vsEpisodeUrl = ffprobe.exe -v quiet -show_entries format_tags='COMMENT' -of default=nw=1:nk=1 "$vsEpisodePathO"
            Write-Output "[Processing] $(Get-DateTime) - Episode Release Date(Orig.): $($vsEpisodeDateO) - $($vsEpisode)"
            if ($null -eq $vsEpisodeDateO -or $vsEpisodeDateO.trim() -eq '') {
                $vsEpisodeDateO = Get-Date -Format 'yyyyMMdd'
            }
            $vsEpisodeDateO = [DateTime]::ParseExact($vsEpisodeDateO, 'yyyyMMdd', $null)
            $vsEpisodeDate = $vsEpisodeDateO.ToString('dddd yyyy-MM-dd')
            Write-Output "[Processing] $(Get-DateTime) - Episode Release Date(Final): $($vsEpisodeDate) - $($vsEpisode)"
            if ($vsOverridePath) {
                $vsDestPath = Join-Path -Path $vsOverridePath -ChildPath $siteHome.Substring(3)
                $vsDestPathBase = Join-Path -Path $vsOverridePath -ChildPath $siteHomeBase.Substring(3)
                $vsDestPathVideo = $vsEpisodePath.Replace($siteSrc, $vsDestPath)
                $vsDestPathDirectory = Split-Path $vsDestPathVideo -Parent
            }
            else {
                $vsDestPath = $siteHome
                $vsDestPathBase = $siteHomeBase
                $vsDestPathVideo = $vsEpisodePath.Replace($siteSrc, $vsDestPath)
                $vsDestPathDirectory = Split-Path $vsDestPathVideo -Parent
                $vsOverridePath = [System.IO.path]::GetPathRoot($vsDestPath)
            }
            Get-ChildItem -Path $siteSrc -Recurse -Include $subType | Where-Object { $_.FullName -match $vsEpisodeRaw } | Select-Object Name, FullName, BaseName, Extension -Unique | ForEach-Object {
                [System.Collections.ArrayList]$episodeSubtitles = @{}
                # original directory
                $origSubPathO = $_.FullName
                $subtitleRawO = $_.BaseName
                $subtitleBasenameExt = $subtitleRawO + $_.Extension
                $origSubPath = Join-Path $vsSeriesDirectory -ChildPath $subtitleBasenameExt
                $subtitleBase = $subtitleBasenameExt
                Write-Output "[Processing] $(Get-DateTime) - Subtitle Found: $subtitleBase"
                if ($vsOverridePath) {
                    $overrideSubPath = $origSubPath.Replace($siteSrc, $vsDestPath)
                }
                else {
                    $overrideSubPath = $origSubPath
                }
                $episodeSubtitles = @{ subtitleBase = $subtitleBase; origSubPath = $origSubPath; overrideSubPath = $overrideSubPath }
                [void]$vsEpisodeSubtitle.Add($episodeSubtitles)
                # Move subs from original directory to new directory
                Invoke-ExpressionConsole -SCMFN 'Processing' -SCMFP "Move-Item -Path `"$origSubPathO`" -Destination `"$origSubPath`" -Verbose"
            }
            $vsFinalPathVideo = ''
            $vsFilebotTitle = ''
            # Move video file after all subs are moved from original directory to new directory
            Invoke-ExpressionConsole -SCMFN 'Processing' -SCMFP "Move-Item -Path `"$vsEpisodePathO`" -Destination `"$vsEpisodePath`" -Verbose"
            foreach ($i in $_) {
                $VideoStatus = [VideoStatus]::new($vsSite, $vsSeries, $vsEpisode, $vsEpisodeDuration, $vsEpisodeSize, $vsEpisodeRes, $vsEpisodeVideoCodec, $vsEpisodeAudioCodec, `
                        $vsSeriesDirectory, $vsEpisodeRaw, $vsEpisodeTemp, $vsEpisodePath, $vsOverridePath, $vsDestPathBase, $vsDestPath, $vsDestPathDirectory, `
                        $vsDestPathVideo, $vsFinalPathVideo, $vsFilebotTitle, $vsEpisodeUrl, $vsEpisodeDate, $vsEpisodeSubtitle, $vsSECompleted, $vsMKVCompleted, $vsMoveCompleted, $vsFBCompleted, $vsErrored)
                [void]$vsCompletedFilesList.Add($VideoStatus)
            }
            $lf = (Get-ChildItem -Path $vsSeriesDirectoryO -File).Count
            if ($lf -eq 0 ) {
                Invoke-ExpressionConsole -SCMFN 'Processing' -SCMFP "Remove-Item -Path `"$vsSeriesDirectoryO`""
            }
        }
        $vsCompletedFilesList | ForEach-Object {
            if ($_._vsEpisodeSubtitle.count -eq 0) {
                $vsEpisodeRaw = $_._vsEpisodeRaw
                Set-VideoStatus -searchKey '_vsEpisodeRaw' -searchValue $vsEpisodeRaw -nodeKey '_vsErrored' -nodeValue $($true)
            }
        }
    }
    else {
        Write-Output "[Processing] $(Get-DateTime) - No files to process."
    }
    $vsvTotCount = ($vsCompletedFilesList | Measure-Object).Count
    $vsvErrorCount = ($vsCompletedFilesList | Where-Object { $_._vsErrored -eq $true } | Measure-Object).Count
    # SubtitleEdit, MKVMerge, Filebot
    if ($vsvTotCount -gt 0) {
        # Subtitle Edit
        if ($subtitleEdit) {
            $vsCompletedFilesList | Where-Object { $_._vsErrored -ne $true } | Select-Object -ExpandProperty _vsEpisodeSubtitle | ForEach-Object {
                $seSubtitle = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
                Write-Output "[SubtitleEdit] $(Get-DateTime) - Fixing $seSubtitle subtitle."
                While ($True) {
                    if ((Test-Lock $seSubtitle) -eq $True) {
                        continue
                    }
                    else {
                        Invoke-ExpressionConsole -SCMFN 'SubtitleEdit' -SCMFP "powershell `"SubtitleEdit /convert `'$seSubtitle`' AdvancedSubStationAlpha /overwrite /MergeSameTimeCodes /FixCommonErrors /FixCommonErrors`""
                        Set-VideoStatus -searchKey '_vsEpisodeSubtitle' -searchValue $seSubtitle -nodeKey '_vsSECompleted' -nodeValue $($true)
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            Write-Output "[SubtitleEdit] $(Get-DateTime) - Not running."
        }
        if ($subFontName) {
            $vsCompletedFilesList | Select-Object -ExpandProperty _vsEpisodeSubtitle | ForEach-Object {
                $sp = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
                While ($True) {
                    if ((Test-Lock $sp) -eq $True) {
                        continue
                    }
                    else {
                        Write-Output "[SubtitleRegex] $(Get-DateTime) - Regex through $sp file with $subFontName."
                        Update-SubtitleStyle -SubtitleFilePath $sp -subFontName $subFontName
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            Write-Output "[SubtitleRegex] $(Get-DateTime) - No Font specified for subtitle files."
        }
        # MKVMerge
        if ($mkvMerge) {
            $vsCompletedFilesList | Select-Object _vsEpisodeRaw, _vsEpisode, _vsEpisodeTemp, _vsEpisodePath, _vsEpisodeSubtitle, _vsErrored | `
                Where-Object { $_._vsErrored -eq $false } | ForEach-Object {
                $mkvVidInput = $_._vsEpisodePath
                $mkvVidBaseName = $_._vsEpisodeRaw
                $mkvVidSubtitle = $_._vsEpisodeSubtitle
                $mkvVidTempOutput = $_._vsEpisodeTemp
                Invoke-MKVMerge $mkvVidInput $mkvVidBaseName $mkvVidSubtitle $mkvVidTempOutput
            }
            $overrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsMKVCompleted -eq $true -and $_._vsErrored -eq $false } | Select-Object _vsSeriesDirectory, _vsDestPath, _vsDestPathBase -Unique
        }
        else {
            Write-Output "[MKVMerge] $(Get-DateTime) - MKVMerge not running. Moving to next step."
            $overrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsErrored -eq $false } | Select-Object _vsSeriesDirectory, _vsDestPath, _vsDestPathBase -Unique
        }
        $vsvMKVCount = ($vsCompletedFilesList | Where-Object { $_._vsMKVCompleted -eq $true -and $_._vsErrored -eq $false } | Measure-Object).Count
        $vsMoveCount = ($vsCompletedFilesList | Where-Object { $_._vsMoveCompleted -eq $true -and $_._vsErrored -eq $false } | Measure-Object).Count
        # FileMoving
        if (!($mkvMerge) -and $vsvErrorCount -eq 0) {
            Write-Output "[FileMoving] $(Get-DateTime) - All files had matching subtitle file"
            foreach ($orDriveList in $overrideDriveList) {
                Write-Output "[FileMoving] $(Get-DateTime) - $($orDriveList._vsSeriesDirectory) contains files. Moving to $($orDriveList._vsDestPath)."
                if (!(Test-Path -Path $orDriveList._vsDestPath)) {
                    Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "New-Folder -newFolderFullPath `"$($orDriveList._vsDestPath)`" -Verbose"
                }
                Write-Output "[FileMoving] $(Get-DateTime) - Moving $($orDriveList._vsSeriesDirectory) to $($orDriveList._vsDestPath)."
                Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($orDriveList._vsSeriesDirectory)`" -Destination `"$($orDriveList._vsDestPath)`" -Force -Verbose"
                if (!(Test-Path -Path $orDriveList._vsSeriesDirectory)) {
                    Write-Output "[FileMoving] $(Get-DateTime) - Move completed for $($orDriveList._vsSeriesDirectory)."
                    Set-VideoStatus -searchKey '_vsSeriesDirectory' -searchValue $orDriveList._vsSeriesDirectory -nodeKey '_vsMoveCompleted' -nodeValue $($true)
                }
            }
        }
        elseif (($vsvMKVCount -eq $vsvTotCount) -and ($vsvMKVCount -eq $vsMoveCount) -and $vsvErrorCount -eq 0) {
            Write-Output "[FileMoving] $(Get-DateTime) - All files completed with MKVMerge."
        }
        else {
            Write-Output "[FileMoving] $(Get-DateTime) - $siteSrc contains file(s) with error(s). Not moving files."
        }
        # Filebot
        $filebotOverrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsErrored -eq $false } | Select-Object _vsDestPath -Unique
        if (($filebot -and $vsvMKVCount -eq $vsvTotCount) -or ($filebot -and !($mkvMerge))) {
            foreach ($fbORDriveList in $filebotOverrideDriveList) {
                Write-Output "[Filebot] $(Get-DateTime) - Renaming files in $($fbORDriveList._vsDestPath)."
                Invoke-Filebot -filebotPath $fbORDriveList._vsDestPath -filebotContentType $SiteContentType
            }
        }
        elseif (!($filebot) -and !($mkvMerge) -or (!($filebot) -and $vsvMKVCount -eq $vsvTotCount)) {
            $moveManualList = $vsCompletedFilesList | Where-Object { $_._vsMoveCompleted -eq $true } | Select-Object _vsDestPathDirectory, _vsOverridePath -Unique
            foreach ($mmFiles in $moveManualList) {
                $mmOverrideDrive = $mmFiles._vsOverridePath
                $moveRootDirectory = $mmOverrideDrive + $siteParentFolder
                $moveFolder = Join-Path -Path $moveRootDirectory -ChildPath $siteSubFolder
                Write-Output "[FileMoving] $(Get-DateTime) - Moving $($mmFiles._vsDestPathDirectory) to $moveFolder."
                Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($mmFiles._vsDestPathDirectory)`" -Destination `"$moveFolder`" -Force -Verbose"
            }
        }
        else {
            Write-Output "[FileMoving] $(Get-DateTime) - Issue with files in $siteSrc."
        }
        # Sonarr
        $useSonarr = $false
        if ($sonarrHost -and $sonarrToken -and ($useSonarr -eq $True)) {
            # Get a list of all series
            $series = Invoke-SonarrApi -url 'series'
            $vsCompletedFilesList | Where-Object { $_._vsFBCompleted -eq $true } | ForEach-Object {
                $v = $_
                $videoPath = ($v._vsFinalPathVideo).ToString()
                $s = $null
                Write-Output "[Sonarr] $(Get-DateTime) - Checking: $($videoPath)"
                while (-not (Test-Path -LiteralPath $videoPath)) {
                    Write-Output "[Sonarr] $(Get-DateTime) - Waiting: $($videoPath)"
                    Start-Sleep -Milliseconds 500
                }
                While ($True) {
                    if ((Test-Lock $videoPath -literal) -eq $True) {
                        continue
                    }
                    else {
                        Write-Output "[Sonarr] $(Get-DateTime) - Processing File: $($videoPath)"
                        $parts = $videoPath -split '\\'
                        $sonarrSeriesFolderPath = ($parts[0..3]) -join '\'
                        #$seriesFolderPath = ($sonarrSeriesFolderPath) + '\'
                        #$relativePath = ($videoPath.Substring($seriesFolderPath.Length)).Trim('\')
                        if ($series) {
                            $s = $series | Where-Object { $_.path -eq $sonarrSeriesFolderPath }
                            $seriesId = ($s.id).ToString()
                            $seriesTitle = $s.title
                            $seriesPath = $s.path
                            Write-Output "[Sonarr] $(Get-DateTime) - Series Path Found: $($seriesTitle) - $seriesPath"
                            if ($s) {
                                # Trigger a rescan and rename for each series
                                $i = 0
                                do {
                                    $rescanCommand = @{
                                        name      = 'RescanSeries'
                                        seriesId  = $seriesId
                                        immediate = $true
                                    }
                                    $refreshCommand = @{
                                        name      = 'RefreshSeries'
                                        seriesId  = $seriesId
                                        immediate = $true
                                    }
                                    $rfc = Invoke-SonarrApi -url 'command' -method 'POST' -body ($refreshCommand | ConvertTo-Json)
                                    foreach ($fc in $rfc) {
                                        Write-Output "[Sonarr] $(Get-DateTime) - $($fc.commandName): $($fc.body.completionMessage) - $($seriesTitle) - $($seriesPath)"
                                    }
                                    $rsc = Invoke-SonarrApi -url 'command' -method 'POST' -body ($rescanCommand | ConvertTo-Json)
                                    foreach ($sc in $rsc) {
                                        Write-Output "[Sonarr] $(Get-DateTime) - $($sc.commandName): $($sc.body.completionMessage) - $($seriesTitle) - $($seriesPath)"
                                    }
                                    Start-Sleep 1
                                    $i++
                                } until (
                                    $i -eq 2
                                )
                                # Fetch episodes after rescan
                                $sonarrRenameEpisodes = Invoke-SonarrApi -url "rename?seriesId=$($seriesId)"
                                $sreCount = $sonarrRenameEpisodes.count
                                Write-Output "[Sonarr] $(Get-DateTime) - Total episodes to rename: $sreCount"
                                # Rename all files found
                                foreach ($e in $sonarrRenameEpisodes) {
                                    $fileId = ($e.episodeFileId).ToString()
                                    $existingPath = $e.existingPath
                                    $newPath = $e.newPath
                                    Write-Output "[Sonarr] $(Get-DateTime) - Series Episode Found: $($seriesTitle) - $fileId"
                                    Write-Output "[Sonarr] $(Get-DateTime) - Rename Episode From: $existingPath To: $newPath"
                                    $renameFile = @{
                                        name      = 'RenameFiles'
                                        seriesId  = $seriesId
                                        files     = @($fileId)
                                        immediate = $true
                                    }
                                    Invoke-SonarrApi -url 'command' -method 'POST' -Body ($renameFile | ConvertTo-Json)
                                }
                            }
                            else {
                                Write-Output "[Sonarr] - No series found in Sonarr for $sonarrSeriesFolderPath"
                            }
                        }
                        else {
                            Write-Output '[Sonarr] - No series found in Sonarr'
                        }
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            Write-Output "[Sonarr] $(Get-DateTime) - Not using Sonarr."
        }

        # Plex
        if ($plexHost -and $plexToken -and $($usePlex -eq $True)) {
            Write-Output "[PLEX] $(Get-DateTime) - Fetching library ID for `"$siteParentFolder\$siteSubFolder`"."
            [xml]$plexLibCheck = Invoke-WebRequest "$plexHost/library/sections/all?X-Plex-Token=$plexToken"
            $plexlibID = $plexLibCheck.mediacontainer.Directory | Where-Object { $_.location.path -match ".*$siteParentFolder\\$siteSubFolder$" } | Select-Object key, title -Unique
            Write-Output "[PLEX] $(Get-DateTime) - Updating Plex Library - $($plexlibID.title) - $($plexlibID.key)."
            $plexUrl = "$plexHost/library/sections/$($plexlibID.key)/refresh?X-Plex-Token=$plexToken"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $plexUrl | Out-Null
            $ProgressPreference = 'Continue'
        }
        else {
            Write-Output "[PLEX] $(Get-DateTime) - Not using Plex."
        }
        $vsvFBCount = ($vsCompletedFilesList | Where-Object { $_._vsFBCompleted -eq $true } | Measure-Object).Count
        # Discord
        $occurrenceTotal = $vsvTotCount
        if ($sendDiscord) {
            Write-Output "[Discord] $(Get-DateTime) - Preparing Discord message."
            $vsCompletedFilesList | ForEach-Object {
                $occurrenceCount++
                $fieldSite = $_._vsSite
                $fieldSeries = $_._vsSeries
                $fieldEpisode = $_._vsEpisode
                $fieldSubtitle = $_._vsEpisodeSubtitle.value
                $fieldSize = $_._vsEpisodeSize
                $fieldQuality = $_._vsEpisodeRes
                $fieldVideoCodec = $_._vsEpisodeVideoCodec
                $fieldAudioCodec = $_._vsEpisodeAudioCodec
                $fieldDuration = $_._vsEpisodeDuration
                $fieldEpisodeUrl = $_._vsEpisodeUrl
                $fieldEpisodeDate = $_._vsEpisodeDate
                if ($_._vsErrored -eq $true) {
                    $fieldRelease = "Error: $($_._vsEpisode)"
                    $discordUrl = $discordHookErrorURL
                }
                else {
                    if ($filebot) {
                        $fieldRelease = $_._vsFilebotTitle
                    }
                    else {
                        $fieldRelease = $($_._vsEpisode)
                    }
                    $discordUrl = $discordHookURL
                }
                $filePath = if ($_._vsFinalPathVideo.trim() -ne '') {
                    $_._vsFinalPathVideo
                }
                Else {
                    $_._vsDestPathVideo
                }
                $f = Get-DiskSpace $filePath
                $fieldFooterText = "$($f.RootName): $($f.FreeSpaceFormatted)/$($f.TotalSpaceFormatted) - $($f.PercentageFree)% Free/$($f.PercentageUsed)% Used"
                Write-Output "[Discord] $(Get-DateTime) - Sending $occurrenceCount/$occurrenceTotal."
                Invoke-Discord -site $fieldSite -series $fieldSeries -episode $fieldEpisode -siteIcon $discordSiteIcon -color $discordSiteColor -subtitle $fieldSubtitle `
                    -episodeSize $fieldSize -quality $fieldQuality -videoCodec $fieldVideoCodec -audioCodec $fieldAudioCodec -duration $fieldDuration -release $fieldRelease `
                    -episodeURL $fieldEpisodeUrl -episodeDate $fieldEpisodeDate -siteFooterIcon $discordFooterIconDefault -siteFooterText $fieldFooterText -DiscordSiteUrl $discordUrl
            }
        }
        Write-Output "[Processing] $(Get-DateTime) - Errored Files: $vsvErrorCount/$vsvTotCount"
        $vsCompletedFilesList | Select-Object * | ForEach-Object {
            $e = $_._vsEpisodeRaw
            $_ | Select-Object -ExpandProperty _vsEpisodeSubtitle | ForEach-Object {
                $r = $_ | Where-Object { $_.name -notin 'origSubPath', 'overrideSubPath' }
                Set-VideoStatus -searchKey '_vsEpisodeRaw' -searchValue $e -nodeKey '_vsEpisodeSubtitle' -nodeValue $r
            }
        }
        # File status
        $vsFileStatus = @{Label = 'Episode'; Expression = { "$($_._vsSeries): $($_._vsEpisode)" } }, @{Label = 'SECompleted'; Expression = { $_._vsSECompleted } }, `
        @{Label = 'MKVCompleted'; Expression = { $_._vsMKVCompleted } }, @{Label = 'MoveCompleted'; Expression = { $_._vsMoveCompleted } }, `
        @{Label = 'FBCompleted'; Expression = { $_._vsFBCompleted } }, @{Label = 'Errored'; Expression = { $_._vsErrored } }
        # File paths
        $vsFilesDetail = @{Label = 'Episode'; Expression = { "$($_._vsSeries): $($_._vsEpisode)" } }, @{Label = 'Final Path'; Expression = { if ($_._vsFinalPathVideo.trim() -ne '') { $_._vsFinalPathVideo } Else { $_._vsDestPathVideo } } }
        # File Subtitles
        $vsFileSubtitles = @{Label = 'Episode'; Expression = { "$($_._vsSeries): $($_._vsEpisode)" } }, @{Label = 'Subtitle'; Expression = { ($_._vsEpisodeSubtitle.value | Join-String -Separator "`r`n") } }
        # File Info
        $vsFileInfo = @{Label = 'Episode'; Expression = { "$($_._vsSeries): $($_._vsEpisode)" } }, @{Label = 'Size'; Expression = { $_._vsEpisodeSize } }, `
        @{Label = 'Resolution'; Expression = { $_._vsEpisodeRes } }, @{Label = 'Video Codec'; Expression = { $_._vsEpisodeVideoCodec } }, @{Label = 'Audio Codec'; Expression = { $_._vsEpisodeAudioCodec } }, `
        @{Label = 'Duration'; Expression = { $_._vsEpisodeDuration } }, @{Label = 'Title'; Expression = { if ($_._vsFilebotTitle.trim() -ne '') { $_._vsFilebotTitle } else { "Error: $($_._vsEpisode)" } } }
        # Store table data for output later
        $vsCompletedFilesTable = $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsFileStatus -AutoSize -Wrap
        $vsCompletedFilesTable += $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsFileInfo -AutoSize -Wrap
        $vsCompletedFilesTable += $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsFilesDetail -AutoSize -Wrap
        $vsCompletedFilesTable += $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsFileSubtitles -AutoSize -Wrap
        # Output additional table data if debugScript = true
        if ($debugScript) {
            # File Debug Info
            $vsDebugInfoV1 = @{Label = 'Episode'; Expression = { "$($_._vsSeries): $($_._vsEpisode)" } }, @{Label = 'TempPath'; Expression = { $_._vsEpisodeTemp } }, @{Label = 'DestinationPath'; Expression = { $_._vsDestPathVideo } }
            $vsDebugInfoV2 = @{Label = 'Episode'; Expression = { "$($_._vsSeries): $($_._vsEpisode)" } }, @{Label = 'Final Date'; Expression = { $_._vsEpisodeDate } }, @{Label = 'Episode URL'; Expression = { $_._vsEpisodeUrl } }
            $vsCompletedFilesTable += $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsDebugInfoV1 -AutoSize -Wrap
            $vsCompletedFilesTable += $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsDebugInfoV2 -AutoSize -Wrap
        }
        # Backup if downloaded a video
        $sharedBackups = $archiveFile, $cookieFile, $batFile, $configPath, $subFontDir, $siteConfig
        foreach ($sb in $sharedBackups) {
            if (($sb -ne 'None') -and ($sb.trim() -ne '')) {
                if ($sb -eq $subFontDir) {
                    Write-Output "[FileBackup] $(Get-DateTime) - Copying $sb to $srcDriveSharedFonts."
                    Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$srcDriveSharedFonts`" -PassThru -Verbose"
                }
                elseif ($sb -eq $configPath) {
                    Write-Output "[FileBackup] $(Get-DateTime) - Copying $sb to $srcBackup."
                    Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$srcBackup`" -PassThru -Verbose"
                }
                elseif ($sb -in $siteConfig, $archiveFile, $cookieFile, $batFile) {
                    Write-Output "[FileBackup] $(Get-DateTime) - Copying $sb to $siteConfigBackup."
                    Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$siteConfigBackup`" -PassThru -Verbose"
                }
            }
        }
    }
    Exit-Script
}