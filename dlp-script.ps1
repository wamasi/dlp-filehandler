<#
.Synopsis
   Script to run yt-dlp, mkvmerge, subtitle edit, filebot and a python script for downloading and processing videos
.EXAMPLE
   Runs the script using the mySiteHere as a manual run with the defined config using login, cookies, archive file, mkvmerge, and sends out a telegram message
   D:\_DL\dlp-script.ps1 -sn mySiteHere -l -c -mk -a -st
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
                    $validSites = (([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName) -join ', '
                    throw "The following Sites are valid: $validSites"
                }
            }
            else {
                throw "No valid config.xml found in $PSScriptRoot. Run ($PSScriptRoot\dlp-script.ps1 -nc) for a new config file."
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $True)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $True)]
    [string]$site,
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
    [Alias('ST')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$sendTelegram,
    [Alias('AL')]
    [ValidateScript({
            $langValues = 'ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und'
            if ($_ -in $langValues) {
                $true
            }
            else {
                throw "Value '{0}' is invalid. The following languages are valid: {1}." -f $_, $($langValues -join ', ')
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$audioLang,
    [Alias('SL')]
    [ValidateScript({
            $langValues = 'ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und'
            if ($_ -in $langValues ) {
                $true
            }
            else {
                
                throw "Value '{0}' is invalid. The following languages are valid: {1}." -f $_, $($langValues -join ', ')
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$subtitleLang,
    [Alias('T')]
    [Parameter(ParameterSetName = 'Test', Mandatory = $true)]
    [switch]$testScript
)
# Timer for script
$scriptStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
# Setting styling to remove error characters and width
$psStyle.OutputRendering = 'Host'
$width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($width, 9999)
mode con cols=9999

function Get-Day {
    return (Get-Date -Format 'yy-MM-dd')
}

function Get-TimeStamp {
    return (Get-Date -Format 'yy-MM-dd HH-mm-ss')
}

function Get-Time {
    return (Get-Date -Format 'MMddHHmmssfff')
}

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
        Write-Output "[$scmFunctionName] $(Get-TimeStamp) - $_"
    }
}

# Test if file is available to interact with
function Test-Lock {
    Param(
        [parameter(Mandatory = $true)]
        $testLockFilename
    )
    $testLockFile = Get-Item -Path (Resolve-Path $testLockFilename) -Force
    if ($testLockFile -is [IO.FileInfo]) {
        trap {
            Write-Output "[FileLockCheck] $(Get-Timestamp) - $testLockFile File locked. Waiting."
            return $true
            continue
        }
        $testLockStream = New-Object system.IO.StreamReader $testLockFile
        if ($testLockStream) { $testLockStream.Close() }
    }
    Write-Output "[FileLockCheck] $(Get-Timestamp) - $testLockFile File unlocked. Continuing."
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

function New-SuppFiles {
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
        if (($filledLogFiles | Measure-Object).Count -gt 0) {
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
        if (($emptyLogFiles | Measure-Object).Count -gt 0) {
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

function Exit-Script {
    $scriptStopWatch.Stop()
    # Cleanup folders
    Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Folders -removeFolder `"$($siteTemp)`" -removeFolderBaseMatch `"$($siteTempBaseMatch)`""
    Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Folders -removeFolder `"$($siteSrc)`" -removeFolderBaseMatch `"$($siteSrcBaseMatch)`""
    Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Folders -removeFolder `"$($siteHome)`" -removeFolderBaseMatch `"$($siteHomeBaseMatch)`""
    if ($overrideDriveList.count -gt 0) {
        foreach ($orDriveList in $overrideDriveList) {
            $orDriveListBaseMatch = ($orDriveList._vsDestPathBase).Replace('\', '\\')
            Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Folders -removeFolder `"$($orDriveList._vsDestPath)`" -removeFolderBaseMatch `"$($orDriveListBaseMatch)`""
        }
    }
    # Cleanup Log Files
    Invoke-ExpressionConsole -SCMFN 'LogCleanup' -SCMFP 'Remove-Logfiles'
    Write-Output "[END] $(Get-Timestamp) - Script completed. Total Elapsed Time: $($scriptStopWatch.Elapsed.ToString())"
    Stop-Transcript
    ((Get-Content -Path $logFile | Select-Object -Skip 5) | Select-Object -SkipLast 4) | Set-Content -Path $logFile
    Remove-Spaces -removeSpacesFile $logFile
    $logTemp = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime-Temp.log"
    New-Item -Path $logTemp -ItemType File | Out-Null
    $asciiLogo | Out-File -FilePath $logTemp -Width 9999
    if ($vsvTotCount -gt 0) {
        $vsCompletedFilesTable | Out-File -FilePath $logTemp -Width 9999 -Append
    }
    Get-Content $logFile -ReadCount 5000 | ForEach-Object {
        $_ | Add-Content -Path $logTemp
    }
    Remove-Item -Path $logFile
    if ($vsvTotCount -gt 0) {
        Rename-Item -Path $logTemp -NewName "$dateTime-T$vsvTotCount-E$vsvErrorCount.log"
    }
    elseif ($testScript) {
        Rename-Item -Path $logTemp -NewName "$dateTime-DEBUG.log"
    }
    else {
        Rename-Item -Path $logTemp -NewName $logFile
    }
    exit
}

# Update SE/MKV/FB true false
function Set-VideoStatus {
    param (
        [parameter(Mandatory = $true)]
        [string]$svsKey,
        [parameter(Mandatory = $true)]
        [string]$svsValue,
        [parameter(Mandatory = $false)]
        [switch]$svsSE,
        [parameter(Mandatory = $false)]
        [switch]$svsMKV,
        [parameter(Mandatory = $false)]
        [switch]$svsMove,
        [parameter(Mandatory = $false)]
        [switch]$svsFP,
        [parameter(Mandatory = $false)]
        [switch]$svsER
    )
    $vsCompletedFilesList | Where-Object { $_.$svsKey -eq $svsValue } | ForEach-Object {
        if ($svsSE) {
            $_._vsSECompleted = $svsSE
        }
        if ($svsMKV) {
            $_._vsMKVCompleted = $svsMKV
        }
        if ($svsMove) {
            $_._vsMoveCompleted = $svsMove
        }
        if ($svsFP) {
            $_._vsFBCompleted = $svsFP
        }
        if ($svsER) {
            $_._vsErrored = $svsER
        }
    }
}

# Looks at language code in filename to match a pair of vars
function Get-SubtitleLanguage {
    param ($subFiles)
    switch -Regex ($subFiles) {
        '\.ar\.' { $stLang = 'ar'; $stTrackName = 'Arabic Sub' }
        '\.de\.' { $stLang = 'de'; $stTrackName = 'Deutsch Sub' }
        '\.en.*\.' { $stLang = 'en'; $stTrackName = 'English Sub' }
        '\.es\.' { $stLang = 'es'; $stTrackName = 'Spanish(Latin America) Sub' }
        '\.es-es\.' { $stLang = 'es-es'; $stTrackName = 'Spanish(Spain) Sub' }
        '\.fr\.' { $stLang = 'fr'; $stTrackName = 'French Sub' }
        '\.it\.' { $stLang = 'it'; $stTrackName = 'Italian Sub' }
        '\.ja\.' { $stLang = 'ja'; $stTrackName = 'Japanese Sub' }
        '\.pt-br\.' { $stLang = 'pt-br'; $stTrackName = 'Português (Brasil) Sub' }
        '\.pt-pt\.' { $stLang = 'pt-pt'; $stTrackName = 'Português (Portugal) Sub' }
        '\.ru\.' { $stLang = 'ru'; $stTrackName = 'Russian Video' }
        default { $stLang = 'und'; $stTrackName = 'und sub' }
    }
    $return = @($stLang, $stTrackName)
    return $return
}

# Getting list of Site, Series, and Episodes for Telegram messages
function Get-SiteSeriesEpisode {
    $seriesEpisodeList = $vsCompletedFilesList | Group-Object -Property _vsSite, _vsSeries |
    Select-Object @{n = 'Site'; e = { $_.Values[0] } }, `
    @{ n = 'Series'; e = { $_.Values[1] } }, `
    @{n = 'Episode'; e = { $_.Group | Select-Object _vsEpisode } }
    $telegramMessage = '<b>Site:</b> ' + $siteNameRaw + "`n"
    $SeriesMessage = ''
    $seriesEpisodeList | ForEach-Object {
        $epList = ''
        foreach ($i in $_) {
            $epList = $_.Episode._vsEpisode | Out-String
        }
        $seriesMessage = '<strong>Series:</strong> ' + $_.Series + "`n<strong>Episode:</strong>`n" + $epList
        $telegramMessage += $seriesMessage + "`n"
    }
    return $telegramMessage
}

# Sending To telegram for new file notifications
function Invoke-Telegram {
    Param(
        [Parameter( Mandatory = $true)]
        [String]$sendTelegramMessage)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    switch ($telegramNotification.ToLower()) {
        true { $telegramRequest = "https://api.telegram.org/bot$($telegramToken)/sendMessage?chat_id=$($telegramChatID)&text=$($sendTelegramMessage)&parse_mode=html&disable_notification=true" }
        false { $telegramRequest = "https://api.telegram.org/bot$($telegramToken)/sendMessage?chat_id=$($telegramChatID)&text=$($sendTelegramMessage)&parse_mode=html" }
        Default { Write-Output 'Improper configured value. Accepted values: true/false' }
    }
    Invoke-ExpressionConsole -scmFunctionName 'Telegram' -scmFunctionParams "Write-Output `"$sendTelegramMessage`""
    $ProgressPreference = 'SilentlyContinue'
    Invoke-WebRequest -Uri $telegramRequest | Out-Null
    $ProgressPreference = 'Continue'
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
    Write-Output "[MKVMerge] $(Get-Timestamp) - Default Video = $videoLang/$videoTrackName - Default Audio Language = $audioLang/$audioTrackName - Default Subtitle = $subtitleLang/$subtitleTrackName."
    $mkvVidSubtitle | ForEach-Object {
        $sp = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
        While ($True) {
            if ((Test-Lock $sp) -eq $True) {
                continue
            }
            else {
                if ($subFontName -ne 'None') {
                    Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - Python - Regex through $sp file with $subFontName."
                    Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "python `"$subtitleRegex`" `"$sp`" `"$subFontName`""
                    break
                }
                else {
                    Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - No Font specified for $sp file."
                }
            }
            Start-Sleep -Seconds 1
        }
    }
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
                Write-Output "[MKVMerge] $(Get-Timestamp) - MKV subtitle params: `"$mkvCMD`""
                Write-Output "[MKVMerge] $(Get-Timestamp) - Combining $sblist and $mkvVidInput files with $subFontDir."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$mkvVidTempOutput`" --language 0:`"$videoLang`" --track-name 0:`"$videoTrackName`" --language 1:`"$audioLang`" --track-name 1:`"$audioTrackName`" ( `"$mkvVidInput`" ) $mkvCMD --attach-file `"$subFontDir`" --attachment-mime-type application/x-truetype-font"
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) - Merging as-is. No Font specified for $sblist and $mkvVidInput files with $subFontDir."
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
            Write-Output "[MKVMerge] $(Get-Timestamp) - Removing $mkvVidInput file."
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$mkvVidInput`" -Verbose"
            $mkvVidSubtitle | ForEach-Object {
                $sp = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
                While ($True) {
                    if ((Test-Lock $sp) -eq $True) {
                        continue
                    }
                    else {
                        Write-Output "[MKVMerge] $(Get-Timestamp) - Removing $sp file."
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
            Write-Output "[MKVMerge] $(Get-Timestamp) - Renaming $mkvVidTempOutput to $mkvVidInput."
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
    Set-VideoStatus -svsKey '_vsEpisodeRaw' -svsValue $mkvVidBaseName -svsMKV
}

# Function to process video files through FileBot
function Invoke-Filebot {
    param (
        [parameter(Mandatory = $true)]
        [string]$filebotPath
    )
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to rename and move to final folder."
    $filebotVideoList = $vsCompletedFilesList | Where-Object { $_._vsDestPath -eq $filebotPath } | Select-Object _vsDestPath, _vsEpisodeFBPath, _vsEpisodeRaw, _vsEpisodeSubtitle, _vsOverridePath #_vsEpisodeSubFBPath
    
    foreach ($filebotFiles in $filebotVideoList) {
        $filebotVidInput = $filebotFiles._vsEpisodeFBPath
        $filebotSubInput = $filebotFiles._vsEpisodeSubtitle
        $filebotVidBaseName = $filebotFiles._vsEpisodeRaw
        $filebotOverrideDrive = $filebotFiles._vsOverridePath
        if ($siteParentFolder.trim() -ne '' -or $siteSubFolder.trim() -ne '') {
            $FilebotRootFolder = $filebotOverrideDrive + $siteParentFolder
            $filebotParams = Join-Path -Path (Join-Path -Path $FilebotRootFolder -ChildPath $siteSubFolder) -ChildPath $filebotArgument
            $filebotSubParams = $filebotParams + "{'.'+lang.ISO2}"
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($filebotVidInput). Renaming video and moving files to final folder. Using path($filebotParams)."
            Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$filebotVidInput`" -r --db TheTVDB -non-strict --format `"$filebotParams`" --apply date tags clean --log info"
            if (!($mkvMerge)) {
                $filebotSubInput | ForEach-Object {
                    $filebotSubParams = $filebotParams
                    $osp = $_ | Where-Object { $_.key -eq 'overrideSubPath' } | Select-Object -ExpandProperty value
                    $StLang = Get-SubtitleLanguage -subFiles $osp
                    $filebotSubParams = $filebotSubParams + "{'.$($StLang[0])'}"
                    Write-Output "[Filebot] $(Get-Timestamp) - Files found($osp). Renaming subtitle and moving files to final folder. Using path($filebotSubParams)."
                    Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$osp`" -r --db TheTVDB -non-strict --format `"$filebotSubParams`" --apply date tags clean --log info"
                }
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($filebotVidInput). ParentFolder or Subfolder path not specified. Renaming files in place  using path($filebotArgument)."
            $filebotSubInput | ForEach-Object {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($osp). ParentFolder or Subfolder path not specified. Renaming files in place."
                Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$osp`" -r --db TheTVDB -non-strict --format `"$filebotArgument`" --apply date tags clean --log info"
            }
            if (!($mkvMerge)) {
                $filebotSubInput | ForEach-Object {
                    $filebotSubParams = $filebotArgument
                    $osp = $_ | Where-Object { $_.key -eq 'overrideSubPath' } | Select-Object -ExpandProperty value
                    $StLang = Get-SubtitleLanguage -subFiles $osp
                    $filebotSubParams = $filebotSubParams + "{'.$($StLang[0])'}"
                    $filebotSubParams
                    Write-Output "[Filebot] $(Get-Timestamp) - Files found($osp). Renaming subtitle($($StLang[0])) and renaming files in place using path($filebotSubParams)."
                    Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$osp`" -r --db TheTVDB -non-strict --format `"$filebotSubParams`" --apply date tags clean --log info"
                }
            }
        }
        if (!(Test-Path -Path $filebotVidInput)) {
            Write-Output "[Filebot] $(Get-Timestamp) - Setting file($filebotVidInput) as completed."
            Set-VideoStatus -svsKey '_vsEpisodeRaw' -svsValue $filebotVidBaseName -svsFP
        }
    }
    $vsvFBCount = ($vsCompletedFilesList | Where-Object { $_._vsFBCompleted -eq $true } | Measure-Object).Count
    if ($vsvFBCount -eq $vsvTotCount ) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($vsvFBCount) = ($vsvTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup."
        Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -script fn:cleaner `"$siteHome`" --log info"
    }
    else {
        Write-Output "[Filebot] $(Get-Timestamp) - Filebot($vsvFBCount) and Total Video($vsvTotCount) count mismatch. Manual check required."
    }
    if ($vsvFBCount -ne $vsvTotCount) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
    }
}

# Setting up arraylist for MKV and Filebot lists
class VideoStatus {
    [string]$_vsSite
    [string]$_vsSeries
    [string]$_vsEpisode
    [string]$_vsSeriesDirectory
    [string]$_vsEpisodeRaw
    [string]$_vsEpisodeTemp
    [string]$_vsEpisodePath
    [array]$_vsEpisodeSubtitle
    [string]$_vsEpisodeFBPath
    [string]$_vsOverridePath
    [string]$_vsDestPathDirectory
    [string]$_vsDestPath
    [string]$_vsDestPathBase
    [bool]$_vsSECompleted
    [bool]$_vsMKVCompleted
    [bool]$_vsMoveCompleted
    [bool]$_vsFBCompleted
    [bool]$_vsErrored
    
    VideoStatus([string]$vsSite, [string]$vsSeries, [string]$vsEpisode, [string]$vsSeriesDirectory, [string]$vsEpisodeRaw, [string]$vsEpisodeTemp, [string]$vsEpisodePath, [array]$vsEpisodeSubtitle, `
            [string]$vsEpisodeFBPath, [string]$vsOverridePath, [string]$vsDestPathDirectory, [string]$vsDestPath, [string]$vsDestPathBase, [bool]$vsSECompleted, [bool]$vsMKVCompleted, `
            [bool]$vsMoveCompleted, [bool]$vsFBCompleted, [bool]$vsErrored) {
        $this._vsSite = $vsSite
        $this._vsSeries = $vsSeries
        $this._vsEpisode = $vsEpisode
        $this._vsSeriesDirectory = $vsSeriesDirectory
        $this._vsEpisodeRaw = $vsEpisodeRaw
        $this._vsEpisodeTemp = $vsEpisodeTemp
        $this._vsEpisodePath = $vsEpisodePath
        $this._vsEpisodeSubtitle = $vsEpisodeSubtitle
        $this._vsEpisodeFBPath = $vsEpisodeFBPath
        $this._vsOverridePath = $vsOverridePath
        $this._vsDestPathDirectory = $vsDestPathDirectory
        $this._vsDestPath = $vsDestPath
        $this._vsDestPathBase = $vsDestPathBase
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
$dlpScript = Join-Path -Path $scriptDirectory -ChildPath 'dlp-script.ps1'
$subtitleRegex = Join-Path -Path $scriptDirectory -ChildPath 'subtitle_regex.py'
$configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
$sharedFolder = Join-Path -Path $scriptDirectory -ChildPath 'shared'
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
    <Telegram>
        <!-- Telegram Bot tokenId/ your group ChatId, disable sound notification -->
        <token tokenId="" chatid="" disableNotification="true" />
    </Telegram>
    <credentials>
        <!-- Where you store the Site name, username/password, plexlibraryid, folder in library, and a custom font used to embed into video/sub
            <site sitename="MySiteHere">
                <username>MyUserName</username>
                <password>MyPassword</password>
                <plexlibraryid>4</plexlibraryid>
                <parentfolder>Video</parentfolder>
                <subfolder>A</subfolder>
                <font>Marker SD.ttf</font>
            </site>
        -->
        <site sitename="">
            <username></username>
            <password></password>
            <plexplexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
        </site>
        <site sitename="">
            <username></username>
            <password></password>
            <plexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
        </site>
        <site sitename="">
            <username></username>
            <password></password>
            <plexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
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
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
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
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$vrvConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
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
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$crunchyrollConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
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
--match-filter "season!~='\(.* Dub\)'"
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv[height>=1080]+ba[height>=1080] / bv+ba / b*'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$funimationConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
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
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$hidiveConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
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
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$paramountPlusConfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
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
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$testTelegramMessage = @'
<b>Site:</b> TestSite
<strong>Series:</strong> TestSeries1
<strong>Episode:</strong>
S1E1 TestSeries - TestingEpisode

<strong>Series:</strong> TestSeries2
<strong>Episode:</strong>
S2E3 TestSeries2- TestingEpisode

Test Message.
'@
$asciiLogo = @'
######   ##       #######           #######  ######   ##       #######  ##   ##  #######  ###  ##  ######   ##       #######  #######
     ##  ##            ##                      ##     ##                ##   ##       ##  #### ##       ##  ##                     ##
##   ##  ##       #######   #####   ####       ##     ##       ####     #######  #######  ## ####  ##   ##  ##       ####     #######
##   ##  ##   ##  ##       # # #    ##         ##     ##   ##  ##       ##   ##  ##   ##  ##  ###  ##   ##  ##   ##  ##       ##  ##
######   #######  ##                ##       ######   #######  #######  ##   ##  ##   ##  ##   ##  ######   #######  #######  ##   ##
_____________________________________________________________________________________________________________________________________
'@

if (!(Test-Path -Path "$scriptDirectory\config.xml" -PathType Leaf)) {
    New-Item -Path "$scriptDirectory\config.xml" -ItemType File -Force
}
if ($help) {
    Show-Markdown -Path "$scriptDirectory\README.md" -UseBrowser
    exit
}
if ($newConfig) {
    if (!(Test-Path -Path $configPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-Content -Path $configPath))) {
        
        New-Item -Path $configPath -ItemType File -Force
        Write-Output "$configPath File Created successfully."
        $xmlConfig | Set-Content -Path $configPath
    }
    else {
        Write-Output "$configPath File Exists."
        
    }
    exit
}
if ($supportFiles) {
    $suppFileFolders = $fontFolder, $sharedFolder
    foreach ($sff in $suppFileFolders) {
        Invoke-ExpressionConsole -SCMFN 'SupportFolders' -SCMFP "New-Folder -newFolderFullPath `"$sff`""
    }
    $configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
    [xml]$configFile = Get-Content -Path $configPath
    $sitenameXML = $configFile.configuration.credentials.site | Where-Object { $_.siteName.trim() -ne '' } | Select-Object 'siteName' -ExpandProperty siteName
    $sitenameXML | ForEach-Object {
        $sn = New-Object -Type PSObject -Property @{
            sn = $_.siteName
        }
        $scc = Join-Path -Path $scriptDirectory -ChildPath 'sites'
        $scf = Join-Path -Path $scc -ChildPath $sn.SN
        $scdf = $scf.TrimEnd('\') + '_D'
        $siteSupportFolders = $scc, $scf, $scdf
        foreach ($F in $siteSupportFolders) {
            Invoke-ExpressionConsole -SCMFN 'SiteFolders' -SCMFP "New-Folder -newFolderFullPath `"$F`""
        }
        $sadf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_D_A"
        $sbdf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_D_B"
        $sbdc = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_D_C"
        $saf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_A"
        $sbf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_B"
        $sbc = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_C"
        $siteSuppFiles = $sadf, $sbdf, $sbdc, $saf, $sbf, $sbc
        foreach ($S in $siteSuppFiles) {
            Invoke-ExpressionConsole -SCMFN 'SiteFiles' -SCMFP "New-SuppFiles -newSupportFiles `"$S`""
        }
        $scfDCD = $scf.TrimEnd('\') + '_D'
        $scfDC = Join-Path -Path $scfDCD -ChildPath 'yt-dlp.conf'
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
    if (Test-Path -Path $subtitleRegex) {
        Write-Output "[SETUP] $(Get-Timestamp) - $dlpScript, $subtitle_regex do exist in $scriptDirectory folder."
    }
    else {
        Write-Output "[SETUP] $(Get-Timestamp) - subtitle_regex.py does not exist or was not found in $scriptDirectory folder. Exiting."
        Exit
    }
    $date = Get-Day
    $dateTime = Get-TimeStamp
    $time = Get-Time
    $site = $site.ToLower()
    # Reading from XML
    $configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
    [xml]$configFile = Get-Content -Path $configPath
    $siteParams = $configFile.configuration.credentials.site | Where-Object { $_.siteName -ne '' -or $_.sitename -ne $null } | Select-Object 'siteName', 'username', 'password', 'plexlibraryid', 'parentfolder', 'subfolder', 'font'
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
    $siteName = $siteNameParams.siteName.ToLower()
    $siteNameRaw = $siteNameParams.siteName
    $siteFolderDirectory = Join-Path -Path $scriptDirectory -ChildPath 'sites'
    if ($daily) {
        $siteType = $siteName + '_D'
        $siteFolder = Join-Path -Path $siteFolderDirectory -ChildPath $siteType
        $logFolderBase = Join-Path -Path $siteFolder -ChildPath 'log'
        $logFolderBaseDate = Join-Path -Path $logFolderBase -ChildPath $date
        $logFile = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime.log"
        Start-Transcript -Path $logFile -UseMinimalHeader
        Write-Output "[Setup] $(Get-Timestamp) - $siteNameRaw"
    }
    else {
        $siteType = $siteName
        $siteFolder = Join-Path -Path $siteFolderDirectory -ChildPath $siteType
        $logFolderBase = Join-Path -Path $siteFolder -ChildPath 'log'
        $logFolderBaseDate = Join-Path -Path $logFolderBase -ChildPath $date
        $logFile = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime.log"
        Start-Transcript -Path $logFile -UseMinimalHeader
        Write-Output "[Setup] $(Get-Timestamp) - $siteNameRaw"
    }
    $siteUser = $siteNameParams.username
    $sitePass = $siteNameParams.password
    $siteLibraryID = $siteNameParams.plexlibraryid
    $siteParentFolder = $siteNameParams.parentfolder
    $siteSubFolder = $siteNameParams.subfolder
    $subFontExtension = $siteNameParams.font
    $backupDrive = $configFile.configuration.Directory.backup.location
    $tempDrive = $configFile.configuration.Directory.temp.location
    $srcDrive = $configFile.configuration.Directory.src.location
    $destDrive = $configFile.configuration.Directory.dest.location
    $ffmpeg = $configFile.configuration.Directory.ffmpeg.location
    [int]$emptyLogs = $configFile.configuration.Logs.keeplog.emptylogskeepdays
    [int]$filledLogs = $configFile.configuration.Logs.keeplog.filledlogskeepdays
    $plexHost = $configFile.configuration.Plex.plexcred.plexUrl
    $plexToken = $configFile.configuration.Plex.plexcred.plexToken
    $filebotArgument = $configFile.configuration.Filebot.fbfolder.fbArgument
    $overrideSeriesList = $configFile.configuration.OverrideSeries.override | Where-Object { $_.orSeriesId -ne '' -and $_.orSrcdrive -ne '' }
    $telegramToken = $configFile.configuration.Telegram.token.tokenId
    $telegramChatID = $configFile.configuration.Telegram.token.chatid
    $telegramNotification = $configFile.configuration.Telegram.token.disableNotification
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
        es { $videoLang = $audioLang; $audioTrackName = 'Spanish(Latin America) Audio'; $videoTrackName = 'Spanish(Latin America) Video' }
        es-es { $videoLang = $audioLang; $audioTrackName = 'Spanish(Spain) Audio'; $videoTrackName = 'Spanish(Spain) Video' }
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
        es { $subtitleTrackName = 'Spanish(Latin America) Sub' }
        es-es { $subtitleTrackName = 'Spanish(Spain) Sub' }
        fr { $subtitleTrackName = 'French Sub' }
        it { $subtitleTrackName = 'Italian Sub' }
        ja { $subtitleTrackName = 'Japanese Sub' }
        pt-br { $subtitleTrackName = 'Português (Brasil) Sub' }
        pt-pt { $subtitleTrackName = 'Português (Portugal) Sub' }
        ru { $subtitleTrackName = 'Russian Video' }
        ja { $subtitleTrackName = 'Japanese Sub' }
        en { $subtitleTrackName = 'English Sub' }
        und { $subtitleLang = 'und'; $subtitleTrackName = 'und sub' }
    }
    # End reading from XML
    if ($subFontExtension.Trim() -ne '') {
        $subFontDir = Join-Path -Path $fontFolder -ChildPath $subFontExtension
        if (Test-Path -Path $subFontDir) {
            $subFontName = [System.Io.Path]::GetFileNameWithoutExtension($subFontExtension)
            Write-Output "[Setup] $(Get-Timestamp) - $subFontExtension set for $siteName."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $subFontExtension specified in $configFile is missing from $fontFolder. Exiting."
            Exit-Script
        }
    }
    else {
        $subFontDir = 'None'
        $subFontName = 'None'
        $subFontExtension = 'None'
        Write-Output "[Setup] $(Get-Timestamp) - $subFontExtension - No font set for $siteName."
    }
    $siteShared = Join-Path -Path $scriptDirectory -ChildPath 'shared'
    $srcBackup = Join-Path -Path $backupDrive -ChildPath '_Backup'
    $srcBackupDriveShared = Join-Path -Path $srcBackup -ChildPath 'shared'
    $srcDriveSharedFonts = Join-Path -Path $srcBackup -ChildPath 'fonts'
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
        $siteConfig = Join-Path -Path $siteFolder -ChildPath 'yt-dlp.conf'
        if ($srcDrive -eq $tempDrive) {
            Write-Output "[Setup] $(Get-Timestamp) - Src($srcDrive) and Temp($tempDrive) Directories cannot be the same"
            Exit-Script
        }
        if ((Test-Path -Path $siteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $siteConfig file found. Continuing."
            $dlpParams = $dlpParams + " --config-location $siteConfig -P temp:$siteTemp -P home:$siteSrc"
            $dlpArray += "`"--config-location`"", "`"$siteConfig`"", "`"-P`"", "`"temp:$siteTemp`"", "`"-P`"", "`"home:$siteSrc`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $siteConfig does not exist. Exiting."
            Exit-Script
        }
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
        $siteConfig = $siteFolder + '\yt-dlp.conf'
        if ((Test-Path -Path $siteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $siteConfig file found. Continuing."
            $dlpParams = $dlpParams + " --config-location $siteConfig -P temp:$siteTemp -P home:$siteSrc"
            $dlpArray += "`"--config-location`"", "`"$siteConfig`"", "`"-P`"", "`"temp:$siteTemp`"", "`"-P`"", "`"home:$siteSrc`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $siteConfig does not exist. Exiting."
            Exit-Script
        }
    }
    $siteConfigBackup = Join-Path -Path (Join-Path -Path $srcBackup -ChildPath 'sites') -ChildPath $siteType
    $cookieFile = Join-Path -Path $siteShared -ChildPath "$($siteType)_C"
    if ($login) {
        if ($siteUser -and $sitePass) {
            Write-Output "[Setup] $(Get-Timestamp) - Login is true and SiteUser/Password is filled. Continuing."
            $dlpParams = $dlpParams + " -u $siteUser -p $sitePass"
            $dlpArray += "`"-u`"", "`"$siteUser`"", "`"-p`"", "`"$sitePass`""
            if ($cookies) {
                if ((Test-Path -Path $cookieFile)) {
                    Write-Output "[Setup] $(Get-Timestamp) - Cookies is true and $cookieFile file found. Continuing."
                    $dlpParams = $dlpParams + " --cookies $cookieFile"
                    $dlpArray += "`"--cookies`"", "`"$cookieFile`""
                }
                else {
                    Write-Output "[Setup] $(Get-Timestamp) - $cookieFile does not exist. Exiting."
                    Exit-Script
                }
            }
            else {
                $cookieFile = 'None'
                Write-Output "[Setup] $(Get-Timestamp) - Login is true and Cookies is false. Continuing."
            }
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - Login is true and Username/Password is Empty. Exiting."
            Exit-Script
        }
    }
    else {
        if ((Test-Path -Path $cookieFile)) {
            Write-Output "[Setup] $(Get-Timestamp) - $cookieFile file found. Continuing."
            $dlpParams = $dlpParams + " --cookies $cookieFile"
            $dlpArray += "`"--cookies`"", "`"$cookieFile`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $cookieFile does not exist. Exiting."
            Exit-Script
        }
    }
    if ($ffmpeg) {
        Write-Output "[Setup] $(Get-Timestamp) - $ffmpeg file found. Continuing."
        $dlpParams = $dlpParams + " --ffmpeg-location $ffmpeg"
        $dlpArray += "`"--ffmpeg-location`"", "`"$ffmpeg`""
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - FFMPEG: $ffmpeg missing. Exiting."
        Exit-Script
    }
    $batFile = Join-Path -Path $siteShared -ChildPath "$($siteType)_B"
    if ((Test-Path -Path $batFile)) {
        Write-Output "[Setup] $(Get-Timestamp) - $batFile file found. Continuing."
        if (![String]::IsNullOrWhiteSpace((Get-Content -Path $batFile))) {
            Write-Output "[Setup] $(Get-Timestamp) - $batFile not empty. Continuing."
            $dlpParams = $dlpParams + " -a $batFile"
            $dlpArray += "`"-a`"", "`"$batFile`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $batFile is empty. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - BAT: $batFile missing. Exiting."
        Exit-Script
    }
    if ($archive) {
        $archiveFile = Join-Path -Path $siteShared -ChildPath "$($siteType)_A"
        if ((Test-Path -Path $archiveFile)) {
            Write-Output "[Setup] $(Get-Timestamp) - $archiveFile file found. Continuing."
            $dlpParams = $dlpParams + " --download-archive $archiveFile"
            $dlpArray += "`"--download-archive`"", "`"$archiveFile`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - Archive file missing. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - Using --no-download-archive. Continuing."
        $archiveFile = 'None'
        $dlpParams = $dlpParams + ' --no-download-archive'
        $dlpArray += "`"--no-download-archive`""
    }
    $videoType = Select-String -Path $siteConfig -Pattern '--remux-video.*' | Select-Object -First 1
    if ($null -ne $videoType) {
        $vidType = '*.' + ($videoType -split ' ')[1]
        $vidType = $vidType.Replace("'", '').Replace('"', '')
        if ($vidType -eq '*.mkv') {
            Write-Output "[Setup] $(Get-Timestamp) - Using $vidType. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - VidType(mkv) is missing. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - --remux-video parameter is missing. Exiting."
        Exit-Script
    }
    if (Select-String -Path $siteConfig '--write-subs' -SimpleMatch -Quiet) {
        Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit or MKVMerge is true and --write-subs is in config. Continuing."
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting."
        Exit-Script
    }
    $subtitleType = Select-String -Path $siteConfig -Pattern '--convert-subs.*' | Select-Object -First 1
    if ($null -ne $subtitleType) {
        $subType = '*.' + ($subtitleType -split ' ')[1]
        $subType = $subType.Replace("'", '').Replace('"', '')
        if ($subType -eq '*.ass') {
            Write-Output "[Setup] $(Get-Timestamp) - Using $subType. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - Subtype(ass) is missing. Exiting."
            Exit-Script
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - --convert-subs parameter is missing. Exiting."
        Exit-Script
    }
    $writeSub = Select-String -Path $siteConfig -Pattern '--write-subs.*' | Select-Object -First 1
    if ($null -ne $writeSub) {
        Write-Output "[Setup] $(Get-Timestamp) - --write-subs is in config. Continuing."
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - SubSubWrite is missing. Exiting."
        Exit-Script
    }
    $debugVars = [ordered]@{Site = $siteName; IsDaily = $daily; UseLogin = $login; UseCookies = $cookies; UseArchive = $archive; SubtitleEdit = $subtitleEdit; `
            MKVMerge = $mkvMerge; VideoTrackName = $videoTrackName; AudioLang = $audioLang; AudioTrackName = $audioTrackName; SubtitleLang = $subtitleLang; SubtitleTrackName = $subtitleTrackName; Filebot = $filebot; `
            SiteNameRaw = $siteNameRaw; SiteType = $siteType; SiteUser = $siteUser; SitePass = $sitePass; SiteFolderIdName = $siteFolderIdName[0]; SiteFolder = $siteFolder; SiteParentFolder = $siteParentFolder; `
            SiteSubFolder = $siteSubFolder; SiteLibraryId = $siteLibraryID; SiteTemp = $siteTemp; SiteSrcBase = $siteSrcBase; SiteSrc = $siteSrc; SiteHomeBase = $siteHomeBase; `
            SiteHome = $siteHome; SiteDefaultPath = $siteDefaultPath; SiteConfig = $siteConfig; CookieFile = $cookieFile; Archive = $archiveFile; Bat = $batFile; Ffmpeg = $ffmpeg; SubFontName = $subFontName; SubFontExtension = $subFontExtension; `
            SubFontDir = $subFontDir; SubType = $subType; VidType = $vidType; Backup = $srcBackup; BackupShared = $srcBackupDriveShared; BackupFont = $srcDriveSharedFonts; `
            SiteConfigBackup = $siteConfigBackup; PlexHost = $plexHost; PlexToken = $plexToken; telegramToken = $telegramToken; TelegramChatId = $telegramChatID; ConfigPath = $configPath; `
            ScriptDirectory = $scriptDirectory; dlpParams = $dlpParams
    }
    if ($testScript) {
        Write-Output "[START] $dateTime - $siteNameRaw - DEBUG Run"
        $debugVars
        $overrideSeriesList
        Write-Output 'dlpArray:'
        $dlpArray
        if ($sendTelegram) {
            Invoke-Telegram -sendTelegramMessage $testTelegramMessage
        }
        Exit-Script
    }
    else {
        $debugVarsRemove = 'SiteUser' , 'SitePass', 'PlexToken', 'telegramToken', 'TelegramChatId', 'dlpParams'
        foreach ($dbv in $debugVarsRemove) {
            $debugVars.Remove($dbv)
        }
        if ($daily) {
            Write-Output "[START] $dateTime - $siteNameRaw - Daily Run"
        }
        else {
            Write-Output "[START] $dateTime - $siteNameRaw - Manual Run"
        }
        Write-Output '[START] Debug Vars:'
        $debugVars
        Write-Output '[START] Series Drive Overrides:'
        $overrideSeriesList | Sort-Object orSrcdrive, orSeriesName | Format-Table
        # Create folders
        $createFolders = $tempDrive, $srcDrive, $backupDrive, $srcBackup, $siteConfigBackup, $srcBackupDriveShared, $srcDriveSharedFonts, $destDrive, $siteTemp, $siteSrc, $siteHome
        foreach ($cf in $createFolders) {
            Invoke-ExpressionConsole -SCMFN 'START' -SCMFP "New-Folder -newFolderFullPath `"$cf`""
        }
        # Log cleanup
        Invoke-ExpressionConsole -SCMFN 'LogCleanup' -SCMFP 'Remove-Logfiles'
    }
    # yt-dlp
    & yt-dlp.exe $dlpArray *>&1 | Out-Host
    # Post-processing
    $totalDownloaded = (Get-ChildItem -Path $siteSrc -Recurse -Force -File -Include "$vidType" | Select-Object * | Measure-Object).Count
    if ($totalDownloaded -gt 0) {
        Get-ChildItem -Path $siteSrc -Recurse -Include "$vidType" | Sort-Object LastWriteTime | Select-Object -Unique | Get-Unique | ForEach-Object {
            [System.Collections.ArrayList]$vsEpisodeSubtitle = @{}
            $vsSite = $siteNameRaw
            $vsSeries = (Get-Culture).TextInfo.ToTitleCase(("$(Split-Path (Split-Path $_ -Parent) -Leaf)").Replace('_', ' ').Replace('-', ' ')) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $vsOverridePath = $overrideSeriesList | Where-Object { $_.orSeriesName.ToLower() -eq $vsSeries.ToLower() } | Select-Object -ExpandProperty orSrcdrive
            $vsEpisode = (Get-Culture).TextInfo.ToTitleCase( ($_.BaseName.Replace('_', ' ').Replace('-', ' '))) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $vsSeriesDirectory = $_.DirectoryName
            $vsEpisodeRaw = $_.BaseName
            $vsEpisodeTemp = Join-Path -Path $vsSeriesDirectory -ChildPath "$($vsEpisodeRaw).temp$($_.Extension)"
            $vsEpisodePath = $_.FullName
            if ($null -ne $vsOverridePath) {
                $vsDestPath = Join-Path -Path $vsOverridePath -ChildPath $siteHome.Substring(3)
                $vsDestPathBase = Join-Path -Path $vsOverridePath -ChildPath $siteHomeBase.Substring(3)
                $vsEpisodeFBPath = $vsEpisodePath.Replace($siteSrc, $vsDestPath)
                $vsDestPathDirectory = Split-Path $vsEpisodeFBPath -Parent
            }
            else {
                $vsDestPath = $siteHome
                $vsDestPathBase = $siteHomeBase
                $vsEpisodeFBPath = $vsEpisodePath.Replace($siteSrc, $vsDestPath)
                $vsDestPathDirectory = Split-Path $vsEpisodeFBPath -Parent
                $vsOverridePath = [System.IO.path]::GetPathRoot($vsDestPath)
                
            }
            Get-ChildItem -Path $siteSrc -Recurse -Include $subType | Where-Object { $_.FullName -match $vsEpisodeRaw } | Select-Object Name, FullName, BaseName -Unique | ForEach-Object {
                [System.Collections.ArrayList]$episodeSubtitles = @{}
                $origSubPath = $_.FullName
                $subtitleBase = $_.name
                if ($null -ne $vsOverridePath) {
                    $overrideSubPath = $origSubPath.Replace($siteSrc, $vsDestPath)
                }
                else {
                    $overrideSubPath = $origSubPath
                }
                $episodeSubtitles = @{ subtitleBase = $subtitleBase; origSubPath = $origSubPath; overrideSubPath = $overrideSubPath }
                [void]$vsEpisodeSubtitle.Add($episodeSubtitles)
            }
            foreach ($i in $_) {
                $VideoStatus = [VideoStatus]::new($vsSite, $vsSeries, $vsEpisode, $vsSeriesDirectory, $vsEpisodeRaw, $vsEpisodeTemp, $vsEpisodePath, $vsEpisodeSubtitle, `
                        $vsEpisodeFBPath, $vsOverridePath, $vsDestPathDirectory, $vsDestPath, $vsDestPathBase, $vsSECompleted, $vsMKVCompleted, $vsMoveCompleted, $vsFBCompleted, $vsErrored)
                [void]$vsCompletedFilesList.Add($VideoStatus)
            }
        }
        $vsCompletedFilesList | Select-Object -ExpandProperty _vsEpisodeSubtitle | ForEach-Object {
            $seSubtitle = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
            if ($null -eq $seSubtitle) {
                Set-VideoStatus -svsKey '_vsEpisodeRaw' -svsValue $vsEpisodeRaw -svsER
            }
        }
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files to process."
    }
    $vsCompletedFilesList
    $vsvTotCount = ($vsCompletedFilesList | Measure-Object).Count
    $vsvErrorCount = ($vsCompletedFilesList | Where-Object { $_._vsErrored -eq $true } | Measure-Object).Count
    # SubtitleEdit, MKVMerge, Filebot
    if ($vsvTotCount -gt 0) {
        # Subtitle Edit
        if ($subtitleEdit) {
            $vsCompletedFilesList | Where-Object { $_._vsErrored -ne $true } | Select-Object -ExpandProperty _vsEpisodeSubtitle | ForEach-Object {
                $seSubtitle = $_ | Where-Object { $_.key -eq 'origSubPath' } | Select-Object -ExpandProperty value
                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $seSubtitle subtitle."
                While ($True) {
                    if ((Test-Lock $seSubtitle) -eq $True) {
                        continue
                    }
                    else {
                        Invoke-ExpressionConsole -SCMFN 'SubtitleEdit' -SCMFP "powershell `"SubtitleEdit /convert `'$seSubtitle`' AdvancedSubStationAlpha /overwrite /MergeSameTimeCodes`""
                        Set-VideoStatus -svsKey '_vsEpisodeSubtitle' -svsValue $seSubtitle -svsSE
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Not running."
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
            Write-Output "[MKVMerge] $(Get-Timestamp) - MKVMerge not running. Moving to next step."
            $overrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsErrored -eq $false } | Select-Object _vsSeriesDirectory, _vsDestPath, _vsDestPathBase -Unique
        }
        $vsvMKVCount = ($vsCompletedFilesList | Where-Object { $_._vsMKVCompleted -eq $true -and $_._vsErrored -eq $false } | Measure-Object).Count
        # FileMoving
        if (($vsvMKVCount -eq $vsvTotCount -and $vsvErrorCount -eq 0) -or (!($mkvMerge) -and $vsvErrorCount -eq 0)) {
            Write-Output "[FileMoving] $(Get-Timestamp) - All files had matching subtitle file"
            foreach ($orDriveList in $overrideDriveList) {
                Write-Output "[FileMoving] $(Get-Timestamp) - $($orDriveList._vsSeriesDirectory) contains files. Moving to $($orDriveList._vsDestPath)."
                if (!(Test-Path -Path $orDriveList._vsDestPath)) {
                    Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "New-Folder -newFolderFullPath `"$($orDriveList._vsDestPath)`" -Verbose"
                }
                Write-Output "[FileMoving] $(Get-Timestamp) - Moving $($orDriveList._vsSeriesDirectory) to $($orDriveList._vsDestPath)."
                Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($orDriveList._vsSeriesDirectory)`" -Destination `"$($orDriveList._vsDestPath)`" -Force -Verbose"
                if (!(Test-Path -Path $orDriveList._vsSeriesDirectory)) {
                    Write-Output "[FileMoving] $(Get-Timestamp) - Move completed for $($orDriveList._vsSeriesDirectory)."
                    Set-VideoStatus -svsKey '_vsSeriesDirectory' -svsValue $orDriveList._vsSeriesDirectory -svsMove
                    
                }
            }
        }
        else {
            Write-Output "[FileMoving] $(Get-Timestamp) - $siteSrc contains file(s) with error(s). Not moving files."
        }
        # Filebot
        $filebotOverrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsErrored -eq $false } | Select-Object _vsDestPath -Unique
        if (($filebot -and $vsvMKVCount -eq $vsvTotCount) -or ($filebot -and !($mkvMerge))) {
            foreach ($fbORDriveList in $filebotOverrideDriveList) {
                Write-Output "[Filebot] $(Get-Timestamp) - Renaming files in $($fbORDriveList._vsDestPath)."
                Invoke-Filebot -filebotPath $orDriveList._vsDestPath
            }
        }
        elseif (!($filebot) -and !($mkvMerge) -or (!($filebot) -and $vsvMKVCount -eq $vsvTotCount)) {
            $moveManualList = $vsCompletedFilesList | Where-Object { $_._vsMoveCompleted -eq $true } | Select-Object _vsDestPathDirectory, _vsOverridePath -Unique
            foreach ($mmFiles in $moveManualList) {
                $mmOverrideDrive = $mmFiles._vsOverridePath
                $moveRootDirectory = $mmOverrideDrive + $siteParentFolder
                $moveFolder = Join-Path -Path $moveRootDirectory -ChildPath $siteSubFolder
                Write-Output "[FileMoving] $(Get-Timestamp) - Moving $($mmFiles._vsDestPathDirectory) to $moveFolder."
                Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($mmFiles._vsDestPathDirectory)`" -Destination `"$moveFolder`" -Force -Verbose"
            }
        }
        else {
            Write-Output "[FileMoving] $(Get-Timestamp) - Issue with files in $siteSrc."
        }
        # Plex
        if ($plexHost -and $plexToken -and $siteLibraryID ) {
            Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
            $plexUrl = "$plexHost/library/sections/$siteLibraryID/refresh?X-Plex-Token=$plexToken"
            $ProgressPreference = 'SilentlyContinue'
            Invoke-WebRequest -Uri $plexUrl | Out-Null
            $ProgressPreference = 'Continue'
        }
        else {
            Write-Output "[PLEX] $(Get-Timestamp) - Not using Plex."
        }
        $vsvFBCount = ($vsCompletedFilesList | Where-Object { $_._vsFBCompleted -eq $true } | Measure-Object).Count
        # Telegram
        if ($sendTelegram) {
            Write-Output "[Telegram] $(Get-Timestamp) - Preparing Telegram message."
            $tm = Get-SiteSeriesEpisode
            if ($plexHost -and $plexToken -and $siteLibraryID) {
                if ($filebot -or $mkvMerge) {
                    if (($vsvFBCount -gt 0 -and $vsvMKVCount -gt 0 -and $vsvFBCount -eq $vsvMKVCount) -or (!($filebot) -and $mkvMerge -and $vsvMKVCount -gt 0)) {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $siteHome. Success."
                        $tm += 'All files added to PLEX.'
                        Invoke-Telegram -sendTelegramMessage $tm
                    }
                    else {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $siteHome. Failure."
                        $tm += 'Not all files added to PLEX.'
                        Invoke-Telegram -sendTelegramMessage $tm
                    }
                }
            }
            else {
                Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $siteHome."
                $tm += 'Added files to folders.'
                Invoke-Telegram -sendTelegramMessage $tm
            }
        }
        Write-Output "[VideoList] $(Get-Timestamp) - Total Files: $vsvTotCount"
        Write-Output "[VideoList] $(Get-Timestamp) - Errored Files: $vsvErrorCount"
        $vsCompletedFilesListHeaders = @{Label = 'Series'; Expression = { $_._vsSeries } }, @{Label = 'Episode'; Expression = { $_._vsEpisode } }, 
        @{Label = 'SECompleted'; Expression = { $_._vsSECompleted } }, @{Label = 'MKVCompleted'; Expression = { $_._vsMKVCompleted } }, @{Label = 'MoveCompleted'; Expression = { $_._vsMoveCompleted } }, `
        @{Label = 'FBCompleted'; Expression = { $_._vsFBCompleted } }, @{Label = 'Errored'; Expression = { $_._vsErrored } }, @{Label = 'Subtitle'; Expression = { $_._vsEpisodeSubtitle.value } }, `
        @{Label = 'Drive'; Expression = { $_._vsOverridePath } }, @{Label = 'SrcDirectory'; Expression = { $_._vsSeriesDirectory } }, @{Label = 'DestBase'; Expression = { $_._vsDestPathBase } }, `
        @{Label = 'DestPath'; Expression = { $_._vsDestPath } }, @{Label = 'DestPathDirectory'; Expression = { $_._vsDestPathDirectory } }
        if ($vsvTotCount -gt 12) {
            $vsCompletedFilesTable = $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsCompletedFilesListHeaders -Expand CoreOnly
        }
        else {
            $vsCompletedFilesTable = $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-List $vsCompletedFilesListHeaders -Expand CoreOnly
        }
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files downloaded. Skipping other defined steps."
    }
    # Backup
    $sharedBackups = $archiveFile, $cookieFile, $batFile, $configPath, $subFontDir, $siteConfig
    foreach ($sb in $sharedBackups) {
        if (($sb -ne 'None') -and ($sb.trim() -ne '')) {
            if ($sb -eq $subFontDir) {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $srcDriveSharedFonts."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$srcDriveSharedFonts`" -PassThru -Verbose"
            }
            elseif ($sb -eq $configPath) {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $srcBackup."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$srcBackup`" -PassThru -Verbose"
            }
            elseif ($sb -eq $siteConfig) {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $siteConfigBackup."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$siteConfigBackup`" -PassThru -Verbose"
            }
            else {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $srcBackupDriveShared."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$srcBackupDriveShared`" -PassThru -Verbose"
            }
        }
    }
    Exit-Script
}