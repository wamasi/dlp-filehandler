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
    $iecObject | Where-Object { $_.length -gt 0 } | ForEach-Object {
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
        New-Item -Path $newFolderFullPath -ItemType Directory -Force | Out-Null
        Write-Output "[SetFolder] - $(Get-Timestamp) - $newFolderFullPath missing. Creating."
    }
    else {
        Write-Output "[SetFolder] - $(Get-Timestamp) - $newFolderFullPath already exists."
    }
}

function New-SuppFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string] $newSupportFiles
    )
    if (!(Test-Path -Path $newSupportFiles -PathType Leaf)) {
        New-Item -Path $newSupportFiles -ItemType File | Out-Null
        Write-Output "[NewSupportFiles] - $(Get-Timestamp) - $newSupportFiles file missing. Creating."
    }
    else {
        Write-Output "[NewSupportFiles] - $(Get-Timestamp) - $newSupportFiles file already exists."
    }
}

function New-Config {
    param (
        [Parameter(Mandatory = $true)]
        [string] $newConfigs
    )
    New-Item -Path $newConfigs -ItemType File -Force | Out-Null
    Write-Output "[NewConfigFiles] - $(Get-Timestamp) - Creating $newConfigs"
    if ($newConfigs -match 'vrv') {
        $vrvConfig | Set-Content -Value $newConfigs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $newConfigs created with VRV values."
    }
    elseif ($newConfigs -match 'crunchyroll') {
        $crunchyrollConfig | Set-Content -Value $newConfigs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $newConfigs created with Crunchyroll values."
    }
    elseif ($newConfigs -match 'funimation') {
        $funimationConfig | Set-Content -Value $newConfigs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $newConfigs created with Funimation values."
    }
    elseif ($newConfigs -match 'hidive') {
        $hidiveConfig | Set-Content -Value $newConfigs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $newConfigs created with Hidive values."
    }
    elseif ($newConfigs -match 'paramountplus') {
        $paramountPlusConfig | Set-Content -Value $newConfigs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $newConfigs created with ParamountPlus values."
    }
    else {
        $defaultConfig | Set-Content -Value $newConfigs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $newConfigs created with default values."
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
        if (($removeFolder -match '\\tmp\\') -and ($removeFolder -match $removeFolderBaseMatch) -and (Test-Path -Path $removeFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $removeFolder folders/files."
            Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Item -Path `"$removeFolder`" -Recurse -Force -Verbose"
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteTemp($removeFolder) folder already deleted. Nothing to remove."
        }
    }
    else {
        if (!(Test-Path -Path $removeFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($removeFolder) already deleted."
        }
        elseif ((Test-Path -Path $removeFolder) -and (Get-ChildItem -Path $removeFolder -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($removeFolder) is empty. Deleting folder."
            & $deleteRecursion -deleteRecursionPath $removeFolder
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($removeFolder) contains files. Manual attention needed."
        }
    }
}

function Remove-Logfiles {
    # Log cleanup
    $filledLogsLimit = (Get-Date).AddDays(-$filledLogs)
    $emptyLogsLimit = (Get-Date).AddDays(-$emptyLogs)
    if (!(Test-Path -Path $logFolderBase)) {
        Write-Output "[LogCleanup] $(Get-Timestamp) - $logFolderBase is missing. Skipping log cleanup."
    }
    else {
        Write-Output "[LogCleanup] $(Get-Timestamp) - $logFolderBase found. Starting Filledlog($filledLogs days) cleanup."
        Get-ChildItem -Path $logFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -match '.*-Total-.*' -and $_.FullName -ne $logFile -and $_.CreationTime -lt $filledLogsLimit } | `
            ForEach-Object {
            $removeLog = $_.FullName
            Invoke-ExpressionConsole -SCMFN 'LogCleanup' -SCMFP "`"$removeLog`" | Remove-Item -Recurse -Force -Verbose"
        }
        Write-Output "[LogCleanup] $(Get-Timestamp) - $logFolderBase found. Starting emptylog($emptyLogs days) cleanup."
        Get-ChildItem -Path $logFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -notmatch '.*-Total-.*' -and $_.FullName -ne $logFile -and $_.CreationTime -lt $emptyLogsLimit } | `
            ForEach-Object {
            $removeLog = $_.FullName
            Invoke-ExpressionConsole -SCMFN 'LogCleanup' -SCMFP "`"$removeLog`" | Remove-Item -Recurse -Force -Verbose"
        }
        & $deleteRecursion -deleteRecursionPath $logFolderBase
    }
}

$deleteRecursion = {
    param(
        $deleteRecursionPath
    )
    foreach ($DeleteRecursionDirectory in Get-ChildItem -LiteralPath $deleteRecursionPath -Directory -Force) {
        & $deleteRecursion -deleteRecursionPath $DeleteRecursionDirectory.FullName
    }
    $DRcurrentChildren = Get-ChildItem -LiteralPath $deleteRecursionPath -Force
    $DeleteRecursionEmpty = $DRcurrentChildren -eq $null
    if ($DeleteRecursionEmpty) {
        Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting '${deleteRecursionPath}' folders/files if empty."
        Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Item -Force -LiteralPath `"$deleteRecursionPath`" -Verbose"
    }
}

function Remove-Spaces {
    param (
        [Parameter(Mandatory = $true)]
        [string] $removeSpacesFile
    )
    (Get-Content -Path $removeSpacesFile) | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } | Set-Content -Value $removeSpacesFile
    $removeSpacesContent = [System.IO.File]::ReadAllText($removeSpacesFile)
    $removeSpacesContent = $removeSpacesContent.Trim()
    [System.IO.File]::WriteAllText($removeSpacesFile, $removeSpacesContent)
}

function Exit-Script {
    param (
        [alias('Exit')]
        [switch]$exitScript
    )
    $scriptStopWatch.Stop()
    # Cleanup folders
    Remove-Folders -removeFolder $siteTemp -removeFolderMatch '\\tmp\\' -removeFolderBaseMatch $siteTempBaseMatch
    Remove-Folders -removeFolder $siteSrc -removeFolderMatch '\\src\\' -removeFolderBaseMatch $siteSrcBaseMatch
    Remove-Folders -removeFolder $siteHome -removeFolderMatch '\\tmp\\' -removeFolderBaseMatch $siteHomeBaseMatch
    if ($overrideDriveList.count -gt 0) {
        foreach ($orDriveList in $overrideDriveList) {
            $orDriveListBaseMatch = ($orDriveList._vsDestPathBase).Replace('\', '\\')
            Remove-Folders -removeFolder $orDriveList._vsDestPath -removeFolderMatch '\\tmp\\' -removeFolderBaseMatch $orDriveListBaseMatch
        }
    }
    # Cleanup Log Files
    Remove-Logfiles
    Write-Output "[END] $(Get-Timestamp) - Script completed. Total Elapsed Time: $($scriptStopWatch.Elapsed.ToString())"
    Stop-Transcript
    ((Get-Content -Path $logFile | Select-Object -Skip 5) | Select-Object -SkipLast 4) | Set-Content -Value $logFile
    Remove-Spaces -removeSpacesFile $logFile
    if ($exitScript -and !($testScript)) {
        $logTemp = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime-Temp.log"
        New-Item -Path $logTemp -ItemType File | Out-Null
        $asciiLogo | Out-File -FilePath $logTemp -Width 9999
        Get-Content -Path $logFile -ReadCount 5000 | ForEach-Object {
            $_ | Add-Content -Value "$logTemp"
        }
        Remove-Item -Path $logFile
        Rename-Item -Path $logTemp -NewName $logFile
        exit
    }
    elseif ($testScript) {
        $logTemp = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime-Temp.log"
        New-Item -Path $logTemp -ItemType File | Out-Null
        $asciiLogo | Out-File -FilePath$logTemp -Width 9999
        Get-Content -Path $logFile -ReadCount 5000 | ForEach-Object {
            $_ | Add-Content -Value "$logTemp"
        }
        Remove-Item -Path $logFile
        Rename-Item -Path $logTemp -NewName "$dateTime-DEBUG.log"
        exit
    }
    else {
        if ($vsvTotCount -gt 0) {
            $logTemp = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime-Temp.log"
            New-Item -Path $logTemp -ItemType File | Out-Null
            $asciiLogo | Out-File -FilePath$logTemp -Width 9999
            $vsCompletedFilesTable | Out-File -FilePath$logTemp -Width 9999 -Append
            Get-Content -Path $logFile -ReadCount 5000 | ForEach-Object {
                $_ | Add-Content -Value "$logTemp"
            }
            Remove-Item -Path $logFile
            Rename-Item -Path $logTemp -NewName "$dateTime-Total-$vsvTotCount.log"
        }
        else {
            $logTemp = Join-Path -Path $logFolderBaseDate -ChildPath "$dateTime-Temp.log"
            New-Item -Path $logTemp -ItemType File | Out-Null
            $asciiLogo | Out-File -FilePath$logTemp -Width 9999
            Get-Content -Path $logFile -ReadCount 5000 | ForEach-Object {
                $_ | Add-Content -Value "$logTemp"
            }
            Remove-Item -Path $logFile
            Rename-Item -Path $logTemp -NewName $logFile
        }
    }
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
    Invoke-WebRequest -Uri $telegramRequest | Out-Null
}

# Run MKVMerge process
function Invoke-MKVMerge {
    param (
        [parameter(Mandatory = $true)]
        [string]$mkvVidInput,
        [parameter(Mandatory = $true)]
        [string]$mkvVidBaseName,
        [parameter(Mandatory = $true)]
        [string]$mkvVidSubtitle,
        [parameter(Mandatory = $true)]
        [string]$mkvVidTempOutput
    )
    Write-Output "[MKVMerge] $(Get-Timestamp) - Video = $videoLang/$videoTrackName - Audio Language = $audioLang/$audioTrackName - Subtitle = $subtitleLang/$subtitleTrackName."
    Write-Output "[MKVMerge] $(Get-Timestamp) - Replacing Styling in $mkvVidSubtitle."
    While ($True) {
        if ((Test-Lock $mkvVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($subFontName -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - Python - Regex through $mkvVidSubtitle file with $subFontName."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "python `"$subtitleRegex`" `"$mkvVidSubtitle`" `"$subFontName`""
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - No Font specified for $mkvVidSubtitle file."
            }
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $mkvVidInput) -eq $True -and (Test-Lock $mkvVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($subFontDir -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - Combining $mkvVidSubtitle and $mkvVidInput files with $subFontDir."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$mkvVidTempOutput`" --language 0:`"$videoLang`" --track-name 0:`"$videoTrackName`" --language 1:`"$audioLang`" --track-name 1:`"$audioTrackName`" ( `"$mkvVidInput`" ) --language 0:`"$subtitleLang`" --track-name 0:`"$subtitleTrackName`" `"$mkvVidSubtitle`" --attach-file `"$subFontDir`" --attachment-mime-type application/x-truetype-font"
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) -  Merging as-is. No Font specified for $mkvVidSubtitle and $mkvVidInput files with $subFontDir."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$mkvVidTempOutput`" --language 0:`"$videoLang`" --track-name 0:`"$videoTrackName`" --language 1:`"$audioLang`" --track-name 1:`"$audioTrackName`" ( `"$mkvVidInput`" ) --language 0:`"$subtitleLang`" --track-name 0:`"$subtitleTrackName`" `"$mkvVidSubtitle`""
            }
        }
        Start-Sleep -Seconds 1
    }
    While (!(Test-Path -Path $mkvVidTempOutput -ErrorAction SilentlyContinue)) {
        Start-Sleep 1.5
    }
    While ($True) {
        if (((Test-Lock $mkvVidInput) -eq $True) -and ((Test-Lock $mkvVidSubtitle) -eq $True ) -and ((Test-Lock $mkvVidTempOutput) -eq $True)) {
            continue
        }
        else {
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$mkvVidInput`" -Verbose"
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$mkvVidSubtitle`" -Verbose"
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $mkvVidTempOutput) -eq $True) {
            continue
        }
        else {
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
    $filebotVideoList = $vsCompletedFilesList | Where-Object { $_._vsDestPath -eq $filebotPath } | Select-Object _vsDestPath, _vsEpisodeFBPath, _vsEpisodeRaw, _vsEpisodeSubFBPath, _vsOverridePath
    foreach ($filebotFiles in $filebotVideoList) {
        $filebotVidInput = $filebotFiles._vsEpisodeFBPath
        $filebotSubInput = $filebotFiles._vsEpisodeSubFBPath
        $filebotVidBaseName = $filebotFiles._vsEpisodeRaw
        $filebotOverrideDrive = $filebotFiles._vsOverridePath
        if ($siteParentFolder.trim() -ne '' -or $siteSubFolder.trim() -ne '') {
            $FilebotRootFolder = $filebotOverrideDrive + $siteParentFolder
            $filebotParams = Join-Path -Path (Join-Path -Path $FilebotRootFolder -ChildPath $siteSubFolder) -ChildPath $filebotArgument
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($filebotVidInput). Renaming video and moving files to final folder. Using path($filebotParams)."
            Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$filebotVidInput`" -r --db TheTVDB -non-strict --format `"$filebotParams`" --apply date tags clean --log info"
            if (!($mkvMerge)) {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($filebotSubInput). Renaming subtitle and moving files to final folder. Using path($filebotParams)."
                Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$filebotSubInput`" -r --db TheTVDB -non-strict --format `"$filebotParams`" --apply date tags clean --log info"
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($filebotVidInput). ParentFolder or Subfolder path not specified. Renaming files in place."
            Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$filebotVidInput`" -r --db TheTVDB -non-strict --format `"$filebotArgument`" --apply date tags clean --log info"
            if (!($mkvMerge)) {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($filebotSubInput). Renaming subtitle and moving files to final folder. Using path($filebotParams)."
                Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$filebotSubInput`" -r --db TheTVDB -non-strict --format `"$filebotArgument`" --apply date tags clean --log info"
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
    [string]$_vsEpisodeSubtitle
    [string]$_vsEpisodeSubtitleBase
    [string]$_vsEpisodeFBPath
    [string]$_vsEpisodeSubFBPath
    [string]$_vsOverridePath
    [string]$_vsDestPathDirectory
    [string]$_vsDestPath
    [string]$_vsDestPathBase
    [bool]$_vsSECompleted
    [bool]$_vsMKVCompleted
    [bool]$_vsMoveCompleted
    [bool]$_vsFBCompleted
    [bool]$_vsErrored
    
    VideoStatus([string]$vsSite, [string]$vsSeries, [string]$vsEpisode, [string]$vsSeriesDirectory, [string]$vsEpisodeRaw, [string]$vsEpisodeTemp, [string]$vsEpisodePath, [string]$vsEpisodeSubtitle, `
            [string]$vsEpisodeSubtitleBase, [string]$vsEpisodeFBPath, [string]$vsEpisodeSubFBPath, [string]$vsOverridePath, [string]$vsDestPathDirectory, [string]$vsDestPath, [string]$vsDestPathBase, `
            [bool]$vsSECompleted, [bool]$vsMKVCompleted, [bool]$vsMoveCompleted, [bool]$vsFBCompleted, [bool]$vsErrored) {
        $this._vsSite = $vsSite
        $this._vsSeries = $vsSeries
        $this._vsEpisode = $vsEpisode
        $this._vsSeriesDirectory = $vsSeriesDirectory
        $this._vsEpisodeRaw = $vsEpisodeRaw
        $this._vsEpisodeTemp = $vsEpisodeTemp
        $this._vsEpisodePath = $vsEpisodePath
        $this._vsEpisodeSubtitle = $vsEpisodeSubtitle
        $this._vsEpisodeSubtitleBase = $vsEpisodeSubtitleBase
        $this._vsEpisodeFBPath = $vsEpisodeFBPath
        $this._vsEpisodeSubFBPath = $vsEpisodeSubFBPath
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
        $xmlConfig | Set-Content -Value $configPath
    }
    else {
        Write-Output "$configPath File Exists."
        
    }
    exit
}
if ($supportFiles) {
    New-Folder $fontFolder
    New-Folder $sharedFolder
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
            New-Folder $F
        }
        $sadf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_D_A"
        $sbdf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_D_B"
        $sbdc = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_D_C"
        $saf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_A"
        $sbf = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_B"
        $sbc = Join-Path -Path $sharedFolder -ChildPath "$($sn.SN)_C"
        $siteSuppFiles = $sadf, $sbdf, $sbdc, $saf, $sbf, $sbc
        foreach ($S in $siteSuppFiles) {
            New-SuppFiles -newSupportFiles $S
        }
        $scfDCD = $scf.TrimEnd('\') + '_D'
        $scfDC = Join-Path -Path $scfDCD -ChildPath 'yt-dlp.conf'
        $scfC = Join-Path -Path $scf -ChildPath 'yt-dlp.conf'
        $siteConfigFiles = $scfDC , $scfC
        foreach ($cf in $siteConfigFiles) {
            New-Config -newConfigs $cf
            Remove-Spaces -removeSpacesFile $cf
        }
    }
    exit
}
if ($site) {
    if (Test-Path -Path $subtitleRegex) {
        Write-Output "$(Get-Timestamp) - $dlpScript, $subtitle_regex do exist in $scriptDirectory folder."
    }
    else {
        Write-Output "$(Get-Timestamp) - subtitle_regex.py does not exist or was not found in $scriptDirectory folder. Exiting."
        Exit
    }
    $date = Get-Day
    $dateTime = Get-TimeStamp
    $time = Get-Time
    $site = $site.ToLower()
    # Reading from XML
    $configPath = Join-Path -Path $scriptDirectory -ChildPath 'config.xml'
    [xml]$configFile = Get-Content -Path $configPath
    #$siteParams = $configFile.configuration.credentials.site | Where-Object { $_.siteName.ToLower() -eq $site } | Select-Object 'siteName', 'username', 'password', 'plexlibraryid', 'parentfolder', 'subfolder', 'font' -First 1
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
            Exit-Script -Exit
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
            Exit-Script -Exit
        }
        if ((Test-Path -Path $siteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $siteConfig file found. Continuing."
            $dlpParams = $dlpParams + " --config-location $siteConfig -P temp:$siteTemp -P home:$siteSrc"
            $dlpArray += "`"--config-location`"", "`"$siteConfig`"", "`"-P`"", "`"temp:$siteTemp`"", "`"-P`"", "`"home:$siteSrc`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $siteConfig does not exist. Exiting."
            Exit-Script -Exit
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
            Exit-Script -Exit
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
                    Exit-Script -Exit
                }
            }
            else {
                $cookieFile = 'None'
                Write-Output "[Setup] $(Get-Timestamp) - Login is true and Cookies is false. Continuing."
            }
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - Login is true and Username/Password is Empty. Exiting."
            Exit-Script -Exit
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
            Exit-Script -Exit
        }
    }
    if ($ffmpeg) {
        Write-Output "[Setup] $(Get-Timestamp) - $ffmpeg file found. Continuing."
        $dlpParams = $dlpParams + " --ffmpeg-location $ffmpeg"
        $dlpArray += "`"--ffmpeg-location`"", "`"$ffmpeg`""
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - FFMPEG: $ffmpeg missing. Exiting."
        Exit-Script -Exit
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
            Exit-Script -Exit
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - BAT: $batFile missing. Exiting."
        Exit-Script -Exit
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
            Exit-Script -Exit
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
            Exit-Script -Exit
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - --remux-video parameter is missing. Exiting."
        Exit-Script -Exit
    }
    if ($subtitleEdit -or $mkvMerge) {
        if (Select-String -Path $siteConfig '--write-subs' -SimpleMatch -Quiet) {
            Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit or MKVMerge is true and --write-subs is in config. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting."
            Exit-Script -Exit
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
                Exit-Script -Exit
            }
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - --convert-subs parameter is missing. Exiting."
            Exit-Script -Exit
        }
        $writeSub = Select-String -Path $siteConfig -Pattern '--write-subs.*' | Select-Object -First 1
        if ($null -ne $writeSub) {
            Write-Output "[Setup] $(Get-Timestamp) - --write-subs is in config. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - SubSubWrite is missing. Exiting."
            Exit-Script -Exit
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit is false. Continuing."
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
        Write-Output "[END] $dateTime - Debugging enabled. Exiting."
        Exit-Script -Exit
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
        Write-Output '[START] Creating Folders:'
        $createFolders = $tempDrive, $srcDrive, $backupDrive, $srcBackup, $siteConfigBackup, $srcBackupDriveShared, $srcDriveSharedFonts, $destDrive, $siteTemp, $siteSrc, $siteHome
        foreach ($cf in $createFolders) {
            New-Folder $cf
        }
        # Log cleanup
        Remove-Logfiles
    }
    # yt-dlp
    & yt-dlp.exe $dlpArray *>&1 | Out-Host
    # Post-processing
    if ((Get-ChildItem -Path $siteSrc -Recurse -Force -File -Include "$vidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem -Path $siteSrc -Recurse -Include "$vidType" | Sort-Object LastWriteTime | Select-Object -Unique | Get-Unique | ForEach-Object {
            $vsSite = $siteNameRaw
            $vsSeries = (Get-Culture).TextInfo.ToTitleCase(("$(Split-Path (Split-Path $_ -Parent) -Leaf)").Replace('_', ' ').Replace('-', ' ')) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $vsEpisode = (Get-Culture).TextInfo.ToTitleCase( ($_.BaseName.Replace('_', ' ').Replace('-', ' '))) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $vsSeriesDirectory = $_.DirectoryName
            $vsEpisodeRaw = $_.BaseName
            $vsEpisodeTemp = Join-Path -Path $vsSeriesDirectory -ChildPath "$($vsEpisodeRaw).temp$($_.Extension)"
            $vsEpisodePath = $_.FullName
            $vsEpisodeSubtitle = (Get-ChildItem -Path $siteSrc -Recurse -File -Include "$subType" | Where-Object { $_.FullName -match $vsEpisodeRaw } | Select-Object -First 1 ).FullName
            $vsOverridePath = $overrideSeriesList | Where-Object { $_.orSeriesName.ToLower() -eq $vsSeries.ToLower() } | Select-Object -ExpandProperty orSrcdrive
            if ($null -ne $vsOverridePath) {
                $vsDestPath = Join-Path -Path $vsOverridePath -ChildPath $siteHome.Substring(3)
                $vsDestPathBase = Join-Path -Path $vsOverridePath -ChildPath $siteHomeBase.Substring(3)
                $vsEpisodeFBPath = $vsEpisodePath.Replace($siteSrc, $vsDestPath)
                $vsDestPathDirectory = Split-Path $vsEpisodeFBPath -Parent
                if ($vsEpisodeSubtitle -ne '') {
                    $vsEpisodeSubtitleBase = (Get-ChildItem -Path $siteSrc -Recurse -File -Include "$subType" | Where-Object { $_.FullName -match $vsEpisodeRaw } | Select-Object -First 1 ).Name
                    $vsEpisodeSubFBPath = $vsEpisodeSubtitle.Replace($siteSrc, $vsDestPath)
                }
                else {
                    $vsEpisodeSubtitleBase = ''
                    $vsEpisodeSubFBPath = ''
                }
            }
            else {
                $vsDestPath = $siteHome
                $vsDestPathBase = $siteHomeBase  #_vsDestPathDirectory = $vsDestPathDirectory
                $vsEpisodeFBPath = $vsEpisodePath.Replace($siteSrc, $vsDestPath)
                $vsDestPathDirectory = Split-Path $vsEpisodeFBPath -Parent
                $vsOverridePath = [System.IO.path]::GetPathRoot($vsDestPath)
                if ($vsEpisodeSubtitle -ne '') {
                    $vsEpisodeSubtitleBase = (Get-ChildItem -Path $siteSrc -Recurse -File -Include "$subType" | Where-Object { $_.FullName -match $vsEpisodeRaw } | Select-Object -First 1 ).Name
                    $vsEpisodeSubFBPath = $vsEpisodeSubtitle.Replace($siteSrc, $vsDestPath)
                }
                else {
                    $vsEpisodeSubtitleBase = ''
                    $vsEpisodeSubFBPath = ''
                }
            }
            foreach ($i in $_) {
                $VideoStatus = [VideoStatus]::new($vsSite, $vsSeries, $vsEpisode, $vsSeriesDirectory, $vsEpisodeRaw, $vsEpisodeTemp, $vsEpisodePath, $vsEpisodeSubtitle, $vsEpisodeSubtitleBase, $vsEpisodeFBPath, `
                        $vsEpisodeSubFBPath, $vsOverridePath, $vsDestPathDirectory, $vsDestPath, $vsDestPathBase, $vsSECompleted, $vsMKVCompleted, $vsMoveCompleted, $vsFBCompleted, $vsErrored)
                [void]$vsCompletedFilesList.Add($VideoStatus)
            }
        }
        $vsCompletedFilesList | Select-Object _vsEpisodeSubtitle | ForEach-Object {
            if (($_._vsEpisodeSubtitle.Trim() -eq '') -or ($null -eq $_._vsEpisodeSubtitle)) {
                Set-VideoStatus -svsKey '_vsEpisodeRaw' -svsValue $vsEpisodeRaw -svsER
            }
        }
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files to process."
    }
    $vsvTotCount = ($vsCompletedFilesList | Measure-Object).Count
    $vsvErrorCount = ($vsCompletedFilesList | Where-Object { $_._vsErrored -eq $true } | Measure-Object).Count
    Write-Output "[VideoList] $(Get-Timestamp) - Total Files: $vsvTotCount"
    Write-Output "[VideoList] $(Get-Timestamp) - Errored Files: $vsvErrorCount"
    # SubtitleEdit, MKVMerge, Filebot
    if ($vsvTotCount -gt 0) {
        # Subtitle Edit
        if ($subtitleEdit) {
            $vsCompletedFilesList | Select-Object _vsEpisodeSubtitle | Where-Object { $_._vsErrored -ne $true } | ForEach-Object {
                $seSubtitle = $_._vsEpisodeSubtitle
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
                    Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "New-Folder `"$($orDriveList._vsDestPath)`" -Verbose"
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
        $filebotOverrideDriveList = $vsCompletedFilesList | Where-Object { $_._vsMKVCompleted -eq $true -and $_._vsErrored -eq $false } | Select-Object _vsDestPath -Unique
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
                Write-Output "[FileMoving] $(Get-Timestamp) -  Moving $($mmFiles._vsDestPathDirectory) to $moveFolder."
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
            Invoke-WebRequest -Uri $plexUrl | Out-Null
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
                        Write-Output $tm
                        Invoke-Telegram -sendTelegramMessage $tm
                    }
                    else {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $siteHome. Failure."
                        $tm += 'Not all files added to PLEX.'
                        Write-Output $tm
                        Invoke-Telegram -sendTelegramMessage $tm
                    }
                }
            }
            else {
                Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $siteHome."
                $tm += 'Added files to folders.'
                Write-Output $tm
                Invoke-Telegram -sendTelegramMessage $tm
            }
        }
        Write-Output "[VideoList] $(Get-Timestamp) - Total videos downloaded: $vsvTotCount"
        $vsCompletedFilesListHeaders = @{Label = 'Series'; Expression = { $_._vsSeries } }, @{Label = 'Episode'; Expression = { $_._vsEpisode } }, `
        @{Label = 'Subtitle'; Expression = { $_._vsEpisodeSubtitleBase } }, @{Label = 'Drive'; Expression = { $_._vsOverridePath } }, @{Label = 'SrcDirectory'; Expression = { $_._vsSeriesDirectory } }, `
        @{Label = 'DestBase'; Expression = { $_._vsDestPathBase } }, @{Label = 'DestPath'; Expression = { $_._vsDestPath } }, @{Label = 'DestPathDirectory'; Expression = { $_._vsDestPathDirectory } }, `
        @{Label = 'SECompleted'; Expression = { $_._vsSECompleted } }, @{Label = 'MKVCompleted'; Expression = { $_._vsMKVCompleted } }, @{Label = 'MoveCompleted'; Expression = { $_._vsMoveCompleted } }, `
        @{Label = 'FBCompleted'; Expression = { $_._vsFBCompleted } }, @{Label = 'Errored'; Expression = { $_._vsErrored } }
        if ($vsvTotCount -gt 12) {
            $vsCompletedFilesTable = $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-Table $vsCompletedFilesListHeaders -AutoSize -Wrap
        }
        else {
            $vsCompletedFilesTable = $vsCompletedFilesList | Sort-Object _vsSeries, _vsEpisode | Format-List $vsCompletedFilesListHeaders
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