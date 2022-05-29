<#
.Synopsis
   Script to run yt-dlp, mkvmerge, subtitle edit, filebot and a python script for downloading and processing videos
.EXAMPLE
   Runs the script using the crunchyroll as a manual run with the defined config using login, cookies, archive file, mkvmerge, and sends out a telegram message
   D:\_DL\dlp-script.ps1 -sn crunchyroll -l -c -mk -a -st
.EXAMPLE
   Runs the script  using the crunchyroll as a daily run with the defined config using login, cookies, no archive file, and filebot/plex
   D:\_DL\dlp-script.ps1 -sn crunchyroll -d -l -c -f
.NOTES
   See https://github.com/wamasi/dlp-filehandler for full details
   Script was designed to be ran via powershell console on a cronjob. copying and pasting into powershell console will not work.
#>
param(
    [Parameter(Mandatory = $false)]
    [Alias('H')]
    [switch]$Help,
    [Parameter(Mandatory = $false)]
    [Alias('NC')]
    [switch]$NewConfig,
    [Parameter(Mandatory = $false)]
    [Alias('SU')]
    [switch]$SupportFiles,
    [Parameter(Mandatory = $false)]
    [Alias('SN')]
    [ValidateScript({ if (Test-Path -Path "$PSScriptRoot\config.xml") {
                if (([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName -contains $_ ) {
                    $true
                }
                else {
                    throw ([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName
                }
            }
            else {
                throw "No valid config.xml found in $PSScriptRoot. Run ($PSScriptRoot\dlp-script.ps1 -nc) for a new config file."
            }
        })]
    [string]$Site,
    [Parameter(Mandatory = $false)]
    [Alias('D')]
    [switch]$Daily,
    [Parameter(Mandatory = $false)]
    [Alias('L')]
    [switch]$Login,
    [Parameter(Mandatory = $false)]
    [Alias('C')]
    [switch]$Cookies,
    [Parameter(Mandatory = $false)]
    [Alias('A')]
    [switch]$Archive,
    [Parameter(Mandatory = $false)]
    [Alias('SE')]
    [switch]$SubtitleEdit,
    [Parameter(Mandatory = $false)]
    [Alias('MK')]
    [switch]$MKVMerge,
    [Parameter(Mandatory = $false)]
    [Alias('F')]
    [switch]$Filebot,
    [Parameter(Mandatory = $false)]
    [Alias('ST')]
    [switch]$SendTelegram,
    [Parameter(Mandatory = $false)]
    [ValidateSet('ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', IgnoreCase = $true)]
    [Alias('AL')]
    [string]$AudioLang,
    [Parameter(Mandatory = $false)]
    [ValidateSet('ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', IgnoreCase = $true)]
    [Alias('SL')]
    [string]$SubtitleLang,
    [Parameter(Mandatory = $false)]
    [Alias('T')]
    [switch]$TestScript
)
$ScriptStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
$PSStyle.OutputRendering = 'Host'
$Width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Width, 9999)
mode con cols=9999
function Get-Day {
    return (Get-Date -Format 'yy-MM-dd')
}
function Get-TimeStamp {
    return (Get-Date -Format 'yy-MM-dd HH-mm-ss')
}
function Get-Time {
    return (Get-Date -Format 'MMddHHmmss')
}
function Set-Folders {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Fullpath
    )
    if (!(Test-Path -Path $Fullpath)) {
        New-Item -ItemType Directory -Path $Fullpath -Force | Out-Null
        Write-Output "[SetFolder] - $(Get-Timestamp) - $Fullpath missing. Creating..."
    }
    else {
        Write-Output "[SetFolder] - $(Get-Timestamp) - $Fullpath already exists."
    }
}
function Set-SuppFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string] $SuppFiles
    )
    if (!(Test-Path $SuppFiles -PathType Leaf)) {
        New-Item $SuppFiles -ItemType File | Out-Null
        Write-Output "$SuppFiles file missing. Creating..."
    }
    else {
        Write-Output "$SuppFiles already file exists."
    }
}
function Resolve-Configs {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Configs
    )
    New-Item $Configs -ItemType File -Force | Out-Null
    Write-Output "Creating $Configs"
    if ($Configs -match 'vrv') {
        $vrvconfig | Set-Content $Configs
        Write-Output "$Configs created with VRV values."
    }
    elseif ($Configs -match 'crunchyroll') {
        $crunchyrollconfig | Set-Content $Configs
        Write-Output "$Configs created with Crunchyroll values."
    }
    elseif ($Configs -match 'funimation') {
        $funimationconfig | Set-Content $Configs
        Write-Output "$Configs created with Funimation values."
    }
    elseif ($Configs -match 'hidive') {
        $hidiveconfig | Set-Content $Configs
        Write-Output "$Configs created with Hidive values."
    }
    elseif ($Configs -match 'paramountplus') {
        $paramountplusconfig | Set-Content $Configs
        Write-Output "$Configs created with ParamountPlus values."
    }
    else {
        $defaultconfig | Set-Content $Configs
        Write-Output "$Configs created with default values."
    }
}
function Remove-Spaces {
    param (
        [Parameter(Mandatory = $true)]
        [string] $File
    )
    (Get-Content $File) | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } | Set-Content $File
    $content = [System.IO.File]::ReadAllText($File)
    $content = $content.Trim()
    [System.IO.File]::WriteAllText($File, $content)
}
# Setting up arraylist for MKV and Filebot lists
class VideoStatus {
    [string]$_VSSite
    [string]$_VSSeries
    [string]$_VSEpisode
    [string]$_VSSeriesDirectory
    [string]$_VSEpisodeRaw
    [string]$_VSEpisodeTemp
    [string]$_VSEpisodePath
    [string]$_VSEpisodeSubtitle
    [string]$_VSEpisodeSubtitleBase
    [string]$_VSEpisodeFBPath
    [string]$_VSEpisodeSubFBPath
    [string]$_VSOverridePath
    [string]$_VSDestPath
    [string]$_VSDestPathBase
    [bool]$_VSSECompleted
    [bool]$_VSMKVCompleted
    [bool]$_VSFBCompleted
    [bool]$_VSErrored
    
    VideoStatus([string]$VSSite, [string]$VSSeries, [string]$VSEpisode, [string]$VSSeriesDirectory, [string]$VSEpisodeRaw, [string]$VSEpisodeTemp, [string]$VSEpisodePath, [string]$VSEpisodeSubtitle, `
            [string]$VSEpisodeSubtitleBase, [string]$VSEpisodeFBPath, [string]$VSEpisodeSubFBPath, [string]$VSOverridePath, [string]$VSDestPath, [string]$VSDestPathBase, [bool]$VSSECompleted, [bool]$VSMKVCompleted, [bool]$VSFBCompleted, [bool]$VSErrored) {
        $this._VSSite = $VSSite
        $this._VSSeries = $VSSeries
        $this._VSEpisode = $VSEpisode
        $this._VSSeriesDirectory = $VSSeriesDirectory
        $this._VSEpisodeRaw = $VSEpisodeRaw
        $this._VSEpisodeTemp = $VSEpisodeTemp
        $this._VSEpisodePath = $VSEpisodePath
        $this._VSEpisodeSubtitle = $VSEpisodeSubtitle
        $this._VSEpisodeSubtitleBase = $VSEpisodeSubtitleBase
        $this._VSEpisodeFBPath = $VSEpisodeFBPath
        $this._VSEpisodeSubFBPath = $VSEpisodeSubFBPath
        $this._VSOverridePath = $VSOverridePath
        $this._VSDestPath = $VSDestPath
        $this._VSDestPathBase = $VSDestPathBase
        $this._VSSECompleted = $VSSECompleted
        $this._VSMKVCompleted = $VSMKVCompleted
        $this._VSFBCompleted = $VSFBCompleted
        $this._VSErrored = $VSErrored
    }
}
[System.Collections.ArrayList]$VSCompletedFilesList = @()
# Update SE/MKV/FB true false
function Set-VideoStatus {
    param (
        [parameter(Mandatory = $true)]
        [string]$SVSKey,
        [parameter(Mandatory = $true)]
        [string]$SVSValue,
        [parameter(Mandatory = $false)]
        [bool]$SVSSE,
        [parameter(Mandatory = $false)]
        [bool]$SVSMKV,
        [parameter(Mandatory = $false)]
        [bool]$SVSFP,
        [parameter(Mandatory = $false)]
        [bool]$SVSER
    )
    $VSCompletedFilesList | Where-Object { $_.$SVSKey -eq $SVSValue } | ForEach-Object {
        if ($SVSSE) {
            $_._VSSECompleted = $SVSSE
        }
        if ($SVSMKV) {
            $_._VSMKVCompleted = $SVSMKV
        }
        if ($SVSFP) {
            $_._VSFBCompleted = $SVSFP
        }
        if ($SVSER) {
            $_._VSErrored = $SVSER
        }
    }
}
# Getting list of Site, Series, and Episodes for Telegram messages
function Get-SiteSeriesEpisode {
    $SEL = $VSCompletedFilesList | Group-Object -Property _VSSite, _VSSeries |
    Select-Object @{n = 'Site'; e = { $_.Values[0] } }, `
    @{ n = 'Series'; e = { $_.Values[1] } }, `
    @{n = 'Episode'; e = { $_.Group | Select-Object _VSEpisode } }
    $Telegrammessage = '<b>Site:</b> ' + $SiteNameRaw + "`n"
    $SeriesMessage = ''
    $SEL | ForEach-Object {
        $EpList = ''
        foreach ($i in $_) {
            $EpList = $_.Episode._VSEpisode | Out-String
        }
        $SeriesMessage = '<b>Series:</b> ' + $_.Series + "`n<b>Episode:</b>`n" + $EpList
        $Telegrammessage += $SeriesMessage + "`n"
    }
    return $Telegrammessage
}
# Sending To telegram for new file notifications
function Send-Telegram {
    Param(
        [Parameter( Mandatory = $true)]
        [String]$STMessage)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($STMessage)&parse_mode=html" | Out-Null
}
# Run MKVMerge process
function Start-MKVMerge {
    param (
        [parameter(Mandatory = $true)]
        [string]$MKVVidInput,
        [parameter(Mandatory = $true)]
        [string]$MKVVidBaseName,
        [parameter(Mandatory = $true)]
        [string]$MKVVidSubtitle,
        [parameter(Mandatory = $true)]
        [string]$MKVVidTempOutput
    )
    # Video track title inherits from Audio language code
    switch ($AudioLang) {
        ar { $VideoLang = $AudioLang; $ALTrackName = 'Arabic Audio'; $VTrackName = 'Arabic Video' }
        de { $VideoLang = $AudioLang; $ALTrackName = 'Deutsch Audio'; $VTrackName = 'Deutsch Video' }
        en { $VideoLang = $AudioLang; $ALTrackName = 'English Audio'; $VTrackName = 'English Video' }
        es { $VideoLang = $AudioLang; $ALTrackName = 'Spanish(Latin America) Audio'; $VTrackName = 'Spanish(Latin America) Video' }
        es-es { $VideoLang = $AudioLang; $ALTrackName = 'Spanish(Spain) Audio'; $VTrackName = 'Spanish(Spain) Video' }
        fr { $VideoLang = $AudioLang; $ALTrackName = 'French Audio'; $VTrackName = 'French Video' }
        it { $VideoLang = $AudioLang; $ALTrackName = 'Italian Audio'; $VTrackName = 'Italian Video' }
        ja { $VideoLang = $AudioLang; $ALTrackName = 'Japanese Audio'; $VTrackName = 'Japanese Video' }
        pt-br { $VideoLang = $AudioLang; $ALTrackName = 'Português (Brasil) Audio'; $VTrackName = 'Português (Brasil) Video' }
        pt-pt { $VideoLang = $AudioLang; $ALTrackName = 'Português (Portugal) Audio'; $VTrackName = 'Português (Portugal) Video' }
        ru { $VideoLang = $AudioLang; $ALTrackName = 'Russian Audio'; $VTrackName = 'Russian Video' }
        Default { $VideoLang = 'und'; $AudioLang = 'und'; $ALTrackName = 'und audio'; $VTrackName = 'und Video' }
    }
    switch ($SubtitleLang) {
        ar { $STTrackName = 'Arabic Sub' }
        de { $STTrackName = 'Deutsch Sub' }
        en { $STTrackName = 'English Sub' }
        es { $STTrackName = 'Spanish(Latin America) Sub' }
        es-es { $STTrackName = 'Spanish(Spain) Sub' }
        fr { $STTrackName = 'French Sub' }
        it { $STTrackName = 'Italian Sub' }
        ja { $STTrackName = 'Japanese Sub' }
        pt-br { $STTrackName = 'Português (Brasil) Sub' }
        pt-pt { $STTrackName = 'Português (Portugal) Sub' }
        ru { $STTrackName = 'Russian Video' }
        ja { $STTrackName = 'Japanese Sub' }
        en { $STTrackName = 'English Sub' }
        Default { $SubtitleLang = 'und'; $STTrackName = 'und sub' }
    }
    Write-Output "[MKVMerge] $(Get-Timestamp) - Video = $VideoLang/$VTrackName - Audio Language = $AudioLang/$ALTrackName - Subtitle = $SubtitleLang/$STTrackName."
    Write-Output "[MKVMerge] $(Get-Timestamp) - Replacing Styling in $MKVVidSubtitle."
    While ($True) {
        if ((Test-Lock $MKVVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($SF -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - Python - Regex through $MKVVidSubtitle file with $SF."
                python $SubtitleRegex $MKVVidSubtitle $SF *>&1 | Out-Host
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - No Font specified for $MKVVidSubtitle file."
            }
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True -and (Test-Lock $MKVVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($SubFontDir -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - Combining $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput --language 0:$VideoLang --track-name 0:$VTrackName --language 1:$AudioLang --track-name 1:$ALTrackName ( $MKVVidInput ) --language 0:$SubtitleLang --track-name 0:$STTrackName ( $MKVVidSubtitle ) --attach-file $SubFontDir --attachment-mime-type application/x-truetype-font *>&1 | Out-Host
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) -  Merging as-is. No Font specified for $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput --language 0:$VideoLang --track-name 0:$VTrackName --language 1:$AudioLang --track-name 1:$ALTrackName ( $MKVVidInput ) --language 0:$SubtitleLang --track-name 0:$STTrackName ( $MKVVidSubtitle ) *>&1 | Out-Host
            }
        }
        Start-Sleep -Seconds 1
    }
    While (!(Test-Path $MKVVidTempOutput -ErrorAction SilentlyContinue)) {
        Start-Sleep 1.5
    }
    While ($True) {
        if (((Test-Lock $MKVVidInput) -eq $True) -and ((Test-Lock $MKVVidSubtitle) -eq $True ) -and ((Test-Lock $MKVVidTempOutput) -eq $True)) {
            continue
        }
        else {
            Remove-Item -Path $MKVVidInput -Confirm:$false -Verbose
            Remove-Item -Path $MKVVidSubtitle -Confirm:$false -Verbose
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidTempOutput) -eq $True) {
            continue
        }
        else {
            Rename-Item -Path $MKVVidTempOutput -NewName $MKVVidInput -Confirm:$false -Verbose
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True) {
            continue
        }
        else {
            mkvpropedit $MKVVidInput --edit track:s1 --set flag-default=1 *>&1 | Out-Host
            break
        }
        Start-Sleep -Seconds 1
    }
    Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $MKVVidBaseName -SVSMKV $true
}
# Function to process video files through FileBot
function Start-Filebot {
    param (
        [parameter(Mandatory = $true)]
        [string]$FBPath
    )
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to rename and move to final folder."
    $FBVideoList = $VSCompletedFilesList | Where-Object { $_._VSDestPath -eq $FBPath } | Select-Object _VSDestPath, _VSEpisodeFBPath, _VSEpisodeRaw, _VSEpisodeSubFBPath, _VSOverridePath
    foreach ($FBFiles in $FBVideoList) {
        $FBVidInput = $FBFiles._VSEpisodeFBPath
        $FBSubInput = $FBFiles._VSEpisodeSubFBPath
        $FBVidBaseName = $FBFiles._VSEpisodeRaw
        $FBOverrideDrive = $FBFiles._VSOverridePath
        if ($PlexLibPath) {
            $FBParams = $FBOverrideDrive + "$FBBaseFolder\$PlexLibPath\$FBArgument"
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBVidInput). Renaming video and moving files to final folder. Using path($FBParams)."
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "$FBParams" --log info *>&1 | Out-Host
            if (!($MKVMerge)) {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBSubInput). Renaming subtitle and moving files to final folder. Using path($FBParams)."
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format "$FBParams" --log info *>&1 | Out-Host
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBVidInput). Plex path not specified. Renaming files in place."
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "$FBArgument" --log info *>&1 | Out-Host
            if (!($MKVMerge)) {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBSubInput). Renaming subtitle and moving files to final folder. Using path($FBParams)."
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format "$FBArgument" --log info *>&1 | Out-Host
            }
        }
        if (!(Test-Path $FBVidInput)) {
            Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $FBVidBaseName -SVSFP $true
        }
    }
    $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
    if ($VSVFBCount -eq $VSVTotCount ) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($VSVFBCount) = ($VSVTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup."
        filebot -script fn:cleaner "$SiteHome" --log info
    }
    else {
        Write-Output "[Filebot] $(Get-Timestamp) - Filebot($VSVFBCount) and Total Video($VSVTotCount) count mismatch. Manual check required."
    }
    if ($VSVFBCount -ne $VSVTotCount) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
    }
}
# Delete Tmp/Src/Home folder logic
function Remove-Folders {
    param (
        [parameter(Mandatory = $true)]
        [string]$RFFolder,
        [parameter(Mandatory = $false)]
        [string]$RFMatch,
        [parameter(Mandatory = $true)]
        [string]$RFBaseMatch
    )
    if ($RFFolder -eq $SiteTemp) {
        if (($RFFolder -match '\\tmp\\') -and ($RFFolder -match $RFBaseMatch) -and (Test-Path $RFFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $RFFolder folders/files."
            Remove-Item $RFFolder -Recurse -Force -Confirm:$false -Verbose | Out-Null
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteTemp($RFFolder) folder already deleted. Nothing to remove."
        }
    }
    else {
        if (!(Test-Path $RFFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) already deleted."
        }
        elseif ((Test-Path $RFFolder) -and (Get-ChildItem $RFFolder -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) is empty. Deleting folder."
            & $DeleteRecursion -DRPath $RFFolder
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) contains files. Manual attention needed."
        }
    }
}
# Test if file is available to interact with
function Test-Lock {
    Param(
        [parameter(Mandatory = $true)]
        $TLfilename
    )
    $TLfile = Get-Item (Resolve-Path $TLfilename) -Force
    if ($TLfile -is [IO.FileInfo]) {
        trap {
            Write-Output "[FileLockCheck] $(Get-Timestamp) - $TLfile File locked. Waiting..."
            return $true
            continue
        }
        $TLstream = New-Object system.IO.StreamReader $TLfile
        if ($TLstream) { $TLstream.Close() }
    }
    Write-Output "[FileLockCheck] $(Get-Timestamp) - $TLfile File unlocked. Continuing..."
    return $false
}
function Remove-Logfiles {
    # Log cleanup
    $FilledLogslimit = (Get-Date).AddDays(-$FilledLogs)
    $EmptyLogslimit = (Get-Date).AddDays(-$EmptyLogs)
    if (!(Test-Path $LFolderBase)) {
        Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase is missing. Skipping log cleanup..."
    }
    else {
        Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting Filledlog($FilledLogs) cleanup..."
        Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -match '.*-Total-.*' -and $_.FullName -ne $LFile -and $_.CreationTime -lt $FilledLogslimit } | `
            ForEach-Object {
            $_.FullName | Remove-Item -Recurse -Force -Confirm:$false -Verbose
        }
        Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting emptylog($EmptyLogs) cleanup..."
        Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -notmatch '.*-Total-.*' -and $_.FullName -ne $LFile -and $_.CreationTime -lt $EmptyLogslimit } | `
            ForEach-Object {
            $_.FullName | Remove-Item -Recurse -Force -Confirm:$false -Verbose
        }
        & $DeleteRecursion -DRPath $LFolderBase
    }
}
function Exit-Script {
    param (
        [alias('ES')]
        [switch]$ExitScript
    )
    $ScriptStopWatch.Stop()
    Write-Output "[END] $(Get-Timestamp) - Script completed. Total Elapsed Time: $($ScriptStopWatch.Elapsed.ToString())"
    Stop-Transcript
    ((Get-Content $LFile | Select-Object -Skip 5) | Select-Object -SkipLast 4) | Set-Content $LFile
    Remove-Spaces $LFile
    # Cleanup folders
    Remove-Folders -RFFolder $SiteTemp -RFMatch '\\tmp\\' -RFBaseMatch $SiteTempBaseMatch
    Remove-Folders -RFFolder $SiteSrc -RFMatch '\\src\\' -RFBaseMatch $SiteSrcBaseMatch
    Remove-Folders -RFFolder $SiteHome -RFMatch '\\tmp\\' -RFBaseMatch $SiteHomeBaseMatch
    if ($OverrideDriveList.count -gt 0) {
        foreach ($ORDriveList in $OverrideDriveList) {
            $ORDriveListBaseMatch = ($ORDriveList._VSDestPathBase).Replace('\', '\\')
            Remove-Folders -RFFolder $ORDriveList._VSDestPath -RFMatch '\\tmp\\' -RFBaseMatch $ORDriveListBaseMatch
        }
    }
    # Cleanup Log Files
    Remove-Logfiles
    if ($ExitScript) {
        Rename-Item -Path $LFile -NewName "$DateTime-DEBUG.log"
        exit
    }
    else {
        if ($VSVTotCount -gt 0) {
            Rename-Item -Path $LFile -NewName "$DateTime-Total-$VSVTotCount.log"
        }
    }
}
$DeleteRecursion = {
    param(
        $DRPath
    )
    foreach ($DRchildDirectory in Get-ChildItem -Force -LiteralPath $DRPath -Directory) {
        & $DeleteRecursion -DRPath $DRchildDirectory.FullName
    }
    $DRcurrentChildren = Get-ChildItem -Force -LiteralPath $DRPath
    $DRisEmpty = $DRcurrentChildren -eq $null
    if ($DRisEmpty) {
        Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting '${DRPath}' folders/files if empty."
        Remove-Item -Force -LiteralPath $DRPath -Verbose
    }
}
$ScriptDirectory = $PSScriptRoot
$DLPScript = "$ScriptDirectory\dlp-script.ps1"
$SubtitleRegex = "$ScriptDirectory\subtitle_regex.py"
$ConfigPath = "$ScriptDirectory\config.xml"
$SharedF = "$ScriptDirectory\shared"
$FontFolder = "$ScriptDirectory\fonts"
$xmlconfig = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <Directory>
        <backup location="" />
        <temp location="" />
        <src location="" />
        <dest location="" />
        <ffmpeg location="" />
    </Directory>
    <Logs>
        <keeplog emptylogskeepdays="0" filledlogskeepdays="7" />
    </Logs>
    <Plex>
        <plexcred plexUrl="" plexToken="" />
        <library libraryid="" folder="" />
        <library libraryid="" folder="" />
        <library libraryid="" folder="" />
    </Plex>
    <Filebot>
        <fbfolder fbFolderName="Videos" fbargument="{ plex.tail }" />
    </Filebot>
    <OverrideSeries>
        <override orSeriesName="" orSrcdrive="" />
        <override orSeriesName="" orSrcdrive="" />
        <override orSeriesName="" orSrcdrive="" />
    </OverrideSeries>
    <Telegram>
        <token tokenId="" chatid="" />
    </Telegram>
    <credentials>
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
        <site siteName="" username="" password="" libraryid="" font="" />
    </credentials>
</configuration>
'@
$defaultconfig = @'
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
$vrvconfig = @'
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
$crunchyrollconfig = @'
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
$funimationconfig = @'
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
-N 32
--downloader aria2c
--extractor-args 'funimation:language=japanese'
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$hidiveconfig = @'
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
--sleep-requests "2"
--sleep-subtitles "3"
--min-sleep-interval "60"
--max-sleep-interval "120"
--retries "10"
--file-access-retries "10"
--fragment-retries "10"
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$paramountplusconfig = @'
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
if (!(Test-Path "$ScriptDirectory\config.xml" -PathType Leaf)) {
    New-Item "$ScriptDirectory\config.xml" -ItemType File -Force
}
if ($help) {
    Show-Markdown -Path "$ScriptDirectory\README.md" -UseBrowser
    exit
}
if ($NewConfig) {
    if (!(Test-Path $ConfigPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-Content $ConfigPath))) {
        
        New-Item $ConfigPath -ItemType File -Force
        Write-Output "$ConfigPath File Created successfully."
            
        $xmlconfig | Set-Content $ConfigPath
    }
    else {
        Write-Output "$ConfigPath File Exists."
        
    }
}
if ($SupportFiles) {
    Set-Folders $FontFolder
    $ConfigPath = "$ScriptDirectory\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SNfile = $ConfigFile.configuration.credentials.site | Where-Object { $_.siteName.trim() -ne '' } | Select-Object 'siteName' -ExpandProperty siteName
    $SNfile | ForEach-Object {
        $SN = New-Object -Type PSObject -Property @{
            SN = $_.siteName
        }
        $SharedF = "$ScriptDirectory\shared"
        Set-Folders $SharedF
        $SCC = "$ScriptDirectory\sites"
        Set-Folders $SCC
        $SCF = "$SCC\" + $SN.SN
        Set-Folders $SCF
        $SCDF = "$SCF" + '_D'
        Set-Folders $SCDF
        $SADF = "$SharedF\" + $SN.SN + '_D_A'
        Set-SuppFiles $SADF
        $SBDF = "$SharedF\" + $SN.SN + '_D_B'
        Set-SuppFiles $SBDF
        $SBDC = "$SharedF\" + $SN.SN + '_D_C'
        Set-SuppFiles $SBDC
        $SAF = "$SharedF\" + $SN.SN + '_A'
        Set-SuppFiles $SAF
        $SBF = "$SharedF\" + $SN.SN + '_B'
        Set-SuppFiles $SBF
        $SBC = "$SharedF\" + $SN.SN + '_C'
        Set-SuppFiles $SBC
        $SCFDC = "$SCF" + '_D\yt-dlp.conf'
        Resolve-Configs $SCFDC
        Remove-Spaces $SCFDC
        $SCFC = "$SCF\yt-dlp.conf"
        Resolve-Configs $SCFC
        Remove-Spaces $SCFC
    }
}
if ($Site) {
    if (Test-Path -Path $SubtitleRegex) {
        Write-Output "$(Get-Timestamp) - $DLPScript, $subtitle_regex do exist in $ScriptDirectory folder."
    }
    else {
        Write-Output "$(Get-Timestamp) - subtitle_regex.py does not exist or was not found in $ScriptDirectory folder. Exiting..."
        Exit
    }
    $Date = Get-Day
    $DateTime = Get-TimeStamp
    $Time = Get-Time
    $site = $site.ToLower()
    # Reading from XML
    $ConfigPath = "$ScriptDirectory\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SiteParams = $ConfigFile.configuration.credentials.site | Where-Object { $_.siteName.ToLower() -eq $site } | Select-Object 'siteName', 'username', 'password', 'libraryid', 'font' -First 1
    $SiteName = $SiteParams.siteName.ToLower()
    $SiteNameRaw = $SiteParams.siteName
    $SiteFolder = "$ScriptDirectory\sites\"
    if ($Daily) {
        $SiteType = $SiteName + '_D'
        $SiteFolder = "$SiteFolder" + $SiteType
        $LFolderBase = "$SiteFolder\log\"
        $LFile = "$SiteFolder\log\$Date\$DateTime.log"
        Start-Transcript -Path $LFile -UseMinimalHeader
        Write-Output "[Setup] $(Get-Timestamp) - $SiteNameRaw"
    }
    else {
        $SiteType = $SiteName
        $SiteFolder = $SiteFolder + $SiteType
        $LFolderBase = "$SiteFolder\log\"
        $LFile = "$SiteFolder\log\$Date\$DateTime.log"
        Start-Transcript -Path $LFile -UseMinimalHeader
        Write-Output "[Setup] $(Get-Timestamp) - $SiteNameRaw"
    }
    $SiteUser = $SiteParams.username
    $SitePass = $SiteParams.password
    $SiteLib = $SiteParams.libraryid
    $SubFont = $SiteParams.font
    $BackupDrive = $ConfigFile.configuration.Directory.backup.location
    $TempDrive = $ConfigFile.configuration.Directory.temp.location
    $SrcDrive = $ConfigFile.configuration.Directory.src.location
    $DestDrive = $ConfigFile.configuration.Directory.dest.location
    $Ffmpeg = $ConfigFile.configuration.Directory.ffmpeg.location
    [int]$EmptyLogs = $ConfigFile.configuration.Logs.keeplog.emptylogskeepdays
    [int]$FilledLogs = $ConfigFile.configuration.Logs.keeplog.filledlogskeepdays
    $PlexHost = $ConfigFile.configuration.Plex.plexcred.plexUrl
    $PlexToken = $ConfigFile.configuration.Plex.plexcred.plexToken
    $PlexLibrary = $ConfigFile.configuration.plex.library | Where-Object { $_.libraryid -eq $SiteLib } | Select-Object libraryid, folder
    $PlexLibId = $PlexLibrary.libraryid
    $PlexLibPath = $PlexLibrary.folder
    $FBBaseFolder = $ConfigFile.configuration.Filebot.fbfolder.fbFolderName
    $FBArgument = $ConfigFile.configuration.Filebot.fbfolder.fbArgument
    $OverrideSeriesList = $ConfigFile.configuration.OverrideSeries.override | Where-Object { $_.orSeriesId -ne '' -and $_.orSrcdrive -ne '' }
    $Telegramtoken = $ConfigFile.configuration.Telegram.token.tokenId
    $Telegramchatid = $ConfigFile.configuration.Telegram.token.chatid
    # End reading from XML
    if ($SubFont.Trim() -ne '') {
        $SubFontDir = "$FontFolder\$Subfont"
        if (Test-Path $SubFontDir) {
            $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
            Write-Output "$(Get-Timestamp) - $SubFont set for $SiteName."
        }
        else {
            Write-Output "$(Get-Timestamp) - $SubFont specified in $ConfigFile is missing from $FontFolder. Exiting..."
            Exit-Script -es
        }
    }
    else {
        $SubFont = 'None'
        $SubFontDir = 'None'
        $SF = 'None'
        Write-Output "$(Get-Timestamp) - $SubFont - No font set for $SiteName."
    }
    $SiteShared = "$ScriptDirectory\shared\"
    $SrcBackup = "$BackupDrive\_Backup\"
    $SrcDriveShared = "$SrcBackup" + 'shared\'
    $SrcDriveSharedFonts = "$SrcBackup" + 'fonts\'
    $dlpParams = 'yt-dlp'
    $dlpArray = @()
    if ($Daily) {
        $SiteTempBase = "$TempDrive\" + $SiteName.Substring(0, 1)
        $SiteTempBaseMatch = $SiteTempBase.Replace('\', '\\')
        $SiteTemp = "$SiteTempBase\$Time"
        $SiteSrcBase = "$SrcDrive\" + $SiteName.Substring(0, 1)
        $SiteSrcBaseMatch = $SiteSrcBase.Replace('\', '\\')
        $SiteSrc = "$SiteSrcBase\$Time"
        $SiteHomeBase = "$DestDrive\_" + $PlexLibPath + '\' + ($SiteName).Substring(0, 1)
        $SiteHomeBaseMatch = $SiteHomeBase.Replace('\', '\\')
        $SiteHome = "$SiteHomeBase\$Time"
        $SiteConfig = $SiteFolder + '\yt-dlp.conf'
        if ($SrcDrive -eq $TempDrive) {
            Write-Output "[Setup] $(Get-Timestamp) - Src($SrcDrive) and Temp($TempDrive) Directories cannot be the same"
            Exit-Script -es
        }
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig file found. Continuing..."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
            $dlpArray += "`"--config-location`"", "`"$SiteConfig`"", "`"-P`"", "`"temp:$SiteTemp`"", "`"-P`"", "`"home:$SiteSrc`""
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist. Exiting..."
            Exit-Script -es
        }
    }
    else {
        $SiteTempBase = "$TempDrive\" + $SiteName.Substring(0, 1) + 'M'
        $SiteTempBaseMatch = $SiteTempBase.Replace('\', '\\')
        $SiteTemp = "$SiteTempBase\$Time"
        $SiteSrcBase = "$SrcDrive\" + $SiteName.Substring(0, 1) + 'M'
        $SiteSrcBaseMatch = $SiteSrcBase.Replace('\', '\\')
        $SiteSrc = "$SiteSrcBase\$Time"
        $SiteHomeBase = "$DestDrive\_M\" + $SiteName.Substring(0, 1)
        $SiteHomeBaseMatch = $SiteHomeBase.Replace('\', '\\')
        $SiteHome = "$SiteHomeBase\$Time"
        $SiteConfig = $SiteFolder + '\yt-dlp.conf'
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig file found. Continuing..."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
            $dlpArray += "`"--config-location`"", "`"$SiteConfig`"", "`"-P`"", "`"temp:$SiteTemp`"", "`"-P`"", "`"home:$SiteSrc`""
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist. Exiting..."
            Exit-Script -es
        }
    }
    $SiteConfigBackup = $SrcBackup + "sites\$SiteType\"
    $CookieFile = $SiteShared + $SiteType + '_C'
    if ($Login) {
        if ($SiteUser -and $SitePass) {
            Write-Output "$(Get-Timestamp) - Login is true and SiteUser/Password is filled. Continuing..."
            $dlpParams = $dlpParams + " -u $SiteUser -p $SitePass"
            $dlpArray += "`"-u`"", "`"$SiteUser`"", "`"-p`"", "`"$SitePass`""
            if ($Cookies) {
                if ((Test-Path -Path $CookieFile)) {
                    Write-Output "$(Get-Timestamp) - Cookies is true and $CookieFile file found. Continuing..."
                    $dlpParams = $dlpParams + " --cookies $CookieFile"
                    $dlpArray += "`"--cookies`"", "`"$CookieFile`""
                }
                else {
                    Write-Output "$(Get-Timestamp) - $CookieFile does not exist. Exiting..."
                    Exit-Script -es
                }
            }
            else {
                $CookieFile = 'None'
                Write-Output "$(Get-Timestamp) - Login is true and Cookies is false. Continuing..."
            }
        }
        else {
            Write-Output "$(Get-Timestamp) - Login is true and Username/Password is Empty. Exiting..."
            Exit-Script -es
        }
    }
    else {
        if ((Test-Path -Path $CookieFile)) {
            Write-Output "$(Get-Timestamp) - $CookieFile file found. Continuing..."
            $dlpParams = $dlpParams + " --cookies $CookieFile"
            $dlpArray += "`"--cookies`"", "`"$CookieFile`""
        }
        else {
            Write-Output "$(Get-Timestamp) - $CookieFile does not exist. Exiting..."
            Exit-Script -es
        }
    }
    if ($Ffmpeg) {
        Write-Output "$(Get-Timestamp) - $Ffmpeg file found. Continuing..."
        $dlpParams = $dlpParams + " --ffmpeg-location $Ffmpeg"
        $dlpArray += "`"--ffmpeg-location`"", "`"$Ffmpeg`""
    }
    else {
        Write-Output "$(Get-Timestamp) - FFMPEG: $Ffmpeg missing. Exiting..."
        Exit-Script -es
    }
    $BatFile = "$SiteShared" + $SiteType + '_B'
    if ((Test-Path -Path $BatFile)) {
        Write-Output "$(Get-Timestamp) - $BatFile file found. Continuing..."
        if (![String]::IsNullOrWhiteSpace((Get-Content $BatFile))) {
            Write-Output "$(Get-Timestamp) - $BatFile not empty. Continuing..."
            $dlpParams = $dlpParams + " -a $BatFile"
            $dlpArray += "`"-a`"", "`"$BatFile`""
        }
        else {
            Write-Output "$(Get-Timestamp) - $BatFile is empty. Exiting..."
            Exit-Script -es
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - BAT: $Batfile missing. Exiting..."
        Exit-Script -es
    }
    if ($Archive) {
        $ArchiveFile = "$SiteShared" + $SiteType + '_A'
        if ((Test-Path -Path $ArchiveFile)) {
            Write-Output "$(Get-Timestamp) - $ArchiveFile file found. Continuing..."
            $dlpParams = $dlpParams + " --download-archive $ArchiveFile"
            $dlpArray += "`"--download-archive`"", "`"$ArchiveFile`""
        }
        else {
            Write-Output "$(Get-Timestamp) - Archive file missing. Exiting..."
            Exit-Script -es
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - Using --no-download-archive. Continuing..."
        $ArchiveFile = 'None'
        $dlpParams = $dlpParams + ' --no-download-archive'
        $dlpArray += "`"--no-download-archive`""
    }
    if ($SubtitleEdit -or $MKVMerge) {
        if (Select-String -Path $SiteConfig '--write-subs' -SimpleMatch -Quiet) {
            Write-Output "$(Get-Timestamp) - SubtitleEdit or MKVMerge is true and --write-subs is in config. Continuing..."
        }
        else {
            Write-Output "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting..."
            Exit-Script -es
        }
        $SType = Select-String -Path $SiteConfig -Pattern '--convert-subs.*' | Select-Object -First 1
        if ($null -ne $SType) {
            $SubType = '*.' + ($SType -split ' ')[1]
            $SubType = $SubType.Replace("'", '').Replace('"', '')
            if ($SubType -eq '*.ass') {
                Write-Output "$(Get-Timestamp) - Using $SubType. Continuing..."
            }
            else {
                Write-Output "$(Get-Timestamp) - Subtype(ass) is missing. Exiting..."
                Exit-Script -es
            }
        }
        else {
            Write-Output "$(Get-Timestamp) - --convert-subs parameter is missing. Exiting..."
            Exit-Script -es
        }
        $Vtype = Select-String -Path $SiteConfig -Pattern '--remux-video.*' | Select-Object -First 1
        if ($null -ne $Vtype) {
            $VidType = '*.' + ($Vtype -split ' ')[1]
            $VidType = $VidType.Replace("'", '').Replace('"', '')
            if ($VidType -eq '*.mkv') {
                Write-Output "$(Get-Timestamp) - Using $VidType. Continuing..."
            }
            else {
                Write-Output "$(Get-Timestamp) - VidType(mkv) is missing. Exiting..."
                Exit-Script -es
            }
        }
        else {
            Write-Output "$(Get-Timestamp) - --remux-video parameter is missing. Exiting..."
            Exit-Script -es
        }
        $Wsub = Select-String -Path $SiteConfig -Pattern '--write-subs.*' | Select-Object -First 1
        if ($null -ne $Wsub) {
            Write-Output "$(Get-Timestamp) - --write-subs is in config. Continuing..."
        }
        else {
            Write-Output "$(Get-Timestamp) - SubSubWrite is missing. Exiting..."
            Exit-Script -es
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - SubtitleEdit is false. Continuing..."
    }
    $DebugVars = [ordered]@{Site = $SiteName; isDaily = $Daily; UseLogin = $Login; UseCookies = $Cookies; UseArchive = $Archive; SubtitleEdit = $SubtitleEdit; `
            MKVMerge = $MKVMerge; AudioLang = $AudioLang; SubtitleLang = $SubtitleLang; Filebot = $Filebot; SiteNameRaw = $SiteNameRaw; SiteType = $SiteType; SiteUser = $SiteUser; SitePass = $SitePass; `
            SiteFolder = $SiteFolder; SiteTemp = $SiteTemp; SiteTempBaseMatch = $SiteTempBaseMatch; SiteSrc = $SiteSrc; SiteSrcBase = $SiteSrcBase; `
            SiteSrcBaseMatch = $SiteSrcBaseMatch; SiteHome = $SiteHome; SiteHomeBase = $SiteHomeBase; SiteHomeBaseMatch = $SiteHomeBaseMatch; `
            SiteConfig = $SiteConfig; CookieFile = $CookieFile; Archive = $ArchiveFile; Bat = $BatFile; Ffmpeg = $Ffmpeg; SF = $SF; SubFont = $SubFont; `
            SubFontDir = $SubFontDir; SubType = $SubType; VidType = $VidType; Backup = $SrcBackup; BackupShared = $SrcDriveShared; BackupFont = $SrcDriveSharedFonts; `
            SiteConfigBackup = $SiteConfigBackup; PlexHost = $PlexHost; PlexToken = $PlexToken; PlexLibPath = $PlexLibPath; PlexLibId = $PlexLibId; `
            TelegramToken = $TelegramToken; TelegramChatId = $TelegramChatId; ConfigPath = $ConfigPath; ScriptDirectory = $ScriptDirectory; dlpParams = $dlpParams
    }
    if ($TestScript) {
        Write-Output "[START] $DateTime - $SiteNameRaw - DEBUG Run"
        $DebugVars
        $OverrideSeriesList
        Write-Output 'dlpArray:'
        $dlpArray
        Write-Output "[END] $DateTime - Debugging enabled. Exiting..."
        Exit-Script -es
    }
    else {
        $DebugVarRemove = 'SitePass', 'PlexToken', 'TelegramToken', 'TelegramChatId'
        foreach ($dbv in $DebugVarRemove) {
            $DebugVars.Remove($dbv)
        }
        if ($Daily) {
            Write-Output "[START] $DateTime - $SiteNameRaw - Daily Run"
        }
        else {
            Write-Output "[START] $DateTime - $SiteNameRaw - Manual Run"
        }
        Write-Output 'Debug Vars:'
        $DebugVars
        Write-Output 'Series Drive Overrides:'
        $OverrideSeriesList
        Write-Output 'dlpArray:'
        $dlpArray
        # Create folders
        $CreateFolders = $TempDrive, $SrcDrive, $BackupDrive, $SrcBackup, $SiteConfigBackup, $SrcDriveShared, $SrcDriveSharedFonts, $DestDrive, $SiteTemp, $SiteSrc, $SiteHome
        foreach ($c in $CreateFolders) {
            Set-Folders $c
        }
        # Log cleanup
        Remove-Logfiles
    }
    # yt-dlp
    & yt-dlp.exe $dlpArray *>&1 | Out-Host
    # Post-processing
    if ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem $SiteSrc -Recurse -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -Unique | ForEach-Object {
            $VSSite = $SiteNameRaw
            $VSSeries = (Get-Culture).TextInfo.ToTitleCase(("$(Split-Path (Split-Path $_ -Parent) -Leaf)").Replace('_', ' ').Replace('-', ' ')) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $VSEpisode = (Get-Culture).TextInfo.ToTitleCase( ($_.BaseName.Replace('_', ' ').Replace('-', ' '))) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $VSSeriesDirectory = $_.DirectoryName
            $VSEpisodeRaw = $_.BaseName
            $VSEpisodeTemp = $VSSeriesDirectory + "\$VSEpisodeRaw.temp" + $_.Extension
            $VSEpisodePath = $_.FullName
            $VSEpisodeSubtitle = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).FullName
            $VSOverridePath = $OverrideSeriesList | Where-Object { $_.orSeriesName.ToLower() -eq $VSSeries.ToLower() } | Select-Object -ExpandProperty orSrcdrive
            Write-Host "Test $VSOverridePath"
            if ($null -ne $VSOverridePath) {
                Write-Host "[Test] - WITH OVERRIDE $VSOverridePath"
                $VSDestPath = $VSOverridePath + $SiteHome.Substring(3)
                $VSDestPathBase = $VSOverridePath + $SiteHomeBase.Substring(3)
                $VSEpisodeFBPath = $VSEpisodePath.Replace($SiteSrc, $VSDestPath)
                if ($VSEpisodeSubtitle -ne '') {
                    $VSEpisodeSubtitleBase = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).Name
                    $VSEpisodeSubFBPath = $VSEpisodeSubtitle.Replace($SiteSrc, $VSDestPath)
                }
                else {
                    $VSEpisodeSubtitleBase = ''
                    $VSEpisodeSubFBPath = ''
                }
            }
            else {
                $VSDestPath = $SiteHome
                $VSDestPathBase = $SiteHomeBase
                $VSEpisodeFBPath = $VSEpisodePath.Replace($SiteSrc, $VSDestPath)
                $VSOverridePath = [System.IO.path]::GetPathRoot($VSDestPath)
                Write-Host "[Test] - NO OVERRIDE $VSOverridePath"
                if ($VSEpisodeSubtitle -ne '') {
                    $VSEpisodeSubtitleBase = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).Name
                    $VSEpisodeSubFBPath = $VSEpisodeSubtitle.Replace($SiteSrc, $VSDestPath)
                }
                else {
                    $VSEpisodeSubtitleBase = ''
                    $VSEpisodeSubFBPath = ''
                }
            }
            foreach ($i in $_) {
                $VideoStatus = [VideoStatus]::new($VSSite, $VSSeries, $VSEpisode, $VSSeriesDirectory, $VSEpisodeRaw, $VSEpisodeTemp, $VSEpisodePath, $VSEpisodeSubtitle, $VSEpisodeSubtitleBase, $VSEpisodeFBPath, `
                        $VSEpisodeSubFBPath, $VSOverridePath, $VSDestPath, $VSDestPathBase, $VSSECompleted, $VSMKVCompleted, $VSFBCompleted, $VSErrored)
                [void]$VSCompletedFilesList.Add($VideoStatus)
            }
        }
        $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | ForEach-Object {
            if (($_._VSEpisodeSubtitle.Trim() -eq '') -or ($null -eq $_._VSEpisodeSubtitle)) {
                Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $VSEpisodeRaw -SVSER $true
            }
        }
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files to process."
    }
    $VSCompletedFilesList
    $VSVTotCount = ($VSCompletedFilesList | Measure-Object).Count
    $VSVErrorCount = ($VSCompletedFilesList | Where-Object { $_._VSErrored -eq $true } | Measure-Object).Count
    Write-Output "[VideoList] $(Get-Timestamp) - Total Files: $VSVTotCount"
    Write-Output "[VideoList] $(Get-Timestamp) - Errored Files: $VSVErrorCount"
    # SubtitleEdit, MKVMerge, Filebot
    if ($VSVTotCount -gt 0) {
        if ($SubtitleEdit) {
            $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | Where-Object { $_._VSErrored -ne $true } | ForEach-Object {
                $SESubtitle = $_._VSEpisodeSubtitle
                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $SESubtitle subtitle."
                While ($True) {
                    if ((Test-Lock $SESubtitle) -eq $True) {
                        continue
                    }
                    else {
                        $outSE = Invoke-Command -ScriptBlock { powershell "SubtitleEdit /convert '$SESubtitle' AdvancedSubStationAlpha /overwrite /MergeSameTimeCodes" }
                        Write-Output $outSE
                        Set-VideoStatus -SVSKey '_VSEpisodeSubtitle' -SVSValue $SESubtitle -SVSSE $true
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Not running."
        }
        if ($MKVMerge) {
            $VSCompletedFilesList | Select-Object _VSEpisodeRaw, _VSEpisode, _VSEpisodeTemp, _VSEpisodePath, _VSEpisodeSubtitle, _VSErrored | `
                Where-Object { $_._VSErrored -eq $false } | ForEach-Object {
                $MKVVidInput = $_._VSEpisodePath
                $MKVVidBaseName = $_._VSEpisodeRaw
                $MKVVidSubtitle = $_._VSEpisodeSubtitle
                $MKVVidTempOutput = $_._VSEpisodeTemp
                Start-MKVMerge $MKVVidInput $MKVVidBaseName $MKVVidSubtitle $MKVVidTempOutput
            }
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - MKVMerge not running. Moving to next step."
        }
        $OverrideDriveList = $VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Select-Object _VSSeriesDirectory, _VSDestPath, _VSDestPathBase -Unique
        $OverrideDriveList
        $VSVMKVCount = ($VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Measure-Object).Count
        # FileMoving
        if (($VSVMKVCount -eq $VSVTotCount -and $VSVErrorCount -eq 0) -or (!($MKVMerge) -and $VSVErrorCount -eq 0)) {
            Write-Output "[FileMoving] $(Get-Timestamp)- All files had matching subtitle file"
            foreach ($ORDriveList in $OverrideDriveList) {
                Write-Output "[FileMoving] $(Get-Timestamp) - $($ORDriveList._VSSeriesDirectory) contains files. Moving to $($ORDriveList._VSDestPath)..."
                if (!(Test-Path $ORDriveList._VSDestPath)) {
                    Set-Folders $ORDriveList._VSDestPath
                }
                Move-Item -Path $ORDriveList._VSSeriesDirectory -Destination $ORDriveList._VSDestPath -Force -Verbose
            }
        }
        else {
            Write-Output "[FileMoving] $(Get-Timestamp) - $SiteSrc contains file(s) with error(s). Not moving files."
        }
        # Filebot
        if (($Filebot -and $VSVMKVCount -eq $VSVTotCount) -or ($Filebot -and !($MKVMerge))) {
            foreach ($ORDriveList in $OverrideDriveList) {
                Write-Output "[Filebot] $(Get-Timestamp) - Renaming files in $($ORDriveList._VSDestPath)."
                Start-Filebot -FBPath $ORDriveList._VSDestPath
            }
        }
        elseif (($Filebot -and $MKVMerge -and $VSVMKVCount -ne $VSVTotCount)) {
            Write-Output "[Filebot] $(Get-Timestamp) - Files in $SiteSrc need manual attention. Skipping to next step... Incomplete files in $SiteSrc."
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Not running Filebot."
        }
        # Plex
        if ($PlexHost -and $PlexToken -and $PlexLibId) {
            Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
            $PlexUrl = "$PlexHost/library/sections/$PlexLibId/refresh?X-Plex-Token=$PlexToken"
            Invoke-WebRequest -Uri $PlexUrl | Out-Null
        }
        else {
            Write-Output "[PLEX] $(Get-Timestamp) - [END] - Not using Plex."
        }
        $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
        # Telegram
        if ($SendTelegram) {
            Write-Output "[Telegram] $(Get-Timestamp) - Preparing Telegram message."
            $TM = Get-SiteSeriesEpisode
            if ($PlexHost -and $PlexToken -and $PlexLibId) {
                if ($Filebot -or !($Filebot) -and $MKVMerge) {
                    if (($VSVFBCount -gt 0 -and $VSVMKVCount -gt 0 -and $VSVFBCount -eq $VSVMKVCount) -or (!($Filebot) -and $MKVMerge -and $VSVMKVCount -gt 0)) {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $SiteHome. Success."
                        $TM += 'All files added to PLEX.'
                        Write-Output $TM
                        Send-Telegram -STMessage $TM
                    }
                    else {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $SiteHome. Failure."
                        $TM += 'Not all files added to PLEX.'
                        Write-Output $TM
                        Send-Telegram -STMessage $TM
                    }
                }
            }
            else {
                Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $SiteHome."
                Write-Output $TM
                Send-Telegram -STMessage $TM
            }
        }
        Write-Output "[VideoList] $(Get-Timestamp) - Final file status:"
        $VSCompletedFilesList | Format-Table @{Label = 'Series'; Expression = { $_._VSSeries } }, @{Label = 'Episode'; Expression = { $_._VSEpisode } }, `
        @{Label = 'EpisodeSubtitle'; Expression = { $_._VSEpisodeSubtitleBase } }, @{Label = 'EpisodeOverrideDrive'; Expression = { $_._VSOverridePath } }, @{Label = 'SECompleted'; Expression = { $_._VSSECompleted } }, `
        @{Label = 'MKVCompleted'; Expression = { $_._VSMKVCompleted } }, @{Label = 'FBCompleted'; Expression = { $_._VSFBCompleted } }, `
        @{Label = 'Errored'; Expression = { $_._VSErrored } } -AutoSize -Wrap
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files downloaded. Skipping other defined steps."
    }
    # Backup
    $SharedBackups = $ArchiveFile, $CookieFile, $BatFile, $ConfigPath, $SubFontDir, $SiteConfig
    foreach ($sb in $SharedBackups) {
        if (($sb -ne 'None') -and ($sb.trim() -ne '')) {
            if ($sb -eq $SubFontDir) {
                Copy-Item -Path $sb -Destination $SrcDriveSharedFonts -PassThru | Out-Null
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcDriveSharedFonts."
            }
            elseif ($sb -eq $ConfigPath) {
                Copy-Item -Path $sb -Destination $SrcBackup -PassThru | Out-Null
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcBackup."
            }
            elseif ($sb -eq $SiteConfig) {
                Copy-Item -Path $sb -Destination $SiteConfigBackup -PassThru | Out-Null
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SiteConfigBackup."
            }
            else {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcDriveShared."
                Copy-Item -Path $sb -Destination $SrcDriveShared -PassThru | Out-Null
            }
        }
    }
    Exit-Script
}