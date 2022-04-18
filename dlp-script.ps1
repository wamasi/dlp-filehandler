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
    [Alias('T')]
    [switch]$TestScript,
    [Parameter(Mandatory = $false)]
    [Alias('SN')]
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
    [switch]$SendTelegram
)
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
    [string]$_VSEpisodeRaw
    [string]$_VSEpisodeTemp
    [string]$_VSEpisodePath
    [string]$_VSEpisodeSubtitle
    [string]$_VSEpisodeSubtitleBase
    [string]$_VSEpisodeFBPath
    [string]$_VSEpisodeSubFBPath
    [bool]$_VSSECompleted
    [bool]$_VSMKVCompleted
    [bool]$_VSFBCompleted
    [bool]$_VSErrored
    
    VideoStatus([string]$VSSite, [string]$VSSeries, [string]$VSEpisode, [string]$VSEpisodeRaw, [string]$VSEpisodeTemp, [string]$VSEpisodePath, [string]$VSEpisodeSubtitle, `
            [string]$VSEpisodeSubtitleBase, [string]$VSEpisodeFBPath, [string]$VSEpisodeSubFBPath, [bool]$VSSECompleted, [bool]$VSMKVCompleted, [bool]$VSFBCompleted, [bool]$VSErrored) {
        $this._VSSite = $VSSite
        $this._VSSeries = $VSSeries
        $this._VSEpisode = $VSEpisode
        $this._VSEpisodeRaw = $VSEpisodeRaw
        $this._VSEpisodeTemp = $VSEpisodeTemp
        $this._VSEpisodePath = $VSEpisodePath
        $this._VSEpisodeSubtitle = $VSEpisodeSubtitle
        $this._VSEpisodeSubtitleBase = $VSEpisodeSubtitleBase
        $this._VSEpisodeFBPath = $VSEpisodeFBPath
        $this._VSEpisodeSubFBPath = $VSEpisodeSubFBPath
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
    Write-Host $Telegrammessage
    return $Telegrammessage
}
# Sending To telegram for new file notifications
function Send-Telegram {
    Param([Parameter(Mandatory = $true)][String]$STMessage)
    $Telegramtoken
    $Telegramchatid
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
    Write-Output "[MKVMerge] $(Get-Timestamp) - Replacing Styling in $MKVVidSubtitle."
    While ($True) {
        if ((Test-Lock $MKVVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($SF -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - Python - Regex through $MKVVidSubtitle file with $SF."
                python $SubtitleRegex $MKVVidSubtitle $SF
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
                mkvmerge -o $MKVVidTempOutput $MKVVidInput $MKVVidSubtitle --attach-file $SubFontDir --attachment-mime-type application/x-truetype-font
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) -  Merging as-is. No Font specified for $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput $MKVVidInput $MKVVidSubtitle
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
            mkvpropedit $MKVVidInput --edit track:s1 --set flag-default=1
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
    $VSCompletedFilesList | Select-Object _VSEpisodeFBPath, _VSEpisodeRaw, _VSEpisodeSubFBPath | ForEach-Object {
        $FBVidInput = $_._VSEpisodeFBPath
        $FBSubInput = $_._VSEpisodeSubFBPath
        $FBVidBaseName = $_._VSEpisodeRaw
        if ($PlexLibPath) {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found. Renaming and moving files to final folder."
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
            if (!($MKVMerge)) {
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found. Plex path not specified. Renaming files in place."
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format '{ plex.tail }' --log info
            if (!($MKVMerge)) {
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format '{ plex.tail }' --log info
            }
        }
        if (!(Test-Path $FBVidInput)) {
            Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $FBVidBaseName -SVSFP $true
        }
    }
    $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
    if ($VSVFBCount -eq $VSVTotCount ) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($VSVFBCount) = ($VSVTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup."
        filebot -script fn:cleaner "$SiteHome" --log all
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
        <emptylogs keepdays=""/>
        <filledlogs keepdays=""/>
    </Logs>
    <Plex>
        <hosturl url="" />
        <plextoken token="" />
        <library libraryid="" folder="" />
        <library libraryid="" folder="" />
        <library libraryid="" folder="" />
    </Plex>
    <Telegram>
        <token></token>
        <chatid></chatid>
    </Telegram>
    <credentials>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
            <font></font>
        </site>
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
    $SNfile = $ConfigFile.getElementsByTagName('site') | Where-Object { $_.id.trim() -ne '' } | Select-Object 'id' -ExpandProperty id
    $SNfile | ForEach-Object {
        $SN = New-Object -Type PSObject -Property @{
            SN = $_.id
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
    $ConfigPath = "$ScriptDirectory\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SNfile = $ConfigFile.getElementsByTagName('site') | Select-Object 'id', 'username', 'password', 'libraryid', 'font' | Where-Object { $_.id.ToLower() -eq "$site" }
    foreach ($object in $SNfile) {
        $SN = New-Object -TypeName PSObject -Property @{
            SN  = $object.id.ToLower()
            SND = $object.id
            SUN = $object.username
            SPW = $object.password
            SLI = $object.libraryid
            SFT = $object.font
        }
    }
    $SiteName = $SN.SN
    $SiteNameRaw = $SN.SND
    $SiteUser = $SN.SUN
    $SitePass = $SN.SPW
    $SiteLib = $SN.SLI
    $SubFont = $SN.SFT
    if ($site -eq $SiteName) {
        Write-Output "$(Get-Timestamp) - $site = $SiteName. Continuing..."
    }
    else {
        Write-Output "$(Get-Timestamp) - $site != $SiteName. Exiting..."
        exit
    }
    $BackupDrive = $ConfigFile.configuration.Directory.backup.location
    $TempDrive = $ConfigFile.configuration.Directory.temp.location
    $SrcDrive = $ConfigFile.configuration.Directory.src.location
    $DestDrive = $ConfigFile.configuration.Directory.dest.location
    $Ffmpeg = $ConfigFile.configuration.Directory.ffmpeg.location
    [int]$EmptyLogs = $ConfigFile.configuration.Logs.emptylogs.keepdays
    [int]$FilledLogs = $ConfigFile.configuration.Logs.filledlogs.keepdays
    $PlexHost = $ConfigFile.configuration.Plex.hosturl.url
    $PlexToken = $ConfigFile.configuration.Plex.plextoken.token
    $PlexLibrary = $ConfigFile.SelectNodes(('//library[@folder]')) | Where-Object { $_.libraryid -eq $SiteLib }
    $PlexLibPath = $PlexLibrary.Attributes[1].'#text'
    $PlexLibId = $PlexLibrary.Attributes[0].'#text'
    $Telegramtoken = $ConfigFile.configuration.Telegram.token
    $Telegramchatid = $ConfigFile.configuration.Telegram.chatid
    if ($SubFont.Trim() -ne '') {
        $SubFontDir = "$FontFolder\$Subfont"
        if (Test-Path $SubFontDir) {
            $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
            Write-Output "$(Get-Timestamp) - $SubFont set for $SiteName."
        }
        else {
            Write-Output "$(Get-Timestamp) - $SubFont specified in $ConfigFile is missing from $FontFolder. Exiting..."
            Exit
        }
    }
    else {
        $SubFont = 'None'
        $SubFontDir = 'None'
        $SF = 'None'
        Write-Output "$(Get-Timestamp) - $SubFont - No font set for $SiteName."
    }
    $SiteFolder = "$ScriptDirectory\sites\"
    $SiteShared = "$ScriptDirectory\shared\"
    $SrcBackup = "$BackupDrive\_Backup\"
    $SrcDriveShared = "$SrcBackup" + 'shared\'
    $SrcDriveSharedFonts = "$SrcBackup" + 'fonts\'
    $dlpParams = 'yt-dlp'
    $dlpArray = @()
    if ($Daily) {
        $SiteType = $SiteName + '_D'
        $SiteFolder = "$SiteFolder" + $SiteType
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
        $LFolderBase = "$SiteFolder\log\"
        $LFile = "$SiteFolder\log\$Date\$DateTime.log"
        if ($SrcDrive -eq $TempDrive) {
            Write-Output "[Setup] $(Get-Timestamp) - Src($SrcDrive) and Temp($TempDrive) Directories cannot be the same"
            Exit
        }
        Start-Transcript -Path $LFile -UseMinimalHeader
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $SiteNameRaw"
            Write-Output "$(Get-Timestamp) - $SiteConfig file found. Continuing..."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
            $dlpArray += "`"--config-location`"", "`"$SiteConfig`"", "`"-P`"", "`"temp:$SiteTemp`"", "`"-P`"", "`"home:$SiteSrc`""
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist. Exiting..."
            Exit
        }
    }
    else {
        $SiteType = $SiteName
        $SiteFolder = $SiteFolder + $SiteType
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
        $LFolderBase = "$SiteFolder\log\"
        $LFile = "$SiteFolder\log\$Date\$DateTime.log"
        Start-Transcript -Path $LFile -UseMinimalHeader
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $SiteNameRaw"
            Write-Output "$(Get-Timestamp) - $SiteConfig file found. Continuing..."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
            $dlpArray += "`"--config-location`"", "`"$SiteConfig`"", "`"-P`"", "`"temp:$SiteTemp`"", "`"-P`"", "`"home:$SiteSrc`""
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist. Exiting..."
            Exit
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
                    Exit
                }
            }
            else {
                $CookieFile = 'None'
                Write-Output "$(Get-Timestamp) - Login is true and Cookies is false. Continuing..."
            }
        }
        else {
            Write-Output "$(Get-Timestamp) - Login is true and Username/Password is Empty. Exiting..."
            Exit
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
            Exit
        }
    }
    if ($Ffmpeg) {
        Write-Output "$(Get-Timestamp) - $Ffmpeg file found. Continuing..."
        $dlpParams = $dlpParams + " --ffmpeg-location $Ffmpeg"
        $dlpArray += "`"--ffmpeg-location`"", "`"$Ffmpeg`""
    }
    else {
        Write-Output "$(Get-Timestamp) - FFMPEG: $Ffmpeg missing. Exiting..."
        Exit
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
            Exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - BAT: $Batfile missing. Exiting..."
        Exit
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
            Exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - Using --no-download-archive. Continuing..."
        $ArchiveFile = 'None'
        $dlpParams = $dlpParams + ' --no-download-archive'
        $dlpArray += "`"--no-download-archive`""
    }
    if ($SubtitleEdit -or $MKVMerge) {
        Write-Output $SiteConfig
        if (Select-String -Path $SiteConfig '--write-subs' -SimpleMatch -Quiet) {
            Write-Output "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is in config. Continuing..."
        }
        else {
            Write-Output "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting..."
            Exit
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
            exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - --convert-subs parameter is missing. Exiting..."
        exit
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
            exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - --remux-video parameter is missing. Exiting..."
        exit
    }
    $Wsub = Select-String -Path $SiteConfig -Pattern '--write-subs.*' | Select-Object -First 1
    if ($null -ne $Wsub) {
            Write-Output "$(Get-Timestamp) - --write-subs is in config. Continuing..."
        }
        else {
            Write-Output "$(Get-Timestamp) - SubSubWrite is missing. Exiting..."
            exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - SubtitleEdit is false. Continuing..."
    }
    
    $DebugVars = [ordered]@{Site = $SiteName; isDaily = $Daily; UseLogin = $Login; UseCookies = $Cookies; UseArchive = $Archive; SubtitleEdit = $SubtitleEdit; `
            MKVMerge = $MKVMerge; Filebot = $Filebot; SiteNameRaw = $SiteNameRaw; SiteType = $SiteType; SiteUser = $SiteUser; SitePass = $SitePass; `
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
        Write-Output 'dlpArray:'
        $dlpArray
        Write-Output "[End] $DateTime - Debugging enabled. Exiting..."
        exit
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
        $DebugVars
        Write-Output 'dlpArray:'
        $dlpArray
        $CreateFolders = $TempDrive, $SrcDrive, $BackupDrive, $SrcBackup, $SiteConfigBackup, $SrcDriveShared, $SrcDriveSharedFonts, $DestDrive, $SiteTemp, $SiteSrc, $SiteHome
        foreach ($c in $CreateFolders) {
            Set-Folders $c
        }
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
    # yt-dlp
    & yt-dlp.exe $dlpArray *>&1 | Out-Host
    # Post-processing
    if ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem $SiteSrc -Recurse -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -Unique | ForEach-Object {
            $VSSite = $SiteNameRaw
            $VSSeries = (Get-Culture).TextInfo.ToTitleCase(("$(Split-Path (Split-Path $_ -Parent) -Leaf)").Replace('_', ' ').Replace('-', ' ')) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $VSEpisode = (Get-Culture).TextInfo.ToTitleCase( ($_.BaseName.Replace('_', ' ').Replace('-', ' '))) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $VSEpisodeRaw = $_.BaseName
            $VSEpisodeTemp = $_.DirectoryName + "\$VSEpisodeRaw.temp" + $_.Extension
            $VSEpisodePath = $_.FullName
            $VSEpisodeSubtitle = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).FullName
            if ($VSEpisodeSubtitle -ne '') {
                $VSEpisodeSubtitleBase = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).Name
                $VSEpisodeSubFBPath = $VSEpisodeSubtitle.Replace($SiteSrc, $SiteHome)
            }
            else {
                $VSEpisodeSubtitleBase = ''
                $VSEpisodeSubFBPath = ''
            }
            
            $VSEpisodeFBPath = $VSEpisodePath.Replace($SiteSrc, $SiteHome)
            foreach ($i in $_) {
                $VideoStatus = [VideoStatus]::new($VSSite, $VSSeries, $VSEpisode, $VSEpisodeRaw, $VSEpisodeTemp, $VSEpisodePath, $VSEpisodeSubtitle, $VSEpisodeSubtitleBase, $VSEpisodeFBPath, `
                        $VSEpisodeSubFBPath, $VSSECompleted, $VSMKVCompleted, $VSFBCompleted, $VSErrored)
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
        $VSVMKVCount = ($VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Measure-Object).Count
        # FileMoving
        if (($VSVMKVCount -eq $VSVTotCount -and $VSVErrorCount -eq 0) -or (!($MKVMerge) -and $VSVErrorCount -eq 0)) {
            Write-Output "[FileMoving] $(Get-Timestamp)- All files had matching subtitle file"
            Write-Output "[FileMoving] $(Get-Timestamp) - $SiteSrc contains files. Moving to $SiteHomeBase..."
            Move-Item -Path $SiteSrc -Destination $SiteHomeBase -Force -Verbose
        }
        else {
            Write-Output "[FileMoving] $(Get-Timestamp) - $SiteSrc contains file(s) with error(s). Not moving files."
        }
        # Filebot
        if (($Filebot -and $VSVMKVCount -eq $VSVTotCount) -or ($Filebot -and !($MKVMerge))) {
            Start-Filebot -FBPath $SiteHome
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
            Write-Output "[PLEX] $(Get-Timestamp) - [End] - Not using Plex."
        }
        $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
        $TM = Get-SiteSeriesEpisode
        # Telegram
        if ($SendTelegram) {
            Write-Output "[Telegram] $(Get-Timestamp) - Preparing Telegram message."
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
        @{Label = 'EpisodeSubtitle'; Expression = { $_._VSEpisodeSubtitleBase } }, @{Label = 'SECompleted'; Expression = { $_._VSSECompleted } }, `
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
    # Cleanup
    Remove-Folders -RFFolder $SiteTemp -RFMatch '\\tmp\\' -RFBaseMatch $SiteTempBaseMatch
    Remove-Folders -RFFolder $SiteSrc -RFMatch '\\src\\' -RFBaseMatch $SiteSrcBaseMatch
    Remove-Folders -RFFolder $SiteHome -RFMatch '\\tmp\\' -RFBaseMatch $SiteHomeBaseMatch
    Write-Output "[END] $(Get-Timestamp) - Script completed"
    Stop-Transcript
    ((Get-Content $LFile | Select-Object -Skip 5) | Select-Object -SkipLast 4) | Set-Content $LFile
    Remove-Spaces $LFile
    if ($VSVTotCount -gt 0) {
        Rename-Item -Path $LFile -NewName "$DateTime-Total-$VSVTotCount.log"
    }
}