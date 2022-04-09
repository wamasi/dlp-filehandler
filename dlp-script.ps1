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
        Write-Output "[SetFolder] - $(Get-Timestamp) - $Fullpath has been created."
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
        New-Item $SuppFiles -ItemType File
        Write-Output "$SuppFiles file missing. Creating..."
    }
    else {
        Write-Output "$SuppFiles file exists"
    }
}
function Resolve-Configs {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Configs
    )
    New-Item $Configs -ItemType File -Force
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
$ScriptDirectory = $PSScriptRoot
$DLPScript = "$ScriptDirectory\dlp-script.ps1"
$DLPExecScript = "$ScriptDirectory\dlp-exec.ps1"
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
--match-filter "season!~='\(.*\)'"
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
        Write-Output "$ConfigPath File Created successfully"
            
        $xmlconfig | Set-Content $ConfigPath
    }
    else {
        Write-Output "$ConfigPath File Exists"
        
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
    if (!(Test-Path -Path $DLPExecScript) -or !(Test-Path -Path $SubtitleRegex)) {
        Write-Output "$(Get-Timestamp) - dlp-script.ps1, dlp-exec.ps1, subtitle_regex.py does not exist or was not found in $ScriptDirectory folder. Exiting..."
        Exit
    }
    else {
        Write-Output "$(Get-Timestamp) - $DLPScript, $DLPExecScript, $subtitle_regex do exist $ScriptDirectory folder. Exiting..."
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
    if ($site -ne $SiteName) {
        Write-Output "$(Get-Timestamp) - $site != $SiteName. Exiting..."
        exit
    }
    else { 
        Write-Output "$(Get-Timestamp) - $site = $SiteName. Continuing..."
    }
    $BackupDrive = $ConfigFile.configuration.Directory.backup.location
    $TempDrive = $ConfigFile.configuration.Directory.temp.location
    $SrcDrive = $ConfigFile.configuration.Directory.src.location
    $DestDrive = $ConfigFile.configuration.Directory.dest.location
    $Ffmpeg = $ConfigFile.configuration.Directory.ffmpeg.location
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
            Write-Output "$(Get-Timestamp) - $SubFont set for $SiteName"
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
        Write-Output "$(Get-Timestamp) - $SubFont - No font set for $SiteName"
    }
    $SiteFolder = "$ScriptDirectory\sites\"
    $SiteShared = "$ScriptDirectory\shared\"
    $SrcBackup = "$BackupDrive\_Backup\"
    $SrcDriveShared = "$SrcBackup" + 'shared\'
    $SrcDriveSharedFonts = "$SrcBackup" + 'fonts\'
    $dlpParams = 'yt-dlp'
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
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig exists."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist."
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
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig exists."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist."
            Exit
        }
    }
    $SiteConfigBackup = $SrcBackup + "sites\$SiteType\"
    $CookieFile = $SiteShared + $SiteType + '_C'
    if ($Login) {
        if ($SiteUser -and $SitePass) {
            Write-Output "$(Get-Timestamp) - Login is true and SiteUser/Password is filled. Continuing..."
            $dlpParams = $dlpParams + " -u $SiteUser -p $SitePass"
            if ($Cookies) {
                if ((Test-Path -Path $CookieFile)) {
                    Write-Output "$(Get-Timestamp) - Cookies is true and $CookieFile exists. Continuing..."
                    $dlpParams = $dlpParams + " --cookies $CookieFile"
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
            Write-Output "$(Get-Timestamp) - $CookieFile exists. Continuing..."
            $dlpParams = $dlpParams + " --cookies $CookieFile"
        }
        else {
            Write-Output "$(Get-Timestamp) - $CookieFile does not exist. Exiting..."
            Exit
        }
    }
    if ($Ffmpeg) {
        Write-Output "$(Get-Timestamp) - $Ffmpeg file found. Continuing..."
        $dlpParams = $dlpParams + " --ffmpeg-location $Ffmpeg "
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
        }
        else {
            Write-Output "$(Get-Timestamp) - Archive file missing. Exiting..."
            Exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - Using --no-download-archive"
        $ArchiveFile = 'None'
        $dlpParams = $dlpParams + ' --no-download-archive'
    }
    if ($SubtitleEdit) {
        Write-Output $SiteConfig
        if (Select-String -Path $SiteConfig '--write-subs' -SimpleMatch -Quiet) {
            Write-Output "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is in config. Continuing..."
        }
        else {
            Write-Output "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting..."
            Exit
        }
    }
    else {
        Write-Output "$(Get-Timestamp) - SubtitleEdit is false. Continuing..."
    }
    Select-String -Path $SiteConfig -Pattern '--convert-subs.*' | ForEach-Object {
        $SubType = '*.' + ($_ -split ' ')[1]
        $SubType = $SubType.Replace("'", '').Replace('"', '')
        if ($SubType -eq '*.ass') {
            Write-Output "$(Get-Timestamp) - Using $SubType"
        }
        else {
            Write-Output "$(Get-Timestamp) - Subtype(ass) is missing"
            exit
        }
    }
    Select-String -Path $SiteConfig -Pattern '--remux-video.*' | ForEach-Object {
        $VidType = '*.' + ($_ -split ' ')[1]
        $VidType = $VidType.Replace("'", '').Replace('"', '')
        if ($VidType -eq '*.mkv') {
            Write-Output "$(Get-Timestamp) - Using $VidType"
        }
        else {
            Write-Output "$(Get-Timestamp) - VidType(mkv) is missing..."
            exit
        }
    }
    Select-String -Path $SiteConfig -Pattern '--write-subs.*' | ForEach-Object {
        $SubWrite = ($_)
        if ($SubWrite -ne '') {
            Write-Output "$(Get-Timestamp) - SubWrite($SubWrite) is in config."
        }
        else {
            Write-Output "$(Get-Timestamp) - SubSubWrite is missing..."
            exit
        }
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
    $DebugVars
    $LFolderBase = "$SiteFolder\log\"
    $LFile = "$SiteFolder\log\$Date\$DateTime.log"
    New-Item -Path $LFile -ItemType File -Force | Out-Null
    if ($TestScript) {
        Write-Output "[START] $DateTime - $SiteName - DEBUG Run" *>&1 | Out-File -FilePath $LFile -Append -Width 9999
        $DebugVars *>&1 | Out-File -FilePath $LFile -Append -Width 9999
        Write-Output "[End] $DateTime - Debugging enabled. Exiting..." *>&1 | Out-File -FilePath $LFile -Append -Width 9999
        exit
    }
    $DebugVarRemove = 'SitePass', 'PlexToken', 'TelegramToken', 'TelegramChatId'
    foreach ($dbv in $DebugVarRemove) {
        $DebugVars.Remove($dbv)
    }
    if (($Daily) -and (!($TestScript))) {
        Write-Output "[START] $DateTime - $SiteName - Daily Run" *>&1 | Out-File -FilePath $LFile -Append -Width 9999
        $DebugVars *>&1 | Out-File -FilePath $LFile -Append -Width 9999
    }
    elseif (!($Daily) -and (!($TestScript))) {
        Write-Output "[START] $DateTime - $SiteName - Manual Run" *>&1 | Out-File -FilePath $LFile -Append -Width 9999
        $DebugVars *>&1 | Out-File -FilePath $LFile -Append -Width 9999
    }
    & $DLPExecScript -dlpParams $dlpParams -Filebot $Filebot -SubtitleEdit $SubtitleEdit -MKVMerge $MKVMerge `
        -SiteName $SiteName -SiteNameRaw $SiteNameRaw -SF $SF -SubFontDir $SubFontDir -PlexHost $PlexHost -PlexToken $PlexToken -PlexLibId $PlexLibId `
        -LFolderBase $LFolderBase -SiteSrc $SiteSrc -SiteHome $SiteHome -ConfigPath $ConfigPath -SiteTempBaseMatch $SiteTempBaseMatch `
        -SiteSrcBaseMatch $SiteSrcBaseMatch -SiteHomeBaseMatch $SiteHomeBaseMatch -SrcDriveShared $SrcDriveShared `
        -SrcDriveSharedFonts $SrcDriveSharedFonts *>&1 | Out-File -FilePath $LFile -Append -Width 9999
}