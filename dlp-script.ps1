# Switch params for script
param(
    [Parameter(Mandatory = $false)]
    [Alias("H")]
    [switch]$Help,
    [Parameter(Mandatory = $false)]
    [Alias("NC")]
    [switch]$NewConfig,
    [Parameter(Mandatory = $false)]
    [Alias("SU")]
    [switch]$SupportFiles,
    [Parameter(Mandatory = $false)]
    [Alias("T")]
    [switch]$TestScript,
    [Parameter(Mandatory = $false)]
    [Alias("SN")]
    [string]$Site,
    [Parameter(Mandatory = $false)]
    [Alias("D")]
    [switch]$Daily,
    [Parameter(Mandatory = $false)]
    [Alias("A")]
    [switch]$Archive,
    [Parameter(Mandatory = $false)]
    [Alias("L")]
    [switch]$Login,
    [Parameter(Mandatory = $false)]
    [Alias("C")]
    [switch]$Cookies,
    [Parameter(Mandatory = $false)]
    [Alias("F")]
    [switch]$Filebot,
    [Parameter(Mandatory = $false)]
    [Alias("MK")]
    [switch]$MKVMerge,
    [Parameter(Mandatory = $false)]
    [Alias("SE")]
    [switch]$SubtitleEdit,
    [Parameter(Mandatory = $false)]
    [Alias("ST")]
    [switch]$SendTelegram
)
# Removes erroring characters from output logs
$PSStyle.OutputRendering = 'Host'
# Getting date and datetime
function Get-Day {
    return (Get-Date -Format "yy-MM-dd")
}
function Get-TimeStamp {
    return (Get-Date -Format "yy-MM-dd HH-mm-ss")
}
function Get-Time {
    return (Get-Date -Format "MMddHHmmss")
}
# Creating folders
function Set-Folders {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Fullpath
    )
    if (!(Test-Path -Path $Fullpath)) {
        New-Item -ItemType Directory -Path $Fullpath -Force
        Write-Output "$(Get-Timestamp) - $Fullpath has been created."
    }
    else {
        Write-Output "$(Get-Timestamp) - $Fullpath already exists."
    }
}
# Creating empty support files
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
# Create Base Site
function Resolve-Configs {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Configs
    )
    New-Item $Configs -ItemType File -Force
    Write-Output "Creating $Configs"
    if ($Configs -match "vrv") {
        $vrvconfig | Set-Content $Configs
        Write-Output "$Configs created with VRV values."
    }
    elseif ($Configs -match "crunchyroll") {
        $crunchyrollconfig | Set-Content $Configs
        Write-Output "$Configs created with Crunchyroll values."
    }
    elseif ($Configs -match "funimation") {
        $funimationconfig | Set-Content $Configs
        Write-Output "$Configs created with Funimation values."
    }
    elseif ($Configs -match "hidive") {
        $hidiveconfig | Set-Content $Configs
        Write-Output "$Configs created with Hidive values."
    }
    elseif ($Configs -match "paramountplus") {
        $paramountplusconfig | Set-Content $Configs
        Write-Output "$Configs created with ParamountPlus values."
    }
    else {
        $defaultconfig | Set-Content $Configs
        Write-Output "$Configs created with default values."
    }
}
# Removing Empty whitespace
function Remove-Spaces {
    param (
        [Parameter(Mandatory = $true)]
        [string] $File
    )
    (Get-Content $File) | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } | set-content $File
    $content = [System.IO.File]::ReadAllText($File)
    $content = $content.Trim()
    [System.IO.File]::WriteAllText($File, $content)
}
# Setting Script Directory
$ScriptDirectory = $PSScriptRoot
$ConfigPath = "$ScriptDirectory\config.xml"
$SharedF = "$ScriptDirectory\shared"
$FontFolder = "$ScriptDirectory\fonts"
# Base XML config
$xmlconfig = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <Directory>
        <temp location="" />
        <src location="" />
        <home location="" />
        <ffmpeg location="" />
    </Directory>
    <Plex>
        <hosturl url="" />
        <plextoken token="" />
        <library libraryid="" />
        <library libraryid="" />
        <library libraryid="" />
    </Plex>
    <Telegram>
        <token></token>
        <chatid></chatid>
    </Telegram>
    <credentials>
        <site id="dummy">
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
"@
# Base site config parameters if file is not present
$defaultconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season_number,episode" "[$%^@.#+]" "-"
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
"@
$vrvconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season_number,episode" "[$%^@.#+]" "-"
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
"@
$crunchyrollconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season_number,episode" "[$%^@.#+]" "-"
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
--extractor-arg crunchyrollbetashow:type=sub
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv[height>=1080]+ba[height>=1080] / bv+ba / b*'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
"@
$funimationconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season_number,episode" "[$%^@.#+]" "-"
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
"@
$hidiveconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season_number,episode" "[$%^@.#+]" "-"
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
"@
$paramountplusconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season_number,episode" "[$%^@.#+]" "-"
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
"@
# Create empty config if not exists
if (!(Test-Path "$ScriptDirectory\config.xml" -PathType Leaf)) {
    New-Item "$ScriptDirectory\config.xml" -ItemType File -Force
}
# Help text to remind me what I did/it does when set to true or all parameters false
if ($help) {
    Show-Markdown -Path "$ScriptDirectory\README.md" -UseBrowser
    exit
}
# Create config if NewConfig = True
if ($NewConfig) {
    if (!(Test-Path $ConfigPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-content $ConfigPath))) {
        #PowerShell Create directory if not exists
        New-Item $ConfigPath -ItemType File -Force
        Write-Output "$ConfigPath File Created successfully"
            
        $xmlconfig | Set-Content $ConfigPath
    }
    else {
        Write-Output "$ConfigPath File Exists"
        # Perform Delete file from folder operation
    }
}
# Create supporting files if SupportFiles = True
if ($SupportFiles) {
    Set-Folders $FontFolder
    $ConfigPath = "$ScriptDirectory\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SNfile = $ConfigFile.getElementsByTagName("site") | Where-Object { $_.id.trim() -ne "" } | Select-Object "id" -ExpandProperty id
    $SNfile | ForEach-Object {
        $SN = New-Object -Type PSObject -Property @{
            SN = $_.id
        }
        # if site support files (Config, archive, bat, cookie) are missing it will attempt to create an Daily and non-Daily set
        # Creating Shared directory
        $SharedF = "$ScriptDirectory\shared"
        Set-Folders $SharedF
        #Creating Site directory
        $SCC = "$ScriptDirectory\sites"
        Set-Folders $SCC
        # Base/Manual Site root directory variable
        $SCF = "$SCC\" + $SN.SN
        Set-Folders $SCF
        # Daily Site directories
        $SCDF = "$SCF" + "_D"
        Set-Folders $SCDF
        # Daily Archive
        $SADF = "$SharedF\" + $SN.SN + "_D_A"
        Set-SuppFiles $SADF
        # Daily Bat
        $SBDF = "$SharedF\" + $SN.SN + "_D_B"
        Set-SuppFiles $SBDF
        # Daily Cookie
        $SBDC = "$SharedF\" + $SN.SN + "_D_C"
        Set-SuppFiles $SBDC
        # Manual Archive
        $SAF = "$SharedF\" + $SN.SN + "_A"
        Set-SuppFiles $SAF
        # Manual Bat
        $SBF = "$SharedF\" + $SN.SN + "_B"
        Set-SuppFiles $SBF
        # Manual Cookie
        $SBC = "$SharedF\" + $SN.SN + "_C"
        Set-SuppFiles $SBC
        # Daily Site configs
        $SCFDC = "$SCF" + "_D\yt-dlp.conf"
        Resolve-Configs $SCFDC
        Remove-Spaces $SCFDC
        # Manual Congfig
        $SCFC = "$SCF\yt-dlp.conf"
        Resolve-Configs $SCFC
        Remove-Spaces $SCFC
    }
}
# Begin if site parameter is true
if ($Site) {
    # Checking if scripts exist/found
    $DLPScript = "$ScriptDirectory\dlp-script.ps1"
    $DLPExecScript = "$ScriptDirectory\dlp-exec.ps1"
    $SubtitleRegex = "$ScriptDirectory\subtitle_regex.py"
    if (!(Test-Path -Path $DLPScript) -or !(Test-Path -Path $DLPExecScript) -or (!(Test-Path -Path $SubtitleRegex) -and $SubtitleEdit)) {
        Write-Output "$(Get-Timestamp) - dlp-script.ps1, dlp-exec.ps1, subtitle_regex.py does not exist or was not found in $ScriptDirectory folder. Exiting..."
        Exit
    }
    else {
        Write-Output "$(Get-Timestamp) - $DLPScript, $DLPExecScript, $subtitle_regex do exist $ScriptDirectory folder. Exiting..."
    }
    # Setting Date and Datetime variable for Log
    $Date = Get-Day
    $DateTime = Get-TimeStamp
    $Time = Get-Time
    # Start of parsing config xml
    $site = $site.ToLower()
    $ConfigPath = "$ScriptDirectory\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    # Fetching site variables
    $SNfile = $ConfigFile.getElementsByTagName("site") | Select-Object "id", "username", "password", "libraryid", "font" | Where-Object { $_.id.ToLower() -eq "$site" }
    foreach ($object in $SNfile) {
        $SN = New-Object -TypeName PSObject -Property @{
            SN  = $object.id.ToLower()
            SUN = $object.username
            SPW = $object.password
            SLI = $object.libraryid
            SFT = $object.font
        }
    }
    $SiteName = $SN.SN
    $SiteUser = $SN.SUN
    $SitePass = $SN.SPW
    $SiteLib = $SN.SLI
    $SubFont = $SN.SFT
    if ($site -ne $SiteName) {
        Write-Output "$site != $SiteName. Exiting..."
        exit
    }
    else { 
        Write-Output "$site = $SiteName. Continuing..."
    }
    # Setting Temp, Src, home and FFMPEG variables
    $TempDrive = $ConfigFile.configuration.Directory.temp.location
    $SrcDrive = $ConfigFile.configuration.Directory.src.location
    $DestDrive = $ConfigFile.configuration.Directory.dest.location
    $Ffmpeg = $ConfigFile.configuration.Directory.ffmpeg.location
    # Setting Plex variables
    $PlexHost = $ConfigFile.configuration.Plex.hosturl.url
    $PlexToken = $ConfigFile.configuration.Plex.plextoken.token
    $PlexLibrary = $ConfigFile.SelectNodes(("//library[@folder]")) | Where-Object { $_.libraryid -eq $SiteLib }
    # Plex Library folder name
    $PlexLibPath = $PlexLibrary.Attributes[1].'#text'
    # Plex Library ID
    $PlexLibId = $PlexLibrary.Attributes[0].'#text'
    # Telegram Bot credentials
    $Telegramtoken = $ConfigFile.configuration.Telegram.token
    $Telegramchatid = $ConfigFile.configuration.Telegram.chatid
    # Setting fonts per site. These are manually tested to work with embedding and displayin in video files
    if ($SubFont.Trim() -ne "") {
        $SubFontDir = "$ScriptDirectory\fonts\$Subfont"
        if (Test-Path $SubFontDir) {
            $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
            Write-Output "$SubFont set for $SiteName"
        }
        else {
            Write-Output "$SubFont is specified in $ConfigFile and is missing from $ScriptDirectory\fonts. Exiting..."
            Exit
        }
    }
    else {
        $SubFont = "None"
        $SubFontDir = "None"
        $SF = "None"
        Write-Output "$SubFont - No font set for $SiteName"
    }
    # Setting Site/Shared folder
    $SiteFolder = "$ScriptDirectory\sites\"
    $SiteShared = "$ScriptDirectory\shared\"
    $SrcDriveShared = "$SrcDrive\_shared\"
    $SrcDriveSharedFonts = "$SrcDriveShared\fonts\"
    # Base command for yt-dlp
    $dlpParams = 'yt-dlp'

    if ($Daily) {
        # Site folder
        $SiteType = $SiteName + "_D"
        $SiteFolder = "$SiteFolder" + $SiteType
        # Site Temp folder
        $SiteTempBase = "$TempDrive\" + $SiteName.Substring(0, 1)
        $SiteTempBaseMatch = $SiteTempBase.Replace("\", "\\")
        $SiteTemp = "$SiteTempBase\$Time"
        # Site Source folder
        $SiteSrcBase = "$SrcDrive\" + $SiteName.Substring(0, 1)
        $SiteSrcBaseMatch = $SiteSrcBase.Replace("\", "\\")
        $SiteSrc = "$SiteSrcBase\$Time"
        # Site Destination folder
        $SiteHomeBase = "$DestDrive\_" + $PlexLibPath + "\" + ($SiteName).Substring(0, 1)
        $SiteHomeBaseMatch = $SiteHomeBase.Replace("\", "\\")
        $SiteHome = "$SiteHomeBase\$Time"
        # Setting Site config
        $SiteConfig = $SiteFolder + "\yt-dlp.conf"
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig exists."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist."
            Exit
        }
    }
    # Manual
    else {
        # Site folder
        $SiteType = $SiteName
        $SiteFolder = "$SiteFolder" + $SiteType
        # Site Temp folder
        $SiteTempBase = "$TempDrive\" + $SiteName.Substring(0, 1) + "M"
        $SiteTempBaseMatch = $SiteTempBase.Replace("\", "\\")
        $SiteTemp = "$SiteTempBase\$Time"
        # Site Source folder
        $SiteSrcBase = "$SrcDrive\" + $SiteName.Substring(0, 1) + "M"
        $SiteSrcBaseMatch = $SiteSrcBase.Replace("\", "\\")
        $SiteSrc = "$SiteSrcBase\$Time"
        # Site Destination folder
        $SiteHomeBase = "$DestDrive\_M\" + $SiteName.Substring(0, 1)
        $SiteHomeBaseMatch = $SiteHomeBase.Replace("\", "\\")
        $SiteHome = "$SiteHomeBase\$Time"
        # Site Config
        $SiteConfig = $SiteFolder + "\yt-dlp.conf"
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig exists."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist."
            Exit
        }
    }
    # if Login is true then grabs associated login info if true and checks if empty. If not, then grabs cookie file.
    $CookieFile = "$SiteShared" + $SiteType + "_C"
    if ($Login) {
        # Setting cookie variable to none
        # Setting User and Password command
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
                $CookieFile = "None"
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
    # FFMPEG - Always used to handle processing and file moving
    if ($Ffmpeg) {
        Write-Output "$(Get-Timestamp) - $Ffmpeg file found. Continuing..."
        $dlpParams = $dlpParams + " --ffmpeg-location $Ffmpeg "
    }
    else {
        Write-Output "$(Get-Timestamp) - FFMPEG: $Ffmpeg missing. Exiting..."
        Exit
    }
    # BAT - Always used for calling URLS
    $BatFile = "$SiteShared" + $SiteType + "_B"
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
    # Whether archive file is used and which one
    if ($Archive) {
        $ArchiveFile = "$SiteShared" + $SiteType + "_A"
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
        $ArchiveFile = "None"
        $dlpParams = $dlpParams + " --no-download-archive"
    }
    # If SubtitleEdit in use checks if --write-subs is in config otherwise it exits
    if ($SubtitleEdit) {
        Write-Output $SiteConfig
        if (Select-String -Path $SiteConfig "--write-subs" -SimpleMatch -Quiet) {
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
    # If SubetitleEdit or Filebot true then check config for subtitle/video outputs and store in variables
    if ($SubtitleEdit -or $MKVMerge) {
        Select-String -Path $SiteConfig -Pattern "--convert-subs.*" | ForEach-Object {
            $SubType = "*." + ($_ -split " ")[1]
            $SubType = $SubType.Replace("'", "").Replace('"', "")
            if ($SubType -ne "\*.ass") {
                Write-Output "$(Get-Timestamp) - Using $SubType"
            }
            else {
                Write-Output "$(Get-Timestamp) - Subtype(ass) is missing"
                exit
            }
        }
        Select-String -Path $SiteConfig -Pattern "--remux-video.*" | ForEach-Object {
            $VidType = "*." + ($_ -split " ")[1]
            $VidType = $VidType.Replace("'", "").Replace('"', "")
            if ($SubType -ne "\*.mkv") {
                Write-Output "$(Get-Timestamp) - Using $VidType"
            }
            else {
                Write-Output "$(Get-Timestamp) - VidType(mkv) is missing..."
                exit
            }
        }
        Select-String -Path $SiteConfig -Pattern "--write-subs.*" | ForEach-Object {
            $SubWrite = ($_)
            if ($SubWrite -ne "") {
                Write-Output "$(Get-Timestamp) - SubWrite($SubWrite) is in config."
            }
            else {
                Write-Output "$(Get-Timestamp) - SubSubWrite is missing..."
                exit
            }
        }
    }
    
    # Writing all variables
    $DebugVars = [ordered]@{Site = $SiteName; SiteType = $SiteType; Daily = $Daily; SiteUser = $SiteUser; SitePass = $SitePass; Login = $Login; SiteFolder = $SiteFolder; SiteTemp = $SiteTemp; `
            SiteTempBaseMatch = $SiteTempBaseMatch; SiteSrc = $SiteSrc; SiteSrcBaseMatch = $SiteSrcBaseMatch; SiteHome = $SiteHome; SiteHomeBaseMatch = $SiteHomeBaseMatch; SiteConfig = $SiteConfig; CookieFile = $CookieFile; `
            Cookies = $Cookies; Archive = $ArchiveFile; UseDownloadArchive = $Archive; Bat = $BatFile; Ffmpeg = $Ffmpeg; SF = $SF; SubFont = $SubFont; SubFontDir = $SubFontDir; SubType = $SubType; VidType = $VidType; `
            SubtitleEdit = $SubtitleEdit; MKVMerge = $MKVMerge; Filebot = $Filebot; PlexHost = $PlexHost; PlexToken = $PlexToken; PlexLibPath = $PlexLibPath; PlexLibId = $PlexLibId; `
            TelegramToken = $TelegramToken; TelegramChatId = $TelegramChatId; ConfigPath = $ConfigPath; ScriptDirectory = $ScriptDirectory; `
            dlpParams = $dlpParams
    }
    # Creating associated log folder and file
    $LFolderBase = "$SiteFolder\log\"
    $LFile = "$SiteFolder\log\$Date\$DateTime.log"
    New-Item -Path $LFile -ItemType File -Force
    # Will generate log file for site run will all variables for debugging setup
    if ($TestScript) {
        Write-Output "[START] $DateTime - $SiteName - DEBUG Run" *>&1 | Tee-Object -FilePath $LFile -Append
        $DebugVars *>&1 | Tee-Object -FilePath $LFile -Append
        Write-Output "[End] - Debugging enabled. Exiting..." *>&1 | Tee-Object -FilePath $LFile -Append
        exit
    }
    # Will generate log file for site run most variables
    if (($Daily) -and (!($TestScript))) {
        $DebugVars.Remove("SitePass")
        $DebugVars.Remove("PlexToken")
        $DebugVars.Remove("TelegramToken")
        $DebugVars.Remove("TelegramChatId")
        Write-Output "[START] $DateTime - $SiteName - Daily Run" *>&1 | Tee-Object -FilePath $LFile -Append
        $DebugVars *>&1 | Tee-Object -FilePath $LFile -Append
    }
    elseif (!($Daily) -and (!($TestScript))) {
        $DebugVars.Remove("SitePass")
        $DebugVars.Remove("PlexToken")
        $DebugVars.Remove("TelegramToken")
        $DebugVars.Remove("TelegramChatId")
        Write-Output "[START] $DateTime - $SiteName - Manual Run" *>&1 | Tee-Object -FilePath $LFile -Append
        $DebugVars *>&1 | Tee-Object -FilePath $LFile -Append
    }
    # Runs dlp-exec.ps1 execution
    & "$ScriptDirectory\dlp-exec.ps1" -dlpParams $dlpParams -Filebot $Filebot -SubtitleEdit $SubtitleEdit -MKVMerge $MKVMerge `
        -SiteName $SiteName -SF $SF -SubFontDir $SubFontDir -PlexHost $PlexHost -PlexToken $PlexToken -PlexLibId $PlexLibId `
        -LFolderBase $LFolderBase -SiteSrc $SiteSrc -SiteHome $SiteHome -ConfigPath $ConfigPath -SiteTempBaseMatch $SiteTempBaseMatch `
        -SiteSrcBaseMatch $SiteSrcBaseMatch -SiteHomeBaseMatch $SiteHomeBaseMatch -SrcDriveShared $SrcDriveShared `
        -SrcDriveSharedFonts $SrcDriveSharedFonts *>&1 | Tee-Object -FilePath $LFile -Append
}