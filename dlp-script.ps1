# Switch params for script
param(
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidateSet("vrv", "funimation", "hidive", "paramountplus", "twitch", "youtube", ErrorMessage = "Value '{0}' is invalid. Only {1} are allowed")]
    [Alias("S")]
    [string]$site,
    [Parameter(Mandatory = $false)]
    [Alias("D")]
    [switch]$isDaily,
    [Parameter(Mandatory = $false)]
    [Alias("A")]
    [switch]$useArchive,
    [Parameter(Mandatory = $false)]
    [Alias("L")]
    [switch]$useLogin,
    [Parameter(Mandatory = $false)]
    [Alias("F")]
    [switch]$useFilebot,
    [Parameter(Mandatory = $false)]
    [Alias("SE")]
    [switch]$useSubtitleEdit,
    [Parameter(Mandatory = $false)]
    [Alias("B")]
    [switch]$usedebug,
    [Parameter(Mandatory = $false)]
    [Alias("H")]
    [switch]$help,
    [Parameter(Mandatory = $false)]
    [Alias("NC")]
    [switch]$newconfig,
    [Parameter(Mandatory = $false)]
    [Alias("SU")]
    [switch]$createsupportfiles
)
# Help text to remind me what I did/it does. if set then overrides other switches
if ($help) {
    Write-Host @"
This scrip assumes the following:
- Your sites are one of the following:
  - VRV
  - Funimation
  - ParamountPlus
  - Twitch
  - Youtube
- You have the following installed:
  - FFMPEG
  - Aria2
  - yt-dlp
  - Filebot
  - SubtitleEdit

Parameters explained:
- site | s | S = Tells the script what site its working with (required)
- isDaily | d | D = will use different yt-dlp configs and files and temp/home folder structure
- useArchive | a | A = Will tell yt-dlp command to use or not use archive file
- useLogin | l | L = Tells yt-dlp command to use credentials stored in config xml
- useFilebot | f | F  = Tells script to run Filebot. Will take Plex folder name defined in config xml
- useSubtitleEdit | se | SE  = Tells script to run SubetitleEdit to fix common problems with .srt files if they are present.
- useDebug | b | B = Shows minor additional info
- help | h | H  = Shows this text
- newconfig | n | N = Used to generate empty config if none is present
- createsupportfiles | su | SU = creates support files like archives, batch and cookies files for sites in the config
"@
}
# Create config if newconfig = True
elseif ($newconfig) {
    $ConfigPath = "$PSScriptRoot\config.xml" 
    if (!(Test-Path $ConfigPath -PathType Leaf)) {
        #PowerShell Create directory if not exists
        New-Item $ConfigPath -ItemType File 
        Write-Host "$ConfigPath File Created successfully"
        $config = @"
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <Directory>
        <temp location="" />
        <home location="" />
        <ffmpeg location="" />
    </Directory>
    <Plex>
        <hosturl url="" />
        <plextoken token="" />
        <library libraryid="" folder="" libraryname="" />
        <library libraryid="" folder="" libraryname="" />
        <library libraryid="" folder="" libraryname="" />
    </Plex>
    <Credentials>
        <login site="" username="" password="" libraryid="" />
        <login site="" username="" password="" libraryid="" />
        <login site="" username="" password="" libraryid="" />
        <login site="" username="" password="" libraryid="" />
        <login site="" username="" password="" libraryid="" />
        <login site="" username="" password="" libraryid="" />
    </Credentials>
</configuration>
"@
        $config | Set-Content $ConfigPath
    }
    else {
        Write-Host "$ConfigPath File Exists"
        # Perform Delete file from folder operation
    }
}
# Create supporting files if createsupportfiles = True
elseif ($createsupportfiles) {
    $ConfigPath = "$PSScriptRoot\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SNfile = $ConfigFile.SelectNodes(("//*[@site]")) | Where-Object { $_.site -ne $null } 
    $SNfile | ForEach-Object {
        $SN = New-Object -Type PSObject -Property @{
            SN = $_.site
        }
        # Base config parameters if file is not present
        $defaultconfig = @"
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--trim-filenames 248
--add-metadata
--sub-langs "en"
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 32 -s 32 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
"@
        # if site support files (Config, archive, bat, cookie) are missing it will attempt to create an isdaily and non-isDaily set
        $SharedF = "$PSScriptRoot\_shared"
        if (!(Test-Path $SharedF)) {
            New-Item ("$SharedF") -ItemType Directory
        }
        else {
            Write-Host "$SharedF Directory exists"
        }
        $SCF = "$PSScriptRoot\" + $SN.SN
        if (!(Test-Path $SCF)) {
            New-Item ("$SCF") -ItemType Directory
        }
        else {
            Write-Host "$SCF Directory exists"
        }
        $SCFC = "$PSScriptRoot\" + $SN.SN + "\yt-dlp.conf"
        if (!(Test-Path $SCFC -PathType Leaf)) {
            New-Item ("$SCFC") -ItemType File
            $defaultconfig | Set-Content  $SCFC
            Write-Host "$SCFC created with default values."
        }
        else {
            Write-Host "$SCFC File exists"
        }
        $SCDF = "$PSScriptRoot\" + $SN.SN + "_D"
        if (!(Test-Path $SCDF)) {
            New-Item ("$SCDF") -ItemType Directory
        }
        else {
            Write-Host "$SCDF Directory exists"
        }
        $SCFDC = "$PSScriptRoot\" + $SN.SN + "_D\yt-dlp.conf"
        if (!(Test-Path $SCFDC -PathType Leaf)) {
            New-Item ("$SCFDC") -ItemType File
            $defaultconfig | Set-Content  $SCFDC
            Write-Host "$SCFDC created with default values."
        }
        else {
            Write-Host "$SCFDC File exists"
        }
        $SAF = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_A"
        if (!(Test-Path $SAF -PathType Leaf)) {
            New-Item ("$SAF") -ItemType File
        }
        else {
            Write-Host "$SAF File exists"
        }
        $SBF = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_B"
        if (!(Test-Path $SBF -PathType Leaf)) {
            New-Item ("$SBF") -ItemType File
        }
        else {
            Write-Host "$SBF File exists"
        }
        $SBC = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_C"
        if (!(Test-Path $SBC -PathType Leaf)) {
            New-Item ("$SBC") -ItemType File
        }
        else {
            Write-Host "$SBC File exists"
        }
        $SADF = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_D_A"
        if (!(Test-Path $SADF -PathType Leaf)) {
            New-Item ("$SADF") -ItemType File
        }
        else {
            Write-Host "$SADF File exists"
        }
        $SBDF = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_D_B"
        if (!(Test-Path $SBDF -PathType Leaf)) {
            New-Item ("$SBDF") -ItemType File
        }
        else {
            Write-Host "$SBDF File exists"
        }
        $SBDC = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_D_C"
        if (!(Test-Path $SBDC -PathType Leaf)) {
            New-Item ("$SBDC") -ItemType File
        }
        else {
            Write-Host "$SBDC File exists"
        }
    }
}
else {
    # Functions and variables
    # Dates and timestamp
    function Get-Day {
    
        return (Get-Date -Format "yy-MM-dd")
    
    }
    function Get-TimeStamp {
    
        return (Get-Date -Format "yy-MM-dd HH-mm-ss")
    
    }
    $Date = Get-Day
    $DateTime = Get-TimeStamp
    # Start of parsing config xml
    $site = $site.ToLower()
    $ConfigPath = "$PSScriptRoot\config.xml"
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $Credentials = $ConfigFile.SelectNodes(("//*[@site]")) | Where-Object { $_.site -eq "$site" }
    $SiteLib = $Credentials.Attributes[3].'#text'
    $HomeDrive = $ConfigFile.configuration.Directory.home.location
    $TempDrive = $ConfigFile.configuration.Directory.temp.location
    $PlexHost = $ConfigFile.configuration.Plex.hosturl.url
    $PlexToken = $ConfigFile.configuration.Plex.plextoken.token
    $PlexLibrary = $ConfigFile.SelectNodes(("//library[@folder]")) | Where-Object { $_.libraryid -eq $SiteLib }
    $PlexLibPath = $PlexLibrary.Attributes[1].'#text'
    $PlexLibId = $PlexLibrary.Attributes[0].'#text'
    $ReqCookies = Get-Content "$PSScriptRoot\ReqCookies"
    #  Setting fonts per site. These are manually tested to work with embedding and displayin in video files
    if ($site -eq "vrv") {
        $SubFont = "Marker SD.ttf"
    }
    elseif ($site -eq "funimation") {
        $SubFont = "Fuzzy Bubbles.ttf"
        
    }
    elseif ($site -eq "hidive") {
        $SubFont = "Milky Nice Clean.ttf"
        
    }
    elseif ($site -eq "paramountplus") {
        $SubFont = "Coyotris Comic.ttf"
        
    }
    else {
        $SubFont = "Hey Comic.ttf"
    }

    $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
    $SubFontDir = "$PSScriptRoot\fonts\$Subfont"
    Write-Host $SubFont
    Write-Host $SF
    Write-Host $SubFontDir
    # Base command for yt-dlp
    $dlpParams = 'yt-dlp'
    # Depending on if isDaily is set will use appropriate files and setup temp/home directory paths
    if ($isDaily) {
        $SiteType = $site + "_D"
        $SiteFolder = "$PSScriptRoot\" + $SiteType
        $SiteTemp = "$TempDrive\" + $site.Substring(0, 1)
        if (!(Test-Path -Path $SiteTemp)) {
            try {
                New-Item -ItemType Directory -Path $SiteTemp -Force -Verbose
                Write-Output "$(Get-Timestamp) - $SiteTemp has been created."
            }
            catch {
                throw $_.Exception.Message
            }
        }
        else {
            Write-Output "$(Get-Timestamp) -$SiteTemp already exists."
        }
        $SiteHome = "$HomeDrive\_" + $PlexLibPath + "\" + ($site).Substring(0, 1)
        if (!(Test-Path -Path $SiteHome)) {
            try {
                New-Item -ItemType Directory -Path $SiteHome -Force -Verbose
                Write-Output "$(Get-Timestamp) - $SiteHome has been created."
            }
            catch {
                throw $_.Exception.Message
            }
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteHome already exists."
        }
        $SiteConfig = $SiteFolder + "\yt-dlp.conf"
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig exists."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteHome"
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist."
            Exit
        }
    }
    else {
        $SiteType = $site
        $SiteFolder = "$PSScriptRoot\" + $SiteType
        $SiteTemp = "$TempDrive\" + $site.Substring(0, 1) + "M"
        if (!(Test-Path -Path $SiteTemp)) {
            try {
                New-Item -ItemType Directory -Path $SiteTemp -Force 
                Write-Output "$(Get-Timestamp) - $SiteTemp has been created."
            }
            catch {
                throw $_.Exception.Message
            }
        }
        else {
            Write-Output "$(Get-Timestamp) -$SiteTemp already exists."
        }
        $SiteHome = "$HomeDrive\_M\" + $site.Substring(0, 1)
        if (!(Test-Path -Path $SiteHome)) {
            try {
                New-Item -ItemType Directory -Path $SiteHome -Force 
                Write-Output "$(Get-Timestamp) - $SiteHome has been created."
            }
            catch {
                throw $_.Exception.Message
            }
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteHome already exists."
        }

        $SiteConfig = $SiteFolder + "\yt-dlp.conf"
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "$(Get-Timestamp) - $SiteConfig exists."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteHome"
        }
        else {
            Write-Output "$(Get-Timestamp) - $SiteConfig does not exist."
            Exit
        }
    }
    # if useLogin is true then grabs associated login info if true and checks if empty. If not, then grabs cookie file.
    if ($useLogin) {
        $Username = $Credentials.Attributes[1].'#text'
        $Password = $Credentials.Attributes[2].'#text'
        $CookieFile = "None"
        if (($Username -and $Password)) {
            Write-Host "$(Get-Timestamp) - useLogin is true and Username/Password is filled. Continuing..."
            $dlpParams = $dlpParams + " -u $Username -p $Password"
        }
        else {
            Write-Host "$(Get-Timestamp) - useLogin is true and Username/Password is Empty. Exiting..."
            Exit
        }
    }
    else {
        if (!($ReqCookies -like $site)) {
            Write-Host "$(Get-Timestamp) - useLogin is FALSE and site is $ReqCookies. Exiting..."
            Exit
        }
        else {
            $CookieFile = "$PSScriptRoot\_shared\" + $SiteType + "_C"
            if ((Test-Path -Path $CookieFile)) {
                Write-Host "$(Get-Timestamp) - $CookieFile exists. Continuing..."
                $dlpParams = $dlpParams + " --cookies $CookieFile"
            }
            else {
                Write-Host "$(Get-Timestamp) - $CookieFile does not exist. Exiting..."
                Exit
            }
        }
    }
    # FFMPEG - Always used to handle processing and file moving
    $Ffmpeg = $ConfigFile.configuration.Directory.ffmpeg.location
    if ($Ffmpeg) {
        Write-Host "$(Get-Timestamp) - $Ffmpeg file found. Continuing..."
        $dlpParams = $dlpParams + " --ffmpeg-location $Ffmpeg "
    }
    else {
        Write-Host "$(Get-Timestamp) - FFMPEG: $Ffmpeg missing. Exiting..."
        Exit
    }
    # BAT - Always used for calling URLS
    $BatFile = "$PSScriptRoot\_shared\" + $SiteType + "_B"
    if ((Test-Path -Path $BatFile)) {
        Write-Host "$(Get-Timestamp) - $BatFile file found. Continuing..."
        if (![String]::IsNullOrWhiteSpace((Get-Content $BatFile))) {
            Write-Host "$(Get-Timestamp) - $BatFile not empty. Continuing..."
            $dlpParams = $dlpParams + " -a $BatFile"
        }
        else {
            Write-Host "$(Get-Timestamp) - $BatFile is empty. Exiting..."
            Exit
        }
    }
    else {
        Write-Host "$(Get-Timestamp) - BAT: $Batfile missing. Exiting..."
        Exit
    }
    # Whether archive file is used and which one
    if ($useArchive) {
        $ArchiveFile = "$PSScriptRoot\_shared\" + $SiteType + "_A"
        if ((Test-Path -Path $ArchiveFile)) {
            Write-Host "$(Get-Timestamp) - $ArchiveFile file found. Continuing..."
            $dlpParams = $dlpParams + " --download-archive $ArchiveFile"
        }
        else {
            Write-Host "$(Get-Timestamp) - Archive file missing. Exiting..."
            Exit
        }
    }
    else {
        Write-Host "$(Get-Timestamp) - Using --no-download-archive"
        $ArchiveFile = "None"
        $dlpParams = $dlpParams + " --no-download-archive"
    }
    # If SubtitleEdit in use checks if --write-subs is in config otherwise it exits
    if ($useSubtitleEdit) {
        Write-Host $SiteConfig
        if (Select-String -Path $SiteConfig "--write-subs" -SimpleMatch -Quiet) {
            Write-Host "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is in config. Continuing..."
        }
        else {
            Write-Host "$(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting..."
            Exit
        }
    }
    else {
        Write-Host "$(Get-Timestamp) - SubtitleEdit is false. Continuing..."
    }
    # If SubetitleEdit or Filebot true then check config for subtitle/video outputs and store in variables
    if ($UseSubtitleEdit -or $UseFileBot) {
        Select-String -Path $SiteConfig -Pattern "--convert-subs.*" | ForEach-Object {
            $SubType = "*." + ($_ -split " ")[1]
            $SubType = $SubType.Replace("'", "").Replace('"', "")
            Write-Host "$(Get-Timestamp) - Using $SubType"
        }
        Select-String -Path $SiteConfig -Pattern "--remux-video.*" | ForEach-Object {
            $VidType = "*." + ($_ -split " ")[1]
            $VidType = $VidType.Replace("'", "").Replace('"', "")
            Write-Host "$(Get-Timestamp) - Using $VidType"
        }
    }
    # Creating associated log folder and file
    $LFolderBase = "$PSScriptRoot\$SiteType\$SiteType" + "_log\"
    $LFolder = "$LFolderBase\" + $Date
    $LFile = "$LFolder\$SiteType" + "_" + "$DateTime.log"
    New-Item -Path $LFolder -ItemType Directory -Force
    New-Item -Path $LFile -ItemType File
    if ($isDaily) {
        Write-Output "[Start] $DateTime - $site - Daily Run" *>&1 | Tee-Object -FilePath $LFile -Append
    }
    else {
        Write-Output "[Start] $DateTime - $site - Manual Run" *>&1 | Tee-Object -FilePath $LFile -Append
    }
    # Runs execution
    & "$PSScriptRoot\dlp-exec.ps1" -dlpParams $dlpParams -useFilebot $useFilebot -useSubtitleEdit $useSubtitleEdit -site $site -SF $SF *>&1 | Tee-Object -FilePath $LFile -Append
}