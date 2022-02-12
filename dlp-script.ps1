# Switch params for script
param(
    [Parameter(Mandatory = $true)]
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
    [switch]$help
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
- site = Tells the script what site its working with (required)
- isDaily = will use different yt-dlp configs and files and temp/home folder structure
- useArchive = Will tell yt-dlp command to use or not use archive file
- useLogin = Tells yt-dlp command to use credentials stored in config xml
- useFilebot = Tells script to run Filebot. Will take Plex folder name defined in config xml
- useSubtitleEdit = Tells script to run SubetitleEdit to fix common problems with .srt files if they are present.
"@
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