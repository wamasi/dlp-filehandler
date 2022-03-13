# Switch params for script
param(
    [Parameter(Mandatory = $false)]
    [Alias("SN")]
    [string]$Site,
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
    [Alias("C")]
    [switch]$useCookies,
    [Parameter(Mandatory = $false)]
    [Alias("F")]
    [switch]$useFilebot,
    [Parameter(Mandatory = $false)]
    [Alias("SE")]
    [switch]$useSubtitleEdit,
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
Begin {
    # Create empty config if not exists
    if (!(Test-Path "$PSScriptRoot\config.xml" -PathType Leaf)) {
        New-Item "$PSScriptRoot\config.xml" -ItemType File -Force
    }
    # Help text to remind me what I did/it does when set to true or all parameters false
    if ($help -or [string]::IsNullOrEmpty($site) -and $isDaily -eq $false -and $useArchive -eq $false -and $useLogin -eq $false -and $useFilebot -eq $false -and $useSubtitleEdit -eq $false -and $newconfig -eq $false -and $createsupportfiles -eq $false) {
        Show-Markdown -Path "$PSSCriptRoot\README.md" -UseBrowser
    }
    # Create config if newconfig = True
    if ($newconfig) {
        $ConfigPath = "$PSScriptRoot\config.xml"
        if (!(Test-Path $ConfigPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-content $ConfigPath))) {
            #PowerShell Create directory if not exists
            New-Item $ConfigPath -ItemType File -Force
            Write-Output "$ConfigPath File Created successfully"
            $config = @"
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
            $config | Set-Content $ConfigPath
        }
        else {
            Write-Output "$ConfigPath File Exists"
            # Perform Delete file from folder operation
        }
    }
    # Base config parameters if file is not present
    $defaultconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
--sub-format 'ass/vtt/srt'
--sub-format 'vtt/srt'
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
    $vrvconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--trim-filenames 248
--add-metadata
--sub-langs "en-US"
--sub-format 'ass/vtt/srt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 32 -s 32 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
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
--trim-filenames 248
--add-metadata
--sub-langs "en-US"
--sub-format 'ass/vtt/srt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--extractor-arg crunchyrollbetashow:type=sub
--downloader aria2c
--downloader-args aria2c:'-c -j 32 -s 32 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
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
--trim-filenames 248
--add-metadata
--sub-langs 'en.*'
--sub-format 'ass/vtt/srt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--extractor-args 'funimation:language=japanese'
--downloader-args aria2c:'-c -j 32 -s 32 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
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
--trim-filenames 248
--add-metadata
--sub-langs "english-subs"
--sub-format 'ass/vtt/srt'
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
    $paramountplusconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
--sub-format 'ass/vtt/srt'
--sub-format 'vtt/srt'
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
    # Create supporting files if createsupportfiles = True
    if ($createsupportfiles) {
        function Folders {
            param (
                [Parameter(Mandatory = $true)]
                [string] $folder
            )
            if (!(Test-Path $folder)) {
                New-Item ("$folder") -ItemType Directory
                Write-Output "$folder directory missing. Creating..."
            }
            else {
                Write-Output "$folder directory exists"
            }
        }
        function Configs {
            param (
                [Parameter(Mandatory = $true)]
                [string] $Configs
            )
            if (!(Test-Path $Configs -PathType Leaf)) {
                New-Item $Configs -ItemType File
                Write-Output "$Configs file missing. Creating..."
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
            else {
                Write-Output "$Configs file exists"
            }
        }
        function TrimSpaces {
            param (
                [Parameter(Mandatory = $true)]
                [string] $File
            )
            (Get-Content $File) | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } | set-content $File
            $content = [System.IO.File]::ReadAllText($File)
            $content = $content.Trim()
            [System.IO.File]::WriteAllText($File, $content)
        }
        function SuppFiles {
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
        $FontFolder = "$PSScriptRoot\fonts"
        Folders $FontFolder
        $ConfigPath = "$PSScriptRoot\config.xml"
        [xml]$ConfigFile = Get-Content -Path $ConfigPath
        $SNfile = $ConfigFile.getElementsByTagName("site") | Where-Object { $_.id.trim() -ne "" } | Select-Object "id" -ExpandProperty id
        $SNfile | ForEach-Object {
            $SN = New-Object -Type PSObject -Property @{
                SN = $_.id
            }
            # if site support files (Config, archive, bat, cookie) are missing it will attempt to create an isdaily and non-isDaily set
            # Creating Shared directory
            $SharedF = "$PSScriptRoot\shared"
            Folders $SharedF
            #Creating Site directory
            $SCC = "$PSScriptRoot\sites"
            Folders $SCC
            # Base/Manual Site root directory variable
            $SCF = "$SCC\" + $SN.SN
            Folders $SCF
            # Daily Site directories
            $SCDF = "$SCF" + "_D"
            Folders $SCDF
            # Daily Archive
            $SADF = "$SharedF\" + $SN.SN + "_D_A"
            SuppFiles $SADF
            # Daily Bat
            $SBDF = "$SharedF\" + $SN.SN + "_D_B"
            SuppFiles $SBDF
            # Daily Cookie
            $SBDC = "$SharedF\" + $SN.SN + "_D_C"
            SuppFiles $SBDC
            # Manual Archive
            $SAF = "$SharedF\" + $SN.SN + "_A"
            SuppFiles $SAF
            # Manual Bat
            $SBF = "$SharedF\" + $SN.SN + "_B"
            SuppFiles $SBF
            # Manual Cookie
            $SBC = "$SharedF\" + $SN.SN + "_C"
            SuppFiles $SBC
            # Daily Site configs
            $SCFDC = "$SCF" + "_D\yt-dlp.conf"
            Configs $SCFDC
            TrimSpaces $SCFDC
            # Manual Congfig
            $SCFC = "$SCF\yt-dlp.conf"
            Configs $SCFC
            TrimSpaces $SCFC
        }
    }
}
# Begin if site parameter is true
Process {
    if ($Site) {
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
                try {
                    New-Item -ItemType Directory -Path $Fullpath -Force -Verbose
                    Write-Output "$(Get-Timestamp) - $Fullpath has been created."
                }
                catch {
                    throw $_.Exception.Message
                }
            }
            else {
                Write-Output "$(Get-Timestamp) - $Fullpath already exists."
            }
        }
        # Setting Date and Datetime variable for Log
        $Date = Get-Day
        $DateTime = Get-TimeStamp
        $Time = Get-Time
        # Start of parsing config xml
        $site = $site.ToLower()
        $ScriptDirectory = $PSScriptRoot
        $ConfigPath = "$PSScriptRoot\config.xml"
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
        # Setting fonts per site. These are manually tested to work with embedding and displayin in video files
        if ($SubFont.trim() -ne "") {
            $SubFontDir = "$PSScriptRoot\fonts\$Subfont"
            $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
            Write-Output "$SubFont set for $SiteName"
        }
        else {
            $SubFont = "None"
            $SubFontDir = "None"
            $SF = "None"
            Write-Output "$SubFont - No font set for $SiteName"
        }
        # Setting Site/Shared folder
        $SiteFolder = "$PSScriptRoot\sites\"
        $SiteShared = "$PSScriptRoot\shared\"
        $SrcDriveShared = "$SrcDrive\_shared\"
        # Base command for yt-dlp
        $dlpParams = 'yt-dlp'
        # Depending on if isDaily is set will use appropriate files and setup temp/home directory paths
        Set-Folders $TempDrive
        Set-Folders $SrcDrive
        Set-Folders $SrcDriveShared
        Set-Folders $DestDrive
        if ($isDaily) {
            # Site folder
            $SiteType = $SiteName + "_D"
            $SiteFolder = "$SiteFolder" + $SiteType
            # Video Temp folder
            $SiteTempBase = "$TempDrive\" + $SiteName.Substring(0, 1)
            $SiteTemp = "$SiteTempBase\$Time"
            Set-Folders $SiteTemp
            # Video Destination folder
            $SrcBase = "$SrcDrive\" + $SiteName.Substring(0, 1)
            $SiteSrc = "$SrcBase\$Time"
            Set-Folders $SiteSrc
            $SiteHomeBase = "$DestDrive\_" + $PlexLibPath + "\" + ($SiteName).Substring(0, 1)
            $SiteHome = "$SiteHomeBase\$Time"
            Set-Folders $SiteHome
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
            $SiteTemp = "$SiteTempBase\$Time"
            Set-Folders $SiteTemp
            # Site Destination folder
            $SiteSrcBase = "$SrcDrive\" + $SiteName.Substring(0, 1) + "M"
            $SiteSrc = "$SiteSrcBase\$Time"
            Set-Folders $SiteSrc
            $SiteHomeBase = "$DestDrive\_M\" + $SiteName.Substring(0, 1)
            $SiteHome = "$SiteHomeBase\$Time"
            Set-Folders $SiteHome
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
        # if useLogin is true then grabs associated login info if true and checks if empty. If not, then grabs cookie file.
        $CookieFile = "$SiteShared" + $SiteType + "_C"
        if ($useLogin) {
            # Setting cookie variable to none
            # Setting User and Password command
            if (($SiteUser -and $SitePass)) {
                Write-Output "$(Get-Timestamp) - useLogin is true and SiteUser/Password is filled. Continuing..."
                $dlpParams = $dlpParams + " -u $SiteUser -p $SitePass"
                if ($useCookies) {
                    if ((Test-Path -Path $CookieFile)) {
                        Write-Output "$(Get-Timestamp) - useCookies is true and $CookieFile exists. Continuing..."
                        $dlpParams = $dlpParams + " --cookies $CookieFile"
                    }
                    else {
                        Write-Output "$(Get-Timestamp) - $CookieFile does not exist. Exiting..."
                        Exit
                    }
                }
                else {
                    $CookieFile = "None"
                    Write-Output "$(Get-Timestamp) - useLogin is true and useCookies is false. Continuing..."
                }
            }
            else {
                Write-Output "$(Get-Timestamp) - useLogin is true and Username/Password is Empty. Exiting..."
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
        if ($useArchive) {
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
        if ($useSubtitleEdit) {
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
        if ($UseSubtitleEdit -or $UseFileBot) {
            Select-String -Path $SiteConfig -Pattern "--convert-subs.*" | ForEach-Object {
                $SubType = "*." + ($_ -split " ")[1]
                $SubType = $SubType.Replace("'", "").Replace('"', "")
                Write-Output "$(Get-Timestamp) - Using $SubType"
            }
            Select-String -Path $SiteConfig -Pattern "--remux-video.*" | ForEach-Object {
                $VidType = "*." + ($_ -split " ")[1]
                $VidType = $VidType.Replace("'", "").Replace('"', "")
                Write-Output "$(Get-Timestamp) - Using $VidType"
            }
        }
        # Creating associated log folder and file
        $LFolderBase = "$SiteFolder\log\"
        $LFile = "$SiteFolder\log\$Date\$DateTime.log"
        New-Item -Path $LFile -ItemType File -Force
        # Setting log header row based on daily vs Manual
        if ($isDaily) {
            Write-Output "[START] $DateTime - $SiteName - Daily Run" *>&1 | Tee-Object -FilePath $LFile -Append
        }
        else {
            Write-Output "[START] $DateTime - $SiteName - Manual Run" *>&1 | Tee-Object -FilePath $LFile -Append
        }
        # Runs execution
        & "$PSScriptRoot\dlp-exec.ps1" -dlpParams $dlpParams -useFilebot $useFilebot -useSubtitleEdit $useSubtitleEdit `
        -SiteName $SiteName -SF $SF -SubFontDir $SubFontDir -PlexHost $PlexHost -PlexToken $PlexToken -PlexLibId $PlexLibId `
        -LFolderBase $LFolderBase -SiteSrc $SiteSrc -SiteHome $SiteHome -ConfigPath $ConfigPath *>&1 | Tee-Object -FilePath $LFile -Append
    }
}