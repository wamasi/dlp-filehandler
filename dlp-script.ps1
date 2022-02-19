# Switch params for script
param(
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
DynamicParam {
    # Need WFTools to work. Otherwise modify to normal string paramater above.
    New-DynamicParam -Name site -Alias SN -ValidateSet $(([xml](Get-Content "$PSSCriptRoot\config.xml")).getElementsByTagName("site") | Where-Object { $_.id -ne $null } | Select-Object "id" -ExpandProperty id)
}
Begin {
    #This standard block of code loops through bound parameters...
    #If no corresponding variable exists, one is created
    foreach ($param in $PSBoundParameters.Keys) {
        if (-not ( Get-Variable -name $param -scope 0 -ErrorAction SilentlyContinue ) ) {
            New-Variable -Name $param -Value $PSBoundParameters.$param
            Write-Verbose "Adding variable for dynamic parameter '$param' with value '$($PSBoundParameters.$param)'"
        }
    }
    ($MyInvocation.MyCommand.Parameters ).Keys | ForEach-Object {
        $val = (Get-Variable -Name $_ -EA SilentlyContinue).Value
        if ( $val.length -gt 0 ) {
            "($($_)) = ($($val))"
        }
    }
    # Create empty config if not exists
    if (!(Test-Path "$PSScriptRoot\config.xml" -PathType Leaf)) {
        New-Item "$PSScriptRoot\config.xml" -ItemType File -Force
    }
    # Help text to remind me what I did/it does when set to true or all parameters false
    if ($help -or [string]::IsNullOrEmpty($site) -and $isDaily -eq $false -and $useArchive -eq $false -and $useLogin -eq $false -and $useFilebot -eq $false -and $useSubtitleEdit -eq $false -and $usedebug -eq $false -and $newconfig -eq $false -and $createsupportfiles -eq $false) {
        Show-Markdown -Path "$PSSCriptRoot\README.md" -UseBrowser
    }
    # Create config if newconfig = True
    if ($newconfig) {
        $ConfigPath = "$PSScriptRoot\config.xml"
        if (!(Test-Path $ConfigPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-content $ConfigPath))) {
            #PowerShell Create directory if not exists
            New-Item $ConfigPath -ItemType File -Force
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
        <library libraryid="" />
        <library libraryid="" />
        <library libraryid="" />
    </Plex>
    <credentials>
        <site id="dummy">
            <username></username>
            <password></password>
            <libraryid></libraryid>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
        </site>
        <site id="">
            <username></username>
            <password></password>
            <libraryid></libraryid>
        </site>
    </credentials>
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
    if ($createsupportfiles) {
        $ConfigPath = "$PSScriptRoot\config.xml"
        [xml]$ConfigFile = Get-Content -Path $ConfigPath
        $SNfile = $ConfigFile.getElementsByTagName("site") | Where-Object { $_.id.trim() -ne "" } | Select-Object "id" -ExpandProperty id
        $SNfile | ForEach-Object {
            $SN = New-Object -Type PSObject -Property @{
                SN = $_.id
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
            $vrvconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
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
            $funimationconfig = @"
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--trim-filenames 248
--add-metadata
--sub-langs 'en.*,en_Uncut.*'
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
--sub-langs "en.*"
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
            # Creating Shared directory
            $SharedF = "$PSScriptRoot\_shared"
            if (!(Test-Path $SharedF)) {
                New-Item ("$SharedF") -ItemType Directory
                Write-Host "$SharedF directory missing. Creating..."
            }
            else {
                Write-Host "$SharedF directory exists"
            }
            # Creating Site directory
            $SCC = "$PSScriptRoot\sites"
            if (!(Test-Path $SCC)) {
                New-Item ("$SCC") -ItemType Directory
                Write-Host "$SCC directory missing. Creating..."
            }
            else {
                Write-Host "$SCC directory exists"
            }
            # Base Site root directory variable
            $SCF = "$SCC\" + $SN.SN
            # Daily
            # Daily Site directories
            $SCDF = "$SCF" + "_D"
            if (!(Test-Path $SCDF)) {
                New-Item ("$SCDF") -ItemType Directory
                Write-Host "$SCDF directory missing. Creating..."
            }
            else {
                Write-Host "$SCDF directory exists"
            }
            # Daily Site configs
            $SCFDC = "$SCF" + "_D\yt-dlp.conf"
            if (!(Test-Path $SCFDC -PathType Leaf)) {
                New-Item ("$SCFDC") -ItemType File
                Write-Host "$SCFDC file missing. Creating..."
                if ($SCFDC -match "vrv") {
                    $vrvconfig | Set-Content $SCFDC
                    Write-Host "$SCFDC created with VRV values."
                }
                elseif ($SCFDC -match "funimation") {
                    $funimationconfig | Set-Content $SCFDC
                    Write-Host "$SCFDC created with Funimation values."
                }
                elseif ($SCFDC -match "hidive") {
                    $hidiveconfig | Set-Content $SCFDC
                    Write-Host "$SCFDC created with Hidive values."
                }
                elseif ($SCFDC -match "paramountplus") {
                    $paramountplusconfig | Set-Content $SCFDC
                    Write-Host "$SCFDC created with ParamountPlus values."
                }
                else {
                    $defaultconfig | Set-Content $SCFDC
                    Write-Host "$SCFDC created with default values."
                }
            }
            else {
                Write-Host "$SCFDC file exists"
            }
            # Daily Archive
            $SADF = "$PSScriptRoot" + "\_shared\" + $SN.SN + "_D_A"
            if (!(Test-Path $SADF -PathType Leaf)) {
                New-Item ("$SADF") -ItemType File
                Write-Host "$SADF file missing. Creating..."
            }
            else {
                Write-Host "$SADF file exists"
            }
            # Daily Bat
            $SBDF = "$SharedF\" + $SN.SN + "_D_B"
            if (!(Test-Path $SBDF -PathType Leaf)) {
                New-Item ("$SBDF") -ItemType File
                Write-Host "$SBDF file missing. Creating..."
            }
            else {
                Write-Host "$SBDF file exists"
            }
            # Daily Cookie
            $SBDC = "$SharedF\" + $SN.SN + "_D_C"
            if (!(Test-Path $SBDC -PathType Leaf)) {
                New-Item ("$SBDC") -ItemType File
                Write-Host "$SBDC file missing. Creating..."
            }
            else {
                Write-Host "$SBDC file exists"
            }
            # Manual
            # Manual Base directories
            if (!(Test-Path $SCF)) {
                New-Item ("$SCF") -ItemType Directory
                Write-Host "$SCF directory missing. Creating..."
            }
            else {
                Write-Host "$SCF directory exists"
            }
            # Manual Congfig
            $SCFC = "$SCF\yt-dlp.conf"
            if (!(Test-Path $SCFC -PathType Leaf)) {
                New-Item ("$SCFC") -ItemType File
                Write-Host "$SCFC file missing. Creating..."
                if ($SCFC -match "vrv") {
                    $vrvconfig | Set-Content $SCFC
                    Write-Host "$SCFDC created with VRV values."
                }
                elseif ($SCFC -match "funimation") {
                    $funimationconfig | Set-Content $SCFC
                    Write-Host "$SCFC created with Funimation values."
                }
                elseif ($SCFC -match "hidive") {
                    $hidiveconfig | Set-Content $SCFC
                    Write-Host "$SCFC created with Hidive values."
                }
                elseif ($SCFC -match "paramountplus") {
                    $paramountplusconfig | Set-Content $SCFC
                    Write-Host "$SCFC created with ParamountPlus values."
                }
                else {
                    $defaultconfig | Set-Content $SCFC
                    Write-Host "$SCFC created with default values."
                }
            }
            else {
                Write-Host "$SCFC file exists"
            }
            # Manual Archive
            $SAF = "$SharedF\" + $SN.SN + "_A"
            if (!(Test-Path $SAF -PathType Leaf)) {
                New-Item ("$SAF") -ItemType File
                Write-Host "$SAF file missing. Creating..."
            }
            else {
                Write-Host "$SAF file exists"
            }
            # Manual Bat
            $SBF = "$SharedF\" + $SN.SN + "_B"
            if (!(Test-Path $SBF -PathType Leaf)) {
                New-Item ("$SBF") -ItemType File
                Write-Host "$SBF file missing. Creating..."
            }
            else {
                Write-Host "$SBF file exists"
            }
            # Manual Cookie
            $SBC = "$SharedF\" + $SN.SN + "_C"
            if (!(Test-Path $SBC -PathType Leaf)) {
                New-Item ("$SBC") -ItemType File
                Write-Host "$SBC file missing. Creating..."
            }
            else {
                Write-Host "$SBC file exists"
            }
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
        # Setting Date and Datetime variable for Log
        $Date = Get-Day
        $DateTime = Get-TimeStamp
        # Start of parsing config xml
        $site = $site.ToLower()
        $ConfigPath = "$PSScriptRoot\config.xml"
        [xml]$ConfigFile = Get-Content -Path $ConfigPath
        # Fetching site variables
        $SNfile = $ConfigFile.getElementsByTagName("site") | Select-Object "id", "username", "password", "libraryid" | Where-Object { $_.id -eq "$site" }
        $SNfile | ForEach-Object {
            $SN = New-Object -Type PSObject -Property @{
                SN  = $_.id
                SUN = $_.username
                SPW = $_.password
                SLI = $_.libraryid
            }
        }
        $SiteName = $SN.SN
        $SiteUser = $SN.SUN
        $SitePass = $SN.SPW
        $SiteLib = $SN.SLI
        # Setting Home and Temp directory variables
        $HomeDrive = $ConfigFile.configuration.Directory.home.location
        $TempDrive = $ConfigFile.configuration.Directory.temp.location
        # Setting Plex variables
        $PlexHost = $ConfigFile.configuration.Plex.hosturl.url
        $PlexToken = $ConfigFile.configuration.Plex.plextoken.token
        $PlexLibrary = $ConfigFile.SelectNodes(("//library[@folder]")) | Where-Object { $_.libraryid -eq $SiteLib }
        # Plex Library folder name
        $PlexLibPath = $PlexLibrary.Attributes[1].'#text'
        # Plex Library ID
        $PlexLibId = $PlexLibrary.Attributes[0].'#text'
        # Pulling sites that require cookies from text
        $ReqCookies = Get-Content "$PSScriptRoot\ReqCookies"
        #  Setting fonts per site. These are manually tested to work with embedding and displayin in video files
        if ($SiteName -eq "vrv") {
            $SubFont = "Marker SD.ttf"
        }
        elseif ($SiteName -eq "funimation") {
            $SubFont = "Fuzzy Bubbles.ttf"
        }
        elseif ($SiteName -eq "hidive") {
            $SubFont = "Milky Nice Clean.ttf"
        }
        elseif ($SiteName -eq "paramountplus") {
            $SubFont = "Coyotris Comic.ttf"
        }
        else {
            $SubFont = "Hey Comic.ttf"
        }
        $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
        $SubFontDir = "$PSScriptRoot\fonts\$Subfont"
        # Setting Site/Shared folder
        $SiteFolder = "$PSScriptRoot\sites\"
        $SiteShared = "$PSScriptRoot\_shared\"
        # Base command for yt-dlp
        $dlpParams = 'yt-dlp'
        # Depending on if isDaily is set will use appropriate files and setup temp/home directory paths
        if ($isDaily) {
            # Site folder
            $SiteType = $SiteName + "_D"
            $SiteFolder = "$SiteFolder" + $SiteType
            # Video Temp folder
            $SiteTemp = "$TempDrive\" + $SiteName.Substring(0, 1)
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
            # Video Destination folder
            $SiteHome = "$HomeDrive\_" + $PlexLibPath + "\" + ($SiteName).Substring(0, 1)
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
            # Setting Site config
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
        # Manual
        else {
            # Site folder
            $SiteType = $SiteName
            $SiteFolder = "$SiteFolder" + $SiteType
            # Site Temp folder
            $SiteTemp = "$TempDrive\" + $SiteName.Substring(0, 1) + "M"
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
            # Site Destination folder
            $SiteHome = "$HomeDrive\_M\" + $SiteName.Substring(0, 1)
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
            # Site Config
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
            # Setting cookie variable to none
            $CookieFile = "None"
            # Setting User and Password command
            if (($SiteUser -and $SitePass)) {
                Write-Host "$(Get-Timestamp) - useLogin is true and SiteUser/Password is filled. Continuing..."
                $dlpParams = $dlpParams + " -u $SiteUser -p $SitePass"
            }
            else {
                Write-Host "$(Get-Timestamp) - useLogin is true and Username/Password is Empty. Exiting..."
                Exit
            }
        }
        else {
            if (!($ReqCookies -like $SiteName)) {
                Write-Host "$(Get-Timestamp) - useLogin is FALSE and site is not $ReqCookies. Exiting..."
                Exit
            }
            else {
                $CookieFile = "$SiteShared" + $SiteType + "_C"
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
        $BatFile = "$SiteShared" + $SiteType + "_B"
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
            $ArchiveFile = "$SiteShared" + $SiteType + "_A"
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
        $LFolderBase = "$SiteFolder\log\"
        $LFile = "$SiteFolder\log\$Date\$DateTime.log"
        New-Item -Path $LFile -ItemType File -Force
        # Setting log header row based on daily vs Manual
        if ($isDaily) {
            Write-Output "[Start] $DateTime - $SiteName - Daily Run" *>&1 | Tee-Object -FilePath $LFile -Append
        }
        else {
            Write-Output "[Start] $DateTime - $SiteName - Manual Run" *>&1 | Tee-Object -FilePath $LFile -Append
        }
        # Runs execution
        & "$PSScriptRoot\dlp-exec.ps1" -dlpParams $dlpParams -useFilebot $useFilebot -useSubtitleEdit $useSubtitleEdit -SiteName $SiteName -SF $SF -SubFontDir $SubFontDir -PlexHost $PlexHost -PlexToken $PlexToken -PlexLibId $PlexLibId -LFolderBase $LFolderBase *>&1 | Tee-Object -FilePath $LFile -Append
    }
}