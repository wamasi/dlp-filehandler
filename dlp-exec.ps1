param ($dlpParams, $Filebot, $SubtitleEdit, $MKVMerge, $SiteName, $SF, $SubFontDir, $PlexHost, $PlexToken, $PlexLibId, $LFolderBase, $SiteSrc, $SiteHome, $SiteTempBaseMatch, $SiteSrcBaseMatch, $SiteHomeBaseMatch, $ConfigPath )
# Setting up arraylist for MKV and Filebot lists
class VideoStatus {
    [string]$_VSSite
    [string]$_VSSeries
    [string]$_VSEpisode
    [string]$_VSEpisodeRaw
    [string]$_VSEpisodeTemp
    [string]$_VSEpisodePath
    [string]$_VSEpisodeSubtitle
    [string]$_VSEpisodeFBPath
    [string]$_VSEpisodeSubFBPath
    [bool]$_VSSECompleted
    [bool]$_VSMKVCompleted
    [bool]$_VSFBCompleted
    [bool]$_VSErrored

    VideoStatus([string]$VSSite, [string]$VSSeries, [string]$VSEpisode, [string]$VSEpisodeRaw, [string]$VSEpisodeTemp, [string]$VSEpisodePath, [string]$VSEpisodeSubtitle, [string]$VSEpisodeFBPath, [string]$VSEpisodeSubFBPath, [bool]$VSSECompleted, [bool]$VSMKVCompleted, [bool]$VSFBCompleted, [bool]$VSErrored) {
        $this._VSSite = $VSSite
        $this._VSSeries = $VSSeries
        $this._VSEpisode = $VSEpisode
        $this._VSEpisodeRaw = $VSEpisodeRaw
        $this._VSEpisodeTemp = $VSEpisodeTemp
        $this._VSEpisodePath = $VSEpisodePath
        $this._VSEpisodeSubtitle = $VSEpisodeSubtitle
        $this._VSEpisodeFBPath = $VSEpisodeFBPath
        $this._VSEpisodeSubFBPath = $VSEpisodeSubFBPath
        $this._VSSECompleted = $VSSECompleted
        $this._VSMKVCompleted = $VSMKVCompleted
        $this._VSFBCompleted = $VSFBCompleted
        $this._VSErrored = $VSErrored
    }
}
[System.Collections.ArrayList]$VSCompletedFilesList = @()
# Update MKV/FB true false
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
    $Telegrammessage = "<b>Site:</b> " + $SiteNameRaw + "`n"
    $SeriesMessage = ""
    $SEL | ForEach-Object {
        $EpList = ""
        foreach ($i in $_) {
            $EpList = $_.Episode._VSEpisode | Out-String
        }
        $SeriesMessage = "<b>Series:</b> " + $_.Series + "`n<b>Episode:</b>`n" + $EpList
        $Telegrammessage += $SeriesMessage + "`n"
    }
    Write-Host $Telegrammessage
    return $Telegrammessage
}
# Optional sending To telegram for new file notifications
Function Send-Telegram {
    Param([Parameter(Mandatory = $true)][String]$STMessage)
    $Telegramtoken
    $Telegramchatid
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($STMessage)&parse_mode=html"
}
# Function to process video files through MKVMerge
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
            Write-Output "[MKVMerge] $(Get-Timestamp) - $MKVVidSubtitle File locked.  Waiting..."
            continue
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - File not locked. Formatting $MKVVidSubtitle file."
            if ($SF -ne "None") {
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
    Write-Output "[MKVMerge] $(Get-Timestamp) - Found matching  $MKVVidSubtitle and $MKVVidInput files to process."
    #mkmerge command to combine video and subtitle file and set subtitle default
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True -and (Test-Lock $MKVVidSubtitle) -eq $True) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - $MKVVidSubtitle and $MKVVidInput File locked.  Waiting..."
            continue
        }
        else {
            if ($SubFontDir -ne "None") {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [MKVMERGE] - File not locked.  Combining $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput $MKVVidInput $MKVVidSubtitle --attach-file $SubFontDir --attachment-mime-type application/x-truetype-font
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [MKVMERGE] - Merging as-is. No Font specified for $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput $MKVVidInput $MKVVidSubtitle
            }
        }
        Start-Sleep -Seconds 1
    }
    # If file doesn't exist yet then wait
    While (!(Test-Path $MKVVidTempOutput -ErrorAction SilentlyContinue)) {
        Start-Sleep 1.5
    }
    # Wait for files to input, subtitle, and MKVVidTempOutput to be ready
    While ($True) {
        if (((Test-Lock $MKVVidInput) -eq $True) -and ((Test-Lock $MKVVidSubtitle) -eq $True ) -and ((Test-Lock $MKVVidTempOutput) -eq $True)) {
            Write-Output "[MKVMerge] $(Get-Timestamp)- File locked.  Waiting..."
            continue
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - File not locked. Removing $MKVVidInput and $MKVVidSubtitle file."
            # Remove original video/subtitle file
            Remove-Item -Path $MKVVidInput -Confirm:$false -Verbose
            Remove-Item -Path $MKVVidSubtitle -Confirm:$false -Verbose
            break
        }
        Start-Sleep -Seconds 1
    }
    # Rename temp to original MKVVidBaseName
    While ($True) {
        if ((Test-Lock $MKVVidTempOutput) -eq $True) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - $MKVVidTempOutput File locked.  Waiting..."
            continue
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - File not locked. Renaming $MKVVidTempOutput to $MKVVidInput file."
            # Remove original video/subtitle file
            Rename-Item -Path $MKVVidTempOutput -NewName $MKVVidInput -Confirm:$false -Verbose
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True) {
            Write-Output "[MKVMerge] $(Get-Timestamp) -  $MKVVidInput File locked.  Waiting..."
            continue
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - $MKVVidInput File not locked. Setting default subtitle."
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
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to rename and move to final folder"
    $VSCompletedFilesList | Select-Object _VSEpisodeFBPath, _VSEpisodeRaw, _VSEpisodeSubFBPath | ForEach-Object {
        $FBVidInput = $_._VSEpisodeFBPath
        $FBSubInput = $_._VSEpisodeSubFBPath
        $FBVidBaseName = $_._VSEpisodeRaw
        # Filebot command
        if ($PlexLibPath) {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found. Renaming and moving files to final folder"
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
            if (!($MKVMerge)) {
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found. Plex path not specified. Renaming files in place"
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "{ plex.tail }" --log info
            if (!($MKVMerge)) {
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format "{ plex.tail }" --log info
            }
        }
        if (!(Test-Path $FBVidInput)) {
            Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $FBVidBaseName -SVSFP $true
        }
    }
    $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
    if ($VSVFBCount -eq $VSVTotCount ) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($VSVFBCount) = ($VSVTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup. Completed files:"
        $VSCompletedFilesList | Out-String
        filebot -script fn:cleaner "$SiteHome" --log all
    }
    else {
        write-output "[Filebot] $(Get-Timestamp) - Filebot($VSVFBCount) and Total Video($VSVTotCount) count mismatch. Manual check required."
    }
    # Check if folder is empty. If contains a video file file then exit, if not then completed successfully and continues
    if ($VSVFBCount -ne $VSVTotCount) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
        break
    }
    else {
        # If plex values not null then run api call else skip
        if ($PlexHost -and $PlexToken -and $PlexLibId) {
            Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
            $PlexUrl = "$PlexHost/library/sections/$PlexLibId/refresh?X-Plex-Token=$PlexToken"
            Invoke-RestMethod -UseBasicParsing $PlexUrl
        }
        else {
            Write-Output "[PLEX] $(Get-Timestamp) - [End] - Not using Plex."
        }
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
        # Regardless of failures still force delete tmp for clean runs
        if (($RFFolder -match "\\tmp\\") -and ($RFFolder -match $RFFolderBaseMatch) -and (Test-Path $RFFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $RFFolder folders/files"
            Remove-Item $RFFolder -Recurse -Force -Confirm:$false -Verbose
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteTemp($RFFolder) folder already deleted. Nothing to remove."
        }
    }
    else {
        # Clean up SiteSrc/SiteHome folder if empty
        if (($RFFolder -match "$RFMatch") -and ($RFFolder -match $RFFolderBaseMatch) -and (Test-Path $RFFolder) -and (Get-ChildItem $RFFolder -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) is empty. Deleting folder."
            & $DeleteRecursion -DRPath $RFFolder
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) contains files. Manual attention needed."
        }
    }
}
# Function to check if file is locked by process before moving forward
function Test-Lock {
    Param(
        [parameter(Mandatory = $true)]
        $TLfilename
    )
    $TLfile = Get-Item (Resolve-Path $TLfilename) -Force
    if ($TLfile -is [IO.FileInfo]) {
        trap {
            return $true
            continue
        }
        $TLstream = New-Object system.IO.StreamReader $TLfile
        if ($TLstream) { $TLstream.Close() }
    }
    return $false
}
# Function to recurse delete empty parent and subfolders
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
        Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting '${DRPath}' folders/files if empty"
        Remove-Item -Force -LiteralPath $DRPath -Verbose
    }
}
# Will use appropriate files and setup temp/home directory paths based on dlp-script.ps1 param logic
$CreateFolders = $TempDrive, $SrcDrive, $SrcDriveShared, $SrcDriveSharedFonts, $DestDrive, $SiteTemp, $SiteSrc, $SiteHome
foreach ($c in $CreateFolders) {
    Set-Folders $c
}
# Deleting logs older than 1 days
Write-Output "[LogCleanup] $(Get-Timestamp) - Deleting old logfiles"
$limit = (Get-Date).AddDays(-1)
# Delete files older than the $limit.
if (!(Test-Path $LFolderBase)) {
    Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase is missing. Skipping log cleanup..."
    break
}
else {
    Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting log cleanup..."
    # Delete any 1 day old log files.
    Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | ForEach-Object {
        $_.FullName | Remove-Item -Recurse -Force -Confirm:$false -Verbose
    }
    # Delete any empty directories left behind after deleting the old files.
    & $DeleteRecursion -DRPath $LFolderBase
}
# Call to YT-DLP with parameters
Invoke-Expression $dlpParams
# Setting up arraylist values
Write-Output "[VideoList] $(Get-Timestamp) - Fetching raw files for arraylist."
if ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
    Get-ChildItem $SiteSrc -Recurse -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -Unique | ForEach-Object {
        $VSSite = $SiteNameRaw
        $VSSeries = ("$(split-path (split-path $_ -parent) -leaf)").Replace("_", " ")
        $VSEpisode = $_.BaseName.Replace("_", " ")
        $VSEpisodeRaw = $_.BaseName
        $VSEpisodeTemp = $_.DirectoryName + "\$VSEpisodeRaw.temp" + $_.Extension
        $VSEpisodePath = $_.FullName
        $VSEpisodeSubtitle = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1).FullName
        $VSEpisodeFBPath = $VSEpisodePath.Replace($SiteSrc, $SiteHome)
        $VSEpisodeSubFBPath = $VSEpisodeSubtitle.Replace($SiteSrc, $SiteHome)
        foreach ($i in $_) {
            $VideoStatus = [VideoStatus]::new($VSSite, $VSSeries, $VSEpisode, $VSEpisodeRaw, $VSEpisodeTemp, $VSEpisodePath, $VSEpisodeSubtitle, $VSEpisodeFBPath, $VSEpisodeSubFBPath, $VSSECompleted, $VSMKVCompleted, $VSFBCompleted, $VSErrored)
            [void]$VSCompletedFilesList.Add($VideoStatus)
        }
    }
    $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | ForEach-Object {
        if (($_._VSEpisodeSubtitle -eq "") -or ($null -eq $_._VSEpisodeSubtitle)) {
            Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $VSEpisodeRaw -SVSER $true
        }
    }
}
else {
    Write-Output "[VideoList] $(Get-Timestamp) - No files to process"
}
$VSVTotCount = ($VSCompletedFilesList | Measure-Object).Count
$VSVErrorCount = ($VSCompletedFilesList | Where-Object { $_._VSErrored -eq $true } | Measure-Object).Count
Write-Output "[VideoList] $(Get-Timestamp) - Total Files: $VSVTotCount"
Write-Output "[VideoList] $(Get-Timestamp) - Errored Files: $VSVErrorCount"
$VSCompletedFilesList
if ($VSVTotCount -gt 0) {
    # If SubtitleEdit = True then run SubtitleEdit against SiteSrc folder.
    if ($SubtitleEdit) {
        # Fixing subs - SubtitleEdit
        $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | Where-Object { $_._VSErrored -ne $true } | ForEach-Object {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $_ subtitle"
            $SESubtitle = $_._VSEpisodeSubtitle
            # SubtitleEdit does not play well being called within a script
            While ($True) {
                if ((Test-Lock $SESubtitle) -eq $True) {
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - File locked. Waiting..."
                    continue
                }
                else {
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - File not locked. Editing $SESubtitle file."
                    # Remove original video/subtitle file
                    powershell "SubtitleEdit /convert '$SESubtitle' AdvancedSubStationAlpha /overwrite /MergeSameTimeCodes"
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
    # MKVMerge logic. Runs MKVMerge against SiteSrc folder then moves files.
    if ($MKVMerge) {
        $VSCompletedFilesList | Select-Object _VSEpisodeRaw, _VSEpisode, _VSEpisodeTemp, _VSEpisodePath, _VSEpisodeSubtitle, _VSErrored | `
            Where-Object { $_._VSErrored -eq $false } | ForEach-Object {
            $MKVVidInput = $_._VSEpisodePath
            $MKVVidBaseName = $_._VSEpisodeRaw
            $MKVVidSubtitle = $_._VSEpisodeSubtitle
            $MKVVidTempOutput = $_._VSEpisodeTemp
            # Adding custom styling to ASS subtitle
            Start-MKVMerge $MKVVidInput $MKVVidBaseName $MKVVidSubtitle $MKVVidTempOutput
        }
    }
    else {
        Write-Output "[MKVMerge] $(Get-Timestamp) - MKVMerge not running. Moving to next step."
    }
    $VSVMKVCount = ($VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Measure-Object).Count
    # Moving files from Src to Dest
    if (($VSVMKVCount -eq $VSVTotCount -and $VSVErrorCount -eq 0) -or ($VSVErrorCount -eq 0)) {
        Write-Output "[MKVMerge] $(Get-Timestamp)- All files had matching subtitle file"
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files."
        if ($SendTelegram) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [Telegram] - Sending message for files in $SiteSrc."
            $TM = Get-SiteSeriesEpisode
            Send-Telegram -STMessage $TM | Out-Null
        }
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Moving to $SiteHomeBase..."
        Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
    }
    else {
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - No files downloaded. Moving to next step."
    }
    # If Filebot = True then run Filebot aginst SiteHome folder
    if ((($Filebot) -and ($VSVMKVCount -eq $VSVTotCount)) -or ($Filebot -and !($MKVMerge))) {
        Start-Filebot -FBPath $SiteHome
    }
    elseif (($Filebot -and $MKVMerge) -and ($VSVTotCount -ne $VSVMKVCount)) {
        Write-Output "[Filebot] $(Get-Timestamp) - Files in $SiteSrc need manual attention. Skipping to next step... Incomplete files in $SiteSrc`:"
        $VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $false } | Out-String
    }
    else {
        Write-Output "[Filebot] $(Get-Timestamp) - [End] - Not running Filebot"
    }
}
else {
    Write-Output "[VideoList] $(Get-Timestamp) - [End] - No files downloading. Skipping other defined steps."
}
# Backup of Archive, cookie, bat, config.xml, and font files
$SharedBackups = $ArchiveFile, $CookieFile, $BatFile, $ConfigPath, $SubFontDir
foreach ($sb in $SharedBackups) {
    if (($sb -ne "None") -and ($sb.trim() -ne "")) {
        if ($sb -eq $SubFontDir) {
            Copy-Item -Path $sb -Destination $SrcDriveSharedFonts -PassThru | Out-Null
            Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcDriveSharedFonts."
        }
        else {
            Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcDriveShared."
            Copy-Item -Path $sb -Destination $SrcDriveShared -PassThru | Out-Null
        }
    }
}
Remove-Folders -RFFolder $SiteTemp -RFMatch "\\src\\" -RFBaseMatch $SiteTempBaseMatch
Remove-Folders -RFFolder $SiteSrc -RFMatch "\\src\\" -RFBaseMatch $SiteSrcBaseMatch
Remove-Folders -RFFolder $SiteHome -RFMatch "\\tmp\\" -RFBaseMatch $SiteHomeBaseMatch
if ($VSVTotCount -gt 0) {
    Write-Output "[VideoList] $(Get-Timestamp) - Final file status:"
    $VSCompletedFilesList
}
else {
    Write-Output "[VideoList] $(Get-Timestamp) - No files downloaded."
}
# End
Write-Output "[END] $(Get-Timestamp) - Script completed"