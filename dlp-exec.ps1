param ($dlpParams, $Filebot, $SubtitleEdit, $MKVMerge, $SiteName, $SF, $SubFontDir, $PlexHost, $PlexToken, $PlexLibId, $LFolderBase, $SiteSrc, $SiteHome, $SiteTempBaseMatch, $SiteSrcBaseMatch, $SiteHomeBaseMatch, $ConfigPath )
# defining Class for Telegram messages
class SeriesEpisode {
    [string]$_Site
    [string]$_Series
    [string]$_Episode

    SeriesEpisode([string]$sSite, [string]$Series, [string]$Episode) {
        $this._Site = $sSite
        $this._Series = $Series
        $this._Episode = $Episode
    }
}
[System.Collections.ArrayList]$SeriesEpisodeList = @()
# Getting list of Site, Series, and Episodes for Telegram messages
function Get-SiteSeriesEpisode {
    param (
        [parameter(Mandatory = $true)]
        $VPath,
        [parameter(Mandatory = $true)]
        $VType
    )
    Get-ChildItem $VPath -Recurse -Include "$VType" | Sort-Object LastWriteTime | Select-Object -Unique | ForEach-Object {
        $Series = ("$(split-path (split-path $_ -parent) -leaf)").Replace("_", " ")
        $Episodes = $_.BaseName.Replace("_", " ")
        foreach ($i in $_) {
            $SeriesEpisodes = [SeriesEpisode]::new($sSite, $Series, $Episodes)
            [void]$SeriesEpisodeList.Add($SeriesEpisodes)
        }
    }
    $SEL = $SeriesEpisodeList | Group-Object -Property _Site, _Series |
    Select-Object @{n = 'Site'; e = { $_.Values[0] } }, `
    @{ n = 'Series'; e = { $_.Values[1] } }, `
    @{n = 'Episode'; e = { $_.Group | Select-Object _Episode } }
    $Telegrammessage = "<b>Site:</b> " + $SiteNameRaw + "`n"
    $SeriesMessage = ""
    $SEL | ForEach-Object {
        $EpList = ""
        foreach ($i in $_) {
            $EpList = $_.Episode._Episode | Out-String
        }
        $SeriesMessage = "<b>Series:</b> " + $_.Series + "`n<b>Episode:</b>`n" + $EpList
        $Telegrammessage += $SeriesMessage + "`n"
    }
    Write-Host $Telegrammessage
    return $Telegrammessage
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
# Optional sending To telegram for new file notifications
Function Send-Telegram {
    Param([Parameter(Mandatory = $true)][String]$STMessage)
    $Telegramtoken
    $Telegramchatid
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($STMessage)&parse_mode=html"
}
[System.Collections.ArrayList]$MKVCompletedFiles = [ordered]@{}
[System.Collections.ArrayList]$FBCompletedFiles = [ordered]@{}
[System.Collections.ArrayList]$incompleteFiles = [ordered]@{}
[bool] $SiteSrcDeleted = $false
# Depending on if Daily is set will use appropriate files and setup temp/home directory paths
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
# If SubtitleEdit = True then run SubtitleEdit against SiteSrc folder.
if ($SubtitleEdit) {
    # Fixing subs - SubtitleEdit
    if ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$SubType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Sort-Object LastWriteTime | ForEach-Object {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $_ subtitle"
            $SESubtitle = $_.FullName
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
                    break
                }
                Start-Sleep -Seconds 1
            }
        }
    }
    else {
        Write-Output "[SubtitleEdit] $(Get-Timestamp) - No subtitles found"
    }
}
else {
    Write-Output "[SubtitleEdit] $(Get-Timestamp) - Not running"
}
# If MKVMerge = True then run MKVMerge against SiteSrc folder.
if ($MKVMerge) {
    ForEach ($SiteSrcfolder in $SiteSrc) {
        if ((Get-ChildItem $SiteSrcfolder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0 -and (Get-ChildItem $SiteSrcfolder -Recurse -Force -File -Include "$SubType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - Processing files to fix subtitles and then combine with video."
            # Embedding subs into video files - MKVMerge
            Write-Output "[MKVMerge] $(Get-Timestamp) - Checking for subtitle and video files to merge."
            Get-ChildItem $SiteSrcfolder -Recurse -File -Include "$VidType" | Sort-Object LastWriteTime | ForEach-Object {
                # grabbing associated variables needed to pass onto MKVMerge
                $MKVVidInput = $_.FullName
                $MKVVidBaseName = $_.BaseName
                $MKVVidSubtitle = Get-ChildItem $SiteSrcfolder -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $MKVVidBaseName } | Select-Object -First 1
                $MKVVidTempOutput = $_.DirectoryName + "\" + $_.BaseName + ".temp" + $_.Extension
                $MKVVidSubList = [ordered]@{MKVVidInput = $MKVVidInput; MKVVidBaseName = $MKVVidBaseName; MKVVidSubtitle = $MKVVidSubtitle; MKVVidTempOutput = $MKVVidTempOutput }
                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Checking for $MKVVidSubtitle and $MKVVidInput file to merge."
                $MKVVidSubList | Out-String
                # Only process files with matching subtitle
                if ($MKVVidSubtitle) {
                    # Adding custom styling to ASS subtitle
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
                            Rename-Item -Path $MKVVidTempOutput -NewName $_.FullName -Confirm:$false -Verbose
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                    While ($True) {
                        if ((Test-Lock $_.FullName) -eq $True) {
                            Write-Output "[MKVMerge] $(Get-Timestamp) -  $MKVVidInput File locked.  Waiting..."
                            continue
                        }
                        else {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - $MKVVidInput File not locked. Setting default subtitle."
                            mkvpropedit $_.FullName --edit track:s1 --set flag-default=1
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                    [void]$MKVCompletedFiles.Add($MKVVidInput)
                }
                else {
                    [void]$incompleteFiles.Add($MKVVidInput)
                    Write-Output "[MKVMerge] $(Get-Timestamp) - No matching subtitle files to process. Skipping file."
                    
                }
            }
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - No files to process"
        }
    }
    if ($incompleteFiles.count -gt 0) {
        # If $incomplete file is not empty/null then write out what files have an issue
        Write-Output "[MKVMerge] $(Get-Timestamp) - Not moving files. Script completed with ERRORS. The following files did not have matching subtitle file:"
        $incompleteFiles | Out-String
        break
    }
    else {
        Write-Output "[MKVMerge] $(Get-Timestamp)- All files had matching subtitle file"
        # Moving files from SiteSrc to SiteHome for Filebot processing
        if ((Test-Path -path $SiteTempBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc does not have any files. Removing folder..."
            Remove-Item $SiteSrc -Recurse -Force -Confirm:$false -Verbose
            $SiteSrcDeleted = $true
            Write-Output "[MKVMerge] $(Get-Timestamp) - SiteSrcDeleted = $SiteSrcDeleted"
        }
        elseif ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files."
            if ($SendTelegram) {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [Telegram] - Sending message for files in $SiteSrc."
                $TM = Get-SiteSeriesEpisode -VPath $SiteSrc -VType $VidType
                Send-Telegram -STMessage $TM | Out-Null
            }
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Moving to $SiteHomeBase..."
            Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Waiting to delete folder after move..."
        }
    }
}
else {
    Write-Output "[MKVMerge] $(Get-Timestamp) - [End] - Not running"
    # Moving files from SiteSrc to SiteHome for Filebot processing
    if ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -eq 0) {
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc does not have any files. Removing folder..."
        Remove-Item $SiteSrc -Recurse -Force -Confirm:$false -Verbose
    }
    elseif ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files."
        if ($SendTelegram) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [Telegram] - Sending message for files in $SiteSrc."
            $TM = Get-SiteSeriesEpisode -VPath $SiteSrc -VType $VidType
            Send-Telegram -STMessage $TM | Out-Null
        }
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Moving to $SiteHomeBase..."
        Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
    }
}
# If Filebot = True then run Filebot aginst SiteHome folder
if (($Filebot) -and ($incompleteFiles.count -eq 0)) {
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to renaming and move to final folder"
    ForEach ($FBfolder in $SiteHome ) {
        if ((Get-ChildItem $FBfolder -Recurse -Force -File -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Get-ChildItem $FBfolder -Recurse -File -Include "$VidType" | Sort-Object LastWriteTime | ForEach-Object {
                $FBVidInput = $_.FullName
                # Filebot command
                if ($PlexLibPath) {
                    Write-Output "[Filebot] $(Get-Timestamp) - Files found. Renaming and moving files to final folder"
                    filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
                }
                else {
                    Write-Output "[Filebot] $(Get-Timestamp) - Files found. Plex path not specified. Renaming files in place"
                    filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "{ plex.tail }" --log info
                }
                if (!(Test-Path $FBVidInput)) {
                    [void]$FBCompletedFiles.Add($FBVidInput)
                }
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - No files to process"
        }
    }
    $fbc = $FBCompletedFiles.count
    $mkvc = $MKVCompletedFiles.Count
    if ($FBCompletedFiles.count -eq $MKVCompletedFiles.Count) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($fbc) = ($mkvc)MKV Video.n No other files need to be processed. Attempting Filebot cleanup. Completed files:"
        $FBCompletedFiles | Out-String
        filebot -script fn:cleaner "$SiteHome" --log all
    }
    else {
        write-output "[Filebot] $(Get-Timestamp) - Filebot($fbc) and MKV Video($mkvc) count mismatch. Manual check required."
    }
    filebot -script fn:cleaner "$SiteHome" --log all
    # Check if folder is empty. If contains a video file file then exit, if not then completed successfully and continues
    if ((Get-ChildItem $FBfolder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
        if ($Daily) {
            Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - Daily run - Script completed with ERRORS"
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - Manual run - Script completed"
        }
    }
    else {
        if ($FBCompletedFiles.Count -gt 0) {
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
        else {
            Write-Output "[PLEX] $(Get-Timestamp) - No files processed. Skipping PLEX API call."
        }
    }
}
elseif (($Filebot) -and ($incompleteFiles)) {
    Write-Output "[Filebot] $(Get-Timestamp) - Files in $SiteSrc need manual attention. Skipping to next step... Incomplete files in $SiteSrc :"
    $incompleteFiles | Out-String
}
else {
    Write-Output "[Filebot] $(Get-Timestamp) - [End] - Not running Filebot"
}
# Backup of Archive, cookie, bat, config.xml, and font files
$SharedBackups = $ArchiveFile, $CookieFile, $BatFile, $ConfigPath, $SubFontDir
foreach ($sb in $SharedBackups) {
    if (($sb -ne "None") -and ($sb.trim() -ne "")) {
        if ($SubFontDir) {
            Copy-Item -Path $sb -Destination $SrcDriveSharedFonts -PassThru
            Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcDriveSharedFonts."
        }
        else {
            Write-Output "[FileBackup] $(Get-Timestamp) - Copying ($sb) to $SrcDriveShared."
            Copy-Item -Path $sb -Destination $SrcDriveShared -PassThru
        }
    }
}
# Regardless of failures still force delete tmp for clean runs
if (($SiteTemp -match "\\tmp\\") -and ($SiteTemp -match $SiteTempBaseMatch) -and (Test-Path $SiteTemp)) {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $SiteTemp folders/files"
    Remove-Item $SiteTemp -Recurse -Force -Confirm:$false -Verbose
}
else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteTemp($SiteTemp) folder already deleted. Nothing to remove."
}
# Clean up SiteSrc folder if empty
if ($SiteSrcDeleted) {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteSrc($SiteSrc) folder already deleted. Nothing to remove."
}
elseif (!($SiteSrcDeleted) -and ($SiteSrc -match "\\src\\") -and ($SiteSrc -match $SiteSrcBaseMatch) -and (Test-Path $SiteSrc) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteSrc($SiteSrc) contains files."
}
# Clean up SiteHome folder if empty
if (($SiteHome -match "\\tmp\\") -and ($SiteHome -match $SiteHomeBaseMatch) -and (Test-Path $SiteHome)) {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting SiteHome($SiteHome) if still present."
    & $DeleteRecursion -DRPath $SiteHome
}
else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteHome folder already deleted. Nothing to remove."
}
# End
Write-Output "[END] $(Get-Timestamp) - Script completed"