param ($dlpParams, $useFilebot, $useSubtitleEdit, $useMKVMerge, $SiteName, $SF, $SubFontDir, $PlexHost, $PlexToken, $PlexLibId, $LFolderBase, $SiteSrc, $SiteHome, $SiteTempBaseMatch, $SiteSrcBaseMatch, $SiteHomeBaseMatch, $ConfigPath )

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

# Function to check if file is locked by process before moving forward
function Test-Lock {
    Param(
        [parameter(Mandatory = $true)]
        $filename
    )
    $file = Get-Item (Resolve-Path $filename) -Force
    If ($file -is [IO.FileInfo]) {
        trap {
            return $true
            continue
        }
        $stream = New-Object system.IO.StreamReader $file
        If ($stream) { $stream.Close() }
    }
    return $false
}
# Function to recurse delete empty folders
$DeleteRecursion = {
    param(
        $Path
    )
    foreach ($childDirectory in Get-ChildItem -Force -LiteralPath $Path -Directory) {
        & $DeleteRecursion -Path $childDirectory.FullName
    }
    $currentChildren = Get-ChildItem -Force -LiteralPath $Path
    $isEmpty = $currentChildren -eq $null
    if ($isEmpty) {
        Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting '${Path}' folders/files if empty"
        Remove-Item -Force -LiteralPath $Path -Verbose
    }
}
# Getting list of Site, Series, and Episodes for Telegram messages
function Get-SiteSeriesEpisode {
    param (
        [parameter(Mandatory = $true)]
        $Path,
        [parameter(Mandatory = $true)]
        $VidType
    )
    Get-ChildItem $Path -Recurse -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -Unique | ForEach-Object {
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
    $Telegrammessage = "<b>Site:</b> " + $Site
    $Tmessage = ""
    $SEL | ForEach-Object {
        $EpList = ""
        foreach ($i in $_) {
            $EpList = $_.Episode._Episode | Out-String
        }
        $Tmessage = "`n<b>Series:</b> " + $_.Series + "`n<b>Episode:</b>`n" + $EpList
        $Telegrammessage += $Tmessage + "`n"
    }
    Write-Host $Telegrammessage
    return $Telegrammessage
}
# Optional sending To telegram for new file notifications
Function Send-Telegram {
    Param([Parameter(Mandatory = $true)][String]$Message)
    $Telegramtoken
    $Telegramchatid
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($Message)&parse_mode=html"
}
$completedFiles = ""
$incompleteFiles = ""
[bool] $SiteSrcDeleted = $false
# Depending on if isDaily is set will use appropriate files and setup temp/home directory paths
Set-Folders $TempDrive
Set-Folders $SrcDrive
Set-Folders $SrcDriveShared
Set-Folders $SrcDriveSharedFonts
Set-Folders $DestDrive
Set-Folders $SiteTemp
Set-Folders $SiteSrc
Set-Folders $SiteHome
# Deleting logs older than 1 days
Write-Output "[LogCleanup] $(Get-Timestamp) - Deleting old logfiles"
$limit = (Get-Date).AddDays(-1)
# Delete files older than the $limit.
If (!(Test-Path $LFolderBase)) {
    Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase is missing. Skipping log cleanup..."
    break
}
Else {
    Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting log cleanup..."
    # Delete any 1 day old log files.
    Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | ForEach-Object {
        $_.FullName | Remove-Item -Recurse -Force -Confirm:$false -Verbose
    }
    # Delete any empty directories left behind after deleting the old files.
    & $DeleteRecursion -Path $LFolderBase
}
# Call to YT-DLP with parameters
Invoke-Expression $dlpParams
# If useSubtitleEdit = True then run SubtitleEdit against SiteSrc folder.
If ($useSubtitleEdit) {
    # Fixing subs - SubtitleEdit
    If ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$SubType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Sort-Object LastWriteTime | ForEach-Object {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $_ subtitle"
            $subvar = $_.FullName
            # SubtitleEdit does not play well being called within a script
            While ($True) {
                If ((Test-Lock $subvar) -eq $True) {
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - File locked. Waiting..."
                    continue
                }
                Else {
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - File not locked. Editing $subvar file."
                    # Remove original video/subtitle file
                    powershell "SubtitleEdit /convert '$subvar' AdvancedSubStationAlpha /overwrite /MergeSameTimeCodes"
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
Else {
    Write-Output "[SubtitleEdit] $(Get-Timestamp) - Not running"
}
# If useMKVMerge = True then run MKVMerge against SiteSrc folder.
If ($useMKVMerge) {
    ForEach ($folder in $SiteSrc) {
        If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0 -and (Get-ChildItem $folder -Recurse -Force -File -Include "$SubType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - Processing files to fix subtitles and then combine with video."
            # Embedding subs into video files - MKVMerge
            Write-Output "[MKVMerge] $(Get-Timestamp) - Checking for subtitle and video files to merge."
            Get-ChildItem $folder -Recurse -File -Include "$VidType" | Sort-Object LastWriteTime | ForEach-Object {
                # grabbing associated variables needed to pass onto MKVMerge
                $inputs = $_.FullName
                $filename = $_.BaseName
                $subtitle = Get-ChildItem $folder -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $filename } | Select-Object -First 1
                $tempvideo = $_.DirectoryName + "\" + $_.BaseName + ".temp" + $_.Extension
                Write-Output @"
[SubtitleEdit] $(Get-Timestamp) - Checking for $subtitle and $inputs file to merge.
[Input Filename]    = $inputs
[Base Filename]     = $filename
[Subtitle Filename] = $subtitle
[Temp Video]        = $tempvideo
"@
                # Only process files with matching subtitle
                If ($subtitle) {
                    # Adding custom styling to ASS subtitle
                    Write-Output "[MKVMerge] $(Get-Timestamp) - Replacing Styling in $subtitle."
                    While ($True) {
                        If ((Test-Lock $subtitle) -eq $True) {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - $subtitle File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - File not locked. Formatting $subtitle file."
                            If ($SF -ne "None") {
                                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - Python - Regex through $subtitle file with $SF."
                                python $SubtitleRegex $subtitle $SF
                                break
                            }
                            Else {
                                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - No Font specified for $subtitle file."
                            }
                        }
                        Start-Sleep -Seconds 1
                    }
                    Write-Output "[MKVMerge] $(Get-Timestamp) - Found matching  $subtitle and $inputs files to process."
                    #mkmerge command to combine video and subtitle file and set subtitle default
                    While ($True) {
                        If ((Test-Lock $inputs) -eq $True -and (Test-Lock $subtitle) -eq $True) {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - $subtitle and $inputs File locked.  Waiting..."
                            continue
                        }
                        Else {
                            If ($SubFontDir -ne "None") {
                                Write-Output "[MKVMerge] $(Get-Timestamp) - [MKVMERGE] - File not locked.  Combining $subtitle and $inputs files with $SubFontDir."
                                mkvmerge -o $tempvideo $inputs $subtitle --attach-file $SubFontDir --attachment-mime-type application/x-truetype-font
                                break
                            }
                            Else {
                                Write-Output "[MKVMerge] $(Get-Timestamp) - [MKVMERGE] - Merging as-is. No Font specified for $subtitle and $inputs files with $SubFontDir."
                                mkvmerge -o $tempvideo $inputs $subtitle
                            }
                        }
                        Start-Sleep -Seconds 1
                    }
                    # If file doesn't exist yet then wait
                    While (!(Test-Path $tempvideo -ErrorAction SilentlyContinue)) {
                        Start-Sleep 1.5
                    }
                    # Wait for files to input, subtitle, and tempvideo to be ready
                    While ($True) {
                        If (((Test-Lock $inputs) -eq $True) -and ((Test-Lock $subtitle) -eq $True ) -and ((Test-Lock $tempvideo) -eq $True)) {
                            Write-Output "[MKVMerge] $(Get-Timestamp)- File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - File not locked. Removing $inputs and $subtitle file."
                            # Remove original video/subtitle file
                            Remove-Item -Path $inputs -Confirm:$false -Verbose
                            Remove-Item -Path $subtitle -Confirm:$false -Verbose
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                    # Rename temp to original filename
                    While ($True) {
                        If ((Test-Lock $tempvideo) -eq $True) {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - $tempvideo File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - File not locked. Renaming $tempvideo to $inputs file."
                            # Remove original video/subtitle file
                            Rename-Item -Path $tempvideo -NewName $_.FullName -Confirm:$false -Verbose
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                    While ($True) {
                        If ((Test-Lock $_.FullName) -eq $True) {
                            Write-Output "[MKVMerge] $(Get-Timestamp) -  $inputs File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[MKVMerge] $(Get-Timestamp) - $inputs File not locked. Setting default subtitle."
                            mkvpropedit $_.FullName --edit track:s1 --set flag-default=1
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                }
                Else {
                    $incompleteFiles += "$inputs`n"
                    Write-Output "[MKVMerge] $(Get-Timestamp) - No matching subtitle files to process. Skipping file."
                    
                }
            }
        }
        Else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - No files to process"
        }
    }
    If ($incompleteFiles) {
        # If $incomplete file is not empty/null then write out what files have an issue
        Write-Output @"
[MKVMerge] $(Get-Timestamp) - The following files did not have matching subtitle file: $incompleteFiles
[MKVMerge] $(Get-Timestamp) - Not moving files. Script completed with ERRORS
[END] $(Get-Timestamp) - Script finished incompleted.
"@
        break
    }
    Else {
        Write-Output "[MKVMerge] $(Get-Timestamp)- All files had matching subtitle file"
        # Moving files from SiteSrc to SiteHome for Filebot processing
        If ((Test-Path -path $SiteTempBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc does not have any files. Removing folder..."
            Remove-Item $SiteSrc -Recurse -Force -Confirm:$false -Verbose
            $SiteSrcDeleted = $true
            Write-Output "[MKVMerge] $(Get-Timestamp) - SiteSrcDeleted = $SiteSrcDeleted"
        }
        elseIf ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files."
            if ($SendTelegram) {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [Telegram] - Sending message for files in $SiteSrc."
                $TM = Get-SiteSeriesEpisode -Path $SiteSrc -VidType $sVidType
                Send-Telegram -Message $TM | Out-Null
            }
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Moving to $SiteHomeBase..."
            Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Waiting to delete folder after move..."
        }
    }
}
Else {
    Write-Output "[MKVMerge] $(Get-Timestamp) - [End] - Not running"
    # Moving files from SiteSrc to SiteHome for Filebot processing
    If ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -eq 0) {
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc does not have any files. Removing folder..."
        Remove-Item $SiteSrc -Recurse -Force -Confirm:$false -Verbose
    }
    elseIf ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files."
        if ($SendTelegram) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [Telegram] - Sending message for files in $SiteSrc."
            $TM = Get-SiteSeriesEpisode -Path $SiteSrc -VidType $sVidType
            Send-Telegram -Message $TM | Out-Null
        }
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Moving to $SiteHomeBase..."
        Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
    }
}
$incompleteFiles.Trim()
# If Filebot = True then run Filebot aginst SiteHome folder
If (($useFilebot) -and ($incompleteFiles.Trim() -eq "")) {
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to renaming and move to final folder"
    ForEach ($folder in $SiteHome ) {
        If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Get-ChildItem $folder -Recurse -File -Include "$VidType" | Sort-Object LastWriteTime | ForEach-Object {
                $inputs = $_.FullName
                # Filebot command
                If ($PlexLibPath) {
                    Write-Output "[Filebot] $(Get-Timestamp) - Files found. Renaming and moving files to final folder"
                    filebot -rename "$inputs" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
                }
                Else {
                    Write-Output "[Filebot] $(Get-Timestamp) - Files found. Plex path not specified. Renaming files in place"
                    filebot -rename "$inputs" -r --db TheTVDB -non-strict --format "{ plex.tail }" --log info
                }
                $completedFiles += "$inputs`n"
            }
        }
        Else {
            Write-Output "[Filebot] $(Get-Timestamp) - No files to process"
        }
    }
    $completedFiles.Trim()
    if ($completedFiles) {
        write-output @"
[Filebot] $(Get-Timestamp) - Completed files: `n$completedFiles
[Filebot] $(Get-Timestamp) - No other files need to be processed. Attempting Filebot cleanup.
"@
    }
    else {
        write-output "[Filebot] $(Get-Timestamp) - No file to process. Attempting Filebot cleanup."
    }
    filebot -script fn:cleaner "$SiteHome" --log all
    # Check if folder is empty. If contains a video file file then exit, if not then completed successfully and continues
    If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
        If ($isDaily) {
            Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - Daily run - Script completed with ERRORS"
        }
        Else {
            Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - Manual run - Script completed"
        }
    }
    Else {
        If ($completedFiles) {
            # If plex values not null then run api call else skip
            If ($PlexHost -and $PlexToken -and $PlexLibId) {
                Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
                $PlexUrl = "$PlexHost/library/sections/$PlexLibId/refresh?X-Plex-Token=$PlexToken"
                Invoke-RestMethod -UseBasicParsing -Verbose $PlexUrl
            }
            Else {
                Write-Output "[PLEX] $(Get-Timestamp) - [End] - Not using Plex."
            }
        }
        Else {
            Write-Output "[PLEX] $(Get-Timestamp) - No files processed. Skipping PLEX API call."
        }
    }
}
elseif (($useFilebot) -and ($incompleteFiles)) {
    Write-Output @"
[Filebot] $(Get-Timestamp) - Incomplete files in $SiteSrc\: `n$incompleteFiles
[Filebot] $(Get-Timestamp) - [End] - Files in $SiteSrc need manual attention. Skipping to next step...
"@
}
Else {
    Write-Output "[Filebot] $(Get-Timestamp) - [End] - Not running Filebot"
}
# Backup of Archive file
If (($ArchiveFile.trim() -ne "") -and ($ArchiveFile -ne "None")) {
    Write-Output "[FileBackup] $(Get-Timestamp) - Copying Archive($ArchiveFile) to $SrcDrive."
    Copy-Item -Path $ArchiveFile -Destination "$SrcDrive\_shared" -PassThru
}
Else {
    Write-Output "[FileBackup] $(Get-Timestamp) - ArchiveFile is None. Nothing to copy..."
}
# Backup of Cookie file
If (($CookieFile.trim() -ne "") -and ($CookieFile -ne "None")) {
    Write-Output "[FileBackup] $(Get-Timestamp) - Copying Cookie($CookieFile) to $SrcDrive."
    Copy-Item -Path $CookieFile -Destination "$SrcDrive\_shared" -PassThru
}
Else {
    Write-Output "[FileBackup] $(Get-Timestamp) - CookieFile is None. Nothing to copy..."
}
# Backup of font file
if (($SubFontDir.trim() -ne "") -and ($SubFontDir -ne "None")) {
    Write-Output "[FileBackup] $(Get-Timestamp) - Copying Font($SubFontDir) to $SrcDrive."
    Copy-Item -Path $SubFontDir -Destination "$SrcDrive\_shared\fonts" -PassThru
}
else {
    Write-Output "[FileBackup] $(Get-Timestamp) - SubFontDir is None. Nothing to copy..."
}
# Backup of Bat file
Write-Output "[FileBackup] $(Get-Timestamp) - Copying Bat($BatFile) to $SrcDrive."
Copy-Item -Path $BatFile -Destination "$SrcDrive\_shared" -PassThru
# Backup of config.xml file
Write-Output "[FileBackup] $(Get-Timestamp) - Copying Config($ConfigPath) to $SrcDrive."
Copy-Item -Path $ConfigPath -Destination "$SrcDrive\_shared" -PassThru
# Regardless of failures still force delete tmp for clean runs
If (($SiteTemp -match "\\tmp\\") -and ($SiteTemp -match $SiteTempBaseMatch) -and (Test-Path $SiteTemp)) {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $SiteTemp folders/files"
    Remove-Item $SiteTemp -Recurse -Force -Confirm:$false -Verbose
}
Else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Temp folder not matching as expected. Remove not completed"
}
# Clean up SiteSrc folder if empty
If (($SiteSrc -match "\\src\\") -and ($SiteSrc -match $SiteSrcBaseMatch) -and (Test-Path $SiteSrc) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteSrc($SiteSrc) contains files."
}
else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteSrc($SiteSrc) folder not matching as expected. Remove not completed"
}
# Clean up SiteHome folder if empty
If (($SiteHome -match "\\tmp\\") -and ($SiteHome -match $SiteHomeBaseMatch) -and (Test-Path $SiteHome)) {
    & $DeleteRecursion -Path $SiteHome
}
Else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteHome folder not matching as expected. Remove not completed"
}
# End
Write-Output "[END] $(Get-Timestamp) - Script completed"