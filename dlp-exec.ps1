param ($dlpParams, $useFilebot, $useSubtitleEdit, $SiteName, $SF, $SubFontDir, $PlexHost, $PlexToken, $PlexLibId, $LFolderBase, $SiteSrc, $SiteHome )
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
    Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Recurse -Force -Confirm:$false -Verbose
}
# Writing all variables
Write-Output @"
Site = $SiteName
SiteType = $SiteType
isDaily = $isDaily
SiteUser = $SiteUser
SitePass = $SitePass
UseLogin = $useLogin
SiteFolder = $SiteFolder
SiteTemp = $SiteTemp
SiteSrc = $SiteSrc
SiteHome = $SiteHome
SiteConfig = $SiteConfig
CookieFile = $CookieFile
UseCookies = $useCookies
Archive = $ArchiveFile
UseDownloadArchive = $useArchive
Bat = $BatFile
Ffmpeg = $Ffmpeg
SF = $SF
SubFont = $SubFont
SubFontDir = $SubFontDir
UseFilebot = $useFilebot
useSubtitleEdit = $useSubtitleEdit
PlexHost = $PlexHost
PlexLibPath = $PlexLibPath
PlexLibId = $PlexLibId
Script Directory = $PSScriptRoot
dlpParams = $dlpParams
"@
# Call to YT-DLP with parameters
Invoke-Expression $dlpParams
# If SubtitleEdit = True then run SubtitleEdit against SiteHome folder.
If ($useSubtitleEdit) {
    $incompletefile = ""
    ForEach ($folder in $SiteSrc) {
        If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0 -and (Get-ChildItem $folder -Recurse -Force -File -Include "$SubType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Processing files to fix subtitles and then combine with video."
            # Fixing subs - SubtitleEdit
            Get-ChildItem $folder -Recurse -File -Include "$SubType" | ForEach-Object {
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
            # Embedding subs into video files - ffmpeg
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Checking for subtitle and video files to merge."
            Get-ChildItem $folder -Recurse -File -Include "$VidType" | ForEach-Object {
                # grabbing associated variables needed to pass onto FFMPEG
                $inputs = $_.FullName
                $filename = $_.BaseName
                $subtitle = Get-ChildItem $folder -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $filename } | Select-Object -First 1
                $tempvideo = $_.DirectoryName + "\" + $_.BaseName + ".temp" + $_.Extension
                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Checking for $subtitle and $inputs file to merge."
                "[Input Filename]    = " + $inputs
                "[Base Filename]     = " + $filename
                "[Subtitle Filename] = " + $subtitle
                "[Temp Video]        = " + $tempvideo
                # Only process files with matching subtitle
                If ($subtitle) {
                    # Adding custom styling to ASS subtitle
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - Replacing Styling in $subtitle."
                    While ($True) {
                        If ((Test-Lock $subtitle) -eq $True) {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - $subtitle File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - File not locked. Formatting $subtitle file."
                            If ($SF -ne "None") {
                                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Python - Regex through $subtitle file with $SF."
                                python "$PSScriptRoot\subtitle_regex.py" $subtitle $SF
                                break
                            }
                            Else {
                                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Python - No Font specified for $subtitle file."
                            }
                        }
                        Start-Sleep -Seconds 1
                    }
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - Found matching  $subtitle and $inputs files to process."
                    #mkmerge command to combine video and subtitle file and set subtitle default
                    While ($True) {
                        If ((Test-Lock $inputs) -eq $True -and (Test-Lock $subtitle) -eq $True) {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - $subtitle and $inputs File locked.  Waiting..."
                            continue
                        }
                        Else {
                            If ($SubFontDir -ne "None") {
                                Write-Output "[SubtitleEdit] $(Get-Timestamp) - MKVMERGE - File not locked.  Combining $subtitle and $inputs files with $SubFontDir."
                                mkvmerge -o $tempvideo $inputs $subtitle --attach-file $SubFontDir --attachment-mime-type application/x-truetype-font
                                break
                            }
                            Else {
                                Write-Output "[SubtitleEdit] $(Get-Timestamp) - MKVMERGE - No Font specified for $subtitle and $inputs files with $SubFontDir."
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
                            Write-Output "[SubtitleEdit] $(Get-Timestamp)- File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - File not locked. Removing $inputs and $subtitle file."
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
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - $tempvideo File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - File not locked. Renaming $tempvideo to $inputs file."
                            # Remove original video/subtitle file
                            Rename-Item -Path $tempvideo -NewName $_.FullName -Confirm:$false -Verbose
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                    While ($True) {
                        If ((Test-Lock $_.FullName) -eq $True) {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) -  $inputs File locked.  Waiting..."
                            continue
                        }
                        Else {
                            Write-Output "[SubtitleEdit] $(Get-Timestamp) - $inputs File not locked. Setting default subtitle."
                            mkvpropedit $_.FullName --edit track:s1 --set flag-default=1
                            break
                        }
                        Start-Sleep -Seconds 1
                    }
                }
                Else {
                    $incompletefile += "$inputs`n"
                    Write-Output "[SubtitleEdit] $(Get-Timestamp) - No matching subtitle files to process. Skipping file."
                    
                }
            }
        }
        Else {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - No files to process"
        }
    }
    If ($incompletefile) {
        # If $incomplete file is not empty/null then write out what files have an issue
        Write-Output @"
[SubtitleEdit] $(Get-Timestamp) - The following files did not have matching subtitle file: `n$incompletefile
[SubtitleEdit] $(Get-Timestamp) - Script completed with ERRORS
[END] $(Get-Timestamp) - Script finished incompleted.
"@
        Exit
    }
    Else {
        Write-Output "[SubtitleEdit] $(Get-Timestamp)- All files had matching subtitle file"
        # Moving files from SiteSrc to SiteHome for Filebot processing
        If ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - $SiteSrc does not have any files. Removing folder..."
            Remove-Item $SiteSrc -Recurse -Force -Confirm:$false -Verbose
        }
        elseIf ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - $SiteSrc contains foles. Moving to $SiteHomeBase..."
            Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
        }
    }
}
Else {
    Write-Output "[SubtitleEdit] $(Get-Timestamp) - [End] - Not running"
    # Moving files from SiteSrc to SiteHome for Filebot processing
    If ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -eq 0) {
        Write-Output "[SubtitleEdit] $(Get-Timestamp) - $SiteSrc does not have any files. Removing folder..."
        Remove-Item $SiteSrc -Recurse -Force -Confirm:$false -Verbose
    }
    elseIf ((Test-Path -path $SiteHomeBase) -and (Get-ChildItem $SiteSrc -Recurse -File | Measure-Object).Count -gt 0) {
        Write-Output "[SubtitleEdit] $(Get-Timestamp) - $SiteSrc contains foles. Moving to $SiteHomeBase..."
        Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
    }
}
# If Filebot = True then run Filebot aginst SiteHome folder
If ($useFilebot) {
    $completedF = ""
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to renaming and move to final folder"
    ForEach ($folder in $SiteHome ) {
        If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Get-ChildItem $folder -Recurse -File -Include "$VidType" | ForEach-Object {
                $inputs = $_.FullName
                Write-Output "[Filebot] $(Get-Timestamp) - Files found. Renaming and moving files to final folder"
                # Filebot command
                If ($PlexLibPath) {
                    filebot -rename "$inputs" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
                }
                Else {
                    filebot -rename "$inputs" -r --db TheTVDB -non-strict --format "{ plex.tail }" --log info
                }
                $completedF += "$inputs`n"
            }
        }
        Else {
            Write-Output "[Filebot] $(Get-Timestamp) - No files to process"
        }
    }
    $completedF.Trim()
    $incompletefile.Trim()
    write-output @"
[Filebot] $(Get-Timestamp) - Completed files: `n$completedF
[Filebot] $(Get-Timestamp) - Incomplete files: `n$incompletefile
[Filebot] $(Get-Timestamp) - No other files need to be processed. Attempting Filebot cleanup
"@
    filebot -script fn:cleaner "$SiteHome" --log all
    # Check if folder is empty. If contains a video file file then exit, if not then completed successfully and continues
    If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
        If ($isDaily) {
            Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - Daily run - Script completed with ERRORS"
            break
        }
        Else {
            Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - Manual run - Script completed"
        }
    }
    Else {
        If ($completedF) {
            # If plex values not null then run api call else skip
            If ($PlexHost -and $PlexToken -and $PlexLibId) {
                Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
                $PlexUrl = "$PlexHost/library/sections/$PlexLibId/refresh?X-Plex-Token=$PlexToken"
                Invoke-RestMethod -UseBasicParsing -Verbose $PlexUrl
            }
            Else {
                Write-Output "[PLEX] $(Get-Timestamp) - [End] - Not using Plex."
                break
            }
        }
        Else {
            Write-Output "[PLEX] $(Get-Timestamp) - No files processed. Skipping PLEX API call."
        }
    }
}
Else {
    Write-Output "[Filebot] $(Get-Timestamp) - [End] - Not running Filebot"
}
If ($ArchiveFile -ne "None") {
    Write-Output "[FileBackup] $(Get-Timestamp) - Copying $ArchiveFile and $BatFile to $SrcDrive."
    Copy-Item -Path $ArchiveFile -Destination "$SrcDrive\_shared" -PassThru
    Copy-Item -Path $BatFile -Destination "$SrcDrive\_shared" -PassThru
} Else {
    Write-Output "[FileBackup] $(Get-Timestamp) - $ArchiveFile is None. Nothing to copy..."
}
# Regardless of failures still force delete temp for clean runs
If ($SiteTemp -match "\\tmp\\") {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $SiteTemp folders/files"
    Get-ChildItem -Path $SiteTemp -Recurse -Force | ForEach-Object {
        $_.FullName | Remove-Item -Recurse -Force -Confirm:$false -Verbose
    }
    # Delete any empty directories left behind after deleting the old files.
    Get-ChildItem -Path $SiteTempBase -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Recurse -Force -Confirm:$false -Verbose
}
Else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Temp folder not matching as expected. Remove not completed"
}
# Clean up Home destination folder if empty
If ($SiteHomeBase -match "\\tmp\\") {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $SiteHomeBase folders/files if empty"
    Get-ChildItem -Path $SiteHomeBase -Recurse -Force | Where-Object { $_.PSIsContainer -and (Get-ChildItem -Path $_.FullName -Recurse -Force | Where-Object { !$_.PSIsContainer }) -eq $null } | Remove-Item -Recurse -Force -Confirm:$false -Verbose
}
Else {
    Write-Output "[FolderCleanup] $(Get-Timestamp) - Home folder not matching as expected. Remove not completed"
}
# End
Write-Output "[END] $(Get-Timestamp) - Script completed"