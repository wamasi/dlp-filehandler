param ($dlpParams, $useFilebot, $useSubtitleEdit, $useMKVMerge, $SiteName, $SF, $SubFontDir, $PlexHost, $PlexToken, $PlexLibId, $LFolderBase, $SiteSrc, $SiteHome, $SiteTempBaseMatch, $SiteSrcBaseMatch, $SiteHomeBaseMatch, $ConfigPath )
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
SiteTempBaseMatch = $SiteTempBaseMatch
SiteSrc = $SiteSrc
SiteSrcBaseMatch = $SiteSrcBaseMatch
SiteHome = $SiteHome
SiteHomeBaseMatch = $SiteHomeBaseMatch
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
useSubtitleEdit = $useSubtitleEdit
useMKVMerge = $useMKVMerge
UseFilebot = $useFilebot
PlexHost = $PlexHost
PlexLibPath = $PlexLibPath
PlexLibId = $PlexLibId
ConfigPath = $ConfigPath
Script Directory = $ScriptDirectory
dlpParams = $dlpParams
"@
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
$completedF = ""
$incompletefile = ""
[bool] $SiteSrcDeleted = $false
If ($useSubtitleEdit) {
    # Fixing subs - SubtitleEdit
    If ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$SubType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | ForEach-Object {
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
            Get-ChildItem $folder -Recurse -File -Include "$VidType" | ForEach-Object {
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
                    $incompletefile += "$inputs`n"
                    Write-Output "[MKVMerge] $(Get-Timestamp) - No matching subtitle files to process. Skipping file."
                    
                }
            }
        }
        Else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - No files to process"
        }
    }
    If ($incompletefile) {
        # If $incomplete file is not empty/null then write out what files have an issue
        Write-Output @"
[MKVMerge] $(Get-Timestamp) - The following files did not have matching subtitle file: $incompletefile
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
            Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains files. Moving to $SiteHomeBase..."
            Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
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
        Write-Output "[MKVMerge] $(Get-Timestamp) - [FolderCleanup] - $SiteSrc contains foles. Moving to $SiteHomeBase..."
        Move-Item -Path $SiteSrc -Destination $SiteHomeBase -force -Verbose
    }
}
$incompletefile.Trim()
# If Filebot = True then run Filebot aginst SiteHome folder
If (($useFilebot) -and ($incompletefile.Trim() -eq "")) {
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to renaming and move to final folder"
    ForEach ($folder in $SiteHome ) {
        If ((Get-ChildItem $folder -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
            Get-ChildItem $folder -Recurse -File -Include "$VidType" | ForEach-Object {
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
                $completedF += "$inputs`n"
            }
        }
        Else {
            Write-Output "[Filebot] $(Get-Timestamp) - No files to process"
        }
    }
    $completedF.Trim()
    if ($completedF) {
        write-output @"
[Filebot] $(Get-Timestamp) - Completed files: `n$completedF
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
        If ($completedF) {
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
elseif (($useFilebot) -and ($incompletefile)) {
    Write-Output @"
[Filebot] $(Get-Timestamp) - Incomplete files in $SiteSrc\: `n$incompletefile
[Filebot] $(Get-Timestamp) - [End] - Files in $SiteSrc need manual attention. Skipping to next step...
"@
}
Else {
    Write-Output "[Filebot] $(Get-Timestamp) - [End] - Not running Filebot"
}
# Backup of Archive file
If ($ArchiveFile -ne "None") {
    Write-Output "[FileBackup] $(Get-Timestamp) - Copying $ArchiveFile to $SrcDrive."
    Copy-Item -Path $ArchiveFile -Destination "$SrcDrive\_shared" -PassThru
}
Else {
    Write-Output "[FileBackup] $(Get-Timestamp) - ArchiveFile is None. Nothing to copy..."
}
# Backup of Cookie file
If ($CookieFile -ne "None") {
    Write-Output "[FileBackup] $(Get-Timestamp) - Copying $CookieFile to $SrcDrive."
    Copy-Item -Path $CookieFile -Destination "$SrcDrive\_shared" -PassThru
}
Else {
    Write-Output "[FileBackup] $(Get-Timestamp) - CookieFile is None. Nothing to copy..."
}
# Backup of Bat file
Write-Output "[FileBackup] $(Get-Timestamp) - Copying $BatFile to $SrcDrive."
Copy-Item -Path $BatFile -Destination "$SrcDrive\_shared" -PassThru
# Backup of config.xml file
Write-Output "[FileBackup] $(Get-Timestamp) - Copying $ConfigPath to $SrcDrive."
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
If (($SiteSrc -match "\\src\\") -and ($SiteSrc -match $SiteSrcBaseMatch) -and (Test-Path $SiteSrc)) {
    & $DeleteRecursion -Path $SiteSrc
}
Else {
    if ($SiteSrcDeleted) {
        Write-Output "[FolderCleanup] $(Get-Timestamp) - $SiteSrc folder already removed."
    }
    else {
        Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteSrc($SiteSrc) folder not matching as expected. Remove not completed"
    }
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