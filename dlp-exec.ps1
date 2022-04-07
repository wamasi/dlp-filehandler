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
# Update SE/MKV/FB true false
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
    $Telegrammessage = '<b>Site:</b> ' + $SiteNameRaw + "`n"
    $SeriesMessage = ''
    $SEL | ForEach-Object {
        $EpList = ''
        foreach ($i in $_) {
            $EpList = $_.Episode._VSEpisode | Out-String
        }
        $SeriesMessage = '<b>Series:</b> ' + $_.Series + "`n<b>Episode:</b>`n" + $EpList
        $Telegrammessage += $SeriesMessage + "`n"
    }
    Write-Host $Telegrammessage
    return $Telegrammessage
}
# Sending To telegram for new file notifications
Function Send-Telegram {
    Param([Parameter(Mandatory = $true)][String]$STMessage)
    $Telegramtoken
    $Telegramchatid
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    Invoke-WebRequest -Uri "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($STMessage)&parse_mode=html"
}
# Run MKVMerge process
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
            continue
        }
        else {
            if ($SF -ne 'None') {
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
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True -and (Test-Lock $MKVVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($SubFontDir -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - Combining $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput $MKVVidInput $MKVVidSubtitle --attach-file $SubFontDir --attachment-mime-type application/x-truetype-font
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) -  Merging as-is. No Font specified for $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                mkvmerge -o $MKVVidTempOutput $MKVVidInput $MKVVidSubtitle
            }
        }
        Start-Sleep -Seconds 1
    }
    While (!(Test-Path $MKVVidTempOutput -ErrorAction SilentlyContinue)) {
        Start-Sleep 1.5
    }
    While ($True) {
        if (((Test-Lock $MKVVidInput) -eq $True) -and ((Test-Lock $MKVVidSubtitle) -eq $True ) -and ((Test-Lock $MKVVidTempOutput) -eq $True)) {
            continue
        }
        else {
            Remove-Item -Path $MKVVidInput -Confirm:$false -Verbose
            Remove-Item -Path $MKVVidSubtitle -Confirm:$false -Verbose
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidTempOutput) -eq $True) {
            continue
        }
        else {
            Rename-Item -Path $MKVVidTempOutput -NewName $MKVVidInput -Confirm:$false -Verbose
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True) {
            continue
        }
        else {
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
        if ($PlexLibPath) {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found. Renaming and moving files to final folder"
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
            if (!($MKVMerge)) {
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format "{drive}\Videos\$PlexLibPath\{ plex.tail }" --log info
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found. Plex path not specified. Renaming files in place"
            filebot -rename "$FBVidInput" -r --db TheTVDB -non-strict --format '{ plex.tail }' --log info
            if (!($MKVMerge)) {
                filebot -rename "$FBSubInput" -r --db TheTVDB -non-strict --format '{ plex.tail }' --log info
            }
        }
        if (!(Test-Path $FBVidInput)) {
            Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $FBVidBaseName -SVSFP $true
        }
    }
    $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
    if ($VSVFBCount -eq $VSVTotCount ) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($VSVFBCount) = ($VSVTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup."
        filebot -script fn:cleaner "$SiteHome" --log all
    }
    else {
        Write-Output "[Filebot] $(Get-Timestamp) - Filebot($VSVFBCount) and Total Video($VSVTotCount) count mismatch. Manual check required."
    }
    if ($VSVFBCount -ne $VSVTotCount) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
    }
    else {
        if ($PlexHost -and $PlexToken -and $PlexLibId) {
            Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
            $PlexUrl = "$PlexHost/library/sections/$PlexLibId/refresh?X-Plex-Token=$PlexToken"
            Invoke-WebRequest -Uri $PlexUrl
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
        if (($RFFolder -match '\\tmp\\') -and ($RFFolder -match $RFBaseMatch) -and (Test-Path $RFFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $RFFolder folders/files"
            Remove-Item $RFFolder -Recurse -Force -Confirm:$false -Verbose
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteTemp($RFFolder) folder already deleted. Nothing to remove."
        }
    }
    else {
        if (!(Test-Path $RFFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) already deleted files."
        }
        elseif ((Test-Path $RFFolder) -and (Get-ChildItem $RFFolder -Recurse -File | Measure-Object).Count -eq 0) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) is empty. Deleting folder."
            & $DeleteRecursion -DRPath $RFFolder
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) contains files. Manual attention needed."
        }
    }
}
# Test if file is available to interact with
function Test-Lock {
    Param(
        [parameter(Mandatory = $true)]
        $TLfilename
    )
    $TLfile = Get-Item (Resolve-Path $TLfilename) -Force
    if ($TLfile -is [IO.FileInfo]) {
        trap {
            Write-Output "[FileLockCheck] $(Get-Timestamp) - $TLfile File locked. Waiting..."
            return $true
            continue
        }
        $TLstream = New-Object system.IO.StreamReader $TLfile
        if ($TLstream) { $TLstream.Close() }
    }
    Write-Output "[FileLockCheck] $(Get-Timestamp) - $TLfile File unlocked. Continuing..."
    return $false
}
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
$CreateFolders = $TempDrive, $SrcDrive, $SrcDriveShared, $SrcDriveSharedFonts, $DestDrive, $SiteTemp, $SiteSrc, $SiteHome
foreach ($c in $CreateFolders) {
    Set-Folders $c
}
$limit = (Get-Date).AddDays(-1)
if (!(Test-Path $LFolderBase)) {
    Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase is missing. Skipping log cleanup..."
}
else {
    Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting log cleanup..."
    Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.CreationTime -lt $limit } | ForEach-Object {
        $_.FullName | Remove-Item -Recurse -Force -Confirm:$false -Verbose
    }
    & $DeleteRecursion -DRPath $LFolderBase
}
Invoke-Expression $dlpParams
if ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
    Get-ChildItem $SiteSrc -Recurse -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -Unique | ForEach-Object {
        $VSSite = $SiteNameRaw
        $VSSeries = (Get-Culture).TextInfo.ToTitleCase(("$(Split-Path (Split-Path $_ -Parent) -Leaf)").Replace('_', ' ').Replace('-', ' ')) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
        $VSEpisode = (Get-Culture).TextInfo.ToTitleCase( ($_.BaseName.Replace('_', ' ').Replace('-', ' '))) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
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
        if (($_._VSEpisodeSubtitle.Trim() -eq '') -or ($null -eq $_._VSEpisodeSubtitle)) {
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
if ($VSVTotCount -gt 0) {
    if ($SubtitleEdit) {
        $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | Where-Object { $_._VSErrored -ne $true } | ForEach-Object {
            $SESubtitle = $_._VSEpisodeSubtitle
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $SESubtitle subtitle"
            While ($True) {
                if ((Test-Lock $SESubtitle) -eq $True) {
                    continue
                }
                else {
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
    if ($MKVMerge) {
        $VSCompletedFilesList | Select-Object _VSEpisodeRaw, _VSEpisode, _VSEpisodeTemp, _VSEpisodePath, _VSEpisodeSubtitle, _VSErrored | `
            Where-Object { $_._VSErrored -eq $false } | ForEach-Object {
            $MKVVidInput = $_._VSEpisodePath
            $MKVVidBaseName = $_._VSEpisodeRaw
            $MKVVidSubtitle = $_._VSEpisodeSubtitle
            $MKVVidTempOutput = $_._VSEpisodeTemp
            Start-MKVMerge $MKVVidInput $MKVVidBaseName $MKVVidSubtitle $MKVVidTempOutput
        }
    }
    else {
        Write-Output "[MKVMerge] $(Get-Timestamp) - MKVMerge not running. Moving to next step."
    }
    $VSVMKVCount = ($VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Measure-Object).Count
    if (($VSVMKVCount -eq $VSVTotCount -and $VSVErrorCount -eq 0) -or (!($MKVMerge) -and $VSVErrorCount -eq 0)) {
        Write-Output "[MKVMerge] $(Get-Timestamp)- All files had matching subtitle file"
        if ($SendTelegram) {
            Write-Output "[MKVMerge] $(Get-Timestamp) - [Telegram] - Sending message for files in $SiteSrc."
            $TM = Get-SiteSeriesEpisode
            Send-Telegram -STMessage $TM | Out-Null
        }
        Write-Output "[FolderCleanup] $(Get-Timestamp) - $SiteSrc contains files. Moving to $SiteHomeBase..."
        Move-Item -Path $SiteSrc -Destination $SiteHomeBase -Force -Verbose
    }
    else {
        Write-Output "[FolderCleanup] $(Get-Timestamp) - $SiteSrc contains file(s) with error(s). Not moving files."
    }
    if (($Filebot -and $VSVMKVCount -eq $VSVTotCount) -or ($Filebot -and !($MKVMerge))) {
        Start-Filebot -FBPath $SiteHome
    }
    elseif (($Filebot -and $MKVMerge -and $VSVMKVCount -ne $VSVTotCount)) {
        Write-Output "[Filebot] $(Get-Timestamp) - Files in $SiteSrc need manual attention. Skipping to next step... Incomplete files in $SiteSrc."
    }
    else {
        Write-Output "[Filebot] $(Get-Timestamp) - Not running Filebot"
    }
    Write-Output "[VideoList] $(Get-Timestamp) - Final file status:"
    $VSCompletedFilesList | Format-Table @{Label = 'Series'; Expression = { $_._VSSeries } }, @{Label = 'Episode'; Expression = { $_._VSEpisode } } , `
    @{Label = 'SECompleted'; Expression = { $_._VSSECompleted } }, @{Label = 'MKVCompleted'; Expression = { $_._VSMKVCompleted } }, @{Label = 'FBCompleted'; Expression = { $_._VSFBCompleted } }, `
    @{Label = 'Errored'; Expression = { $_._VSErrored } } -AutoSize -Wrap
}
else {
    Write-Output "[VideoList] $(Get-Timestamp) - No files downloaded. Skipping other defined steps."
}
$SharedBackups = $ArchiveFile, $CookieFile, $BatFile, $ConfigPath, $SubFontDir
foreach ($sb in $SharedBackups) {
    if (($sb -ne 'None') -and ($sb.trim() -ne '')) {
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
Remove-Folders -RFFolder $SiteTemp -RFMatch '\\tmp\\' -RFBaseMatch $SiteTempBaseMatch
Remove-Folders -RFFolder $SiteSrc -RFMatch '\\src\\' -RFBaseMatch $SiteSrcBaseMatch
Remove-Folders -RFFolder $SiteHome -RFMatch '\\tmp\\' -RFBaseMatch $SiteHomeBaseMatch
Write-Output "[END] $(Get-Timestamp) - Script completed"