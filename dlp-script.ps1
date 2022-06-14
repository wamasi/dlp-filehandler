<#
.Synopsis
   Script to run yt-dlp, mkvmerge, subtitle edit, filebot and a python script for downloading and processing videos
.EXAMPLE
   Runs the script using the crunchyroll as a manual run with the defined config using login, cookies, archive file, mkvmerge, and sends out a telegram message
   D:\_DL\dlp-script.ps1 -sn crunchyroll -l -c -mk -a -st
.EXAMPLE
   Runs the script using the crunchyroll as a daily run with the defined config using login, cookies, no archive file, and filebot/plex
   D:\_DL\dlp-script.ps1 -sn crunchyroll -d -l -c -f
.NOTES
   See https://github.com/wamasi/dlp-filehandler for full details
   Script was designed to be ran via powershell console on a cronjob. copying and pasting into powershell console will not work.
#>
[CmdletBinding()]
param(
    [Parameter(ParameterSetName = 'Help', Mandatory = $True)]
    [Alias('H')]
    [switch]$Help,
    [Parameter(ParameterSetName = 'NewConfig', Mandatory = $True)]
    [Alias('NC')]
    [switch]$NewConfig,
    [Alias('SU')]
    [Parameter(ParameterSetName = 'SupportFiles', Mandatory = $True)]
    [switch]$SupportFiles,
    [Alias('SN')]
    [ValidateScript({ if (Test-Path -Path "$PSScriptRoot\config.xml") {
                if (([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName -contains $_ ) {
                    $true
                }
                else {
                    $Test = (([xml](Get-Content -Path "$PSScriptRoot\config.xml")).getElementsByTagName('site').siteName) -join ', '
                    throw "The following Sites are valid: $Test"
                }
            }
            else {
                throw "No valid config.xml found in $PSScriptRoot. Run ($PSScriptRoot\dlp-script.ps1 -nc) for a new config file."
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $True)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $True)]
    [string]$Site,
    [Alias('D')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$Daily,
    [Alias('L')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$Login,
    [Alias('C')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$Cookies,
    [Alias('A')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$Archive,
    [Alias('SE')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$SubtitleEdit,
    [Alias('MK')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$MKVMerge,
    [Alias('F')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$Filebot,
    [Alias('ST')]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [switch]$SendTelegram,
    [Alias('AL')]
    [ValidateScript({
            $LangValues = 'ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und'
            if ($_ -in $LangValues) {
                $true
            }
            else {
                throw "Value '{0}' is invalid. The following languages are valid: {1}." -f $_, $($LangValues -join ', ')
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$AudioLang,
    [Alias('SL')]
    [ValidateScript({
            $LangValues = 'ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und'
            if ($_ -in $LangValues ) {
                $true
            }
            else {
                
                throw "Value '{0}' is invalid. The following languages are valid: {1}." -f $_, $($LangValues -join ', ')
            }
        })]
    [Parameter(ParameterSetName = 'Site', Mandatory = $false)]
    [Parameter(ParameterSetName = 'Test', Mandatory = $false)]
    [string]$SubtitleLang,
    [Alias('T')]
    [Parameter(ParameterSetName = 'Test', Mandatory = $true)]
    [switch]$TestScript
)
# Timer for script
$ScriptStopWatch = [System.Diagnostics.Stopwatch]::StartNew()
# Setting styling to remove error characters and width
$PSStyle.OutputRendering = 'Host'
$Width = $host.UI.RawUI.MaxPhysicalWindowSize.Width
$host.UI.RawUI.BufferSize = New-Object System.Management.Automation.Host.size($Width, 9999)
mode con cols=9999
function Get-Day {
    return (Get-Date -Format 'yy-MM-dd')
}
function Get-TimeStamp {
    return (Get-Date -Format 'yy-MM-dd HH-mm-ss')
}
function Get-Time {
    return (Get-Date -Format 'MMddHHmmssfff')
}
function Invoke-ExpressionConsole {
    param (
        [Parameter(Mandatory = $true)]
        [Alias('SCMFN')]
        [string]$SCMFunctionName,
        [Parameter(Mandatory = $true)]
        [Alias('SCMFP')]
        [string]$SCMFunctionParams
    )
    $SCMArg = "$SCMFunctionParams *>&1"
    $SCMObj = Invoke-Expression $SCMArg
    $SCMobj | Where-Object { $_.length -gt 0 } | ForEach-Object {
        Write-Output "[$SCMFunctionName] $(Get-TimeStamp) - $_"
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
            Write-Output "[FileLockCheck] $(Get-Timestamp) - $TLfile File locked. Waiting."
            return $true
            continue
        }
        $TLstream = New-Object system.IO.StreamReader $TLfile
        if ($TLstream) { $TLstream.Close() }
    }
    Write-Output "[FileLockCheck] $(Get-Timestamp) - $TLfile File unlocked. Continuing."
    return $false
}
function New-Folder {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Fullpath
    )
    if (!(Test-Path -Path $Fullpath)) {
        New-Item -ItemType Directory -Path $Fullpath -Force | Out-Null
        Write-Output "[SetFolder] - $(Get-Timestamp) - $Fullpath missing. Creating."
    }
    else {
        Write-Output "[SetFolder] - $(Get-Timestamp) - $Fullpath already exists."
    }
}
function New-SuppFiles {
    param (
        [Parameter(Mandatory = $true)]
        [string] $SuppFiles
    )
    if (!(Test-Path $SuppFiles -PathType Leaf)) {
        New-Item $SuppFiles -ItemType File | Out-Null
        Write-Output "[NewSupportFiles] - $(Get-Timestamp) - $SuppFiles file missing. Creating."
    }
    else {
        Write-Output "[NewSupportFiles] - $(Get-Timestamp) - $SuppFiles file already exists."
    }
}
function New-Config {
    param (
        [Parameter(Mandatory = $true)]
        [string] $Configs
    )
    New-Item $Configs -ItemType File -Force | Out-Null
    Write-Output "[NewConfigFiles] - $(Get-Timestamp) - Creating $Configs"
    if ($Configs -match 'vrv') {
        $vrvconfig | Set-Content $Configs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $Configs created with VRV values."
    }
    elseif ($Configs -match 'crunchyroll') {
        $crunchyrollconfig | Set-Content $Configs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $Configs created with Crunchyroll values."
    }
    elseif ($Configs -match 'funimation') {
        $funimationconfig | Set-Content $Configs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $Configs created with Funimation values."
    }
    elseif ($Configs -match 'hidive') {
        $hidiveconfig | Set-Content $Configs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $Configs created with Hidive values."
    }
    elseif ($Configs -match 'paramountplus') {
        $paramountplusconfig | Set-Content $Configs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $Configs created with ParamountPlus values."
    }
    else {
        $defaultconfig | Set-Content $Configs
        Write-Output "[NewConfigFiles] - $(Get-Timestamp) - $Configs created with default values."
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
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting $RFFolder folders/files."
            #Remove-Item $RFFolder -Recurse -Force -Verbose | Out-Null
            Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Item `"$RFFolder`" -Recurse -Force -Verbose"
        }
        else {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - SiteTemp($RFFolder) folder already deleted. Nothing to remove."
        }
    }
    else {
        if (!(Test-Path $RFFolder)) {
            Write-Output "[FolderCleanup] $(Get-Timestamp) - Folder($RFFolder) already deleted."
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
function Remove-Logfiles {
    # Log cleanup
    $FilledLogslimit = (Get-Date).AddDays(-$FilledLogs)
    $EmptyLogslimit = (Get-Date).AddDays(-$EmptyLogs)
    if (!(Test-Path $LFolderBase)) {
        Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase is missing. Skipping log cleanup."
    }
    else {
        Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting Filledlog($FilledLogs days) cleanup."
        Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -match '.*-Total-.*' -and $_.FullName -ne $LFile -and $_.CreationTime -lt $FilledLogslimit } | `
            ForEach-Object {
            $RLLog = $_.FullName
            Invoke-ExpressionConsole -SCMFN 'LogCleanup' -SCMFP "`"$RLLog`" | Remove-Item -Recurse -Force -Verbose"
        }
        Write-Output "[LogCleanup] $(Get-Timestamp) - $LFolderBase found. Starting emptylog($EmptyLogs days) cleanup."
        Get-ChildItem -Path $LFolderBase -Recurse -Force | Where-Object { !$_.PSIsContainer -and $_.FullName -notmatch '.*-Total-.*' -and $_.FullName -ne $LFile -and $_.CreationTime -lt $EmptyLogslimit } | `
            ForEach-Object {
            $RLLog = $_.FullName
            Invoke-ExpressionConsole -SCMFN 'LogCleanup' -SCMFP "`"$RLLog`" | Remove-Item -Recurse -Force -Verbose"
        }
        & $DeleteRecursion -DRPath $LFolderBase
    }
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
        Write-Output "[FolderCleanup] $(Get-Timestamp) - Force deleting '${DRPath}' folders/files if empty."
        Invoke-ExpressionConsole -SCMFN 'FolderCleanup' -SCMFP "Remove-Item -Force -LiteralPath `"$DRPath`" -Verbose"
    }
}
function Remove-Spaces {
    param (
        [Parameter(Mandatory = $true)]
        [string] $File
    )
    (Get-Content $File) | Where-Object { -not [String]::IsNullOrWhiteSpace($_) } | Set-Content $File
    $content = [System.IO.File]::ReadAllText($File)
    $content = $content.Trim()
    [System.IO.File]::WriteAllText($File, $content)
}
function Exit-Script {
    param (
        [alias('ES')]
        [switch]$ExitScript
    )
    $ScriptStopWatch.Stop()
    # Cleanup folders
    Remove-Folders -RFFolder $SiteTemp -RFMatch '\\tmp\\' -RFBaseMatch $SiteTempBaseMatch
    Remove-Folders -RFFolder $SiteSrc -RFMatch '\\src\\' -RFBaseMatch $SiteSrcBaseMatch
    Remove-Folders -RFFolder $SiteHome -RFMatch '\\tmp\\' -RFBaseMatch $SiteHomeBaseMatch
    if ($OverrideDriveList.count -gt 0) {
        foreach ($ORDriveList in $OverrideDriveList) {
            $ORDriveListBaseMatch = ($ORDriveList._VSDestPathBase).Replace('\', '\\')
            Remove-Folders -RFFolder $ORDriveList._VSDestPath -RFMatch '\\tmp\\' -RFBaseMatch $ORDriveListBaseMatch
        }
    }
    # Cleanup Log Files
    Remove-Logfiles
    Write-Output "[END] $(Get-Timestamp) - Script completed. Total Elapsed Time: $($ScriptStopWatch.Elapsed.ToString())"
    Stop-Transcript
    ((Get-Content $LFile | Select-Object -Skip 5) | Select-Object -SkipLast 4) | Set-Content $LFile
    Remove-Spaces $LFile
    if ($ExitScript -and !($TestScript)) {
        $LogTemp = Join-Path $LFolderBaseDate -ChildPath "$DateTime-Temp.log"
        New-Item -Path $LogTemp -ItemType File | Out-Null
        $asciiLogo | Out-File $LogTemp -Width 9999
        Get-Content $LFile -ReadCount 5000 | ForEach-Object {
            $_ | Add-Content "$LogTemp"
        }
        Remove-Item $LFile
        Rename-Item $LogTemp -NewName $LFile
        exit
    }
    elseif ($TestScript) {
        $LogTemp = Join-Path $LFolderBaseDate -ChildPath "$DateTime-Temp.log"
        New-Item -Path $LogTemp -ItemType File | Out-Null
        $asciiLogo | Out-File $LogTemp -Width 9999
        Get-Content $LFile -ReadCount 5000 | ForEach-Object {
            $_ | Add-Content "$LogTemp"
        }
        Remove-Item $LFile
        Rename-Item -Path $LogTemp -NewName "$DateTime-DEBUG.log"
        exit
    }
    else {
        if ($VSVTotCount -gt 0) {
            $LogTemp = Join-Path $LFolderBaseDate -ChildPath "$DateTime-Temp.log"
            New-Item -Path $LogTemp -ItemType File | Out-Null
            $asciiLogo | Out-File $LogTemp -Width 9999
            $VSCompletedFilesTable | Out-File $LogTemp -Width 9999 -Append
            Get-Content $LFile -ReadCount 5000 | ForEach-Object {
                $_ | Add-Content "$LogTemp"
            }
            Remove-Item $LFile
            Rename-Item $LogTemp -NewName "$DateTime-Total-$VSVTotCount.log"
        }
        else {
            $LogTemp = Join-Path $LFolderBaseDate -ChildPath "$DateTime-Temp.log"
            New-Item -Path $LogTemp -ItemType File | Out-Null
            $asciiLogo | Out-File $LogTemp -Width 9999
            Get-Content $LFile -ReadCount 5000 | ForEach-Object {
                $_ | Add-Content "$LogTemp"
            }
            Remove-Item $LFile
            Rename-Item $LogTemp -NewName $LFile
        }
    }
}
# Update SE/MKV/FB true false
function Set-VideoStatus {
    param (
        [parameter(Mandatory = $true)]
        [string]$SVSKey,
        [parameter(Mandatory = $true)]
        [string]$SVSValue,
        [parameter(Mandatory = $false)]
        [switch]$SVSSE,
        [parameter(Mandatory = $false)]
        [switch]$SVSMKV,
        [parameter(Mandatory = $false)]
        [switch]$SVSMove,
        [parameter(Mandatory = $false)]
        [switch]$SVSFP,
        [parameter(Mandatory = $false)]
        [switch]$SVSER
    )
    $VSCompletedFilesList | Where-Object { $_.$SVSKey -eq $SVSValue } | ForEach-Object {
        if ($SVSSE) {
            $_._VSSECompleted = $SVSSE
        }
        if ($SVSMKV) {
            $_._VSMKVCompleted = $SVSMKV
        }
        if ($SVSMove) {
            $_._VSMoveCompleted = $SVSMove
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
        $SeriesMessage = '<strong>Series:</strong> ' + $_.Series + "`n<strong>Episode:</strong>`n" + $EpList
        $Telegrammessage += $SeriesMessage + "`n"
    }
    return $Telegrammessage
}
# Sending To telegram for new file notifications
function Invoke-Telegram {
    Param(
        [Parameter( Mandatory = $true)]
        [String]$STMessage)
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    switch ($TelegramNotification.ToLower()) {
        true { $TGRequest = "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($STMessage)&parse_mode=html&disable_notification=true" }
        false { $TGRequest = "https://api.telegram.org/bot$($Telegramtoken)/sendMessage?chat_id=$($Telegramchatid)&text=$($STMessage)&parse_mode=html" }
        Default { Write-Output 'Improper configured value. Accepted values: true/false' }
    }
    Invoke-WebRequest -Uri $TGRequest | Out-Null
}
# Run MKVMerge process
function Invoke-MKVMerge {
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
    Write-Output "[MKVMerge] $(Get-Timestamp) - Video = $VideoLang/$VTrackName - Audio Language = $AudioLang/$ALTrackName - Subtitle = $SubtitleLang/$STTrackName."
    Write-Output "[MKVMerge] $(Get-Timestamp) - Replacing Styling in $MKVVidSubtitle."
    While ($True) {
        if ((Test-Lock $MKVVidSubtitle) -eq $True) {
            continue
        }
        else {
            if ($SF -ne 'None') {
                Write-Output "[MKVMerge] $(Get-Timestamp) - [SubtitleRegex] - Python - Regex through $MKVVidSubtitle file with $SF."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "python `"$SubtitleRegex`" `"$MKVVidSubtitle`" `"$SF`""
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
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$MKVVidTempOutput`" --language 0:`"$VideoLang`" --track-name 0:`"$VTrackName`" --language 1:`"$AudioLang`" --track-name 1:`"$ALTrackName`" ( `"$MKVVidInput`" ) --language 0:`"$SubtitleLang`" --track-name 0:`"$STTrackName`" `"$MKVVidSubtitle`" --attach-file `"$SubFontDir`" --attachment-mime-type application/x-truetype-font"
                break
            }
            else {
                Write-Output "[MKVMerge] $(Get-Timestamp) -  Merging as-is. No Font specified for $MKVVidSubtitle and $MKVVidInput files with $SubFontDir."
                Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvmerge.exe -o `"$MKVVidTempOutput`" --language 0:`"$VideoLang`" --track-name 0:`"$VTrackName`" --language 1:`"$AudioLang`" --track-name 1:`"$ALTrackName`" ( `"$MKVVidInput`" ) --language 0:`"$SubtitleLang`" --track-name 0:`"$STTrackName`" `"$MKVVidSubtitle`""
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
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$MKVVidInput`" -Verbose"
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Remove-Item -Path `"$MKVVidSubtitle`" -Verbose"
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidTempOutput) -eq $True) {
            continue
        }
        else {
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "Rename-Item -Path `"$MKVVidTempOutput`" -NewName `"$MKVVidInput`" -Verbose"
            break
        }
        Start-Sleep -Seconds 1
    }
    While ($True) {
        if ((Test-Lock $MKVVidInput) -eq $True) {
            continue
        }
        else {
            Invoke-ExpressionConsole -SCMFN 'MKVMerge' -SCMFP "mkvpropedit `"$MKVVidInput`" --edit track:s1 --set flag-default=1"
            
            break
        }
        Start-Sleep -Seconds 1
    }
    Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $MKVVidBaseName -SVSMKV
}
# Function to process video files through FileBot
function Invoke-Filebot {
    param (
        [parameter(Mandatory = $true)]
        [string]$FBPath
    )
    Write-Output "[Filebot] $(Get-Timestamp) - Looking for files to rename and move to final folder."
    $FBVideoList = $VSCompletedFilesList | Where-Object { $_._VSDestPath -eq $FBPath } | Select-Object _VSDestPath, _VSEpisodeFBPath, _VSEpisodeRaw, _VSEpisodeSubFBPath, _VSOverridePath
    foreach ($FBFiles in $FBVideoList) {
        $FBVidInput = $FBFiles._VSEpisodeFBPath
        $FBSubInput = $FBFiles._VSEpisodeSubFBPath
        $FBVidBaseName = $FBFiles._VSEpisodeRaw
        $FBOverrideDrive = $FBFiles._VSOverridePath
        if ($SiteParentFolder.trim() -ne '' -or $SiteSubFolder.trim() -ne '') {
            $FBRootFolder = $FBOverrideDrive + $SiteParentFolder
            $FBParams = Join-Path (Join-Path $FBRootFolder -ChildPath $SiteSubFolder) -ChildPath $FBArgument
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBVidInput). Renaming video and moving files to final folder. Using path($FBParams)."
            Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$FBVidInput`" -r --db TheTVDB -non-strict --format `"$FBParams`" --apply date tags clean --log info"
            if (!($MKVMerge)) {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBSubInput). Renaming subtitle and moving files to final folder. Using path($FBParams)."
                Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$FBSubInput`" -r --db TheTVDB -non-strict --format `"$FBParams`" --apply date tags clean --log info"
            }
        }
        else {
            Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBVidInput). ParentFolder or Subfolder path not specified. Renaming files in place."
            Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$FBVidInput`" -r --db TheTVDB -non-strict --format `"$FBArgument`" --apply date tags clean --log info"
            if (!($MKVMerge)) {
                Write-Output "[Filebot] $(Get-Timestamp) - Files found($FBSubInput). Renaming subtitle and moving files to final folder. Using path($FBParams)."
                Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -rename `"$FBSubInput`" -r --db TheTVDB -non-strict --format `"$FBArgument`" --apply date tags clean --log info"
            }
        }
        if (!(Test-Path $FBVidInput)) {
            Write-Output "[Filebot] $(Get-Timestamp) - Setting file($FBVidInput) as completed."
            Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $FBVidBaseName -SVSFP
        }
    }
    $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
    if ($VSVFBCount -eq $VSVTotCount ) {
        Write-Output "[Filebot]$(Get-Timestamp) - Filebot($VSVFBCount) = ($VSVTotCount)Total Videos. No other files need to be processed. Attempting Filebot cleanup."
        Invoke-ExpressionConsole -SCMFN 'Filebot' -SCMFP "filebot -script fn:cleaner `"$SiteHome`" --log info"
    }
    else {
        Write-Output "[Filebot] $(Get-Timestamp) - Filebot($VSVFBCount) and Total Video($VSVTotCount) count mismatch. Manual check required."
    }
    if ($VSVFBCount -ne $VSVTotCount) {
        Write-Output "[Filebot] $(Get-Timestamp) - [FolderCleanup] - File needs processing."
    }
}
# Setting up arraylist for MKV and Filebot lists
class VideoStatus {
    [string]$_VSSite
    [string]$_VSSeries
    [string]$_VSEpisode
    [string]$_VSSeriesDirectory
    [string]$_VSEpisodeRaw
    [string]$_VSEpisodeTemp
    [string]$_VSEpisodePath
    [string]$_VSEpisodeSubtitle
    [string]$_VSEpisodeSubtitleBase
    [string]$_VSEpisodeFBPath
    [string]$_VSEpisodeSubFBPath
    [string]$_VSOverridePath
    [string]$_VSDestPathDirectory
    [string]$_VSDestPath
    [string]$_VSDestPathBase
    [bool]$_VSSECompleted
    [bool]$_VSMKVCompleted
    [bool]$_VSMoveCompleted
    [bool]$_VSFBCompleted
    [bool]$_VSErrored
    
    VideoStatus([string]$VSSite, [string]$VSSeries, [string]$VSEpisode, [string]$VSSeriesDirectory, [string]$VSEpisodeRaw, [string]$VSEpisodeTemp, [string]$VSEpisodePath, [string]$VSEpisodeSubtitle, `
            [string]$VSEpisodeSubtitleBase, [string]$VSEpisodeFBPath, [string]$VSEpisodeSubFBPath, [string]$VSOverridePath, [string]$VSDestPathDirectory, [string]$VSDestPath, [string]$VSDestPathBase, `
            [bool]$VSSECompleted, [bool]$VSMKVCompleted, [bool]$VSMoveCompleted, [bool]$VSFBCompleted, [bool]$VSErrored) {
        $this._VSSite = $VSSite
        $this._VSSeries = $VSSeries
        $this._VSEpisode = $VSEpisode
        $this._VSSeriesDirectory = $VSSeriesDirectory
        $this._VSEpisodeRaw = $VSEpisodeRaw
        $this._VSEpisodeTemp = $VSEpisodeTemp
        $this._VSEpisodePath = $VSEpisodePath
        $this._VSEpisodeSubtitle = $VSEpisodeSubtitle
        $this._VSEpisodeSubtitleBase = $VSEpisodeSubtitleBase
        $this._VSEpisodeFBPath = $VSEpisodeFBPath
        $this._VSEpisodeSubFBPath = $VSEpisodeSubFBPath
        $this._VSOverridePath = $VSOverridePath
        $this._VSDestPathDirectory = $VSDestPathDirectory
        $this._VSDestPath = $VSDestPath
        $this._VSDestPathBase = $VSDestPathBase
        $this._VSSECompleted = $VSSECompleted
        $this._VSMKVCompleted = $VSMKVCompleted
        $this._VSMoveCompleted = $VSMoveCompleted
        $this._VSFBCompleted = $VSFBCompleted
        $this._VSErrored = $VSErrored
    }
}
[System.Collections.ArrayList]$VSCompletedFilesList = @()
# Start of script Variable setup
$ScriptDirectory = $PSScriptRoot
$DLPScript = Join-Path $ScriptDirectory -ChildPath 'dlp-script.ps1'
$SubtitleRegex = Join-Path $ScriptDirectory -ChildPath 'subtitle_regex.py'
$ConfigPath = Join-Path $ScriptDirectory -ChildPath 'config.xml'
$SharedF = Join-Path $ScriptDirectory -ChildPath 'shared'
$FontFolder = Join-Path $ScriptDirectory -ChildPath 'fonts'
$xmlconfig = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
    <Directory>
        <!-- Folder to store backup of site config/archive/bat/cookie files
            ex: <backup location="E:\Backup" />
        -->
        <backup location="" />
        <!-- Temp folder used by YT-DLP
            ex: <temp location="D:\tmp" />
        -->
        <temp location="" />
        <!-- Staging folder for post-process after download
            ex: <src location="D:\src" />
        -->
        <src location="" />
        <!-- Default final staging after post-processing. Can be overridden per series/show under OverrideSeries section
            ex: <dest location="J:\tmp" />
        -->
        <dest location="" />
        <!-- Location of YT-DLP FFMPEG version
            ex: <ffmpeg location="D:\Common\ffmpeg\bin" />
        -->
        <ffmpeg location="" />
    </Directory>
    <Logs>
        <!-- How long to keep folders that did or didn't result in a download -->
        <keeplog emptylogskeepdays="0" filledlogskeepdays="7" />
    </Logs>
    <Plex>
        <!-- PLEX local IP and token used to update plex library after success processing -->
        <plexcred plexUrl="" plexToken="" />
    </Plex>
    <Filebot>
        <!-- Uses dest drive with site parentfolder/subfolder and argument to run filebot command. may want to run for 1 episode to see what the series folder name comes out to.
            ex: <fbfolder fbArgument="{n}\{'Season '+s00}\{n} - {s00e00} - {t}" />
        -->
        <fbfolder fbArgument="{n}\{'Season '+s00}\{n} - {s00e00} - {t}" />
    </Filebot>
    <OverrideSeries>
        <!-- Used to move files to different tmp folder than whats defined in defauilt above.
            <override orSeriesName="MySeriesThatIsInADifferentDriveThanDefaultTemp" orSrcdrive="I:\" />
        -->
        <override orSeriesName="" orSrcdrive="" />
        <override orSeriesName="" orSrcdrive="" />
        <override orSeriesName="" orSrcdrive="" />
    </OverrideSeries>
    <Telegram>
        <!-- Telegram Bot tokenId/ your group ChatId, disable sound notification -->
        <token tokenId="" chatid="" disableNotification="true" />
    </Telegram>
    <credentials>
        <!-- Where you store the Site name, username/password, plexlibraryid, folder in library, and a custom font used to embed into video/sub
            <site sitename="Crunchyroll">
                <username>MyUserName</username>
                <password>MyPassword</password>
                <plexlibraryid>4</plexlibraryid>
                <parentfolder>Video</parentfolder>
                <subfolder>A</subfolder>
                <font>Marker SD.ttf</font>
            </site>
        -->
        <site sitename="">
            <username></username>
            <password></password>
            <plexplexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
        </site>
        <site sitename="">
            <username></username>
            <password></password>
            <plexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
        </site>
        <site sitename="">
            <username></username>
            <password></password>
            <plexlibraryid></plexlibraryid>
            <parentfolder></parentfolder>
            <subfolder></subfolder>
            <font></font>
        </site>
    </credentials>
</configuration>
'@
$defaultconfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
--sub-format 'ass/srt/vtt'
--sub-format 'vtt/srt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$vrvconfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
--trim-filenames 248
--add-metadata
--sub-langs "en-US"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv[format_id*=-ja-JP][format_id!*=hardsub][height>=1080]+ba[format_id*=-ja-JP][format_id!*=hardsub] / b[format_id*=-ja-JP][format_id!*=hardsub][height>=1080] / b*[format_id*=-ja-JP][format_id!*=hardsub]'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$crunchyrollconfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
--trim-filenames 248
--add-metadata
--sub-langs "en-US"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--match-filter "season!~='\(.* Dub\)'"
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv[height>=1080]+ba[height>=1080] / bv+ba / b*'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$funimationconfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
--trim-filenames 248
--add-metadata
--sub-langs 'en.*'
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
--no-check-certificate
-N 32
--downloader aria2c
--extractor-args 'funimation:language=japanese'
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$hidiveconfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
--trim-filenames 248
--add-metadata
--sub-langs "english-subs"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$paramountplusconfig = @'
-v
-F
--list-subs
--no-simulate
--restrict-filenames
--windows-filenames
--replace-in-metadata "title,series,season,season_number,episode" "[$%^@.#+]" "-"
--trim-filenames 248
--add-metadata
--sub-langs "en.*"
--sub-format 'ass/srt/vtt'
--convert-subs 'ass'
--write-subs
--embed-metadata
--embed-thumbnail
--convert-thumbnails 'png'
--remux-video 'mkv'
-N 32
--downloader aria2c
--downloader-args aria2c:'-c -j 64 -s 64 -x 16 --file-allocation=none --optimize-concurrent-downloads=true --http-accept-gzip=true'
-f 'bv*[height>=1080]+ba/b[height>=1080] / bv*+ba/w / b'
-o '%(series).110s/S%(season_number)sE%(episode_number)s - %(title).120s.%(ext)s'
'@
$asciiLogo = @'
######   ##       #######           #######  ######   ##       #######  ##   ##  #######  ###  ##  ######   ##       #######  #######
     ##  ##            ##                      ##     ##                ##   ##       ##  #### ##       ##  ##                     ##
##   ##  ##       #######   #####   ####       ##     ##       ####     #######  #######  ## ####  ##   ##  ##       ####     #######
##   ##  ##   ##  ##       # # #    ##         ##     ##   ##  ##       ##   ##  ##   ##  ##  ###  ##   ##  ##   ##  ##       ##  ##
######   #######  ##                ##       ######   #######  #######  ##   ##  ##   ##  ##   ##  ######   #######  #######  ##   ##
_____________________________________________________________________________________________________________________________________
'@
if (!(Test-Path "$ScriptDirectory\config.xml" -PathType Leaf)) {
    New-Item "$ScriptDirectory\config.xml" -ItemType File -Force
}
if ($help) {
    Show-Markdown -Path "$ScriptDirectory\README.md" -UseBrowser
    exit
}
if ($NewConfig) {
    if (!(Test-Path $ConfigPath -PathType Leaf) -or [String]::IsNullOrWhiteSpace((Get-Content $ConfigPath))) {
        
        New-Item $ConfigPath -ItemType File -Force
        Write-Output "$ConfigPath File Created successfully."
            
        $xmlconfig | Set-Content $ConfigPath
    }
    else {
        Write-Output "$ConfigPath File Exists."
        
    }
    exit
}
if ($SupportFiles) {
    New-Folder $FontFolder
    New-Folder $SharedF
    $ConfigPath = Join-Path $ScriptDirectory -ChildPath 'config.xml'
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SNfile = $ConfigFile.configuration.credentials.site | Where-Object { $_.siteName.trim() -ne '' } | Select-Object 'siteName' -ExpandProperty siteName
    $SNfile | ForEach-Object {
        $SN = New-Object -Type PSObject -Property @{
            SN = $_.siteName
        }
        $SCC = Join-Path $ScriptDirectory -ChildPath 'sites'
        $SCF = Join-Path $SCC -ChildPath $SN.SN
        $SCDF = $SCF.TrimEnd('\') + '_D'
        $SiteSupportFolders = $SCC, $SCF, $SCDF
        foreach ($F in $SiteSupportFolders) {
            New-Folder $F
        }
        $SADF = Join-Path $SharedF -ChildPath "$($SN.SN)_D_A"
        $SBDF = Join-Path $SharedF -ChildPath "$($SN.SN)_D_B"
        $SBDC = Join-Path $SharedF -ChildPath "$($SN.SN)_D_C"
        $SAF = Join-Path $SharedF -ChildPath "$($SN.SN)_A"
        $SBF = Join-Path $SharedF -ChildPath "$($SN.SN)_B"
        $SBC = Join-Path $SharedF -ChildPath "$($SN.SN)_C"
        $SiteSuppFiles = $SADF, $SBDF, $SBDC, $SAF, $SBF, $SBC
        foreach ($S in $SiteSuppFiles) {
            New-SuppFiles $S
        }
        $SCFDCD = $SCF.TrimEnd('\') + '_D'
        $SCFDC = Join-Path $SCFDCD -ChildPath 'yt-dlp.conf'
        $SCFC = Join-Path $SCF -ChildPath 'yt-dlp.conf'
        $SiteConfigFiles = $SCFDC , $SCFC
        foreach ($CF in $SiteConfigFiles) {
            New-Config $CF
            Remove-Spaces $CF
        }
    }
    exit
}
if ($Site) {
    if (Test-Path -Path $SubtitleRegex) {
        Write-Output "$(Get-Timestamp) - $DLPScript, $subtitle_regex do exist in $ScriptDirectory folder."
    }
    else {
        Write-Output "$(Get-Timestamp) - subtitle_regex.py does not exist or was not found in $ScriptDirectory folder. Exiting."
        Exit
    }
    $Date = Get-Day
    $DateTime = Get-TimeStamp
    $Time = Get-Time
    $site = $site.ToLower()
    # Reading from XML
    $ConfigPath = Join-Path $ScriptDirectory -ChildPath 'config.xml'
    [xml]$ConfigFile = Get-Content -Path $ConfigPath
    $SiteParams = $ConfigFile.configuration.credentials.site | Where-Object { $_.siteName.ToLower() -eq $site } | Select-Object 'siteName', 'username', 'password', 'plexlibraryid', 'parentfolder', 'subfolder', 'font' -First 1
    $SiteName = $SiteParams.siteName.ToLower()
    $SiteNameRaw = $SiteParams.siteName
    $SiteFolderDirectory = Join-Path $ScriptDirectory -ChildPath 'sites'
    if ($Daily) {
        $SiteType = $SiteName + '_D'
        $SiteFolder = Join-Path $SiteFolderDirectory -ChildPath $SiteType
        $LFolderBase = Join-Path $SiteFolder -ChildPath 'log'
        $LFolderBaseDate = Join-Path $LFolderBase -ChildPath $Date
        $LFile = Join-Path $LFolderBaseDate -ChildPath "$DateTime.log"
        Start-Transcript -Path $LFile -UseMinimalHeader
        Write-Output "[Setup] $(Get-Timestamp) - $SiteNameRaw"
    }
    else {
        $SiteType = $SiteName
        $SiteFolder = Join-Path $SiteFolderDirectory -ChildPath $SiteType
        $LFolderBase = Join-Path $SiteFolder -ChildPath 'log'
        $LFolderBaseDate = Join-Path $LFolderBase -ChildPath $Date
        $LFile = Join-Path $LFolderBaseDate -ChildPath "$DateTime.log"
        Start-Transcript -Path $LFile -UseMinimalHeader
        Write-Output "[Setup] $(Get-Timestamp) - $SiteNameRaw"
    }
    $SiteUser = $SiteParams.username
    $SitePass = $SiteParams.password
    $SiteLibraryId = $SiteParams.plexlibraryid
    $SiteParentFolder = $SiteParams.parentfolder
    $SiteSubFolder = $SiteParams.subfolder
    $SubFont = $SiteParams.font
    $BackupDrive = $ConfigFile.configuration.Directory.backup.location
    $TempDrive = $ConfigFile.configuration.Directory.temp.location
    $SrcDrive = $ConfigFile.configuration.Directory.src.location
    $DestDrive = $ConfigFile.configuration.Directory.dest.location
    $Ffmpeg = $ConfigFile.configuration.Directory.ffmpeg.location
    [int]$EmptyLogs = $ConfigFile.configuration.Logs.keeplog.emptylogskeepdays
    [int]$FilledLogs = $ConfigFile.configuration.Logs.keeplog.filledlogskeepdays
    $PlexHost = $ConfigFile.configuration.Plex.plexcred.plexUrl
    $PlexToken = $ConfigFile.configuration.Plex.plexcred.plexToken
    $FBArgument = $ConfigFile.configuration.Filebot.fbfolder.fbArgument
    $OverrideSeriesList = $ConfigFile.configuration.OverrideSeries.override | Where-Object { $_.orSeriesId -ne '' -and $_.orSrcdrive -ne '' }
    $Telegramtoken = $ConfigFile.configuration.Telegram.token.tokenId
    $Telegramchatid = $ConfigFile.configuration.Telegram.token.chatid
    $TelegramNotification = $ConfigFile.configuration.Telegram.token.disableNotification
    $SiteDefaultPath = Join-Path (Split-Path $DestDrive -Qualifier) -ChildPath ($SiteParentFolder + '\' + $SiteSubFolder)
    # Video track title inherits from Audio language code
    if ($AudioLang -eq '' -or $null -eq $AudioLang) {
        $AudioLang = 'und'
    }
    if ($SubtitleLang -eq '' -or $null -eq $SubtitleLang) {
        $SubtitleLang = 'und'
    }
    switch ($AudioLang) {
        ar { $VideoLang = $AudioLang; $ALTrackName = 'Arabic Audio'; $VTrackName = 'Arabic Video' }
        de { $VideoLang = $AudioLang; $ALTrackName = 'Deutsch Audio'; $VTrackName = 'Deutsch Video' }
        en { $VideoLang = $AudioLang; $ALTrackName = 'English Audio'; $VTrackName = 'English Video' }
        es { $VideoLang = $AudioLang; $ALTrackName = 'Spanish(Latin America) Audio'; $VTrackName = 'Spanish(Latin America) Video' }
        es-es { $VideoLang = $AudioLang; $ALTrackName = 'Spanish(Spain) Audio'; $VTrackName = 'Spanish(Spain) Video' }
        fr { $VideoLang = $AudioLang; $ALTrackName = 'French Audio'; $VTrackName = 'French Video' }
        it { $VideoLang = $AudioLang; $ALTrackName = 'Italian Audio'; $VTrackName = 'Italian Video' }
        ja { $VideoLang = $AudioLang; $ALTrackName = 'Japanese Audio'; $VTrackName = 'Japanese Video' }
        pt-br { $VideoLang = $AudioLang; $ALTrackName = 'Português (Brasil) Audio'; $VTrackName = 'Português (Brasil) Video' }
        pt-pt { $VideoLang = $AudioLang; $ALTrackName = 'Português (Portugal) Audio'; $VTrackName = 'Português (Portugal) Video' }
        ru { $VideoLang = $AudioLang; $ALTrackName = 'Russian Audio'; $VTrackName = 'Russian Video' }
        und { $AudioLang = 'und'; $VideoLang = 'und'; $ALTrackName = 'und audio'; $VTrackName = 'und Video' }
    }
    switch ($SubtitleLang) {
        ar { $STTrackName = 'Arabic Sub' }
        de { $STTrackName = 'Deutsch Sub' }
        en { $STTrackName = 'English Sub' }
        es { $STTrackName = 'Spanish(Latin America) Sub' }
        es-es { $STTrackName = 'Spanish(Spain) Sub' }
        fr { $STTrackName = 'French Sub' }
        it { $STTrackName = 'Italian Sub' }
        ja { $STTrackName = 'Japanese Sub' }
        pt-br { $STTrackName = 'Português (Brasil) Sub' }
        pt-pt { $STTrackName = 'Português (Portugal) Sub' }
        ru { $STTrackName = 'Russian Video' }
        ja { $STTrackName = 'Japanese Sub' }
        en { $STTrackName = 'English Sub' }
        und { $SubtitleLang = 'und'; $STTrackName = 'und sub' }
    }
    # End reading from XML
    if ($SubFont.Trim() -ne '') {
        $SubFontDir = Join-Path $FontFolder -ChildPath $Subfont
        if (Test-Path $SubFontDir) {
            $SF = [System.Io.Path]::GetFileNameWithoutExtension($SubFont)
            Write-Output "[Setup] $(Get-Timestamp) - $SubFont set for $SiteName."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $SubFont specified in $ConfigFile is missing from $FontFolder. Exiting."
            Exit-Script -es
        }
    }
    else {
        $SubFont = 'None'
        $SubFontDir = 'None'
        $SF = 'None'
        Write-Output "[Setup] $(Get-Timestamp) - $SubFont - No font set for $SiteName."
    }
    $SiteShared = Join-Path $ScriptDirectory -ChildPath 'shared'
    $SrcBackup = Join-Path $BackupDrive -ChildPath '_Backup'
    $SrcDriveShared = Join-Path $SrcBackup -ChildPath 'shared'
    $SrcDriveSharedFonts = Join-Path $SrcBackup -ChildPath 'fonts'
    $dlpParams = 'yt-dlp'
    $dlpArray = @()
    if ($Daily) {
        $SiteTempBase = Join-Path $TempDrive -ChildPath $SiteName.Substring(0, 1)
        $SiteTempBaseMatch = $SiteTempBase.Replace('\', '\\')
        $SiteTemp = Join-Path $SiteTempBase -ChildPath $Time
        $SiteSrcBase = Join-Path $SrcDrive -ChildPath $SiteName.Substring(0, 1)
        $SiteSrcBaseMatch = $SiteSrcBase.Replace('\', '\\')
        $SiteSrc = Join-Path $SiteSrcBase -ChildPath $Time
        $SiteHomeBase = Join-Path (Join-Path $DestDrive -ChildPath "_$SiteSubFolder") -ChildPath ($SiteName).Substring(0, 1)
        $SiteHomeBaseMatch = $SiteHomeBase.Replace('\', '\\')
        $SiteHome = Join-Path $SiteHomeBase -ChildPath $Time
        $SiteConfig = Join-Path $SiteFolder -ChildPath 'yt-dlp.conf'
        if ($SrcDrive -eq $TempDrive) {
            Write-Output "[Setup] $(Get-Timestamp) - Src($SrcDrive) and Temp($TempDrive) Directories cannot be the same"
            Exit-Script -es
        }
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $SiteConfig file found. Continuing."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
            $dlpArray += "`"--config-location`"", "`"$SiteConfig`"", "`"-P`"", "`"temp:$SiteTemp`"", "`"-P`"", "`"home:$SiteSrc`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $SiteConfig does not exist. Exiting."
            Exit-Script -es
        }
    }
    else {
        $SiteTempBase = Join-Path $TempDrive -ChildPath "$($SiteName.Substring(0, 1))M"
        $SiteTempBaseMatch = $SiteTempBase.Replace('\', '\\')
        $SiteTemp = Join-Path $SiteTempBase -ChildPath $Time
        $SiteSrcBase = Join-Path $SrcDrive -ChildPath "$($SiteName.Substring(0, 1))M"
        $SiteSrcBaseMatch = $SiteSrcBase.Replace('\', '\\')
        $SiteSrc = Join-Path $SiteSrcBase -ChildPath $Time
        $SiteHomeBase = Join-Path (Join-Path $DestDrive -ChildPath '_M') -ChildPath ($SiteName).Substring(0, 1)
        $SiteHomeBaseMatch = $SiteHomeBase.Replace('\', '\\')
        $SiteHome = Join-Path $SiteHomeBase -ChildPath $Time
        $SiteConfig = $SiteFolder + '\yt-dlp.conf'
        if ((Test-Path -Path $SiteConfig)) {
            Write-Output "[Setup] $(Get-Timestamp) - $SiteConfig file found. Continuing."
            $dlpParams = $dlpParams + " --config-location $SiteConfig -P temp:$SiteTemp -P home:$SiteSrc"
            $dlpArray += "`"--config-location`"", "`"$SiteConfig`"", "`"-P`"", "`"temp:$SiteTemp`"", "`"-P`"", "`"home:$SiteSrc`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $SiteConfig does not exist. Exiting."
            Exit-Script -es
        }
    }
    $SiteConfigBackup = Join-Path (Join-Path $SrcBackup -ChildPath 'sites') -ChildPath $SiteType
    $CookieFile = Join-Path $SiteShared -ChildPath "$($SiteType)_C"
    if ($Login) {
        if ($SiteUser -and $SitePass) {
            Write-Output "[Setup] $(Get-Timestamp) - Login is true and SiteUser/Password is filled. Continuing."
            $dlpParams = $dlpParams + " -u $SiteUser -p $SitePass"
            $dlpArray += "`"-u`"", "`"$SiteUser`"", "`"-p`"", "`"$SitePass`""
            if ($Cookies) {
                if ((Test-Path -Path $CookieFile)) {
                    Write-Output "[Setup] $(Get-Timestamp) - Cookies is true and $CookieFile file found. Continuing."
                    $dlpParams = $dlpParams + " --cookies $CookieFile"
                    $dlpArray += "`"--cookies`"", "`"$CookieFile`""
                }
                else {
                    Write-Output "[Setup] $(Get-Timestamp) - $CookieFile does not exist. Exiting."
                    Exit-Script -es
                }
            }
            else {
                $CookieFile = 'None'
                Write-Output "[Setup] $(Get-Timestamp) - Login is true and Cookies is false. Continuing."
            }
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - Login is true and Username/Password is Empty. Exiting."
            Exit-Script -es
        }
    }
    else {
        if ((Test-Path -Path $CookieFile)) {
            Write-Output "[Setup] $(Get-Timestamp) - $CookieFile file found. Continuing."
            $dlpParams = $dlpParams + " --cookies $CookieFile"
            $dlpArray += "`"--cookies`"", "`"$CookieFile`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $CookieFile does not exist. Exiting."
            Exit-Script -es
        }
    }
    if ($Ffmpeg) {
        Write-Output "[Setup] $(Get-Timestamp) - $Ffmpeg file found. Continuing."
        $dlpParams = $dlpParams + " --ffmpeg-location $Ffmpeg"
        $dlpArray += "`"--ffmpeg-location`"", "`"$Ffmpeg`""
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - FFMPEG: $Ffmpeg missing. Exiting."
        Exit-Script -es
    }
    $BatFile = Join-Path $SiteShared -ChildPath "$($SiteType)_B"
    if ((Test-Path -Path $BatFile)) {
        Write-Output "[Setup] $(Get-Timestamp) - $BatFile file found. Continuing."
        if (![String]::IsNullOrWhiteSpace((Get-Content $BatFile))) {
            Write-Output "[Setup] $(Get-Timestamp) - $BatFile not empty. Continuing."
            $dlpParams = $dlpParams + " -a $BatFile"
            $dlpArray += "`"-a`"", "`"$BatFile`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - $BatFile is empty. Exiting."
            Exit-Script -es
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - BAT: $Batfile missing. Exiting."
        Exit-Script -es
    }
    if ($Archive) {
        $ArchiveFile = Join-Path $SiteShared -ChildPath "$($SiteType)_A"
        if ((Test-Path -Path $ArchiveFile)) {
            Write-Output "[Setup] $(Get-Timestamp) - $ArchiveFile file found. Continuing."
            $dlpParams = $dlpParams + " --download-archive $ArchiveFile"
            $dlpArray += "`"--download-archive`"", "`"$ArchiveFile`""
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - Archive file missing. Exiting."
            Exit-Script -es
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - Using --no-download-archive. Continuing."
        $ArchiveFile = 'None'
        $dlpParams = $dlpParams + ' --no-download-archive'
        $dlpArray += "`"--no-download-archive`""
    }
    $Vtype = Select-String -Path $SiteConfig -Pattern '--remux-video.*' | Select-Object -First 1
    if ($null -ne $Vtype) {
        $VidType = '*.' + ($Vtype -split ' ')[1]
        $VidType = $VidType.Replace("'", '').Replace('"', '')
        if ($VidType -eq '*.mkv') {
            Write-Output "[Setup] $(Get-Timestamp) - Using $VidType. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - VidType(mkv) is missing. Exiting."
            Exit-Script -es
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - --remux-video parameter is missing. Exiting."
        Exit-Script -es
    }
    if ($SubtitleEdit -or $MKVMerge) {
        if (Select-String -Path $SiteConfig '--write-subs' -SimpleMatch -Quiet) {
            Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit or MKVMerge is true and --write-subs is in config. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit is true and --write-subs is not in config. Exiting."
            Exit-Script -es
        }
        $SType = Select-String -Path $SiteConfig -Pattern '--convert-subs.*' | Select-Object -First 1
        if ($null -ne $SType) {
            $SubType = '*.' + ($SType -split ' ')[1]
            $SubType = $SubType.Replace("'", '').Replace('"', '')
            if ($SubType -eq '*.ass') {
                Write-Output "[Setup] $(Get-Timestamp) - Using $SubType. Continuing."
            }
            else {
                Write-Output "[Setup] $(Get-Timestamp) - Subtype(ass) is missing. Exiting."
                Exit-Script -es
            }
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - --convert-subs parameter is missing. Exiting."
            Exit-Script -es
        }
        $Wsub = Select-String -Path $SiteConfig -Pattern '--write-subs.*' | Select-Object -First 1
        if ($null -ne $Wsub) {
            Write-Output "[Setup] $(Get-Timestamp) - --write-subs is in config. Continuing."
        }
        else {
            Write-Output "[Setup] $(Get-Timestamp) - SubSubWrite is missing. Exiting."
            Exit-Script -es
        }
    }
    else {
        Write-Output "[Setup] $(Get-Timestamp) - SubtitleEdit is false. Continuing."
    }
    $DebugVars = [ordered]@{Site = $SiteName; isDaily = $Daily; UseLogin = $Login; UseCookies = $Cookies; UseArchive = $Archive; SubtitleEdit = $SubtitleEdit; `
            MKVMerge = $MKVMerge; VTrackName = $VTrackName; AudioLang = $AudioLang; ALTrackName = $ALTrackName; SubtitleLang = $SubtitleLang; STTrackName = $STTrackName; Filebot = $Filebot; `
            SiteNameRaw = $SiteNameRaw; SiteType = $SiteType; SiteUser = $SiteUser; SitePass = $SitePass; SiteFolder = $SiteFolder; SiteParentFolder = $SiteParentFolder; `
            SiteSubFolder = $SiteSubFolder; SiteLibraryId = $SiteLibraryId; SiteTemp = $SiteTemp; SiteSrcBase = $SiteSrcBase; SiteSrc = $SiteSrc; SiteHomeBase = $SiteHomeBase; `
            SiteHome = $SiteHome; SiteDefaultPath = $SiteDefaultPath; SiteConfig = $SiteConfig; CookieFile = $CookieFile; Archive = $ArchiveFile; Bat = $BatFile; Ffmpeg = $Ffmpeg; SF = $SF; SubFont = $SubFont; `
            SubFontDir = $SubFontDir; SubType = $SubType; VidType = $VidType; Backup = $SrcBackup; BackupShared = $SrcDriveShared; BackupFont = $SrcDriveSharedFonts; `
            SiteConfigBackup = $SiteConfigBackup; PlexHost = $PlexHost; PlexToken = $PlexToken; TelegramToken = $TelegramToken; TelegramChatId = $TelegramChatId; ConfigPath = $ConfigPath; `
            ScriptDirectory = $ScriptDirectory; dlpParams = $dlpParams
    }
    if ($TestScript) {
        Write-Output "[START] $DateTime - $SiteNameRaw - DEBUG Run"
        $DebugVars
        $OverrideSeriesList
        Write-Output 'dlpArray:'
        $dlpArray
        Write-Output "[END] $DateTime - Debugging enabled. Exiting."
        Exit-Script -es
    }
    else {
        $DebugVarRemove = 'SiteUser' , 'SitePass', 'PlexToken', 'TelegramToken', 'TelegramChatId', 'dlpParams'
        foreach ($dbv in $DebugVarRemove) {
            $DebugVars.Remove($dbv)
        }
        if ($Daily) {
            Write-Output "[START] $DateTime - $SiteNameRaw - Daily Run"
            Write-Output '[START] Debug Vars:'
            $DebugVars
            Write-Output '[START] Series Drive Overrides:'
            $OverrideSeriesList | Sort-Object orSrcdrive, orSeriesName | Format-Table
            # Create folders
            Write-Output '[START] Creating'
            $CreateFolders = $TempDrive, $SrcDrive, $BackupDrive, $SrcBackup, $SiteConfigBackup, $SrcDriveShared, $SrcDriveSharedFonts, $DestDrive, $SiteTemp, $SiteSrc, $SiteHome
            foreach ($c in $CreateFolders) {
                New-Folder $c
            }
        }
        else {
            Write-Output "[START] $DateTime - $SiteNameRaw - Manual Run"
            Write-Output '[START] Debug Vars:'
            $DebugVars
            Write-Output '[START] Series Drive Overrides:'
            $OverrideSeriesList | Sort-Object orSrcdrive, orSeriesName | Format-Table
            # Create folders
            Write-Output '[START] Creating'
            $CreateFolders = $TempDrive, $SrcDrive, $BackupDrive, $SrcBackup, $SiteConfigBackup, $SrcDriveShared, $SrcDriveSharedFonts, $DestDrive, $SiteTemp, $SiteSrc, $SiteHome
            foreach ($c in $CreateFolders) {
                New-Folder $c
            }
        }
        # Log cleanup
        Remove-Logfiles
    }
    # yt-dlp
    & yt-dlp.exe $dlpArray *>&1 | Out-Host
    # Post-processing
    if ((Get-ChildItem $SiteSrc -Recurse -Force -File -Include "$VidType" | Select-Object -First 1 | Measure-Object).Count -gt 0) {
        Get-ChildItem $SiteSrc -Recurse -Include "$VidType" | Sort-Object LastWriteTime | Select-Object -Unique | Get-Unique | ForEach-Object {
            $VSSite = $SiteNameRaw
            $VSSeries = (Get-Culture).TextInfo.ToTitleCase(("$(Split-Path (Split-Path $_ -Parent) -Leaf)").Replace('_', ' ').Replace('-', ' ')) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $VSEpisode = (Get-Culture).TextInfo.ToTitleCase( ($_.BaseName.Replace('_', ' ').Replace('-', ' '))) | ForEach-Object { $_.trim() -Replace '\s+', ' ' }
            $VSSeriesDirectory = $_.DirectoryName
            $VSEpisodeRaw = $_.BaseName
            $VSEpisodeTemp = Join-Path $VSSeriesDirectory -ChildPath "$($VSEpisodeRaw).temp$($_.Extension)"
            $VSEpisodePath = $_.FullName
            $VSEpisodeSubtitle = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).FullName
            $VSOverridePath = $OverrideSeriesList | Where-Object { $_.orSeriesName.ToLower() -eq $VSSeries.ToLower() } | Select-Object -ExpandProperty orSrcdrive
            if ($null -ne $VSOverridePath) {
                $VSDestPath = Join-Path $VSOverridePath -ChildPath $SiteHome.Substring(3)
                $VSDestPathBase = Join-Path $VSOverridePath -ChildPath $SiteHomeBase.Substring(3)
                $VSEpisodeFBPath = $VSEpisodePath.Replace($SiteSrc, $VSDestPath)
                $VSDestPathDirectory = Split-Path $VSEpisodeFBPath -Parent
                if ($VSEpisodeSubtitle -ne '') {
                    $VSEpisodeSubtitleBase = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).Name
                    $VSEpisodeSubFBPath = $VSEpisodeSubtitle.Replace($SiteSrc, $VSDestPath)
                }
                else {
                    $VSEpisodeSubtitleBase = ''
                    $VSEpisodeSubFBPath = ''
                }
            }
            else {
                $VSDestPath = $SiteHome
                $VSDestPathBase = $SiteHomeBase  #_VSDestPathDirectory = $VSDestPathDirectory
                $VSEpisodeFBPath = $VSEpisodePath.Replace($SiteSrc, $VSDestPath)
                $VSDestPathDirectory = Split-Path $VSEpisodeFBPath -Parent
                $VSOverridePath = [System.IO.path]::GetPathRoot($VSDestPath)
                if ($VSEpisodeSubtitle -ne '') {
                    $VSEpisodeSubtitleBase = (Get-ChildItem $SiteSrc -Recurse -File -Include "$SubType" | Where-Object { $_.FullName -match $VSEpisodeRaw } | Select-Object -First 1 ).Name
                    $VSEpisodeSubFBPath = $VSEpisodeSubtitle.Replace($SiteSrc, $VSDestPath)
                }
                else {
                    $VSEpisodeSubtitleBase = ''
                    $VSEpisodeSubFBPath = ''
                }
            }
            foreach ($i in $_) {
                $VideoStatus = [VideoStatus]::new($VSSite, $VSSeries, $VSEpisode, $VSSeriesDirectory, $VSEpisodeRaw, $VSEpisodeTemp, $VSEpisodePath, $VSEpisodeSubtitle, $VSEpisodeSubtitleBase, $VSEpisodeFBPath, `
                        $VSEpisodeSubFBPath, $VSOverridePath, $VSDestPathDirectory, $VSDestPath, $VSDestPathBase, $VSSECompleted, $VSMKVCompleted, $VSMoveCompleted, $VSFBCompleted, $VSErrored)
                [void]$VSCompletedFilesList.Add($VideoStatus)
            }
        }
        $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | ForEach-Object {
            if (($_._VSEpisodeSubtitle.Trim() -eq '') -or ($null -eq $_._VSEpisodeSubtitle)) {
                Set-VideoStatus -SVSKey '_VSEpisodeRaw' -SVSValue $VSEpisodeRaw -SVSER
            }
        }
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files to process."
    }
    $VSVTotCount = ($VSCompletedFilesList | Measure-Object).Count
    $VSVErrorCount = ($VSCompletedFilesList | Where-Object { $_._VSErrored -eq $true } | Measure-Object).Count
    Write-Output "[VideoList] $(Get-Timestamp) - Total Files: $VSVTotCount"
    Write-Output "[VideoList] $(Get-Timestamp) - Errored Files: $VSVErrorCount"
    # SubtitleEdit, MKVMerge, Filebot
    if ($VSVTotCount -gt 0) {
        # Subtitle Edit
        if ($SubtitleEdit) {
            $VSCompletedFilesList | Select-Object _VSEpisodeSubtitle | Where-Object { $_._VSErrored -ne $true } | ForEach-Object {
                $SESubtitle = $_._VSEpisodeSubtitle
                Write-Output "[SubtitleEdit] $(Get-Timestamp) - Fixing $SESubtitle subtitle."
                While ($True) {
                    if ((Test-Lock $SESubtitle) -eq $True) {
                        continue
                    }
                    else {
                        Invoke-ExpressionConsole -SCMFN 'SubtitleEdit' -SCMFP "powershell `"SubtitleEdit /convert `'$SESubtitle`' AdvancedSubStationAlpha /overwrite /MergeSameTimeCodes`""
                        Set-VideoStatus -SVSKey '_VSEpisodeSubtitle' -SVSValue $SESubtitle -SVSSE
                        break
                    }
                    Start-Sleep -Seconds 1
                }
            }
        }
        else {
            Write-Output "[SubtitleEdit] $(Get-Timestamp) - Not running."
        }
        # MKVMerge
        if ($MKVMerge) {
            $VSCompletedFilesList | Select-Object _VSEpisodeRaw, _VSEpisode, _VSEpisodeTemp, _VSEpisodePath, _VSEpisodeSubtitle, _VSErrored | `
                Where-Object { $_._VSErrored -eq $false } | ForEach-Object {
                $MKVVidInput = $_._VSEpisodePath
                $MKVVidBaseName = $_._VSEpisodeRaw
                $MKVVidSubtitle = $_._VSEpisodeSubtitle
                $MKVVidTempOutput = $_._VSEpisodeTemp
                Invoke-MKVMerge $MKVVidInput $MKVVidBaseName $MKVVidSubtitle $MKVVidTempOutput
            }
            $OverrideDriveList = $VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Select-Object _VSSeriesDirectory, _VSDestPath, _VSDestPathBase -Unique
        }
        else {
            Write-Output "[MKVMerge] $(Get-Timestamp) - MKVMerge not running. Moving to next step."
            $OverrideDriveList = $VSCompletedFilesList | Where-Object { $_._VSErrored -eq $false } | Select-Object _VSSeriesDirectory, _VSDestPath, _VSDestPathBase -Unique
        }
        $VSVMKVCount = ($VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Measure-Object).Count
        # FileMoving
        if (($VSVMKVCount -eq $VSVTotCount -and $VSVErrorCount -eq 0) -or (!($MKVMerge) -and $VSVErrorCount -eq 0)) {
            Write-Output "[FileMoving] $(Get-Timestamp) - All files had matching subtitle file"
            foreach ($ORDriveList in $OverrideDriveList) {
                Write-Output "[FileMoving] $(Get-Timestamp) - $($ORDriveList._VSSeriesDirectory) contains files. Moving to $($ORDriveList._VSDestPath)."
                if (!(Test-Path $ORDriveList._VSDestPath)) {
                    Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "New-Folder `"$($ORDriveList._VSDestPath)`" -Verbose"
                }
                Write-Output "[FileMoving] $(Get-Timestamp) - Moving $($ORDriveList._VSSeriesDirectory) to $($ORDriveList._VSDestPath)."
                Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($ORDriveList._VSSeriesDirectory)`" -Destination `"$($ORDriveList._VSDestPath)`" -Force -Verbose"
                if (!(Test-Path $ORDriveList._VSSeriesDirectory)) {
                    Write-Output "[FileMoving] $(Get-Timestamp) - Move completed for $($ORDriveList._VSSeriesDirectory)."
                    Set-VideoStatus -SVSKey '_VSSeriesDirectory' -SVSValue $ORDriveList._VSSeriesDirectory -SVSMove
                    
                }
            }
        }
        else {
            Write-Output "[FileMoving] $(Get-Timestamp) - $SiteSrc contains file(s) with error(s). Not moving files."
        }
        # Filebot
        $FBOverrideDriveList = $VSCompletedFilesList | Where-Object { $_._VSMKVCompleted -eq $true -and $_._VSErrored -eq $false } | Select-Object _VSDestPath -Unique
        if (($Filebot -and $VSVMKVCount -eq $VSVTotCount) -or ($Filebot -and !($MKVMerge))) {
            foreach ($FBORDriveList in $FBOverrideDriveList) {
                Write-Output "[Filebot] $(Get-Timestamp) - Renaming files in $($FBORDriveList._VSDestPath)."
                Invoke-Filebot -FBPath $ORDriveList._VSDestPath
            }
        }
        elseif (!($Filebot) -and !($MKVMerge) -or (!($Filebot) -and $VSVMKVCount -eq $VSVTotCount)) {
            $MoveManualList = $VSCompletedFilesList | Where-Object { $_._VSMoveCompleted -eq $true } | Select-Object _VSDestPathDirectory, _VSOverridePath -Unique
            foreach ($MMFiles in $MoveManualList) {
                $MMOverrideDrive = $MMFiles._VSOverridePath
                $MoveRootDirectory = $MMOverrideDrive + $SiteParentFolder
                $MoveFolder = Join-Path $MoveRootDirectory -ChildPath $SiteSubFolder
                Write-Output "[FileMoving] $(Get-Timestamp) -  Moving $($MMFiles._VSDestPathDirectory) to $MoveFolder."
                Invoke-ExpressionConsole -SCMFN 'FileMoving' -SCMFP "Move-Item -Path `"$($MMFiles._VSDestPathDirectory)`" -Destination `"$MoveFolder`" -Force -Verbose"
            }
        }
        else {
            Write-Output "[FileMoving] $(Get-Timestamp) - Issue with files in $SiteSrc."
        }
        # Plex
        if ($PlexHost -and $PlexToken -and $SiteLibraryId ) {
            Write-Output "[PLEX] $(Get-Timestamp) - Updating Plex Library."
            $PlexUrl = "$PlexHost/library/sections/$SiteLibraryId/refresh?X-Plex-Token=$PlexToken"
            Invoke-WebRequest -Uri $PlexUrl | Out-Null
        }
        else {
            Write-Output "[PLEX] $(Get-Timestamp) - Not using Plex."
        }
        $VSVFBCount = ($VSCompletedFilesList | Where-Object { $_._VSFBCompleted -eq $true } | Measure-Object).Count
        # Telegram
        if ($SendTelegram) {
            Write-Output "[Telegram] $(Get-Timestamp) - Preparing Telegram message."
            $TM = Get-SiteSeriesEpisode
            if ($PlexHost -and $PlexToken -and $SiteLibraryId) {
                if ($Filebot -or $MKVMerge) {
                    if (($VSVFBCount -gt 0 -and $VSVMKVCount -gt 0 -and $VSVFBCount -eq $VSVMKVCount) -or (!($Filebot) -and $MKVMerge -and $VSVMKVCount -gt 0)) {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $SiteHome. Success."
                        $TM += 'All files added to PLEX.'
                        Write-Output $TM
                        Invoke-Telegram -STMessage $TM
                    }
                    else {
                        Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $SiteHome. Failure."
                        $TM += 'Not all files added to PLEX.'
                        Write-Output $TM
                        Invoke-Telegram -STMessage $TM
                    }
                }
            }
            else {
                Write-Output "[Telegram] $(Get-Timestamp) - Sending message for files in $SiteHome."
                $TM += 'Added files to folders.'
                Write-Output $TM
                Invoke-Telegram -STMessage $TM
            }
        }
        Write-Output "[VideoList] $(Get-Timestamp) - Total videos downloaded: $VSVTotCount"
        $VSCompletedFilesListHeaders = @{Label = 'Series'; Expression = { $_._VSSeries } }, @{Label = 'Episode'; Expression = { $_._VSEpisode } }, `
        @{Label = 'Subtitle'; Expression = { $_._VSEpisodeSubtitleBase } }, @{Label = 'Drive'; Expression = { $_._VSOverridePath } }, @{Label = 'SrcDirectory'; Expression = { $_._VSSeriesDirectory } }, `
        @{Label = 'DestBase'; Expression = { $_._VSDestPathBase } }, @{Label = 'DestPath'; Expression = { $_._VSDestPath } }, @{Label = 'DestPathDirectory'; Expression = { $_._VSDestPathDirectory } }, `
        @{Label = 'SECompleted'; Expression = { $_._VSSECompleted } }, @{Label = 'MKVCompleted'; Expression = { $_._VSMKVCompleted } }, @{Label = 'MoveCompleted'; Expression = { $_._VSMoveCompleted } }, `
        @{Label = 'FBCompleted'; Expression = { $_._VSFBCompleted } }, @{Label = 'Errored'; Expression = { $_._VSErrored } }
        if ($VSVTotCount -gt 12) {
            $VSCompletedFilesTable = $VSCompletedFilesList | Sort-Object _VSSeries, _VSEpisode | Format-Table $VSCompletedFilesListHeaders -AutoSize -Wrap
        }
        else {
            $VSCompletedFilesTable = $VSCompletedFilesList | Sort-Object _VSSeries, _VSEpisode | Format-List $VSCompletedFilesListHeaders
        }
    }
    else {
        Write-Output "[VideoList] $(Get-Timestamp) - No files downloaded. Skipping other defined steps."
    }
    # Backup
    $SharedBackups = $ArchiveFile, $CookieFile, $BatFile, $ConfigPath, $SubFontDir, $SiteConfig
    foreach ($sb in $SharedBackups) {
        if (($sb -ne 'None') -and ($sb.trim() -ne '')) {
            if ($sb -eq $SubFontDir) {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $SrcDriveSharedFonts."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$SrcDriveSharedFonts`" -PassThru -Verbose"
            }
            elseif ($sb -eq $ConfigPath) {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $SrcBackup."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$SrcBackup`" -PassThru -Verbose"
            }
            elseif ($sb -eq $SiteConfig) {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $SiteConfigBackup."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$SiteConfigBackup`" -PassThru -Verbose"
            }
            else {
                Write-Output "[FileBackup] $(Get-Timestamp) - Copying $sb to $SrcDriveShared."
                Invoke-ExpressionConsole -SCMFN 'FileBackup' -SCMFP "Copy-Item -Path `"$sb`" -Destination `"$SrcDriveShared`" -PassThru -Verbose"
            }
        }
    }
    Exit-Script
}