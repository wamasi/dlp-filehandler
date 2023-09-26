[CmdletBinding()]
param(
    [switch]$dailyRuns,
    [switch]$backup
)

function Get-DateTime {
    param (
        [int]$dateType,
        [switch]$backup
    )
    switch ($dateType) {
        1 { $datetime = Get-Date -Format 'yy-MM-dd' }
        2 { $datetime = Get-Date -Format 'MMddHHmmssfff' }
        3 { $datetime = ($(Get-Date).ToUniversalTime()).ToString('yyyy-MM-ddTHH:mm:ss.fffZ') }
        Default { $datetime = Get-Date -Format 'yy-MM-dd HH-mm-ss' }
    }
    return $datetime
}
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath
    )
    Write-Host $Message
    Add-Content -Path $LogFilePath -Value $Message
}

$scriptRoot = $PSScriptRoot
$configFilePath = Join-Path $scriptRoot -ChildPath 'config.xml'
if (!(Test-Path $configFilePath)) {
    $baseXML = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <Directory>
    <backup location="D:\Backup" />
  </Directory>
  <characters>
    <char>’</char>
  </characters>
  <Logs>
    <keeplog emptylogskeepdays="0" filledlogskeepdays="7" />
    <lastUpdated></lastUpdated>
  </Logs>
  <Discord>
    <hook ScheduleServerUrl="" MediaServerUrl="" />
    <icon default="" author="" footerIcon="" Color="" />
    <sites>
      <site siteName="" emoji="" />
      <site siteName="" emoji="" />
      <site siteName="Default" emoji="" />
    </sites>
  </Discord>
</configuration>
'@
    New-Item $configFilePath
    Set-Content $configFilePath -Value $baseXML
    Write-Host "No config found. Configure xml file: $configFilePath"
    exit
}

$dlpPath = Resolve-Path "$scriptRoot\.."
$dlpScript = Join-Path $dlpPath -ChildPath '\dlp-script.ps1'

[xml]$configData = Get-Content $configFilePath
$backupPath = $configData.configuration.Directory.backup.location
$logFolder = Join-Path $scriptRoot -ChildPath 'log'
$logfile = Join-Path $logFolder -ChildPath 'smartDL.log'
$siteBatFolder = Join-Path $scriptRoot -ChildPath 'batch\site'
$now = Get-Date
$currentDate = Get-Date $now -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$weekday = Get-Date $currentDate -Format 'dddd'
$weekday
$spacer = "`n" + '-' * 120

if (!(Test-Path $logFolder)) {
    New-Item $logFolder -ItemType Directory
    New-Item $logFile -ItemType File
}
elseif (!(Test-Path $logfile)) {
    New-Item $logfile -ItemType File
}
else {
    $logCreation = (Get-Item -LiteralPath $logfile).CreationTime
    $oldDate = (Get-Date).AddDays(-7)
    if ($logCreation -le $oldDate) {
        Remove-Item $logfile
        New-Item $logfile -ItemType File
        Write-Log -Message "$spacer" -LogFilePath $logFile
    }
}


if ($dailyRuns) {
    if ((Test-Path $dlpScript)) { 
        Write-Log -Message "[Start] $(Get-DateTime) - Running for $weekday" -LogFilePath $logFile
        Get-ChildItem $siteBatFolder | ForEach-Object {
            $site = $_.Name
            $batchFile = (Get-ChildItem $_ -Recurse | Where-Object { $_.Directory.Name -eq $site -and ($_.FullName -match "$($site).*$weekday") }).FullName
            if ($null -eq $batchFile) {
                Write-Log -Message "[DLP-Script] $(Get-DateTime) - $site - No File for $weekday" -LogFilePath $logFile
            }
            else {
                switch ($site) {
                    'Crunchyroll' {
                        Write-Log -Message "[DLP-Script] $(Get-DateTime) - $site - $weekday - $batchFile" -LogFilePath $logFile
                        $return = pwsh.exe $dlpScript -sn Crunchyroll -d -l -mk -f -a -al ja -sl en -sd -overrideBatch $batchFile
                        $message = $return -match '\[DLP-Script\]*' | ForEach-Object { $_ }
                        Write-Log -Message "$message" -LogFilePath $logFile
                    }
                    'HIDIVE' {
                        Write-Log -Message "[DLP-Script] $(Get-DateTime) - $site - $weekday - $batchFile" -LogFilePath $logFile
                        $return = pwsh.exe $dlpScript -sn Hidive -d -c -mk -f -se -a -al ja -sl en -sd -overrideBatch $batchFile
                        $message = $return -match '\[DLP-Script\]*' | ForEach-Object { $_ }
                        Write-Log -Message "$message" -LogFilePath $logFile
                    }
                    Default { Write-Log -Message "[DLP-Script] $(Get-DateTime) - No Applicable Site for $batchFile" -LogFilePath $logFile }
                }
            }
        }
    }
    else {
        Write-Log -Message "[Start] $(Get-DateTime) - dlp-FileHandler script not found at $dlpscript. Exiting...$spacer" -LogFilePath $logFile
    }
    Write-Log -Message "[End] - $(Get-DateTime) - End of script.$spacer" -LogFilePath $logFile
    exit
}

if ($backup) {
    $date = Get-DateTime 2
    $smartDLBackupPathRoot = (Join-Path $backupPath -ChildPath 'SmartDL')
    $bPath = Join-Path $smartDLBackupPathRoot -ChildPath "SmartDL_$date"
    Write-Log -Message "[Start] $(Get-DateTime) - Backing up batches to $bPath" -LogFilePath $logFile
    if (!(Test-Path $bPath)) {
        New-Item $bPath -ItemType Directory
    }
    Get-ChildItem -Path "$scriptRoot\*" | Copy-Item -Destination $bPath -Recurse -Exclude '*/Log'
    # deleting old backups
    $deleteCutoff = $currentDate.AddDays(-180)
    $foldersToDelete = Get-ChildItem -Path $smartDLBackupPathRoot -Directory | Where-Object { $_.CreationTime -lt $deleteCutoff }
    $foldersToDelete | ForEach-Object {
        Remove-Item $_.FullName -Recurse -Force
    }
    Write-Log -Message "[End] - $(Get-DateTime) - End of script.$spacer" -LogFilePath $logFile
}
