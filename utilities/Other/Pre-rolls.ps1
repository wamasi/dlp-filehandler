$Path = 'D:\Videos\Pre-Rolls'
$PreRoll = 'D:\Common\yt-dlp\Logs\MasterLists\Pre-Rolls.txt'
New-Item -ItemType File -Force -Path $PreRoll
Get-ChildItem -Path $Path | Get-Unique | Sort-Object | ForEach-Object {
    $_.FullName + ';' | Add-Content -NoNewline $PreRoll
    $_.FullName + ';' | Write-Host -NoNewline
}
[System.Environment]::NewLine
Pause