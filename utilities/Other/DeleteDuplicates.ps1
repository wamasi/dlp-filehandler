$path = 'D:\Common\yt-dlp\Logs\LogFiles\VideoDuplicates.txt'
Get-Content $path | Sort-Object | Select-Object -Unique | ForEach-Object {
    Remove-Item $_ | Write-Host
}
