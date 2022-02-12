Get-Content 'D:\Common\yt-dlp\Logs\WorkingFiles\1_URL.txt' | Get-Unique | Sort-Object | Out-File 'D:\Common\yt-dlp\Logs\WorkingFiles\2_URL.txt'



#  (?=https://www.paramountplus.com/shows/doug/).*(?=.*/doug).*(?=.*").*