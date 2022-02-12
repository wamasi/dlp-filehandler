$Path = "E:\", "F:\", "G:\", "H:\", "I:\"
$DuplicateFile = "D:\Common\yt-dlp\Logs\LogFiles\VideoDuplicates.txt"
Get-ChildItem -Path $Path -Recurse -File -Exclude "*.jpg", "*.png", "*.cbz", "*.srt", "*.vtt", "*.ass", "*.sub", "*.idx","*.mp3" | Group-Object -Property BaseName | Where-Object { $_.Count -gt 1 } | Select-Object -ExpandProperty group | ForEach-Object { $_.FullName } | Out-File $DuplicateFile