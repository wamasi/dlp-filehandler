$Source = "E:\", "F:\", "G:\", "H:\"
$OrphanFile = "D:\Common\yt-dlp\Logs\LogFiles\Other\VideoOrphans.txt"
$IncludedFolders = "\\Videos\\"
$ExcludedFolders = "\\_YT-DLP\\|\\_Staged\\|\\_SeriesYoutube\\"
$Extensions = ("*.torrent", "*.exe", "*.nfo", "*.txt", "*.png", "*.jpeg", "*.jpg", "*.bmp", "*.gif", "*.zip", "*.part", "*.parts", "*.html", "*.dll", "*.lnk", "*.sys", "*.tmp", "*.msi", "*.ico", "*.srt", "*.ass", "*.vtt", "*.sub")
# Collecting Files in Video directories. stripping drive letter from paths and inserting them into text file
Get-ChildItem -Path $Source -Recurse -Depth 3 -Exclude $Extensions -File | Where-Object { $_.FullName -match $IncludedFolders -and ($_.FullName -match '\\SJ\\' -or $_.FullName -match '\\SE\\') -and $_ -notmatch $ExcludedFolders } | Where-Object { $_.FullName -notcontains "\\Season" } | ForEach-Object {
    $_.DirectoryName | Out-File -Append $OrphanFile
}
Set-Content -Path $OrphanFile -Value (Get-Content -Path $OrphanFile | Get-Unique)