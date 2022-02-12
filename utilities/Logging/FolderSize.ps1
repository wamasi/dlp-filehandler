$Source = "E:\", "F:\", "G:\", "H:\"
$ExcludedFolders = "\\tmp"
Get-ChildItem -Path $Source -Recurse -Directory -Depth 1 | Where-Object { $_.FullName -notmatch $ExcludedFolders -and ($_.FullName.length -lt 8 -or $_.FullName.length -gt 9 ) } | ForEach-Object {
    $len = 0
    Get-ChildItem -Recurse -Force $_.fullname -ErrorAction SilentlyContinue | ForEach-Object { $len += $_.length }
    $_.fullname , '{0:N2} GB' -f ($len / 1tb) | Export-Csv -Path "D:\Common\yt-dlp\Logs\TEST.csv" -Delimiter ','
}