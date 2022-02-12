#Assumes you remembered to name your files correctly(BaseName)
$Source = "E:\Videos", "F:\Videos", "G:\Videos", "H:\Videos"
$SortedFiles= "D:\Common\yt-dlp\Logs\LogFiles\SortByFilename.txt"
Get-ChildItem -Path $Source -Recurse -File -Exclude "*.jpg", "*.png", "*.cbz", "*.srt", "*.vtt", "*.ass", "*.sub", "*.idx", "*.mp3"| Where-Object {$_.FullName -notmatch "\\ME\\|\\MJ\\"} | Sort-Object -Property BaseName | ForEach-Object {
    $_.FullName
} | Out-File $SortedFiles