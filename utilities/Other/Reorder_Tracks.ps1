
#Keeping originals and moving to a final folder
$Path = "E:\Videos\SJ\Ah! My Goddess\Season 01", "E:\Videos\SJ\Ah! My Goddess\Season 02"
Get-ChildItem $path -Recurse -File -Include "*.mkv" | ForEach-Object {
    #ffmpeg/powershell aren't playing nice with naming so this is intended jank
    $_.BaseName + $_.Extension | Write-Host
    $_.DirectoryName | Write-Host
    $LFolder = $_.DirectoryName + "\Final\" + $_.BaseName + $_.Extension 
    $LFolder | Write-Host
    #setting default audio and subtitles. adjust accordingly
    # a: = audio
    # s: = subtitles
    ffmpeg -i $_.FullName -map 0 -c copy -disposition:a:0 0 -disposition:a:1 default -disposition:s:0 0 -disposition:s:2 default $LFolder
}

#Spot check files in Final Folder
ffprobe -hide_banner "E:\Videos\SJ\Ah! My Goddess\Season 01\Ah! My Goddess - S01E01 - Ah! Are You a Goddess.mkv"