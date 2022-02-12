$path = "G:\Videos\SJ\Teasing Master Takagi-san\Season 02\Teasing Master Takagi-san - S02E01 - Textbook + Hypnosis + Waking Up + Skipping Stones.mkv"
$SF = (split-path $path -Parent)  + "\" + (Split-Path $path -LeafBase)
$json= $SF + ".json"
Write-Host $SF
ffmpeg -i $path "$SF.ass"
ffprobe -v quiet -print_format json -show_format -show_streams $path > $json