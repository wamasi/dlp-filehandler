function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}
Function Get-VideoDetails {
    param ($targetDirectory)
    $LengthColumn = 27
    $CommentColumn = 24
    $HeightColumn = 178
    $WidthColumn = 176
    $HeightResColumn = 177
    $WidthResColumn = 175
    $URLColumn = 204

    $objShell = New-Object -ComObject Shell.Application
    Get-ChildItem -LiteralPath $targetDirectory -Include *.avi, *.mp4, *.mvk -Recurse -Force | ForEach-Object {
        if ($_.Extension -in ".avi", ".mp4", ".mkv") {
            $objFolder = $objShell.Namespace($_.DirectoryName)
            $objFile = $objFolder.ParseName($_.Name)
            $Duration = $objFolder.GetDetailsOf($objFile, $LengthColumn)
            $Comments = $objFolder.GetDetailsOf($objFile, $CommentColumn)
            $Height = $objFolder.GetDetailsOf($objFile, $HeightColumn)
            $Width = $objFolder.GetDetailsOf($objFile, $WidthColumn)
            $HieghtRes = $objFolder.GetDetailsOf($objFile, $HeightResColumn)
            $WidthRes = $objFolder.GetDetailsOf($objFile, $WidthResColumn)
            $URL = $objFolder.GetDetailsOf($objFile, $URLColumn)
            [PSCustomObject]@{
                Drive     = $_.FullName.Substring(0, 3)
                FilePath  = $_.FullName
                Name      = $_.Name
                Size      = "$([int]($_.length / 1mb)) MB"
                Duration  = $Duration
                Comments  = $Comments
                Width     = $Width
                Height    = $Height
                HeightRes = $HieghtRes
                WidthRes  = $WidthRes
                URL       = $URL
            }
        }
    }
}

$TempCSV = "D:\Common\yt-dlp\Logs\WorkingFiles\FileLength.csv"
$MasterCSV = "D:\Common\yt-dlp\Logs\WorkingFiles\FileLengthSortedMaster.csv"
Remove-Item $MasterCSV
if (-not(Test-Path -Path $TempCSV -PathType Leaf)) {
    try {
        $null = New-Item -ItemType File -Path $TempCSV -Force -ErrorAction Stop
        Write-Output "$(Get-Timestamp) - [$TempCSV] has been created."
    }
    catch {
        throw $_.Exception.Message
    }
}
else {
    Write-Output "$(Get-Timestamp) - [$TempCSV] already exists."
}

$Paths = "E:\Videos\SJ", "F:\Videos\SJ", "G:\Videos\SJ", "H:\Videos\SJ", "I:\Videos\SJ"
Get-VideoDetails $Paths | Export-Csv -Append -Force $TempCSV
Import-Csv $TempCSV | Export-Csv -Path $MasterCSV
Remove-Item $TempCSV