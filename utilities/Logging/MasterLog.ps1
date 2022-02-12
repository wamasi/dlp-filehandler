
#Drives
$Source = "E:\", "F:\", "G:\", "H:\", "I:\"

#Include/Excluded Folders and File Extensions
$IncludedFolders = "\\Videos\\|\\Books\\"
$ExcludedFolders = "\\_YT-DLP\\|\\_Staged\\|\\_SeriesYoutube\\|\\tmp\\"
$Extensions = ("*.torrent", "*.exe", "*.nfo", "*.txt", "*.png", "*.jpeg", "*.jpg", "*.bmp", "*.gif", "*.zip", "*.part", "*.parts", "*.html", "*.dll", "*.lnk", "*.sys", "*.tmp", "*.msi", "*.ico", "*.srt", "*.ass", "*.vtt", "*.sub")

#Log/Inventory Files
$FileInventory = "D:\Common\yt-dlp\Logs\MasterLists\FileInventory.txt"
$InvFileContent = Get-Content $FileInventory
$FileInventoryMaster = "D:\Common\yt-dlp\Logs\MasterLists\FileInventory_Master.txt"
$MasterTree = "D:\Common\yt-dlp\Logs\MasterLists\MasterTree.txt"
$TotalFiles = "D:\Common\yt-dlp\Logs\MasterLists\FileTotals.txt"
$DuplicateFile = "D:\Common\yt-dlp\Logs\LogFiles\VideoDuplicates.txt"
$OrphanFile = "D:\Common\yt-dlp\Logs\LogFiles\VideoOrphans.txt"
#$VideoLength = "D:\Common\yt-dlp\Logs\FileLengthSortedMaster.csv"
$TempDummyDir = "D:\tmp\_TreeStructure\"
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}

#01 Collecting Files in Video directories. stripping drive letter from paths and inserting them into text file
Write-Output "$(Get-Timestamp) - 01/09 - Starting Inventory Update"
if (-not(Test-Path -Path $FileInventory -PathType Leaf)) {
    try {
        $null = New-Item -ItemType File -Path $FileInventory -Force 
        Write-Output "$(Get-Timestamp) - 01/09 - [$FileInventory] has been created."
    }
    catch {
        throw $_.Exception.Message
    }
    else {
        Write-Output "$(Get-Timestamp) - 01/09 - [$FileInventory] already exists."
    }
}

#02
Write-Output "$(Get-Timestamp) - 02/09 - Updating [$FileInventory] and [$FileInventoryMaster] "
Get-ChildItem -Path $Source -Recurse -Exclude $Extensions -File | Where-Object { $_.FullName -match $IncludedFolders -and $_.FullName -notmatch $ExcludedFolders -and $_.FullName.substring(3) -notin $InvFileContent } | ForEach-Object {
    $_.FullName.substring(3) | Out-File -Append "$FileInventory"
    $_.FullName | Out-File -Append "$FileInventoryMaster"
}
Set-Content -Path $FileInventoryMaster -Value (Get-Content -Path $FileInventoryMaster | Get-Unique)
Set-Content -Path $FileInventory -Value (Get-Content -Path $FileInventory | Get-Unique)
Write-Output "$(Get-Timestamp) - 02/09 - [$FileInventory] and [$FileInventoryMaster] Updated"

#03 Generating Dummy Directory from _FileInventory.txt to be used in consolidated Tree text file
Write-Output "$(Get-Timestamp) - 03/09 - Generating Dummy File Directory"
$SourceTree = $InvFileContent
foreach ($files in $SourceTree) {
    $dummy = $TempDummyDir + ${files}
    New-Item -Path $dummy -Force > $null
}
Write-Output "$(Get-Timestamp) - 03/09 - Dummy File Directory Generated"

#04 Output combined Tree file from all drives
Write-Output "$(Get-Timestamp) - 04/09 - Generating MasterTree file"
Show-Tree -Path $TempDummyDir -ShowLeaf | Out-File -Encoding utf8 "$MasterTree"
Write-Output "$(Get-Timestamp) - 04/09 - MasterTree file generated"

#05 Outputting total counts of videos per category
Write-Output "$(Get-Timestamp) - 05/09 - Updating [$TotalFiles] counts"
"Last Updated: " + $(Get-Date -Format 'dddd yyyy-MMM-dd HH:mm K') | Set-Content "$TotalFiles"
$SJ = Select-String -InputObject $InvFileContent -Pattern "Videos\\SJ\\" -AllMatches
("Series(JP) Total: " + $SJ.Matches.Count) | Add-Content "$TotalFiles"
$MJ = Select-String -InputObject $InvFileContent -Pattern "Videos\\MJ\\" -AllMatches
("Movies(JP) Total: " + $MJ.Matches.Count) | Add-Content "$TotalFiles"
$SE = Select-String -InputObject $InvFileContent -Pattern "Videos\\SE\\" -AllMatches 
("Series(EN) Total: " + $SE.Matches.Count) | Add-Content "$TotalFiles"
$ME = Select-String -InputObject $InvFileContent -Pattern "Videos\\ME\\" -AllMatches
("Movies(EN) Total: " + $ME.Matches.Count) | Add-Content "$TotalFiles"
$BB = Select-String -InputObject $InvFileContent -Pattern "Books\\Books\\" -AllMatches
("Books(BK) Total: " + $BB.Matches.Count) | Add-Content "$TotalFiles"
$BM = Select-String -InputObject $InvFileContent -Pattern "Books\\Manga\\" -AllMatches
("Books(BM) Total: " + $BM.Matches.Count) | Add-Content "$TotalFiles"
$BC = Select-String -InputObject $InvFileContent -Pattern "Books\\Comics\\" -AllMatches
("Books(BC) Total: " + $BC.Matches.Count) | Add-Content "$TotalFiles"
"Overall Video Total: " + ($SJ.Matches.Count + $MJ.Matches.Count + $SE.Matches.Count + $ME.Matches.Count) | Add-Content "$TotalFiles"
"Overall Book Total: " + ($BB.Matches.Count + $BM.Matches.Count + $BC.Matches.Count) | Add-Content "$TotalFiles"
"Overall File Total: " + ($SJ.Matches.Count + $MJ.Matches.Count + $SE.Matches.Count + $ME.Matches.Count + $BB.Matches.Count + $BM.Matches.Count + $BC.Matches.Count) | Add-Content "$TotalFiles"
Write-Output "$(Get-Timestamp) - 05/09 - [$TotalFiles] counts update completed"

#06 Cleaning up Dummy Directory
Write-Output "$(Get-Timestamp) - 06/09 - Deleting Temp Dummy Directory"
Remove-Item -Path $TempDummyDir -Recurse -Force
Write-Output "$(Get-Timestamp) - 06/09 - Deleted Temp Dummy Directory"

#07 Looking for Duplicates or Orphan/Mismatched
Write-Output "$(Get-Timestamp) - 07/09 - Searching for Duplicates"
pwsh.exe -executionpolicy remotesigned -File "D:\Common\yt-dlp\utilities\Logging\Duplicates.ps1"
$nlines = 0;
Get-Content $DuplicateFile -read 1000 | ForEach-Object { $nlines += $_.Length } ;
[string]::Format("$(Get-Timestamp) - 07/09 - {0} has {1} Duplicate(s)", $DuplicateFile, $nlines)
Write-Output "$(Get-Timestamp) - 07/09 - Search for Duplicates Complete"

#08 Looking for Orphan/Mismatched
Write-Output "$(Get-Timestamp) - 08/09 - Searching for Orphan/Mismatched files"
pwsh.exe -executionpolicy remotesigned -File "D:\Common\yt-dlp\utilities\Logging\OrphanFiles.ps1"
$nlines = 0;
Get-Content $OrphanFile -read 10000 | ForEach-Object { $nlines += $_.Length };
[string]::Format("$(Get-Timestamp) - 08/09 - {0} has {1} Orphan(s)", $OrphanFile, $nlines)
Write-Output "$(Get-Timestamp) - 08/09 - Search for Orphan/Mismatched files Complete"

#09 Getting Video Lengths
# Write-Output "$(Get-Timestamp) - 10/09 - Getting Video Lengths"
# pwsh.exe -executionpolicy remotesigned -File "D:\Common\yt-dlp\_scripts\CopyFiles\VideoLengths.ps1"
# Get-ChildItem $VideoLength -re -in "*.csv" | ForEach-Object { 
#     $fileStats = Get-Content $_.FullName | Measure-Object -Line
#     $linesInFile = $fileStats.Lines - 1
#     Write-Host "$(Get-Timestamp) $linesInFile lines in $VideoLength" 
# }
# Write-Output "$(Get-Timestamp) - 10/09 - Video Length File Updated"

#10 Complete
Write-Output "$(Get-Timestamp) - 10/10 - Inventory Update Complete"