$Path = ""
Get-ChildItem $Path | Where-Object { $_.Extension -eq '.mkv' } | ForEach-Object {
    Write-Host $_.FullName
    # mkvpropedit $_.FullName --edit track:a2 --set flag-default=1
    # mkvpropedit $_.FullName --edit track:a1 --set flag-default=0
    # mkvpropedit $_.FullName --edit track:s2 --set flag-default=1
    # mkvpropedit $_.FullName --edit track:s1 --set flag-default=0
}