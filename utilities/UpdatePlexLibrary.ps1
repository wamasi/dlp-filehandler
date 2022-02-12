$TempDir = $PSScriptRoot -replace '\\[^\\]*$', ''
$ConfigPath = "$TempDir\config.xml"
Write-Host $ConfigPath
[xml]$ConfigFile = Get-Content -Path $ConfigPath -Raw
$PlexHost = $ConfigFile.configuration.Plex.hosturl.url
$PlexToken = $ConfigFile.configuration.Plex.plextoken.token
$nodelist = $ConfigFile.SelectNodes("//*[@libraryname]")
$nodelist | Where-Object {$_.libraryname -eq "Movies(English)"} | ForEach-Object {
    $libraryId =  $_.libraryid
    $plexurl = "$PlexHost/library/sections/$libraryId/refresh?X-Plex-Token=$PlexToken"
    Write-Host $libraryId
    Write-Host $plexurl
}
