[CmdletBinding()]
param(
    [Alias('A')]
    [switch]$Automated,
    [Alias('G')]
    [switch]$GenerateAnilistFile,
    [Alias('U')]
    [switch]$updateAnilistCSV,
    [Alias('D')]
    [switch]$SetDownload,
    [Alias('B')]
    [switch]$GenerateBatchFile,
    [Alias('SD')]
    [switch]$SendDiscord,
    [Alias('DS')]
    [switch]$DebugScript,
    [Alias('O')]
    [switch]$Override,
    [Alias('I')]
    [string]$MediaID_In,
    [Alias('NI')]
    [string]$MediaID_NotIn,
    [Alias('NSC')]
    [switch]$newShowCheck
)
$scriptRoot = $PSScriptRoot
$configFilePath = Join-Path $scriptRoot -ChildPath 'config.xml'
if (!(Test-Path $configFilePath)) {
    $baseXML = @'
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <Directory>
    <backup location="D:\Backup" />
  </Directory>
  <characters>
    <char>’</char>
  </characters>
  <Logs>
    <keeplog emptylogskeepdays="0" filledlogskeepdays="7" />
    <lastUpdated></lastUpdated>
  </Logs>
  <Discord>
    <hook ScheduleServerUrl="" MediaServerUrl="" />
    <icon default="" author="" footerIcon="" Color="" />
    <sites>
      <site siteName="" emoji="" />
      <site siteName="" emoji="" />
      <site siteName="Default" emoji="" />
    </sites>
  </Discord>
</configuration>
'@
    New-Item $configFilePath
    Set-Content $configFilePath -Value $baseXML
    Write-Host "No config found. Configure xml file: $configFilePath"
    exit
}
[xml]$config = Get-Content $configFilePath
if ($debugScript) {
    $discordHookUrl = $config.configuration.Discord.hook.TestServerUrl
}
else {
    $discordHookUrl = $config.configuration.Discord.hook.ScheduleServerUrl
}
$discordSiteIcon = $config.configuration.Discord.icon.default
$siteFooterIcon = $config.configuration.Discord.icon.footerIcon
$rootPath = Join-Path $scriptRoot -ChildPath 'batch'
$batchArchiveFolder = Join-Path $rootPath -ChildPath '_archive'
$lastUpdated = $config.configuration.logs.lastUpdated
$emojis = $config.configuration.Discord.sites.site
$badchar = $config.configuration.characters.char
$logFolder = Join-Path $scriptRoot -ChildPath 'log'
$logfile = Join-Path $logFolder -ChildPath 'anilistSeason.log'
$csvFolder = Join-Path $scriptRoot -ChildPath 'anilist'
$badURLs = 'https://www.crunchyroll.com/', 'https://www.crunchyroll.com', 'https://www.hidive.com/', 'https://www.hidive.com'
$supportSites = 'Crunchyroll', 'HIDIVE', 'Netflix', 'Hulu'
$SiteCountList = @()
foreach ($ss in $supportSites) { $a = [PSCustomObject]@{Site = $ss; ShowCount = 0 }; $SiteCountList += $a }
$spacer = "`n" + '-' * 120
$now = Get-Date
$today = Get-Date $now -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$cby = (Get-Date $today -Format 'yyyy').ToString()
$pby = (Get-Date(($today).AddYears(-1)) -Format 'yyyy').ToString()
$cbyShort = $cby.Substring(2)
$pbyShort = $pby.Substring(2)
$csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_$($pbyShort)-$($cbyShort).csv"
$allShows = @()
$allEpisodes = @()
$data = @()
$anime = @()
$updatedRecordCalls = @()
$updatedRecords = @()
$chunkSize = 20
$SeasonDates = @(
    [PSCustomObject]@{Season = 'FALL'; Ordinal = 1 ; StartDate = "10/01/$pby"; EndDate = "12/31/$pby" }
    [PSCustomObject]@{Season = 'WINTER'; Ordinal = 2; StartDate = "01/01/$cby"; EndDate = "03/31/$cby" }
    [PSCustomObject]@{Season = 'SPRING'; Ordinal = 3 ; StartDate = "03/01/$cby"; EndDate = "07/31/$cby" }
    [PSCustomObject]@{Season = 'SUMMER'; Ordinal = 4 ; StartDate = "08/01/$cby"; EndDate = "09/30/$cby" }
    [PSCustomObject]@{Season = 'FALL'; Ordinal = 5 ; StartDate = "10/01/$cby"; EndDate = "12/31/$cby" }
)
$csvStartDate = $SeasonDates | Where-Object { $_.Ordinal -eq 1 } | Select-Object -ExpandProperty StartDate
$csvEndDate = $SeasonDates | Where-Object { $_.Ordinal -eq 5 } | Select-Object -ExpandProperty EndDate
$csvStartDate = Get-Date $csvStartDate -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$csvEndDate = Get-Date $csvEndDate -Hour 0 -Minute 0 -Second 0 -Millisecond 0
$seasonOrder = [ordered]@{}
foreach ($season in $SeasonDates | Where-Object { $_.Ordinal -ne 1 }) { $seasonOrder[$season.Season] = $season.Ordinal }
$weekdayOrder = [ordered]@{
    Monday    = 1
    Tuesday   = 2
    Wednesday = 3
    Thursday  = 4
    Friday    = 5
    Saturday  = 6
    Sunday    = 7
}
function Write-Log {
    param (
        [string]$Message,
        [string]$LogFilePath
    )
    Write-Host $Message
    if (!(Test-Path $LogFilePath)) { New-Item $LogFilePath -ItemType File }
    Add-Content -Path $LogFilePath -Value $Message
}
function Get-DateTime {
    param (
        [int]$dateType
    )
    switch ($dateType) {
        1 { $datetime = Get-Date -Format 'yy-MM-dd' }
        2 { $datetime = Get-Date -Format 'MMddHHmmssfff' }
        3 { $datetime = ($(Get-Date).ToUniversalTime()).ToString('yyyy-MM-ddTHH:mm:ss.fffZ') }
        Default { $datetime = Get-Date -Format 'yy-MM-dd HH-mm-ss' }
    }
    return $datetime
}
function Get-UnixToLocal {
    param (
        $unixTime
    )
    $d = (Get-Date -Date ([DateTimeOffset]::FromUnixTimeSeconds($unixTime)).DateTime).ToLocalTime()
    return $d
}
function Get-FullUnixDay {
    param (
        [DateTime]$date
    )
    $endDate = $date.AddDays(1)
    $epochStart = (Get-Date '1970-01-01 00:00:00').ToLocalTime()
    $startDateOffset = if ($date.IsDaylightSavingTime()) { [TimeSpan]::FromHours(1) } else { [TimeSpan]::FromHours(0) }
    $endDateOffset = if ($endDate.IsDaylightSavingTime()) { [TimeSpan]::FromHours(1) } else { [TimeSpan]::FromHours(0) }
    $startDateUnix = [Math]::Floor(($date.Add(-$startDateOffset)).Subtract($epochStart).TotalSeconds)
    $endDateUnix = [Math]::Floor(($endDate.Add(-$endDateOffset)).Subtract($epochStart).TotalSeconds)
    return $startDateUnix, $endDateUnix
}
function Exit-Script {
    Write-Log -Message "[End] $(Get-DateTime) - End of script.$spacer" -LogFilePath $logFile
    exit
}
function Set-CSVFormatting {
    param (
        $csvPath
    )
    if ((Test-Path $csvPath)) {
        $csvd = Import-Csv $csvPath
        foreach ($d in $csvd) {
            $parts = $($d.AiringTime) -split ':'
            $d.AiringTime = "$($parts[0].PadLeft(2, '0')):$($parts[1])"
            $d.AiringFullTime = Get-Date $($d.AiringFullTime) -Format 'MM/dd/yyyy HH:mm:ss'
            $d.episode = '{0:D2}' -f [int]$d.episode
            $d.TotalEpisodes = '{0:D2}' -f [int]$d.TotalEpisodes
            switch ($d.download) {
                'True' { $d.download = 'True' }
                'False' { $d.download = 'False' }
            }
        }
        $csvd | Sort-Object Site, SeasonYear, { $SeasonOrder[$_.Season] }, Title, episode, TotalEpisodes | Select-Object * -Unique | Export-Csv $csvPath -NoTypeInformation
    }
}
function Update-Records {
    param (
        $source,
        $target
    )
    $newDataArray = @()
    foreach ($t in $target) {
        $s = $source | Where-Object { $_.ShowId -eq $t.ShowId -and $_.Episode -eq $t.Episode }
        $newData = [ordered]@{}
        if ($s) {
            foreach ($property in $s.PSObject.Properties.Name) {
                if ($property -eq 'Download') {
                    $newData[$property] = $s.$property
                }
                elseif ($s.$property -eq $t.$property) {
                    $newData[$property] = $t.$property
                }
                else {
                    $newData[$property] = ($t.$property)
                }
            }
        }
        else {
            foreach ($property in $t.PSObject.Properties.Name) {
                $newData[$property] = $t.$property
            }
        }
        $newDataArray += [PSCustomObject]$newData
    }
    foreach ($s in $source) {
        $n = $newDataArray | Where-Object { $_.ShowId -eq $s.ShowId -and $_.Episode -eq $s.Episode }
        if (-not $n) {
            $newData = [ordered]@{}
            foreach ($property in $s.PSObject.Properties.Name) {
                $newData[$property] = $s.$property
            }
            $newDataArray += [PSCustomObject]$newData
        }
    }
    $newDataArrayFiltered = $newDataArray | Where-Object { $null -ne $_.PSObject.Properties.Value -and $_.PSObject.Properties.Value -ne '' }
    return $newDataArrayFiltered
}
function Invoke-Request {
    param (
        $rUrl,
        $rBody
    )
    $maxRetries = 3
    $retryIntervalSec = 5
    for ($i = 0; $i -lt $maxRetries; $i++) {
        try {
            $rResponse = Invoke-WebRequest -Uri $rUrl -Method POST -Body $rBody -ContentType 'application/json'
            if ($rResponse.StatusCode -eq 200) {
                break
            }
            elseif ($rResponse.StatusCode -ge 500 -and $rResponse.StatusCode -lt 600) {
                Write-Host 'Received HTTP 500 error. Retrying...'
                Start-Sleep -Seconds $retryIntervalSec
            }
            else {
                Write-Host "Received unexpected status code $($rResponse.StatusCode). Exiting."
                break
            }
        }
        catch {
            Write-Host "An error occurred: $_. Retrying..."
            Start-Sleep -Seconds $retryIntervalSec
        }
    }
    return $rResponse
}
function Invoke-AnilistApiShowDate {
    param (
        $start,
        $end
    )
    $allSeasonShows = @()
    $pageNum = 1
    $perPageNum = 20
    $url = 'https://graphql.anilist.co'
    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('method', 'POST')
    $headers.Add('authority', 'graphql.anilist.co')
    $headers.Add('Content', 'application/json')
    $headers.Add('Content-Type', 'application/json')
    do {
        $initialBody = [PSCustomObject]@{
            query = @"
{
  Page(page: $pageNum, perPage: $perPageNum) {
    airingSchedules(airingAt_greater: $start, airingAt_lesser: $end, sort: TIME) {
      media {
        id
        title {
          english
          romaji
          native
        }
        episodes
        genres
        status
        season
        seasonYear
        startDate {
          year
          month
          day
        }
        endDate {
          year
          month
          day
        }
        siteUrl
        externalLinks {
          site
          url
        }
      }
    }
    pageInfo {
      hasNextPage
      lastPage
    }
  }
}
"@
        } | ConvertTo-Json
        $response = Invoke-Request -rUrl $url -rBody $initialBody
        $c = $response.Content | ConvertFrom-Json
        [int]$remaining = $response.Headers.'X-RateLimit-Remaining'[0]
        $nextPage = $c.data.Page.pageInfo.hasNextPage
        $allSeasonShows += $c.data.Page.airingSchedules.media
        | Select-Object @{name = 'ShowId'; expression = { $_.id } }, @{name = 'EpisodeId'; expression = { $_.id } }, @{name = 'episodes'; expression = { $_.episodes } }, `
            airingAt, @{name = 'TitleEnglish'; expression = { $_.title.english } }, @{name = 'TitleRomanji'; expression = { $_.title.romaji } }, `
        @{name = 'TitleNative'; expression = { $_.title.native } }, @{name = 'Genres'; expression = { $_.genres -join ', ' } }, `
        @{name = 'Status'; expression = { $_.status } }, @{name = 'Season'; expression = { $_.season } }, @{name = 'SeasonYear'; expression = { $_.seasonYear } } , `
        @{name = 'StartDate'; expression = { "$($_.startDate.month)/$($_.startDate.day)/$($_.startDate.year)" } }, `
        @{name = 'EndDate'; expression = { "$($_.endDate.month)/$($_.endDate.day)/$($_.endDate.year)" } }, externalLinks
        if ($nextPage) { $pageNum++; Write-Host "Next Page: $pageNum - $start/$end" }
        if ($remaining -le 5) {
            Write-Host 'Pausing for 60 second(s).'
            Start-Sleep -Seconds 60
        }
        else {
            $delayBetweenRequests = [math]::Ceiling((60 / ($remaining + 1)))
            Start-Sleep -Seconds $delayBetweenRequests
        }
    } while ( $nextPage )
    return $allSeasonShows
}
function Invoke-AnilistApiShowSeason {
    param (
        $id,
        $year,
        $season
    )
    $allSeasonShows = @()
    $pageNum = 1
    $perPageNum = 20
    $url = 'https://graphql.anilist.co'
    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('method', 'POST')
    $headers.Add('authority', 'graphql.anilist.co')
    $headers.Add('Content', 'application/json')
    $headers.Add('Content-Type', 'application/json')
    if ($id) {
        $mediaFilter = "id: $id"
        Write-Host "Start Page: $id - $pageNum"
    }
    else {
        $mediaFilter = "seasonYear: $year, season: $season"
        Write-Host "Start Page: $season - $year - $pageNum"
    }
    do {
        $initialBody = [PSCustomObject]@{
            query = @"
{
  Page(page: $pageNum, perPage: $perPageNum) {
    media($mediaFilter, type: ANIME) {
      id
      episodes
      title {
        english
        romaji
        native
      }
      genres
      status
      season
      seasonYear
      startDate {
        year
        month
        day
      }
      endDate {
        year
        month
        day
      }
      externalLinks {
        site
        url
      }
      siteUrl
    }
    pageInfo {
      hasNextPage
    }
  }
}
"@
        } | ConvertTo-Json
        $response = Invoke-Request -rUrl $url -rBody $initialBody
        $c = $response.Content | ConvertFrom-Json
        [int]$remaining = $response.Headers.'X-RateLimit-Remaining'[0]
        $nextPage = $c.data.Page.pageInfo.hasNextPage
        $allSeasonShows += $c.data.Page.media
        | Select-Object @{name = 'ShowId'; expression = { $_.id } }, @{name = 'EpisodeId'; expression = { $_.id } }, `
            airingAt, @{name = 'TitleEnglish'; expression = { $_.title.english } }, @{name = 'TitleRomanji'; expression = { $_.title.romaji } }, `
        @{name = 'TitleNative'; expression = { $_.title.native } }, @{name = 'Genres'; expression = { $_.genres -join ', ' } }, `
        @{name = 'Status'; expression = { $_.status } }, @{name = 'Season'; expression = { $_.season } }, @{name = 'SeasonYear'; expression = { $_.seasonYear } } , `
        @{name = 'StartDate'; expression = { "$($_.startDate.month)/$($_.startDate.day)/$($_.startDate.year)" } }, `
        @{name = 'EndDate'; expression = { "$($_.endDate.month)/$($_.endDate.day)/$($_.endDate.year)" } }, externalLinks
        if ($nextPage) {
            $pageNum++
            if ($season) {
                Write-Host "Next Page: $season - $year - $pageNum"
            }
            else {
                Write-Host "Next Page: $id - $pageNum"
            }
        }
        if ($remaining -le 5) {
            Start-Sleep -Seconds 60
        }
        else {
            $delayBetweenRequests = 60 / ($remaining + 1)
            Start-Sleep -Seconds $delayBetweenRequests
        }
    } while ( $nextPage )
    return $allSeasonShows
}
function Invoke-AnilistApiEpisode {
    param (
        $id,
        $eStart,
        $eEnd
    )
    $ae = @()
    $pageNum = 1
    $perPageNum = 20
    $url = 'https://graphql.anilist.co'
    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('method', 'POST')
    $headers.Add('authority', 'graphql.anilist.co')
    $headers.Add('Content', 'application/json')
    $headers.Add('Content-Type', 'application/json')
    Write-Host "Starting page: $pageNum"
    $mediaFilter = "mediaId_in: [$($id)], airingAt_greater: $eStart, airingAt_lesser: $eEnd"
    do {
        $episodesBody = [PSCustomObject]@{
            query         = @"
{
  Page(page: $($pageNum), perPage: $perPageNum) {
    airingSchedules($mediaFilter) {
      id
      episode
      airingAt
      media {
        id
      }
    }
    pageInfo {
      hasNextPage
    }
  }
}
"@
            variables     = $null
            operationName = $null
        } | ConvertTo-Json
        $response = Invoke-Request -rUrl $url -rBody $episodesBody
        [int]$remaining = $response.Headers.'X-RateLimit-Remaining'[0]
        $c = $response.Content | ConvertFrom-Json
        $nextPage = $c.data.page.pageInfo.hasNextPage
        $ae += $c.data.page.airingSchedules | Select-Object @{name = 'ShowId'; expression = { $_.media.id } }, @{name = 'EpisodeId'; expression = { $_.id } }, episode, airingAt
        Write-Host "added: $($s.count)"
        if ($nextPage) { $pageNum++; Write-Host "Next Page: $pageNum" }
        if ($remaining -le 6) {
            Start-Sleep -Seconds 60
        }
        else {
            $delayBetweenRequests = 60 / ($remaining + 2)
            Start-Sleep -Seconds $delayBetweenRequests
        }
    } while ( $nextPage )
    return $ae
}
function Invoke-AnilistApiDateRange {
    param (
        [Parameter(Mandatory = $true)]
        $start,
        [Parameter(Mandatory = $true)]
        $end,
        [Parameter(Mandatory = $false)]
        $ID_IN,
        [Parameter(Mandatory = $false)]
        $ID_NOTIN
    )
    $pageNum = 1
    $perPageNum = 20
    $url = 'https://graphql.anilist.co'
    $headers = New-Object 'System.Collections.Generic.Dictionary[[String],[String]]'
    $headers.Add('method', 'POST')
    $headers.Add('authority', 'graphql.anilist.co')
    $headers.Add('Content', 'application/json')
    $headers.Add('Content-Type', 'application/json')
    $allResponses = @()
    $mediaFilter = "airingAt_greater: $start, airingAt_lesser: $end"
    if ($null -ne $ID_IN) {
        if ($ID_IN -like '*,*') {
            $mediaFilter = $mediaFilter + ", mediaId_in: [$($ID_IN)]"
        }
        else {
            $mediaFilter = $mediaFilter + ", mediaId: $($ID_IN)"
        }
    }
    if ($null -ne $ID_NOTIN) {
        if ($ID_NOTIN -like '*,*') {
            $mediaFilter = $mediaFilter + ", mediaId_not_in: [$($ID_NOTIN)]"
        }
        else {
            $mediaFilter = $mediaFilter + ", mediaId_not: $($ID_NOTIN)"
        }
    }
    Write-Log -Message "[AnilistAPI] $(Get-DateTime) - $start - Starting on page $pageNum" -LogFilePath $logFile
    do {
        $body = [PSCustomObject]@{
            query         = @"
{
  Page(page: $pageNum, perPage: $perPageNum) {
    airingSchedules($mediaFilter) {
      id
      episode
      airingAt
      media {
        id
        episodes
        siteUrl
        title {
          english
          romaji
          native
        }
        genres
        status
        season
        seasonYear
        startDate {
          year
          month
          day
        }
        endDate {
          year
          month
          day
        }
        externalLinks {
          site
          url
        }
      }
    }
    pageInfo {
      hasNextPage
    }
  }
}
"@
            variables     = $null
            operationName = $null
        } | ConvertTo-Json
        $response = Invoke-Request -rUrl $url -rBody $body
        [int]$remaining = $response.Headers.'X-RateLimit-Remaining'[0]
        $c = $response.Content | ConvertFrom-Json
        $s = $c.data.page.airingSchedules
        | Select-Object @{name = 'ShowId'; expression = { $_.media.id } }, @{name = 'EpisodeId'; expression = { $_.id } }, `
            airingAt, episode, @{name = 'episodes'; expression = { $_.media.episodes } }, @{name = 'TitleEnglish'; expression = { $_.media.title.english } }, @{name = 'TitleRomanji'; expression = { $_.media.title.romaji } }, `
        @{name = 'TitleNative'; expression = { $_.media.title.native } }, @{name = 'Genres'; expression = { $_.media.genres -join ', ' } }, `
        @{name = 'Status'; expression = { $_.media.status } }, @{name = 'Season'; expression = { $_.media.season } }, @{name = 'SeasonYear'; expression = { $_.media.seasonYear } } , `
        @{name = 'StartDate'; expression = { "$($_.media.startDate.month)/$($_.media.startDate.day)/$($_.media.startDate.year)" } } , `
        @{name = 'EndDate'; expression = { "$($_.media.endDate.month)/$($_.media.endDate.day)/$($_.media.endDate.year)" } }, @{name = 'externalLinks'; expression = { $_.media.externalLinks } }
        $nextPage = $c.data.page.pageInfo.hasNextPage
        $allResponses += $s
        $nextPage = $aniapi.data.Page.pageInfo.hasNextPage
        if ($nextPage) {
            $pageNum++
            Write-Host "Next Page: $pageNum"
        }
        if ($remaining -le 5) {
            Start-Sleep -Seconds 60
        }
        else {
            $delayBetweenRequests = 60 / ($remaining + 1)
            Start-Sleep -Seconds $delayBetweenRequests
        }
    } while ( $nextPage )
    return $allResponses
}
if (!(Test-Path $logFolder)) {
    New-Item $logFolder -ItemType Directory
    New-Item $logFile -ItemType File
}
elseif (!(Test-Path $logfile)) {
    New-Item $logfile -ItemType File
}
Write-Log -Message "[Start] $(Get-DateTime) - Running for $today" -LogFilePath $logFile
if (!(Test-Path $csvFolder)) { New-Item $csvFolder -ItemType Directory }
if ($GenerateAnilistFile) {
    do {
        $aType = Read-Host 'Run for (S)eason or (D)ate?'
        switch ($aType) {
            'S' { $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_S_$($pbyShort)-$($cbyShort).csv" }
            'D' { $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_D_$($pbyShort)-$($cbyShort).csv" }
            Default { Write-Output 'Enter S or D' }
        }
    } while ( $aType -notin 'S', 'D' )
    if (Test-Path $csvFilePath) { Remove-Item -LiteralPath $csvFilePath }
    $dStartUnix = Get-FullUnixDay $csvStartDate
    $dEndunix = Get-FullUnixDay $csvEndDate
    if ($aType -eq 's') {
        # Fetching One Piece by ID since its season is 1999
        $sShow = Invoke-AnilistApiShowSeason -id 21
        $allShows += $sShow
        # Fetching the rest by CBY/PBY and season
        foreach ($seasonName in $seasonOrder.Keys) {
            if ($seasonName -eq 'FALL') {
                $sShow = Invoke-AnilistApiShowSeason -year $pby -season $seasonName
                $allShows += $sShow
            }
            $sShow = Invoke-AnilistApiShowSeason -year $cby -season $seasonName
            $allShows += $sShow
        }
    }
    else {
        $dShow = Invoke-AnilistApiShowDate -start $dStartUnix[0] -end $dEndunix[1]
        $allShows += $dShow
    }
    foreach ($a in $allShows) {
        if (-not [string]::IsNullOrEmpty($a.TitleEnglish)) {
            $title = $a.TitleEnglish -replace $badchar, "'"
        }
        elseif (-not [string]::IsNullOrEmpty($a.TitleRomanji)) {
            $title = $a.TitleRomanji
        }
        else {
            $title = $a.TitleNative
        }
        $genres = $a.genres -join ', '
        $totalEpisodes = $a.episodes
        if ($totalEpisodes -eq '' -or $null -eq $totalEpisodes) { $totalEpisodes = '00' }
        $episode = '{0:D2}' -f [int]$episode
        $totalEpisodes = '{0:D2}' -f [int]$totalEpisodes
        $season = if (-not [string]::IsNullOrEmpty($a.season)) { ([System.Globalization.CultureInfo]::CurrentCulture).TextInfo.ToTitleCase($($a.season).ToLower()) }
        else {
            ''
        }
        $seasonYear = $a.SeasonYear
        $status = if (-not [string]::IsNullOrEmpty($a.status)) { ([System.Globalization.CultureInfo]::CurrentCulture).TextInfo.ToTitleCase($($a.status).ToLower()) }
        else {
            ''
        }
        $showStartDate = if ($a.startDate -ne '//') { $a.startDate } else { $null }
        $showEndDate = if ($a.endDate -ne '//') { $a.endDate } else { $null }
        $firstPreferredSite = $null
        foreach ($ps in $supportSites) {
            $firstPreferredSite = $a.externalLinks | Where-Object { $_.site -eq $ps } | Select-Object * -First 1
            if ($firstPreferredSite) { break }
        }
        If ($firstPreferredSite) {
            $site = $firstPreferredSite.Site
            $url = $firstPreferredSite.url
            $url = ($url -replace 'http:', 'https:' -replace '^https:\/\/www\.crunchyroll\.com$', 'https://www.crunchyroll.com/' -replace '^https:\/\/www\.hidive\.com$', 'https://www.hidive.com/').Trim()
            if (($url -in $badURLs) -or ($site -eq 'HIDIVE' -and $totalEpisodes -eq '00') -or ($site -in 'Hulu', 'Netflix')) {
                $download = $false
            }
            else {
                $download = $True
            }
            Write-Log -Message "[Generate] $(Get-DateTime) - $site - $title - $download - $url" -LogFilePath $logFile
            $r = [PSCustomObject]@{
                ShowId        = $a.ShowId
                Title         = $title
                TotalEpisodes = [int]$totalEpisodes
                SeasonYear    = $seasonYear
                Season        = $season
                Status        = $status
                StartDate     = $showStartDate
                EndDate       = $showEndDate
                Site          = $site
                URL           = $url
                Genres        = $genres
                Download      = $download
            }
            $anime += $r
        }
    }
    $idList = $anime | Select-Object ShowId -Unique | ForEach-Object { $_.ShowId }
    $chunkedIdList = @()
    for ($i = 0; $i -lt $idList.Count; $i += $chunkSize) {
        $chunkedIdList += ($idList[$i..($i + $chunkSize - 1)] -join ', ')
    }
    foreach ($c in $chunkedIdList) {
        $allEpisodes += Invoke-AnilistApiEpisode -id $c -eStart $dStartUnix[0] -eEnd $dEndunix[1]
    }
    foreach ($s in $anime) {
        $showId = $s.ShowId
        $title = $s.Title
        $totalEpisodes = $s.TotalEpisodes
        $season = $s.season
        $seasonYear = $s.seasonYear
        $Status = $s.Status
        $startDate = $s.startDate
        $endDate = $s.endDate
        $site = $s.site
        $url = $s.url
        $genres = $s.genres
        $download = $s.download
        foreach ($e in ($allEpisodes | Where-Object { $_.ShowId -eq $showId })) {
            $episodeId = $e.EpisodeId
            $episode = $e.episode
            $airingFullTime = Get-UnixToLocal -unixTime $e.airingAt
            $airingFullTimeFixed = Get-Date $airingFullTime -Format 'MM/dd/yyyy HH:mm:ss'
            $airingDate = Get-Date $airingFullTime -Format 'MM/dd/yyyy'
            $airingTime = Get-Date $airingFullTime -Format 'HH:mm'
            $weekday = Get-Date $airingFullTime -Format 'dddd'
            Write-Output "$title - $episode - $airingFullTimeFixed - $genres"
            $a = [PSCustomObject]@{
                ShowId         = $showId
                EpisodeId      = $episodeId
                Title          = $title
                Episode        = [int]$episode
                TotalEpisodes  = [int]$totalEpisodes
                SeasonYear     = $seasonYear
                Season         = $season
                Status         = $Status
                AirDay         = $weekday
                AirDate        = $airingDate
                AiringTime     = $airingTime
                AiringFullTime = $airingFullTimeFixed
                StartDate      = $startDate
                EndDate        = $endDate
                Site           = $site
                URL            = $url
                Genres         = $genres
                Download       = $download
            }
            $data += $a
        }
    }
    Write-Output "[Setup] $(Get-DateTime) - Creating $csvFilePath"
    $data | Sort-Object Site, SeasonYear, { $SeasonOrder[$_.Season] }, Title, episode, TotalEpisodes | Select-Object * -Unique | Export-Csv $csvFilePath
}
if ($updateAnilistCSV) {
    if ($Automated) { $aType = 'D' }
    else {
        do {
            $aType = Read-Host 'Run for (S)eason or (D)ate?'
            if ($aType -notin 'S', 'D') {
                Write-Output 'Enter S or D'
            }
        } while ( $aType -notin 'S', 'D' )
    }
    switch ($aType) {
        'S' { $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_S_$($pbyShort)-$($cbyShort).csv" }
        'D' { $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_D_$($pbyShort)-$($cbyShort).csv" }
    }
    if (!(Test-Path $csvFilePath)) {
        Write-Log -Message "[Check] $(Get-DateTime) - File does not exist:  $csvFilePath" -LogFilePath $logFile
        Exit-Script
    }
    $startDate = $today
    $start = 0
    if ($Automated) { $maxDay = 8 } else { $maxDay = ($csvEndDate - $today).days }
    Set-CSVFormatting -csvPath $csvFilePath
    $csvFileData = Import-Csv $csvFilePath
    $chunkedIdList = @()
    
    if ($MediaID_In) {
        $oShowIdList = ($MediaID_In -replace ' ', '' ) -split ','
    }
    elseif ($MediaID_NotIn) {
        $oShowIdList = ($MediaID_NotIn -replace ' ', '' ) -split ','
    }
    else {
        $oShowIdList = $csvFileData | Where-Object {
            $dateObject = [DateTime]::ParseExact($_.AirDate, 'MM/dd/yyyy', $null)
            $dateObject -ge $today }
        | Select-Object -ExpandProperty ShowId -Unique
    }
    for ($i = 0; $i -lt $oShowIdList.Count; $i += $chunkSize) {
        $chunkedIdList += ($oShowIdList[$i..($i + $chunkSize - 1)] -join ', ')
    }
    while ($start -lt $maxDay) {
        $start++
        $startUnix, $endUnix = Get-FullUnixDay -Date $startDate
        Write-Log -Message "[Generate] $(Get-DateTime) - $($start)/$($maxDay) - $startDate - $startUnix - $endUnix" -LogFilePath $logFile
        foreach ($c in $chunkedIdList) {
            if ($MediaID_NotIn -or $newShowCheck) {
                $updatedRecordCalls += Invoke-AnilistApiDateRange -start $startUnix -end $endUnix -ID_NOTIN $c
            }
            else {
                $updatedRecordCalls += Invoke-AnilistApiDateRange -start $startUnix -end $endUnix -ID_IN $c
            }
        }
        $startDate = $startDate.AddDays(1)
    }
    foreach ($ur in $updatedRecordCalls | Sort-Object ShowId, EpisodeId) {
        $ShowId = $ur.ShowId
        $EpisodeId = $ur.EpisodeId
        $episode = $ur.episode
        $totalEpisodes = $ur.episodes
        if ( $episode -eq '' -or $null -eq $episode) { $episode = '00' }
        if ($totalEpisodes -eq '' -or $null -eq $totalEpisodes) { $totalEpisodes = '00' }
        $episode = '{0:D2}' -f [int]$episode
        $totalEpisodes = '{0:D2}' -f [int]$totalEpisodes
        $genres = $ur.genres -join ', '
        $seasonYear = $ur.seasonYear
        $season = if (-not [string]::IsNullOrEmpty($ur.season)) { ([System.Globalization.CultureInfo]::CurrentCulture).TextInfo.ToTitleCase($($ur.season).ToLower()) }
        else {
            ''
        }
        $status = if (-not [string]::IsNullOrEmpty($ur.status)) { ([System.Globalization.CultureInfo]::CurrentCulture).TextInfo.ToTitleCase($($ur.status).ToLower()) }
        else {
            ''
        }
        $showStartDate = if ($ur.startDate -ne '//') { $ur.startDate } else { $null }
        $showEndDate = if ($ur.endDate -ne '//') { $ur.endDate } else { $null }
        if (-not [string]::IsNullOrEmpty($ur.TitleEnglish)) {
            $title = $ur.TitleEnglish -replace $badchar, "'"
        }
        elseif (-not [string]::IsNullOrEmpty($ur.TitleRomanji)) {
            $title = $ur.TitleRomanji
        }
        else {
            $title = $ur.TitleNative
        }
        $airingFullTime = Get-UnixToLocal -unixTime $ur.airingAt
        $airingFullTimeFixed = Get-Date $airingFullTime -Format 'MM/dd/yyyy HH:mm:ss'
        $airingDate = Get-Date $airingFullTime -Format 'MM/dd/yyyy'
        $airingTime = Get-Date $airingFullTime -Format 'HH:mm'
        $weekday = Get-Date $airingFullTime -Format 'dddd'
        # Find the first preferred site
        $firstPreferredSite = $null
        foreach ($ps in $supportSites) {
            $firstPreferredSite = $ur.externalLinks | Where-Object { $_.site -eq $ps } | Select-Object * -First 1
            if ($firstPreferredSite) { break }
        }
        If ($firstPreferredSite) {
            $site = $firstPreferredSite.site
            $url = ($firstPreferredSite.url -replace 'http:', 'https:' -replace '^https:\/\/www\.crunchyroll\.com$', 'https://www.crunchyroll.com/' -replace '^https:\/\/www\.hidive\.com$', 'https://www.hidive.com/').Trim()
            if (($url -in $badURLs) -or ($site -eq 'HIDIVE' -and $totalEpisodes -eq '00') -or ($site -in 'Hulu', 'Netflix')) { $download = $false } else { $download = $True }
            Write-Log -Message "[Generate] $(Get-DateTime) - $site - $title - $episode/$totalEpisodes - $download - $url" -LogFilePath $logFile
            $ue = [PSCustomObject]@{
                ShowId         = $showId
                EpisodeId      = $episodeId
                Title          = $title
                Episode        = $episode
                TotalEpisodes  = $totalEpisodes
                SeasonYear     = $seasonYear
                Season         = $season
                Status         = $Status
                AirDay         = $weekday
                AirDate        = $airingDate
                AiringTime     = $airingTime
                AiringFullTime = $airingFullTimeFixed
                StartDate      = $showStartDate
                EndDate        = $showEndDate
                Site           = $site
                URL            = $url
                Genres         = $genres
                Download       = $download
            }
            $updatedRecords += $ue
        }
    }
    $newData = Update-Records -source $csvFileData -target $updatedRecords
    Write-Output "[Generate] $(Get-DateTime) - Updating $csvFilePath"
    $newData | Sort-Object Site, SeasonYear, { $SeasonOrder[$_.Season] }, Title, episode, TotalEpisodes | Select-Object * -Unique | Export-Csv $csvFilePath
}
If ($setDownload) {
    do {
        $sdType = Read-Host 'Set downloads by (S)eason or (D)ate?'
        switch ($sdType) {
            'S' { $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_S_$($pbyShort)-$($cbyShort).csv" }
            'D' { $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_D_$($pbyShort)-$($cbyShort).csv" }
            Default { Write-Output 'Enter S or D' }
        }
        if (!(Test-Path $csvFilePath)) {
            $sdType = 'NA'
            Write-Host "File does not exist: $csvFilePath"
        }
    } while ( $sdType -notin 'S', 'D' )
    Set-CSVFormatting $csvFilePath
    $pContent = Import-Csv $csvFilePath
    do {
        $f = Read-Host 'Update All(A), True(T), False(F)'
        switch ($f) {
            'a' { $sContent = $pContent | Select-Object Title, TotalEpisodes, Site, URL, Download -Unique }
            't' { $sContent = $pContent | Where-Object { $_.download -eq $True } | Select-Object Title, TotalEpisodes, Site, URL, Download -Unique }
            'f' { $sContent = $pContent | Where-Object { $_.download -eq $false } | Select-Object Title, TotalEpisodes, Site, URL, Download -Unique }
            Default { Write-Output 'Enter a valid response(a/t/f)' }
        }
    } while ( $f -notin 'a', 't', 'f' )
    [int]$counter = 1
    $counts = $sContent.count
    Write-Host "Total items: $counts"
    foreach ($c in $sContent) {
        do {
            if (($c.url -in $badURLs ) -or ($c.site -eq 'HIDIVE' -and $c.TotalEpisodes -eq '00')) {
                Write-Output "[Debug] $(Get-DateTime) - $counter/$counts - URL or site condition met: $($c.URL), $($c.site), $($c.TotalEpisodes)"
                Write-Output "$($spacer)`nTotal Episodes: $($c.TotalEpisodes) - URL: $($c.URL)"
                $r = 'no'
            }
            else {
                $r = Read-Host "$($spacer)`n$($counter)/$($counts) - Total Episodes: $($c.TotalEpisodes) - URL: $($c.URL)`nAdd $($c.title) [$($c.download)]?(Y/N/Blank)"
            }
            switch ($r) {
                { 'y', 'yes' -contains $_ } { $c.download = $True; $counter++ }
                { 'n', 'no' -contains $_ } { $c.download = $false; $counter++ }
                { '' -contains $_ } { [bool]$c.download = $c.download; $counter++ }
                Default { Write-Output 'Enter a valid response(y/n/blank)' }
            }
            Write-Output "[Setup] $(Get-DateTime) - Setting $($c.title) to [$($c.download)]$spacer"
        } while ( $r -notin 'y', 'yes', 'n', 'no' , '' )
    }
    foreach ($record in $pContent) {
        $matchingRecord = $sContent | Where-Object { $_.title -eq $record.title -and $_.URL -eq $record.URL }
        if ($matchingRecord) {
            foreach ($m in $matchingRecord) {
                $record.download = $m.download
            }
        }
    }
    Write-Output "[Setup] $(Get-DateTime) - Updating $csvFilePath"
    $pContent | Export-Csv -Path $csvFilePath -NoTypeInformation
}
If ($generateBatchFile) {
    do {
        if ($automated) {
            $daily = $True
        }
        else {
            $daily = Read-Host 'Generate Daily(D) or Season(S) or Both(B)'
        }
        switch ($daily) {
            'D' { $dateBatch = $True; $seasonBatch = $false }
            'S' { $dateBatch = $false; $seasonBatch = $True }
            'B' { $dateBatch = $True; $seasonBatch = $True }
            Default { Write-Host 'Enter D, S, or B' }
        }
    } while ( $daily -notin 'D', 'S', 'B' )
    $cutoffStart = $today.AddDays(-7)
    $cutoffEnd = $today.AddDays(7)
    Write-Host "D: $dateBatch; S: $seasonBatch"
    $date = Get-DateTime 2
    $newBackupPath = Join-Path $batchArchiveFolder -ChildPath "batch_$date"
    if (!(Test-Path $newBackupPath)) { New-Item $newBackupPath -ItemType Directory }
    $sourceDirectory = $rootPath
    $destinationDirectory = $newBackupPath
    if (!(Test-Path $destinationDirectory)) { New-Item -Path $destinationDirectory -ItemType Directory }
    $itemsToCopy = Get-ChildItem -Path $sourceDirectory | Where-Object { $_.Name -ne '_archive' }
    $itemsToCopy | ForEach-Object {
        $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $_.Name
        Copy-Item -Path $_.FullName -Destination $destinationPath -Recurse
    }
    if ($dateBatch -eq $True) {
        $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_D_$($pbyShort)-$($cbyShort).csv"
        Set-CSVFormatting $csvFilePath
        $rawData = (Import-Csv $csvFilePath) | Where-Object {
            $_.site -notin 'Hulu', 'Netflix' -and (
                [DateTime]::ParseExact($_.airingFullTime, 'MM/dd/yyyy HH:mm:ss', $null) -ge $cutoffStart -and
                [DateTime]::ParseExact($_.airingFullTime, 'MM/dd/yyyy HH:mm:ss', $null) -le $cutoffEnd
            )
        }
        $itemsToCopy | Where-Object { $_.Name -eq 'Site' } | ForEach-Object {
            Remove-Item $_.FullName -Recurse
        }
        $baseData = $rawData | Where-Object { $_.download -eq $true }
        $uniqueSites = $baseData | Select-Object -ExpandProperty Site -Unique
        foreach ($site in $uniqueSites) {
            $siteData = $baseData | Where-Object { $_.Site -eq $site }
            $basepath = Join-Path $rootPath -ChildPath 'site'
            $siteRootPath = Join-Path $basepath -ChildPath $site
            if (!(Test-Path $siteRootPath ) ) {
                Write-Log -Message "[Batch] $(Get-DateTime) - Creating folder for $siteRootPath" -LogFilePath $logFile
                New-Item $siteRootPath -ItemType Directory | Out-Null
            }
            else {
                Write-Log -Message "[Batch] $(Get-DateTime) - Creating folder for $siteRootPath" -LogFilePath $logFile
                Remove-Item $siteRootPath -Recurse -Force
                New-Item $siteRootPath -ItemType Directory | Out-Null
            }
            $uniqueWeekdays = $siteData | Select-Object -ExpandProperty AirDay -Unique
            foreach ($weekday in $uniqueWeekdays) {
                $weekdayData = $siteData | Where-Object { $_.AirDay -eq $weekday }
                $dayNum = $weekdayOrder[$weekday]
                $urls = @()
                foreach ($entry in $weekdayData) {
                    $showUrl = $entry.URL
                    $totalEpisodes = $entry.TotalEpisodes
                    if ($totalEpisodes -eq 'NA') { $totalEpisodes = 0 }
                    if ($site -eq 'HIDIVE' -and $totalEpisodes -gt 0) {
                        for ($i = 1; $i -le $totalEpisodes; $i++) {
                            $lastPart = $showUrl -replace '/season-\d+', '' -replace '.*/', ''
                            $v = "https://www.hidive.com/stream/$lastPart"
                            $seasonNumber = [regex]::Match($showUrl, '/season-(\d+)').Groups[1].Value
                            $seasonNumber = '{0:D2}' -f [int]$seasonNumber
                            $formattedNumber = '{0:D3}' -f $i
                            if ($seasonNumber -ne '00') {
                                $modifiedUrl = "$v/s$($seasonNumber)e$($formattedNumber)"
                            }
                            else {
                                $modifiedUrl = "$v/s01e$($formattedNumber)"
                            } 
                            $urls += $modifiedUrl
                        }
                    }
                    else {
                        $urls += $showUrl
                    }
                }
                $urls = $urls | Select-Object -Unique
                # Write to a text file named by the Site-Weekday combination
                $fileName = Join-Path $siteRootPath -ChildPath "$($site)-$($dayNum)-$($weekday)"
                Write-Log -Message "[Batch] $(Get-DateTime) - $($site) - $weekday - $($urls.count) - $fileName" -LogFilePath $logFile
                $urls | Out-File -FilePath $fileName
            }
        }
    }
    if ($seasonBatch -eq $True) {
        $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_S_$($pbyShort)-$($cbyShort).csv"
        Set-CSVFormatting $csvFilePath
        $baseSeasonPath = Join-Path $rootPath -ChildPath 'season'
        if (!(Test-Path $baseSeasonPath)) {
            New-Item $baseSeasonPath -ItemType Directory
        }
        else {
            Get-ChildItem $baseSeasonPath | Remove-Item -Recurse
        }
        $rawData = (Import-Csv $csvFilePath) | Where-Object { $_.site -notin 'Hulu', 'Netflix' }
        | Select-Object SeasonYear, Season, Title, TotalEpisodes, AirDay, Site, URL, Download -Unique | Sort-Object Site, { $weekdayOrder[$_.AirDay] }, AiringTime, Title
        $baseData = $rawData | Where-Object { $_.download -eq $true }
        $years = $baseData | Select-Object -ExpandProperty SeasonYear -Unique
        foreach ($y in $years) {
            $seasons = $baseData | Where-Object { $_.seasonYear -eq $y } | Select-Object -ExpandProperty season -Unique
            foreach ($s in $seasons) {
                $sdata = $baseData | Where-Object { $_.seasonYear -eq $y -and $_.season -eq $s }
                $sitesData = $sdata | Select-Object -ExpandProperty site -Unique
                foreach ($sd in $sitesData) {
                    $a = $sdata | Where-Object { $_.site -eq $sd }
                    $uniqueWeekdays = $a | Select-Object -ExpandProperty AirDay -Unique
                    foreach ($weekday in $uniqueWeekdays) {
                        $batch = $a | Where-Object { $_.AirDay -eq $weekday }
                        $dayNum = $weekdayOrder[$weekday]
                        $urls = @()
                        foreach ($b in $batch) {
                            $site = $sd
                            $p = Join-Path $baseSeasonPath -ChildPath "$y\$s\$sd"
                            if (!(Test-Path $p)) { New-Item $p -ItemType Directory }
                            $url = $b.URL
                            $totalEpisodes = $b.TotalEpisodes
                            if ($totalEpisodes -eq 'NA') { $totalEpisodes = 0 }
                            if ($site -eq 'HIDIVE' -and $totalEpisodes -gt 0) {
                                for ($i = 1; $i -le $totalEpisodes; $i++) {
                                    $lastPart = $url -replace '/season-\d+', '' -replace '.*/', ''
                                    $v = "https://www.hidive.com/stream/$lastPart"
                                    $seasonNumber = [regex]::Match($url, '/season-(\d+)').Groups[1].Value
                                    $seasonNumber = '{0:D2}' -f [int]$seasonNumber
                                    $formattedNumber = '{0:D3}' -f $i
                                    if ($seasonNumber -ne '00') {
                                        $modifiedUrl = "$v/s$($seasonNumber)e$($formattedNumber)"
                                    }
                                    else {
                                        $modifiedUrl = "$v/s01e$($formattedNumber)"
                                    }
                                    $urls += $modifiedUrl
                                }
                            }
                            else {
                                $urls += $url
                            }
                            $urls = $urls | Select-Object -Unique
                            $fileName = Join-Path $p -ChildPath "$($site)-$($y)-$($s)-$($dayNum)-$($weekday)"
                            Write-Log -Message "[Batch] $(Get-DateTime) - $($site) - $weekday - $($urls.count) - $fileName" -LogFilePath $logFile
                            $urls | Out-File -FilePath $fileName
                        }
                    }
                }
            }
        }
    }
}
If ($sendDiscord) {
    $emojiList = @()
    $emojis | ForEach-Object {
        $de = [PSCustomObject]@{
            siteName  = $_.siteName
            siteEmoji = $_.emoji
        }
        $emojiList += $de
    }
    Write-Log -Message "[Discord] $(Get-DateTime) - Current lastUpdated: $lastUpdated " -LogFilePath $logFile
    $newLastUpdated = Get-Date -Format 'MM-dd-yyyy HH:mm:ss'
    $csvFilePath = Join-Path $csvFolder -ChildPath "anilistSeason_D_$($pbyShort)-$($cbyShort).csv"
    Set-CSVFormatting $csvFilePath
    if (Test-Path $csvFilePath ) {
        $origin = (Import-Csv $csvFilePath) | Where-Object { $_.Download -eq $True } | Group-Object -Property Site
        $siteNames = $origin.Name
        $siteList = @()
        foreach ($s in $siteNames) {
            $b = [PSCustomObject]@{
                Site      = $s
                ShowCount = 0
            }
            $siteList += $b
        }
        Write-Log -Message "[Discord] $(Get-DateTime) - Reading Anilist Date CSV $csvFilePath" -LogFilePath $logFile
        $fieldObjects = @()
        $showCount = 0
        $sdWeekday = Get-Date -Format 'dddd dd-MMM'
        $origin | ForEach-Object {
            $fieldValueRelease = @()
            $sdSite = $($_.Name)
            $sCount = 0
            $_.Group | Sort-Object AiringTime, Title | ForEach-Object {
                $airdate = Get-Date $_.AiringFullTime -Format 'MM/dd/yyyy' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
                $nowDate = Get-Date -Format 'MM/dd/yyyy' -Hour 0 -Minute 0 -Second 0 -Millisecond 0
                if ($airdate -eq $nowDate) {
                    $airingtime = Get-Date $($_.AiringTime) -Format 'HH:mm'
                    if ([int]::TryParse($_.Episode, [ref]$null)) {
                        $episode = 'Episode ' + ('{0:D2}' -f [int]$_.Episode)
                    }
                    else {
                        $episode = 'Episode??'
                    }
                    if ([int]::TryParse($_.TotalEpisodes, [ref]$null)) {
                        $totalEpisodes = '{0:D2}' -f [int]$_.TotalEpisodes
                    }
                    else {
                        $totalEpisodes = '??'
                    }
                    $v = '```' + ("$($airingtime)`n$($_.Title)`n$episode/$totalEpisodes") + '```'
                    Write-Log -Message "[Discord] $(Get-DateTime) - $($airingtime) - $($_.Title) - $episode/$totalEpisodes" -LogFilePath $logFile
                    $fieldValueRelease += $v
                    $showCount++
                    $sCount++
                }
                $SiteCountList | Where-Object { $_.Site -eq $sdSite } | ForEach-Object { $_.ShowCount = $sCount }
            }
            switch ($sdSite) {
                { 'Hidive', 'Crunchyroll', 'Hulu', 'Netflix' -contains $_ } { $emoji = $emojiList | Where-Object { $_.siteName -eq $sdSite } | Select-Object -ExpandProperty siteEmoji }
                Default { $emoji = $emojiList | Where-Object { $_.siteName -eq $sdSite } | Select-Object -ExpandProperty siteEmoji }
            }
            $sdSite = $emoji + ' **' + $sdSite + '** ' + $emoji
            if ($fieldValueRelease.count -gt 0) {
                $fieldValueRelease = $fieldValueRelease
                $fObj = [PSCustomObject]@{
                    name   = $sdSite
                    value  = $fieldValueRelease
                    inline = $false
                }
                $fieldObjects += $fObj
            }
        }
        if ($showCount -eq 0) {
            $fObj = [PSCustomObject]@{
                name   = 'No monitored shows airing today'
                value  = ":'("
                inline = $false
            }
            $fieldObjects += $fObj
        }
        $description = "**No. of Shows:** $($showCount)"
        $thumbnailObject = [PSCustomObject]@{
            url = $discordSiteIcon
        }
        $siteFooterText = @()
        foreach ($f in $SiteCountList) {
            $siteFooterText += "$($f.Site) - $($f.Showcount)"
        }
        $siteFooterText = $siteFooterText -join ' | '
        $footerObject = [PSCustomObject]@{
            icon_url = $siteFooterIcon
            text     = "$siteFooterText"
        }
        [System.Collections.ArrayList]$embedArray = @()
        $embedObject = [PSCustomObject]@{
            color       = 15879747
            title       = "Airing Today: $sdWeekday"
            description = $description
            thumbnail   = $thumbnailObject
            fields      = $fieldObjects
            footer      = $footerObject
            timestamp   = $(Get-DateTime 3)
        }
        $payload = [PSCustomObject]@{
            embeds = $embedArray
        }
        $embedArray.Add($embedObject) | Out-Null
        $payloadJson = $payload | ConvertTo-Json -Depth 4
        if ($fieldObjects.count -gt 0) {
            Write-Log -Message "[Discord] $(Get-DateTime) - Sending Discord Message for $sdWeekday" -LogFilePath $logFile
            Invoke-WebRequest -Uri $discordHookUrl -Body $payloadJson -Method Post -ContentType 'application/json' | Select-Object -ExpandProperty Headers | Out-Null
        }
        else {
            Write-Log -Message "[Discord] $(Get-DateTime) - No shows airing for $sdWeekday" -LogFilePath $logFile
        }
    }
    else {
        Write-Log -Message "[Discord] $(Get-DateTime) - Anilist Season CSV not exist at: $csvFilePath" -LogFilePath $logFile
        Exit-Script
    }
    $config.configuration.logs.lastUpdated = $newLastUpdated
    Write-Log -Message "[Discord] $(Get-DateTime) - New lastUpdated: $newLastUpdated" -LogFilePath $logFile
    $config.Save($configFilePath)
    Start-Sleep -Seconds 2
}
Exit-Script