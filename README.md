# Powershell script for yt-dlp downloads and file handling with Subtitle Edit, Mkvtoolnix, Filebot, and Plex
Script used to fetch files with yt-dlp, fix subtitles, embed into video file, rename and move video files into media server folders.
## Prerequisites

Windows:
- All file formatting is based on Windows

Powershell 7+:
- Scripts probably work with earlier versions, but haven't tested them

yt-dlp:
- Downloading videos
- https://github.com/yt-dlp/yt-dlp
  - Read this documenation first to understand how to setup your supporting files

FFMPEG:
- Video downloader used with yt-dlp, remux, copy and move files with yt-dlp
- https://github.com/BtbN/FFmpeg-Builds/releases

Aria2c:
- Video downloader used with yt-dlp
- https://github.com/aria2/aria2

Mkvtoolnix:
- Embeds subtitle and fonts
- https://mkvtoolnix.download/

Filebot:
- File renamer based on tvdb entries
- https://www.filebot.net/

Subtitle Edit:
- Fixes common subtitle issues like text timings
- https://www.nikse.dk/SubtitleEdit/

Plex (optional)
- Media Server
- https://www.plex.tv/
- Getting Plex token:
  - https://support.plex.tv/articles/204059436-finding-an-authentication-token-x-plex-token/
     - Navigate to a item in your library
     - Right click "Get Info"
     - Select "View XML"
     - End of url will have `&X-Plex-Token=<YourPLEXTokenHere>`
     - Getting Plex Library Id:
       - Filling and paste into browser:
         - http://`<PlexIP>`:`<PlexPORT>`/library/sections?X-Plex-Token=`<PlexToken>`
       - Library Id = `key`
          - ex: `key = "1"`

Discord (optional):
- Used to send out a message of videos downloaded to discord chat via webhook

## Script Setup/Execution Walkthrough:
1. Install or download prereqs and map to PATH or configure as needed.
2. Run: `path\to\dlp-script.ps1 -NC` to generate base XML
   - Creates base XML config in root directory
3. Fill out template in your path\to\the\config.xml
   - Paths to temp and destination folder for yt-dlp
   - Path the ffmpeg location
   - Url, Token, LibraryIds to to plex
   - Site credentials and corresponding LibraryId for final destination
   - Optional:
     - Fill out Discord section with token and chatid to send notification of new video with `-sd` parameter.
4. Run: `path\to\dlp-script.ps1 -SU` to generate supporting files
   - Generates
      - `fonts` folder
        - Place to store custom fonts used in `SubtitleEdit` and `mkvmerge` to display custom font in subtitle file.
      - `shared` folder
        - Empty archive, bat, and cookie files.
      - `sites` folder
        - Base site configs
        - Log files
5. Setup up configs, batch, cookie, and font files as needed
   - You'll end up with a set of manual and daily files per site
   - Pre-defined `yt-dlp.config` for `VRV`, `Crunchyroll`, `Funimation`, `ParamountPlus`, `Youtube`, `Twitch` available.
     - Other sites will have a generic `yt-dlp.config` created.
6. Run: `path\to\dlp-script.ps1 {ARGS}` with applicable arguments
   - Ex. `D:\Folder\dlp-script.ps1 -SN youtube -D -SE -A` runs yt-dlp for youtube with the daily(_D suffix) files using the associated cookie and archive file along with running SubtitleEdit afterwards.
   - Script execution steps:
      - `dlp-script.ps1` gathers initial variables based on parameters passed into script.
         - Log file is generated and located in `site` folder based on date/datetime
         - Creates base folder structure for `tmp`, `src`, `dest`.
         - Runs `YT-DLP`, `SubtitleEdit`, `mkvmerge`, `Discord` command based on parameters
         - Each run puts files in folder based on timestamp of execution.
         - Temp files store in `tmp` location defined in  `config.xml`.
           - Then moved into `src` location defined in `config.xml`
           - Will grab associated video and subtitle file in `src` and use this list to run files against later steps.
         - If `SubtitleEdit = True`
            - Will run to fix common issues in subtitle files using SubtitleEdit.
            - If font value is set and file exists:
               - `Update-SubtitleFont` will regex through file to update font name and remove .
               - mkvmerge will embed custom font.
            - If file runs into an issue this will be outputted in the log and all files from this run will not be moved to `dest`.
         - Final postprocessed files will be moved from `src` to `dest` location defined in `config.xml`.
         - If `FileBot = True` will run FileBot against files in `dest` location.
            - Will then move all files with updated name to `Plex Library Folder` set from `config.xml`.
            - If FileBot runs into an issue this will be outputted in the log and all files from this run will not be moved to `Plex Library Folder`.
            - if `Plex Token` and `Plex libraryId` supplied will run API call to update folder
         - If font, archive or cookie file is define then copy `Font`, `Archive` and `Cookie` files for that site to `src` directory.
         - `Bat` file for that site and `config.xml` will always be copied to `src` directory.
         - Clean up of `tmp`, `src`, and `dest` folders.
            - `tmp` will always be deleted at the end of a run
            - `src` and `dest` will only be deleted if empty

# dlp-script Parameters explained:
| Arguments/Switches | Abbreviation | Description | Notes |
 :--- | :--- | :--- | :--- |
|-Help|-h/-H|Shows MD file.| If help is true or all parameters false/null then displays readme. |
|-NewConfig|-nc/-NC|Used to generate empty config if none is present.| |
|-SupportFiles|-su/-SU|Creates support files like archives, batch, some basic configs and cookies files for sites in the `config.xml`.| |
|-TestScript|-t/-T| Runs setup based on commands and values `config.xml` to generate a log file in the site to display proposed values.| If `true`, `dlp-exec.ps1` is not ran and only log of vars is produced. |
|-Site|-sn/-SN|Tells the script what site its working with|Hardcoded acceptable values. Reads from root\config.xml file for list of applicable values|
|-Daily|-d/-D|Will use different yt-dlp configs and files and temp/home folder structure.| If -D = true then it will use the \_D suffix named files.|
|-Archive|-a/-A|Will tell yt-dlp command to use or not use archive file.| If -A = true then it will use the \_A suffix named files.|
|-Login|-l/-L|Tells yt-dlp command to use credentials stored in config xml.| uses stored credentials for site in `config.xml`.|
|-Cookies|-c/-C|Uses cookie file in yt-dlp param even if `-Login = True`.| Optional switch will use the \_C suffix named files with browser cookies.|
|-Filebot|-f/-F|Tells script to run Filebot. Will take Plex folder name defined in config xml.| Outputs file with \{n}\{'Season '+s00}\{n\} - \{s00e00\} - \{t\}|
|-MKVMerge|-mk/-MK|Tells script to run `MKVMerge` against video and subtitle files to edit and embed subtitle if available| Expects presence of `.mkv` and `.ass` file.|
|-SubtitleEdit|-se/-SE|Tells script to run `SubetitleEdit` to fix common problems with `.srt` files if they are present.| Expects presence of `.ass` subtitle file.|
|-SendDiscord|-sd/-SD|If Discord `webhook` and `icon` and `color` filled out in `config.xml` will send out `embed` message to discord channel of `webhook` of new videos.| Outputs video information<br/><br/>Site Name and icon = `config.xml`<br/>Series = original `Directory` name<br/>Episode = original `Basename` of file<br/>Subtitles = takes lang tag(s) in original `Subtitle` filename <br/>Video/Audio codec(s), duration, quality = `ffprobe.exe` cli|
|-AudioLang|-al/-AL|Sets audio and video track to a given language code ('ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und')|ex: -AL ja |
|-SubtitleLang|-sl/-SL|Sets default subtitle track to a given language code for track that matches ('ar', 'de', 'en', 'es', 'es-es', 'fr', 'it', 'ja', 'pt-br', 'pt-pt', 'ru', 'und')| Ex: -SL en |
|-ArchiveTemp|-AT| Will use a separate hardcoded archive file instead of the normal archive | If both Archive and ArchiveTemp are specified AT will tak precedence |
|-OverrideBatch|-OD| Will use a separate batch file specified in the command instead | Ex: -OD `path\to\my\other\batch\file` |

# AnilistData Parameters explained:
| Arguments/Switches | Abbreviation | Description | Notes |
 :--- | :--- | :--- | :--- |
|Automated|-A| Allows automated run of script based on hardcoded values||
|GenerateAnilistFile|-G|Generates either a csv file of anime based on Season or Dates||
|updateAnilistCSV|-U|Updates the generated csv for any changed data since created||
|updateAnilistURLs|-UU|Updates only the URLs in the CSV if they have changed||
|SetDownload|-D|Allows users to specify whether to track a given show||
|GenerateBatchFile|-B|Creates batch files used for separate script based on days of the week||
|SendDiscord|-SD|Sends list of shows airing that day that are marked as watching||
|MediaID_In|-I|Used to override the update parameters to specified anilist Id(s) when updating records||
|MediaID_NotIn|-NI|Used to override the update parameters to specified anilist Id(s) not in the CSV when updating records||
|newShowCheck|-NSC|Similar to -NI except no need to specify ID(s)||
|DailyRuns|-DR|Run dlp-script based on batch files created using anilistData script||
|Backup|-BU|Backups files generated from anilistData script to a specified location||