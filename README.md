# Powershell and Python scripts for yt-dlp downloads and file handling with Subtitle Edit, Filebot, and Plex
Script used to fetch files with yt-dlp, fix subtitles, embed into video file, rename and move video files into media server folders.
## Prerequisites

YT-dlp
- Downloading videos
- https://github.com/yt-dlp/yt-dlp
  - Read this documenation first to understand how to setup your supporting files

Ffmpeg
- Video downloader used with yt-dlp, remux, copy and move files with yt-dlp
- https://github.com/BtbN/FFmpeg-Builds/releases

Aria2c
- Video downloader used with yt-dlp
- https://github.com/aria2/aria2

Filebot
- File renamer based on tvdb entries
- https://www.filebot.net/

Subtitle Edit
- Fixes common subtitle issues like text timings
- https://www.nikse.dk/SubtitleEdit/

Python
- Used for script to regex through subtitle file to edit fonts use
- https://www.python.org/downloads/
- Modules used:
  - sys
  - fileinput
  - time
  - regex

Mkvtoolnix
- Embeds subtitle and fonts
- https://mkvtoolnix.download/

Plex
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
## Steps to run:
1. Install or download prereqs and map to PATH as needed.
2. Run: `path\to\dlp-script.ps1 -NC` to generate base XML
   - Creates base XML config in root directory
3. Fill out template in your path\to\the\config.xml
   - Paths to temp and destination folder for yt-dlp
   - Path the ffmpeg location
   - Url, Token, LibraryIds to to plex
   - Site credentials and corresponding LibraryId for final destination
4. Run: `path\to\dlp-script.ps1 -SU` to generate supporting files
   - Generates
      - `fonts` folder
        - Place to store custom fonts used in `SubtitleEdit`, `subtitle_regex.py`, and `mkvmerge` to display custom font in subtitle file.
      - `shared` folder
        - Empty archive, bat, and cookie files.
      - `sites` folder
        - Base site configs
        - Log files
5. Setup up configs, batch, cookie, and font files as needed
   - You'll end up with a set of manual and daily files per site
6. Run: `path\to\dlp-script.ps1 {ARGS}` with applicable arguments
   - Ex. `D:\Folder\dlp-script.ps1 -SN youtube -D -SE -A` runs yt-dlp for youtube with the daily(_D suffix) files using the associated cookie and archive file along with running SubtitleEdit afterwards.
   - Script execution steps:
      - `dlp-script.ps1` gathers initial variables based on parameters passed into script.
         - Log file is generated and located in `site` folder based on date/datetime
         - Creates base folder structure for `tmp`, `src`, `dest`.
      - `dlp-script.ps1` runs `YT-DLP`, `SubtitleEdit`, `subtitle_regex.py`, `mkvmerge` command based on parameters
         - Each run puts files in folder based on timestamp of execution.
         - Temp files store in `tmp` location set from `config.xml`.
         - Final files from `YT-DLP` will be moved to `src` location set from `config.xml`.
         - If `SubtitleEdit = True`
            - Will run to fix common issues in subtitle files using SubtitleEdit.
            - If font value is set and file exists:
               - `subtitle_regex.py` will regex through file to update font name.
               - mkvmerge will embed custom font.
            - If file runs into an issue this will be outputted in the log and all files from this run will not be moved to `dest`.
         - Moves final file to `dest` location set from `config.xml`.
         - If `FileBot = True` will run FileBot against files in `dest` location.
            - Will then move all files with updated name to `Plex Library Folder` set from `config.xml`.
            - If FileBot runs into an issue this will be outputted in the log and all files from this run will not be moved to `Plex Library Folder`.
            - if `Plex Token` and `Plex libraryId` supplied will run API call to update folder
         - If `UseDownloadArchive = True` then copy `Archive` and `Bat` for that site to `src` directory.
         - Clean up of `tmp`, `src`, and `dest` folders.
            - `tmp` will always be deleted at the end of a run
# Parameters explained:
| Arguments/Switches | Abbreviation | Description | Notes |
 :--- | :--- | :--- | :--- |
|-site|-sn/-SN|Tells the script what site its working with|Hardcoded acceptable values. Reads from root\config.xml file for list of applicable values|
|-isDaily|-d/-D|Will use different yt-dlp configs and files and temp/home folder structure.| If -D = true then it will use the \_D suffix named files.|
|-useArchive|-a/-A|Will tell yt-dlp command to use or not use archive file.| If -A = true then it will use the \_A suffix named files.|
|-useLogin|-l/-L|Tells yt-dlp command to use credentials stored in config xml.| If -L = false then it will use the \_C suffix named files. Will check ReqCookies file if site matches in text then will throw error.|
|-useFilebot|-f/-F|Tells script to run Filebot. Will take Plex folder name defined in config xml.| Outputs file with \{ plex.tail \}|
|-useMKVMerge|-mk/-MK|Tells script to run `subtitle_regex.py` and MKVMerge against file| Expects presence of mkv and ass file.|
|-useSubtitleEdit|-se/-SE|Tells script to run SubetitleEdit to fix common problems with .srt files if they are present.| Expects presence of mkv and ass file.|
|-newconfig|-nc/-NC|Used to generate empty config if none is present.| |
|-createsupportfiles|-su/-SU|Creates support files like archives, batch, some basic configs and cookies files for sites in the `config.xml`.| |
|-help|-h/-H|Shows MD file.| If help is true or all parameters false/null then displays readme. |
