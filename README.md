# Powershell and Python scripts for yt-dlp downloads and file handling with Subtitle Edit, Filebot, and Plex
Script used to fetch files with yt-dlp, fix subtitles, embed into video file, rename and move video files into media server folders.
# Prerequisites

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

Mkvtoolnix
- Embeds subtitle and fonts
- https://mkvtoolnix.download/

Plex
- Media Server
- https://www.plex.tv/
## Steps to run:
1. Install or download prereqs and map to PATH as needed.
2. Run: `path\to\dlp-script.ps1 -NC` to generate base XML
   - Creates base XML config in root directory
3. Fill out template in your path\to\the\config.xml
   - Paths to temp and destination folder for yt-dlp
   - Path the ffmpeg location
   - Url, Token, LibraryIds to to plex, 
   - Site credentials and corresponding LibraryId for final destination
4. Run: `path\to\dlp-script.ps1 -SU` to generate supporting files
5. Setup up configs, batch, and cookie files as needed
   - You'll end up with a set of manual and daily files per site
6. Run: `path\to\dlp-script.ps1 {ARGS}` with applicable arguments
   - Ex. `D:\Folder\dlp-script.ps1 -SN youtube -D -SE -A` runs yt-dlp for youtube with the daily(_D suffix) files using the archive file along with running SubtitleEdit afterwards.
   - As your script is running it will generate a log in the related site folder
# Parameters explained:
| Arguments/Switches | Abbreviation | Description|Notes|
 :--- | :--- | :--- | :--- |
|-site|-sn/-SN|Tells the script what site its working with|Hardcoded acceptable values.| Reads from root\config.xml file for list of applicable values|
|-isDaily|-d/-D|Will use different yt-dlp configs and files and temp/home folder structure.| If -D = true then it will use the \_D suffix named files.|
|-useArchive|-a/-A|Will tell yt-dlp command to use or not use archive file.| If -A = true then it will use the \_A suffix named files.|
|-useLogin|-l/-L|Tells yt-dlp command to use credentials stored in config xml.| If -L = false then it will use the \_C suffix named files. Will check ReqCookies file if site matches in text then will throw error.|
|-useFilebot|-f/-F|Tells script to run Filebot. Will take Plex folder name defined in config xml.| |
|-useSubtitleEdit|-se/-SE|Tells script to run SubetitleEdit to fix common problems with .srt files if they are present.| Expects presence of mkv and ass file.|
|-useDebug|-b/-B| Shows minor additional info.| |
|-help|-h/-H|Shows this text.| If help is true or all parameters false/null then displays readme. |
|-newconfig|-nc/-NC|Used to generate empty config if none is present.| |
|-createsupportfiles|-su/-SU|Creates support files like archives, batch and cookies files for sites in the `config.xml`.| |