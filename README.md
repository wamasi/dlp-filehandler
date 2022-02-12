# Powershell and Python scripts for yt-dlp downloads and file handling with Subtitle Edit, Filebot, and Plex

Script used to fetch files with yt-dlp, fix subtitles, embed into video file, rename and move video files into media server folders.

## Getting Started

Fill out config with:
- Paths to temp and destination folder for yt-dlp
- Path the ffmpeg location
- Url, Token, LibraryIds to to plex, 
- Site credentials and corresponding LibraryId for final destination

Install or download prereqs and map to PATH as needed.

### Expected folder structure:
- For yt-dlp temp: {drive}\tmp\
- For yt-dlp home: {drive}\tmp\
- For Filebot/SubtitleEdit staging: {drive}\tmp\
- Final destination: {drive}\videos\{PlexLibraryFolder}\

## Prerequisites

YT-dlp
- Downloading videos
- https://github.com/yt-dlp/yt-dlp

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
