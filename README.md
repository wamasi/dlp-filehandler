# Powershell and Python scripts for yt-dlp downloads and file handling with Subtitle Edit, Filebot, and Plex

Script used to fetch files with yt-dlp, fix subtitles, embed into video file, rename and move video files into media server folders.

## Getting Started

Fill out config

Install or download prereqs and map to PATH as needed.

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
