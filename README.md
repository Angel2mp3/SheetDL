# SheetDL

<p align="center">
  <strong>A powerful music downloader that syncs with Google Sheets</strong>
</p>

<p align="center">
  A Python-based desktop tool for Windows that lets you download entire music collections<br>
  tracked inside a Google Sheet. Supports multiple download sources and protocols.
</p>

<p align="center">
  <strong>Current Version: 2.1.0</strong>
</p>

---

## ✨ Features

- **Auto-Update** - Checks for updates on startup silently; updates with one click when available
- **Download Queue** - Add multiple sheets to queue while downloading; they'll process automatically
- **Pause & Resume** - Pause downloads and resume where you left off
- **Google Sheets Integration** - Connect to any public Google Sheet tracker containing the music files you want to download
- **Automatic Dependency Installation** - Missing Python packages are detected and can be installed automatically
- **Multi-Tab Support** - Select and download from different sheet tabs
- **Hyperlink Extraction** - Automatically extracts URLs from HYPERLINK() formulas in Google Sheets
- **Smart Organization** - Organize downloads by Artist, Album, or keep flat
- **Album Art Download** - Automatically downloads cover art when available
- **Detailed Metadata Export** - Save comprehensive track info to text files (see below)
- **Rip Format Selection** - When ripping from certain services choose between audio formats (M4A, MP3) or video (MP4)
- **Dark Modern UI** - Custom rounded window design with animated elements

### 📄 Metadata Export

SheetDL can generate a detailed `.txt` file for each track containing all available information from your sheet. This feature is **enabled by default** but can be toggled on/off in the settings.

Example output:
```
Title: Song Name
Artist: Artist Name
Album/Project: Album Name
Genre/Category: Genre
Notes: Any notes from the sheet
File Date: Date information
Surface/Release Date: Dates for those
Type: Track type
Format: Audio quality
Cover Source: Cover art URL
Cover Saved: Yes/No
Download Links:
  - https://example.com/download-link
Generated: Date & Time
```

### 📥 Supported Download Sources

#### 🎵 Audio & Video Providers

| Source | Status | Notes |
|--------|--------|-------|
| **YouTube** | ✅ | Audio & Video formats |
| **SoundCloud** | ✅ | M4A & MP3 formats |
| **Google Drive** | ✅ | Public files |
| **MEGA.nz** | ✅ | Encrypted downloads |
| **KrakenFiles** | ✅ | Attempts original, falls back to M4A if CloudFlare blocks |
| **Pixeldrain** | ✅ | Direct downloads |
| **FileDitch** | ✅ | All file types |
| **Pillows.su** | ✅ | Including legacy's: plwcse.top, pillowcase.zip, pillowcase.su |
| **Froste.lol** | ✅ | Audio files |
| **BumpWorthy** | ✅ | Video/Audio bumps |
| **imgur.gg** | ✅ | Audio/Video files |
| **Gofile.io** | ✅ | File hosting |
| **MediaFire** | ✅ | File hosting |
| **Amazon Web Services (S3)** | ✅ | Public files |
| **Archive.org** | ✅ | Public archives |
| **Streamable** | ✅ | Video hosting |
| **Bandcamp** | ✅ | 128kbs MP3's |
| **Catbox.moe** | ✅ | Audio/Images |
| **Direct URLs** | ✅ | Any direct file link |

#### 🖼️ Image & Cover Providers

| Source | Status | Notes |
|--------|--------|-------|
| **Imgur.com** | ✅ | Images |
| **imgbb.com / ibb.co** | ✅ | Image hosting |
| **Dump.li** | ✅ | Image hosting |

---

## 🚀 Getting Started

### Prerequisites

- **Python 3.10+** (tested on Python 3.13)
- **yt-dlp** (auto-installed on first run, or install manually via `pip install yt-dlp`)
- **FFmpeg** (optional, for audio conversion)

### Installation

1. **Clone or download this repository**
   ```bash
   git clone https://github.com/Angel2mp3/SheetDL
   cd SheetDL
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```
   
   Or let SheetDL install them automatically on first run!

3. **Run SheetDL**
   ```bash
   python SheetDL.py
   ```

### Dependencies

All dependencies are automatically checked and can be installed on first run!

```
gspread>=5.0.0
google-auth>=2.0.0
requests>=2.25.0
yt-dlp>=2023.0.0
beautifulsoup4>=4.9.0
pycryptodome>=3.10.0
cloudscraper>=1.2.0
```

**Optional:**
- `Pillow` - For icon generation
- `FFmpeg` - For audio format conversion

---

## 📖 Usage

### 1. Prepare Your Google Sheet

Your sheet should have columns for:
- **Title/Name** - Song name
- **Artist** - Artist name
- **URL/Link** - Download link(s)
- **Album** (optional) - For organization
- **Cover** (optional) - Album art URL

The app auto-detects common column names, or you can manually map them.

### 2. Get the Sheet URL

1. Open your Google Sheet
2. Make sure it's **publicly viewable** (Share → Anyone with the link)
3. Copy the URL from your browser

### 3. Configure SheetDL

1. Paste the sheet URL
2. Select output folder
3. Choose organization method (Artist/Album/Flat)
4. Select format preferences
5. Click **Start Download**

---

## 🎨 Interface

SheetDL features a modern dark-themed UI with:
- Custom rounded window design with title bar
- Real-time download progress log with auto-scrolling
- Download queue management
- Format selection for YouTube and SoundCloud
- Pause/Resume controls
- Sheet tab selector dropdown

---

## ⚙️ Configuration

Settings are automatically saved to `config.json`:

- Sheet URL and GID
- Output folder path
- Organization preferences
- Column mappings
- Format preferences

---

## 🔧 Troubleshooting

### Updates
SheetDL automatically checks for updates on startup. If a new version is available, you'll be prompted to update. You can also manually check by clicking the **"Check for Updates"** button in the app.

### Sheet not found / Connection fails
If the program can't find your Google Sheet even though the link looks correct, check if your URL ends with `gid=` followed by numbers (e.g., `gid=0` or `gid=123456789`). 

**If your URL doesn't have a GID:**
1. Open your Google Sheet in a browser
2. Navigate to the specific tab you want to download
3. Copy the full URL from your browser - it should now include `gid=...`
4. Paste that complete URL into SheetDL

### "Missing Dependencies" on startup
Click "Yes" to auto-install, or manually run:
```bash
pip install gspread google-auth requests yt-dlp beautifulsoup4 pycryptodome
```

### MEGA downloads not working
Ensure `pycryptodome` is installed:
```bash
pip install pycryptodome
```

### YouTube downloads failing
Update yt-dlp to the latest version:
```bash
pip install -U yt-dlp
```

### Audio conversion not working
Install FFmpeg and ensure it's in your system PATH.

---

## 🐛 Known Issues

### Metadata Embedding Limitations
**Album covers and metadata may not embed properly** from certain sources like YouTube and SoundCloud. This is a limitation of the underlying tools (yt-dlp, FFmpeg) and how these platforms provide metadata.

**Workarounds:**
- Album art files should still be downloaded separately and saved alongside tracks
- Metadata text files contain all information even if embedding fails
- Consider using dedicated audio tagging software (like Mp3tag) for batch metadata editing

### Sheet Tab Selector - Limited Functionality
The in-app sheet tab dropdown may not always switch between tabs reliably.

**Recommended Workaround:** Navigate to the desired tab/section in Google Sheets and copy the URL directly from your browser while on that tab. The URL will contain the correct GID parameter for that specific tab.

- **UI scrolling**: May not be perfectly smooth in some cases
- **KrakenFiles CloudFlare**: May be blocked by CloudFlare protection in some cases

---

## 📋 Requirements File

Create or use the provided `requirements.txt`:
```
gspread>=5.0.0
google-auth>=2.0.0
requests>=2.25.0
yt-dlp>=2023.0.0
beautifulsoup4>=4.9.0
pycryptodome>=3.10.0
cloudscraper>=1.2.0
```

---

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest new features
- Add support for new download sources
- Improve documentation

---

## 📜 License

This open source project is provided as-is for personal use only under the fair use clause of the Digital Millennium Copyright Act. Please respect all copyright laws and only download content you have the legal right to access.

---

## ⚠️ Disclaimer

SheetDL is an open‑source utility for downloading publicly available content that you are legally entitled to access. It does NOT bypass any DRM, copyright protections, paywalls, or any other encryption mechanisms; all downloadable files originate from freely open, public sources. Users are solely responsible for confirming they have the legal right to download any material, and the developers assume no liability for any unlawful use of the software.

---

### Made with ❤️ for music collectors
