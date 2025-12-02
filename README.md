# SheetDL

<p align="center">
  <strong>A powerful music downloader that syncs with Google Sheets</strong>
</p>

<p align="center">
  A Python-based desktop tool for Windows that lets you download entire music collections<br>
  tracked inside a Google Sheet. Supports multiple download sources and protocols.
</p>

<p align="center">
  Made by <a href="https://github.com/Angel2mp3">Angel2mp3</a>
</p>

---

## ✨ Features

- **Google Sheets Integration** - Connect to any public Google Sheet tracker containing the music files you want to download
- **Multi-Tab Support** - Select and download from different sheet tabs
- **Smart Organization** - Organize downloads by Artist, Album, or keep flat
- **Album Art Download** - Automatically downloads cover art when available
- **Metadata Files** - Optionally save track metadata as text files
- **Rip Format Selection** - When ripping choose between audio formats (M4A, MP3) or video (MP4) 

### 📥 Supported Download Sources

| Source | Status | Notes |
|--------|--------|-------|
| **YouTube** | ✅ | Audio & Video formats |
| **SoundCloud** | ✅ | Including short links |
| **Google Drive** | ✅ | Public files |
| **MEGA.nz** | ✅ | Encrypted downloads |
| **Pillows.su** | ✅ | Including legacy plwcse.top links |
| **KrakenFiles** | ➖ | M4A files (No OG Files Due To CloudFlare) |
| **Pixeldrain** | ✅ | Direct downloads |
| **FileDitch** | ✅ | All file types |
| **Froste.lol** | ✅ | Music files |
| **BumpWorthy** | ✅ | Video/Audio bumps |
| **Direct URLs** | ✅ | Any direct file link |

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

```
gspread
google-auth
requests
yt-dlp
beautifulsoup4
pycryptodome
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

SheetDL features a semi-modern maroon themed UI with:
- Custom rounded window design
- Real-time download progress log
- Sheet tab selector (Broken)
- Format selection for YouTube and SoundCloud

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

### Sheet Tab Selector Not Working
The in-app sheet tab dropdown doesn't currently switch between tabs properly.

**Workaround:** Navigate to the desired tab/section in Google Sheets (e.g., "Unreleased", "Albums", etc.) and copy the URL directly from your browser while on that tab. The URL will contain the correct GID parameter for that specific tab.

### Double Clicking Icon Doesnt Open/Close Program

**Workaround:** just click on the icon to open and the minimize button to minimize it instead

### Enlarge Window Button Does Not Work

---

## 📋 Requirements File

Create a `requirements.txt`:
```
gspread>=5.0.0
google-auth>=2.0.0
requests>=2.25.0
yt-dlp>=2023.0.0
beautifulsoup4>=4.9.0
pycryptodome>=3.10.0
```

---

## 👤 Credits

**Created by [Angel2mp3](https://github.com/Angel2mp3)**

---

## 🤝 Contributing

Contributions are welcome! Feel free to:
- Report bugs
- Suggest new features
- Add support for new download sources
- Improve documentation

---

## 📜 License

This project is provided as-is for personal use. Please respect copyright laws and only download content you have the right to access.

---

## ⚠️ Disclaimer

SheetDL is a tool for downloading content you have legitimate access to. Users are responsible for ensuring they have the right to download any content. The developers are not responsible for any misuse of this software.

---


### Made with ❤️ for music collectors



