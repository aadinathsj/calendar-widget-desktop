# Outlook Calendar Widget

A simple desktop widget that displays your Outlook calendar events and allows you to take notes for each meeting. **No cloud setup or registration required** - works directly with your local Outlook installation.

## Features

- 📅 View your Outlook calendar events directly from your local Outlook
- 🕐 See meeting times, locations, and organizers
- 🎥 Quick access to Teams meeting links
- 📝 Take and save notes for each meeting (stored as Markdown files)
- ⏭️ Navigate through days to see upcoming meetings
- 🚀 Auto-starts when you power on your computer
- 📍 Sits in the corner of your desktop

## Requirements

- Windows 10/11
- Microsoft Outlook installed and configured
- Node.js (v16 or later)

## Quick Setup

**No Azure registration needed!** The app works directly with your installed Outlook.

```bash
# Install dependencies
npm install

# Run the app
npm start
```

That's it! The widget will:
1. Check for Outlook on your system
2. Connect to your local Outlook calendar
3. Display your meetings

## Usage

1. **View Calendar**: The widget displays today's meetings automatically
2. **Navigate Days**: Use the ← → arrow buttons to view meetings for other days
3. **Refresh**: Click the ↻ button to refresh your calendar
4. **Take Notes**: Click on any meeting card to open the notes editor
5. **Save Notes**: Notes are automatically saved as Markdown files

## Notes Storage

Meeting notes are stored as Markdown files in:
- Windows: `C:\Users\<YourUsername>\AppData\Roaming\outlook-calendar-widget\meeting-notes\`

Each note file includes:
- Meeting title
- Date and time
- Location
- Teams link (if available)
- Organizer information
- Your custom notes

## How It Works

The widget uses Windows COM automation to interact with your local Outlook installation through PowerShell. This means:
- ✅ No cloud authentication needed
- ✅ No app registration required
- ✅ No API tokens or credentials
- ✅ Works offline (if Outlook is running)
- ✅ Automatically stays in sync with your Outlook calendar

## Auto-Start

The widget is configured to automatically start when you log in to Windows. To disable:
1. Press `Ctrl+Shift+Esc` to open Task Manager
2. Go to the "Startup" tab
3. Find "Calendar Widget"
4. Right-click and select "Disable"

## Building for Production

To create a standalone executable:

```bash
npm run build:win
```

The installer will be in the `dist` folder. Install it and the widget will run without needing Node.js.

## Troubleshooting

### "Outlook not found"
- Ensure Microsoft Outlook is installed on your computer
- Open Outlook at least once to complete initial setup
- Make sure Outlook is running (it can be minimized)

### "Error connecting to Outlook"
- Check if Outlook is responding (not frozen or updating)
- Try restarting Outlook
- Click the "Retry Connection" button in the widget

### Calendar events not loading
- Verify you have events in your Outlook calendar for the selected date
- Click the refresh (↻) button
- Check if Outlook is syncing with your email server

### Widget not starting on boot
- Check Windows Task Manager → Startup tab
- Ensure "Calendar Widget" is enabled
- If using the development version (`npm start`), auto-start won't work - build the production version first

## Development

```bash
# Run with developer tools
npm run dev
```

## Technical Details

- **Framework**: Electron (desktop app)
- **Outlook Integration**: Windows COM API via PowerShell
- **Storage**: Local filesystem (Markdown files)
- **UI**: Vanilla JavaScript, CSS

## Privacy

All data stays on your local machine:
- Calendar data is read directly from your local Outlook
- Notes are stored in your user directory
- No data is sent to any cloud service
- No authentication or tracking

## License

MIT
