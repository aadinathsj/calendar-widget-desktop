# Setup Guide for Calendar Widget

## Overview

This calendar widget works **directly with your local Outlook installation** using Windows COM automation. No cloud setup, no Azure registration, no authentication required!

## System Requirements

### Required
- **Operating System**: Windows 10 or Windows 11
- **Microsoft Outlook**: Any recent version (2016, 2019, 2021, or Microsoft 365)
- **Node.js**: Version 16 or later ([Download here](https://nodejs.org/))

### Verification

Check if you have the requirements:

```powershell
# Check Node.js version
node --version

# Check if Outlook is installed (should return path)
Get-ItemProperty HKLM:\Software\Microsoft\Windows\CurrentVersion\App` Paths\OUTLOOK.EXE -ErrorAction SilentlyContinue
```

## Installation Steps

### 1. Clone or Download the Project

If you haven't already, navigate to your project directory:
```bash
cd "c:\Users\FS150419\OneDrive - First Solar\Documents\Projects\calendar-widget"
```

### 2. Install Dependencies

```bash
npm install
```

This installs:
- Electron (desktop app framework)
- electron-store (settings storage)
- markdown-it (Markdown rendering, if needed)

### 3. Ensure Outlook is Running

**Important**: The widget needs Outlook to be running to access your calendar.

1. Open Microsoft Outlook
2. Make sure your email account is configured
3. Outlook can be minimized (it doesn't need to be visible)

### 4. Run the Widget

```bash
npm start
```

The widget will:
1. Open in the top-right corner of your screen
2. Check for Outlook
3. Connect to your calendar
4. Display today's meetings

## How It Works

The widget uses **Windows COM Automation** to communicate with Outlook:

```
Widget (Electron) → PowerShell Script → Outlook COM API → Your Calendar
```

### What is COM?
- COM (Component Object Model) is a Windows technology
- It allows programs to interact with each other
- Outlook exposes its data through COM
- No internet or cloud services required

### PowerShell Bridge
The app runs PowerShell scripts that:
1. Create an Outlook.Application COM object
2. Access the calendar folder
3. Retrieve appointments for the specified date range
4. Extract meeting details (title, time, location, Teams links)
5. Return the data as JSON

See [src/services/outlookService.js](src/services/outlookService.js) for the implementation.

## Configuration

### Widget Position and Size

Edit [src/main.js](src/main.js#L14):

```javascript
const windowWidth = 400;   // Change width
const windowHeight = 600;  // Change height

// Position calculation
x: width - windowWidth - 20,  // 20px from right edge
y: 20,                         // 20px from top
```

### Auto-Start Behavior

Edit [src/main.js](src/main.js#L44):

```javascript
app.setLoginItemSettings({
  openAtLogin: true,   // Set to false to disable auto-start
  path: app.getPath('exe')
});
```

### Notes Storage Location

By default, notes are stored in:
```
C:\Users\<YourUsername>\AppData\Roaming\outlook-calendar-widget\meeting-notes\
```

To change this, edit [src/main.js](src/main.js#L69):

```javascript
const notesDir = path.join(app.getPath('userData'), 'meeting-notes');
// Change to custom location:
// const notesDir = 'C:\\My Notes\\meetings';
```

## Building for Production

### Create Standalone Executable

```bash
npm run build:win
```

This creates:
- An installer in `dist/outlook-calendar-widget Setup x.x.x.exe`
- No need for Node.js after installation
- Auto-start will work properly

### Install the Built App

1. Run the installer from the `dist` folder
2. The app installs to `C:\Users\<You>\AppData\Local\Programs\outlook-calendar-widget\`
3. A shortcut is added to your Start Menu
4. Auto-start is configured

## Development

### Run with Developer Tools

```bash
npm run dev
```

This opens the widget with Chrome DevTools for debugging.

### Project Structure

```
calendar-widget/
├── src/
│   ├── main.js                 # Electron main process
│   │                          # - Creates window
│   │                          # - Handles IPC communication
│   │                          # - Manages notes storage
│   │
│   ├── preload.js             # Security bridge between main/renderer
│   │
│   ├── services/
│   │   └── outlookService.js  # Outlook COM automation
│   │                          # - PowerShell script execution
│   │                          # - Calendar data retrieval
│   │
│   └── renderer/
│       ├── index.html         # UI structure
│       ├── styles.css         # Styling
│       └── app.js             # Frontend logic
│
├── package.json               # Project configuration
└── README.md                  # Documentation
```

## Troubleshooting

### Issue: "Outlook not found"

**Causes:**
- Outlook is not installed
- Outlook is not running
- Windows is blocking COM access

**Solutions:**
1. Install Microsoft Outlook
2. Open Outlook at least once to complete setup
3. Keep Outlook running (can be minimized)
4. Check Windows security settings aren't blocking PowerShell

### Issue: "Error connecting to Outlook"

**Causes:**
- Outlook is frozen or updating
- PowerShell execution policy
- COM registration issues

**Solutions:**
1. Restart Outlook
2. Check PowerShell execution policy:
   ```powershell
   Get-ExecutionPolicy
   # Should be RemoteSigned or Unrestricted
   ```
3. Re-register Outlook COM (as admin):
   ```cmd
   cd "C:\Program Files\Microsoft Office\root\Office16"
   outlook.exe /regserver
   ```

### Issue: No calendar events showing

**Causes:**
- No events on selected date
- Calendar not synced
- Wrong calendar being read (if multiple)

**Solutions:**
1. Check Outlook directly for events on that date
2. Wait for Outlook to sync with email server
3. Click refresh (↻) button in widget
4. Navigate to a different day and back

### Issue: Teams links not appearing

**Causes:**
- Meeting isn't a Teams meeting
- Teams link is in a non-standard format
- Outlook version doesn't support online meeting properties

**Solutions:**
- Teams links are extracted from meeting body and online meeting properties
- Some older Outlook versions may not populate the `IsOnlineMeeting` property
- The link should still appear if it's in the meeting body

### Issue: Can't save notes

**Causes:**
- No write permission to user directory
- Disk full
- Antivirus blocking file creation

**Solutions:**
1. Check available disk space
2. Verify write permissions to `%APPDATA%`
3. Check antivirus logs for blocked operations
4. Try running as administrator (once to create directory)

### Issue: Widget not auto-starting

**Causes:**
- Using development mode (`npm start`)
- Production build not installed
- Auto-start disabled in Task Manager

**Solutions:**
1. Build production version: `npm run build:win`
2. Install the built app
3. Check Task Manager → Startup tab
4. Enable "Calendar Widget" if disabled

## Security Considerations

### PowerShell Execution
The app runs PowerShell with `-ExecutionPolicy Bypass` to ensure scripts work. This is safe because:
- Scripts are embedded in the app code
- No external scripts are loaded
- Only COM commands are executed
- No system modifications are made

### Data Privacy
- **All data stays local**: Nothing is sent to the cloud
- **No telemetry**: No usage tracking or analytics
- **No authentication**: No passwords or tokens stored
- **Read-only calendar access**: The app only reads calendar data, never modifies it

### Notes Security
- Notes are stored as plain text Markdown files
- Stored in your Windows user profile directory
- Protected by Windows file permissions
- Not encrypted by default (encrypt the folder if needed)

## Advanced Customization

### Change Update Interval

The calendar refreshes when you click the refresh button. To add auto-refresh, edit [src/renderer/app.js](src/renderer/app.js):

```javascript
// Add after init() function
setInterval(loadEvents, 5 * 60 * 1000); // Refresh every 5 minutes
```

### Add Multiple Calendar Support

Currently reads the default calendar. To support multiple calendars, modify [src/services/outlookService.js](src/services/outlookService.js).

### Custom Theming

Edit [src/renderer/styles.css](src/renderer/styles.css):

```css
/* Change primary color */
.window-controls,
header {
  background: #0078d4; /* Change this color */
}

/* Change accent color */
.event-card {
  border-left: 4px solid #0078d4; /* Change this color */
}
```

## Performance

### Memory Usage
- Typical: 50-100 MB RAM
- Outlook COM calls are released after each query
- Garbage collection is forced after COM operations

### CPU Usage
- Idle: ~0% CPU
- Fetching events: Brief spike, then back to idle
- PowerShell overhead is minimal

### Outlook Impact
- Minimal impact on Outlook performance
- COM access is read-only
- No continuous polling (only when refreshing)

## Support

If you encounter issues:
1. Check this guide's troubleshooting section
2. Review the [README.md](README.md) for common questions
3. Open the app with dev tools (`npm run dev`) to see console errors
4. Check Windows Event Viewer for COM errors

## Upgrading

To update to a new version:

```bash
# Pull latest code
git pull

# Reinstall dependencies
npm install

# Rebuild
npm run build:win
```

## Uninstalling

### Development Version
Just delete the project folder.

### Production Version
1. Open Windows Settings → Apps
2. Find "Calendar Widget"
3. Click Uninstall

Notes will remain in:
```
C:\Users\<You>\AppData\Roaming\outlook-calendar-widget\
```

Delete this folder manually if you want to remove notes.

---

Enjoy your Outlook calendar widget! 📅
