# Quick Start Guide

## ⚡ Get Started in 2 Minutes

**No cloud setup required!** Works directly with your local Outlook.

### Prerequisites

✅ Windows 10 or 11  
✅ Microsoft Outlook installed  
✅ Node.js installed  

### Installation

```bash
# 1. Install dependencies
npm install

# 2. Run the app
npm start
```

That's it! The widget will open in the top-right corner.

## First Run

When you first launch the app:

1. **Automatic Connection**: The widget checks for Outlook on your system
2. **If Outlook Found**: Your calendar loads immediately
3. **If Not Found**: You'll see a message to open Outlook first

### Make sure Outlook is running
- Open Microsoft Outlook (it can be minimized to the system tray)
- The widget needs Outlook to be running to access your calendar
- Click "Retry Connection" in the widget if needed

## Features At a Glance

### View Your Calendar
- Today's meetings appear automatically
- See time, location, organizer, and Teams links
- Clean, simple interface

### Navigate Days
- Click **←** to see yesterday's meetings
- Click **→** to see tomorrow's meetings
- Current date is displayed at the top

### Take Meeting Notes
1. Click any meeting card
2. A notes editor opens
3. Type your notes
4. Click "Save Notes"
5. Notes are saved as Markdown files with meeting metadata

### Where Are My Notes?
```
C:\Users\FS150419\AppData\Roaming\outlook-calendar-widget\meeting-notes\
```

Each note includes:
- Meeting details (title, time, location)
- Teams link (if available)
- Your custom notes in Markdown format

## Auto-Start

The widget is configured to start automatically when you log in to Windows. You can find it in:
- Task Manager → Startup tab → "Calendar Widget"

## Building Standalone App

Want to share it or run without Node.js?

```bash
npm run build:win
```

This creates an installer in the `dist` folder. Install it and you're done!

## Troubleshooting

### "Outlook not found"
**Solution**: Open Microsoft Outlook, then click "Retry Connection"

### "Error connecting to Outlook"  
**Solution**: 
1. Restart Outlook
2. Make sure Outlook isn't frozen or updating
3. Try restarting the widget

### No events showing
**Solution**:
1. Check you have events in Outlook for that day
2. Click the refresh (↻) button
3. Navigate to a different day and back

### Widget won't auto-start
**Solution**: If running with `npm start`, auto-start won't work. Build the production version with `npm run build:win` and install it.

## Key Differences from Cloud Approach

| Feature | This App | Cloud/Azure Approach |
|---------|----------|---------------------|
| Setup | Just run it | Register Azure app |
| Authentication | None needed | OAuth login required |
| Data | Local only | Cloud API calls |
| Privacy | 100% local | Data goes to Microsoft |
| Offline | Works offline | Needs internet |
| Complexity | Simple | Complex |

## Next Steps

- Customize the widget position and size (see [src/main.js](src/main.js#L14))
- Change auto-start behavior (see [src/main.js](src/main.js#L44))
- Modify the UI styling (see [src/renderer/styles.css](src/renderer/styles.css))

Enjoy your calendar widget! 📅
