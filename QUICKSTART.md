# Quick Start Guide

## ⚡ Get Started in 5 Minutes

### 1. Register Azure AD App (One-time setup)

**You MUST do this first** - the app needs permission to access your Outlook calendar.

1. Go to https://portal.azure.com
2. Search for "App registrations" → Click "New registration"
3. Fill in:
   - **Name**: `Calendar Widget`
   - **Account types**: Select "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: Select "Public client/native" and enter `http://localhost`
4. Click "Register"
5. **COPY** the "Application (client) ID" (you'll need this next!)
6. Click "Authentication" → Enable "Allow public client flows" → Save

### 2. Add Calendar Permissions

1. Still in Azure Portal, click "API permissions"
2. Click "+ Add a permission"
3. Select "Microsoft Graph" → "Delegated permissions"
4. Search and add: `Calendars.Read`
5. Click "Add permissions"

### 3. Configure the App

Open [src/services/outlookService.js](src/services/outlookService.js#L10) and replace:

```javascript
clientId: 'YOUR_CLIENT_ID'  // Replace with your Application ID from step 1
```

### 4. Run the App

```bash
npm start
```

### 5. Sign In

- Click "Sign in with Microsoft"
- Grant permissions
- Your calendar will load!

## Features

✅ **View meetings** - See all your Outlook events for the day  
✅ **Navigate days** - Use ← → arrows to see upcoming meetings  
✅ **Teams links** - Quick access to join meetings  
✅ **Take notes** - Click any meeting to add notes (saved as .md files)  
✅ **Auto-start** - Opens automatically when you log in to Windows  
✅ **Always visible** - Stays on top in the corner of your screen  

## Where are my notes?

```
C:\Users\FS150419\AppData\Roaming\outlook-calendar-widget\meeting-notes\
```

Each note includes meeting metadata + your custom notes in Markdown format.

## Troubleshooting

**"Authentication failed"**
- Double-check your client ID in `outlookService.js`
- Verify you enabled "public client flows" in Azure
- Make sure you added Calendars.Read permission

**"No events showing"**
- Click the refresh (↻) button
- Check internet connection
- Verify calendar permissions in Azure

**Need help?** See [SETUP_GUIDE.md](SETUP_GUIDE.md) for detailed instructions.
