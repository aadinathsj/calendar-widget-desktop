# Setup Guide for Calendar Widget

## Step 1: Azure AD App Registration

To access your Outlook calendar, you need to register an application in Azure Active Directory.

### Detailed Steps:

1. **Navigate to Azure Portal**
   - Go to https://portal.azure.com
   - Sign in with your Microsoft account (use your First Solar account: aadinath.sanjeevjayanthi@firstsolar.com)

2. **Create App Registration**
   - In the search bar, type "Azure Active Directory" and select it
   - Click "App registrations" in the left sidebar
   - Click "+ New registration" at the top

3. **Configure Registration**
   - **Name**: Enter "Calendar Widget" (or any name you prefer)
   - **Supported account types**: Select "Accounts in any organizational directory and personal Microsoft accounts"
   - **Redirect URI**: 
     - Select platform: "Public client/native (mobile & desktop)"
     - Enter URI: `http://localhost`
   - Click "Register"

4. **Copy Application ID**
   - On the Overview page, find "Application (client) ID"
   - Copy this ID (it looks like: `12345678-1234-1234-1234-123456789abc`)
   - You'll need this in the next step

5. **Enable Public Client Flow**
   - Click "Authentication" in the left sidebar
   - Scroll down to "Advanced settings"
   - Under "Allow public client flows", toggle "Yes"
   - Click "Save" at the top

6. **Verify API Permissions**
   - Click "API permissions" in the left sidebar
   - You should see "Microsoft Graph" with "User.Read" permission
   - Click "+ Add a permission"
   - Select "Microsoft Graph" → "Delegated permissions"
   - Search for and check: "Calendars.Read"
   - Click "Add permissions"

## Step 2: Configure the Application

1. **Open the Project**
   - Open `src/services/outlookService.js` in your editor

2. **Update Client ID**
   - Find the line: `clientId: 'YOUR_CLIENT_ID',`
   - Replace `'YOUR_CLIENT_ID'` with the Application ID you copied
   - Example: `clientId: '12345678-1234-1234-1234-123456789abc',`
   - Save the file

## Step 3: Install Dependencies

```bash
npm install
```

This will install all required packages including:
- Electron (desktop app framework)
- Microsoft Graph Client (for calendar access)
- Azure MSAL (for authentication)

## Step 4: Run the Application

```bash
npm start
```

Or for development with dev tools:

```bash
npm run dev
```

## Step 5: First Login

1. The widget will open in the top-right corner of your screen
2. Click "Sign in with Microsoft"
3. A browser window will open
4. Sign in with your Microsoft account
5. Grant permissions when prompted
6. The widget will load your calendar

## Step 6: Using the Widget

### Viewing Meetings
- Today's meetings are shown by default
- Use ← and → arrows to navigate days
- Click the ↻ button to refresh

### Taking Notes
- Click on any meeting card
- The notes editor will open
- Type your notes
- Click "Save Notes"
- Notes are saved as Markdown files with meeting metadata

### Finding Your Notes
Your notes are stored in:
```
C:\Users\FS150419\AppData\Roaming\outlook-calendar-widget\meeting-notes\
```

Each note file includes:
- Meeting title, date, time
- Location and Teams link
- Organizer information
- Your custom notes

## Troubleshooting

### "Authentication failed"
- Verify your client ID is correct in `outlookService.js`
- Check that you enabled public client flows in Azure
- Ensure you granted Calendar permissions

### "No events loading"
- Click the refresh button
- Check your internet connection
- Try signing out and back in
- Verify you have calendar permissions in Azure AD

### Widget not starting on boot
- Check Windows Task Manager → Startup tab
- Look for "Calendar Widget"
- Enable it if disabled

### Can't find my notes
- Run this command to see the notes directory:
  ```bash
  npm start
  ```
- Then press Ctrl+Shift+I to open dev tools
- In console, type: `await window.electronAPI.getNotesDirectory()`

## Building for Production

To create a standalone executable:

```bash
npm run build:win
```

The installer will be in the `dist` folder.

## Auto-Start Configuration

The app is configured to start automatically on Windows login. To disable:
1. Press Ctrl+Shift+Esc to open Task Manager
2. Go to the "Startup" tab
3. Find "Calendar Widget"
4. Right-click and select "Disable"

## Support

If you encounter issues:
1. Check the console (Ctrl+Shift+I in the app)
2. Review the README.md for common solutions
3. Verify your Azure AD setup matches this guide
