# Outlook Calendar Widget

A desktop widget that displays your Outlook calendar events and allows you to take notes for each meeting.

## Features

- 📅 View your Outlook calendar events
- 🕐 See meeting times, locations, and organizers
- 🎥 Quick access to Teams meeting links
- 📝 Take and save notes for each meeting (stored as Markdown files)
- ⏭️ Navigate through days to see upcoming meetings
- 🚀 Auto-starts when you power on your computer
- 📍 Sits in the corner of your desktop

## Setup

### Prerequisites

1. Node.js (v16 or later)
2. An Azure AD application registration

### Azure AD App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" → "App registrations" → "New registration"
3. Name: "Calendar Widget" (or your choice)
4. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
5. Redirect URI: Select "Public client/native (mobile & desktop)" and enter: `http://localhost`
6. Click "Register"
7. Copy the "Application (client) ID" from the Overview page
8. Go to "Authentication" → "Advanced settings" → Enable "Allow public client flows" → Save
9. Update `src/services/outlookService.js` and replace `YOUR_CLIENT_ID` with your Application ID

### Installation

```bash
# Install dependencies
npm install

# Run in development mode
npm run dev

# Build for production
npm run build:win
```

## Usage

1. **First Launch**: Sign in with your Microsoft account when prompted
2. **View Calendar**: The widget will display today's meetings
3. **Navigate Days**: Use the arrow buttons to view meetings for other days
4. **Take Notes**: Click on any meeting to open the notes editor
5. **Save Notes**: Notes are automatically saved as Markdown files in your user data directory

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

## Auto-Start

The widget is configured to automatically start when you log in to Windows. You can disable this in the Windows Task Manager under the "Startup" tab.

## Development

```bash
# Run with dev tools open
npm run dev
```

## Troubleshooting

### Authentication Issues
- Ensure your Azure AD app is properly configured
- Check that the client ID in `outlookService.js` is correct
- Try signing out and signing in again

### Calendar Not Loading
- Verify you have an active internet connection
- Check that you've granted calendar permissions to the app
- Try clicking the refresh button

### Auto-Start Not Working
- Check Windows Task Manager → Startup tab
- Ensure the app has permission to run at startup

## License

MIT
