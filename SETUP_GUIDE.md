# Setup Guide — Calendar Widget

A desktop widget that surfaces your Outlook calendar directly on your Windows desktop — no cloud setup, no Azure registration, no authentication tokens required.

---

## Table of Contents

1. [Prerequisites](#1-prerequisites)
2. [Clone the Repository](#2-clone-the-repository)
3. [Install Dependencies](#3-install-dependencies)
4. [Launch the Widget](#4-launch-the-widget)
5. [Create a Desktop Shortcut](#5-create-a-desktop-shortcut)
6. [Feature Walkthrough](#6-feature-walkthrough)
7. [Configuration](#7-configuration)
8. [Build a Production Installer](#8-build-a-production-installer)
9. [Development Mode](#9-development-mode)
10. [Project Structure](#10-project-structure)
11. [Troubleshooting](#11-troubleshooting)
12. [Security & Privacy](#12-security--privacy)
13. [Advanced Customization](#13-advanced-customization)
14. [Upgrading & Uninstalling](#14-upgrading--uninstalling)

---

## 1. Prerequisites

| Requirement | Version / Notes |
|---|---|
| **OS** | Windows 10 or Windows 11 |
| **Microsoft Outlook** | 2016, 2019, 2021, or Microsoft 365 — must be **installed and running** |
| **Node.js** | v16 or later — [Download](https://nodejs.org/) |
| **Git** | Any recent version — [Download](https://git-scm.com/) |

### Verify your environment

Open **PowerShell** and run:

```powershell
# Node.js — should print v16.x.x or higher
node --version

# npm — comes with Node.js
npm --version

# Confirm Outlook is installed
Get-ItemProperty "HKLM:\Software\Microsoft\Windows\CurrentVersion\App Paths\OUTLOOK.EXE" -ErrorAction SilentlyContinue
```

If `node --version` fails, install Node.js and restart your terminal before continuing.

---

## 2. Clone the Repository

```bash
git clone https://github.com/aadinathsj/calendar-widget-desktop.git
cd calendar-widget-desktop
```

If you already have the folder, just navigate into it:

```bash
cd path\to\calendar-widget-desktop
```

---

## 3. Install Dependencies

```bash
npm install
```

This installs:

| Package | Purpose |
|---|---|
| `electron` | Desktop app framework (dev dependency) |
| `electron-builder` | Packaging/installer toolchain (dev dependency) |
| `markdown-it` | Markdown rendering for meeting notes |

> **Proxy / corporate network users:** If `npm install` hangs, configure your proxy:
> ```bash
> npm config set proxy http://proxy.example.com:8080
> npm config set https-proxy http://proxy.example.com:8080
> ```

---

## 4. Launch the Widget

Make sure **Microsoft Outlook is open** (it can be minimized), then:

```bash
npm start
```

The widget will:
1. Appear in the **top-right corner** of your primary display
2. Check that Outlook is accessible via COM
3. Load today's calendar events automatically

> If Outlook is not running you will see a "Retry Connection" prompt — open Outlook and click **Retry**.

---

## 5. Create a Desktop Shortcut

To launch the widget with a double-click — without opening a terminal — run:

```bash
npm run shortcut
```

This PowerShell script:
- Generates a custom calendar icon (`assets/icon.ico` + `assets/icon.png`) with today's date
- Places a **"Calendar Widget"** shortcut on your Desktop pointing to the local Electron binary

Double-click the shortcut any time to start the widget.

> To have the widget **launch at login**, see [Auto-Start at Login](#auto-start-at-login) under Advanced Customization.

---

## 6. Feature Walkthrough

### Calendar Tab

| Control | Action |
|---|---|
| `←` / `→` buttons | Navigate to the previous / next day |
| `↻` (refresh) button | Re-fetch events from Outlook |
| Click an event card | Open the notes panel for that meeting |
| `↑` / `↓` arrow keys | Keyboard-navigate between event cards |
| `Enter` | Open the focused event card |
| `Escape` | Close notes panel / dismiss modals |

**Event cards display:**
- Meeting title and time range
- Location (room or address)
- Organizer name
- Teams meeting button (if the meeting has a Teams link)
- Countdown to the next upcoming meeting

### Actions Tab

The **Actions** tab is a personal quick-launch board — store URLs you open often (dashboards, tickets, wikis, etc.).

| Control | Action |
|---|---|
| `+` button | Add a new action (title + URL) |
| Click an action card | Open the URL in your default browser |
| Expand / collapse | Toggle full details on an action card |
| Delete button | Remove an action permanently |

Actions are persisted locally between sessions in your Windows user-data folder.

### Meeting Notes

1. Click any event card to open the notes panel.
2. Type your notes (plain text or Markdown).
3. Notes **auto-save** as you type (with a short debounce delay) — no manual save needed.
4. Click **Save** to force an immediate save.
5. Notes are written as `.md` files to:
   ```
   C:\Users\<YourUsername>\AppData\Roaming\outlook-calendar-widget\meeting-notes\
   ```
   Each file is named after the meeting ID and contains the meeting metadata header plus your notes.

### Teams Links

When a meeting contains a Teams link (in the meeting body or `OnlineMeetingURL` property), a **Join Teams** button appears on the event card. Clicking it opens the link in your default browser.

---

## 7. Configuration

### Window Size and Position

Edit [src/main.js](src/main.js) (around line 11):

```javascript
const windowWidth  = 400;  // pixel width of the widget
const windowHeight = 600;  // pixel height of the widget

// Positioned 20 px from the right and top edges of the primary display
x: width - windowWidth - 20,
y: 20,
```

### Notes Storage Location

Edit [src/main.js](src/main.js) (the `save-note` IPC handler):

```javascript
const notesDir = path.join(app.getPath('userData'), 'meeting-notes');
// Override example:
// const notesDir = 'D:\\My Notes\\meetings';
```

### Accent / Theme Colors

Edit [src/renderer/styles.css](src/renderer/styles.css):

```css
/* Header and window controls */
header { background: rgba(88, 28, 220, 0.85); }

/* Event card left border accent */
.event-card { border-left: 4px solid #7c3aed; }
```

---

## 8. Build a Production Installer

To distribute the widget or use it without Node.js:

```bash
npm run build:win
```

Output: `dist/Calendar Widget Setup x.x.x.exe`

Run the installer — it:
- Installs the app to `%LOCALAPPDATA%\Programs\outlook-calendar-widget\`
- Adds a Start Menu shortcut
- Does **not** require Node.js on the target machine

---

## 9. Development Mode

```bash
npm run dev
```

Opens the widget with **Chrome DevTools** attached. All console logs, network requests, and JS errors are visible in the DevTools panel.

---

## 10. Project Structure

```
calendar-widget-desktop/
├── assets/
│   ├── icon.ico               # App icon (auto-generated by npm run shortcut)
│   └── icon.png               # PNG version of the icon
│
├── scripts/
│   └── create-shortcut.ps1    # Desktop shortcut + icon generator
│
├── src/
│   ├── main.js                # Electron main process
│   │                          #   • Creates the BrowserWindow
│   │                          #   • Handles all IPC channels
│   │                          #   • Reads/writes notes and actions to disk
│   │
│   ├── preload.js             # Context-isolated bridge (main ↔ renderer)
│   │                          #   • Exposes a safe electronAPI to the UI
│   │
│   ├── services/
│   │   └── outlookService.js  # Outlook COM automation via PowerShell
│   │                          #   • checkOutlookAvailable()
│   │                          #   • getCalendarEvents(startDate, endDate)
│   │
│   └── renderer/
│       ├── index.html         # Widget HTML shell
│       ├── styles.css         # Glassmorphic styles
│       └── app.js             # All UI logic (tabs, notes, actions, keyboard nav)
│
├── package.json
├── README.md
└── SETUP_GUIDE.md             # ← You are here
```

### IPC Channels Reference

| Channel | Direction | Purpose |
|---|---|---|
| `check-outlook` | renderer → main | Verify Outlook COM is accessible |
| `get-events` | renderer → main | Fetch calendar events for a date range |
| `save-note` | renderer → main | Write a meeting note to disk |
| `get-note` | renderer → main | Read a saved meeting note |
| `get-actions` | renderer → main | Load the actions list |
| `save-actions` | renderer → main | Persist the actions list |
| `add-action` | renderer → main | Append one action |
| `delete-action` | renderer → main | Remove one action by ID |
| `minimize-window` | renderer → main | Minimize the widget |
| `close-window` | renderer → main | Quit the app |
| `open-external` | renderer → main | Open a URL in the system browser |
| `get-notes-directory` | renderer → main | Return the notes folder path |

---

## 11. Troubleshooting

### "Outlook not found" / "Retry Connection"

| Cause | Fix |
|---|---|
| Outlook is not installed | Install Microsoft Outlook |
| Outlook is not running | Open Outlook (can be minimized) |
| First-time Outlook setup not completed | Open Outlook and finish account setup |
| Windows blocking PowerShell COM access | See PowerShell execution policy fix below |

**Check / fix PowerShell execution policy:**
```powershell
# View current policy
Get-ExecutionPolicy

# If it's 'Restricted', change it (run PowerShell as Administrator):
Set-ExecutionPolicy RemoteSigned -Scope CurrentUser
```

**Re-register Outlook COM objects (run as Administrator):**
```cmd
cd "C:\Program Files\Microsoft Office\root\Office16"
outlook.exe /regserver
```

---

### No Calendar Events Showing

1. Open Outlook and confirm events exist on the selected date.
2. Wait for Outlook to finish syncing with your mail server.
3. Click **↻ Refresh** in the widget.
4. Navigate away to another day, then back.

---

### Teams Join Button Not Appearing

- The meeting must be a Teams meeting (not a Skype or in-person meeting).
- Some older Outlook versions don't populate `IsOnlineMeeting` — the button still appears if the Teams URL is present in the meeting body.

---

### Notes Not Saving

| Cause | Fix |
|---|---|
| No write permission to `%APPDATA%` | Run the widget once as Administrator to create the folder |
| Disk full | Free up space |
| Antivirus blocking file creation | Add an exception for `%APPDATA%\outlook-calendar-widget\` |

---

### `npm install` Fails

- Ensure Node.js ≥ v16 is installed (`node --version`).
- On a corporate network, configure the proxy (see [Step 3](#3-install-dependencies)).
- Delete `node_modules` and `package-lock.json`, then retry:
  ```bash
  Remove-Item -Recurse -Force node_modules
  Remove-Item package-lock.json
  npm install
  ```

---

### Shortcut Not Created / Icon Missing

Run the shortcut script with verbose output:
```powershell
powershell -ExecutionPolicy Bypass -File scripts\create-shortcut.ps1
```
Errors will be printed to the console. The most common cause is a missing `System.Drawing` assembly — this is included in all modern Windows versions but may require .NET Framework to be enabled.

---

## 12. Security & Privacy

### PowerShell Execution
The app spawns PowerShell with `-ExecutionPolicy Bypass` **only** to run the inline COM script embedded in `outlookService.js`. No external scripts are downloaded or executed.

### Data Privacy
- **All data is local** — nothing leaves your machine.
- **No telemetry** — zero usage tracking or analytics.
- **No credentials** — no passwords, API keys, or tokens are stored.
- **Read-only calendar access** — the app never creates, modifies, or deletes Outlook items.

### Notes Security
- Stored as plain-text Markdown under `%APPDATA%\outlook-calendar-widget\meeting-notes\`.
- Protected by standard Windows NTFS user permissions.
- Not encrypted by default — use Windows BitLocker or EFS if you need encryption at rest.

### Electron Security
- `nodeIntegration` is **disabled** in the renderer.
- `contextIsolation` is **enabled**; the renderer can only call the explicitly whitelisted `electronAPI` methods defined in `preload.js`.

---

## 13. Advanced Customization

### Auto-Start at Login

The development build (`npm start`) does not register an auto-start entry. To launch the widget at Windows login:

**Option A — Task Scheduler (recommended)**
1. Open **Task Scheduler** → *Create Basic Task*.
2. Trigger: **At log on**.
3. Action: **Start a program** → browse to `node_modules\electron\dist\electron.exe`.
4. Add argument: `.` (dot = current directory).
5. Set *Start in* to your project folder.

**Option B — Startup folder**
1. Press `Win + R`, type `shell:startup`, press Enter.
2. Copy the Desktop shortcut (created by `npm run shortcut`) into this folder.

**Option C — Production installer**
Build and install the app with `npm run build:win`. The installer registers an auto-start entry automatically.

---

### Add Auto-Refresh

The widget refreshes only on demand (click ↻). To add a timed refresh, append this to `src/renderer/app.js` after the `init()` call:

```javascript
// Refresh calendar every 5 minutes
setInterval(loadEvents, 5 * 60 * 1000);
```

---

### Multiple Calendar Support

`outlookService.js` currently reads `GetDefaultFolder(9)` (the primary Calendar folder). To read additional calendars, iterate over `namespace.Folders` and merge appointments from each calendar folder.

---

### Custom Theming

```css
/* src/renderer/styles.css */

/* Background overlay */
body { background: rgba(15, 10, 40, 0.6); }

/* Card glassmorphism */
.event-card {
  background: rgba(255, 255, 255, 0.12);
  backdrop-filter: blur(20px);
  border-left: 4px solid #7c3aed; /* accent color */
}

/* Header */
header { background: rgba(88, 28, 220, 0.85); }
```

---

## 14. Upgrading & Uninstalling

### Upgrade (development build)

```bash
git pull
npm install
```

### Upgrade (production build)

```bash
git pull
npm install
npm run build:win
# Run the new installer from dist/
```

### Uninstall

**Development build:** Delete the project folder.

**Production build:**
1. Open *Windows Settings → Apps*.
2. Search for **Calendar Widget** → Uninstall.

In both cases, meeting notes and actions remain at:
```
C:\Users\<YourUsername>\AppData\Roaming\outlook-calendar-widget\
```
Delete this folder manually if you want to remove all saved data.

---

*Questions or issues? Open the widget in dev mode (`npm run dev`) to inspect the console, then check the [README.md](README.md) for additional context.*
