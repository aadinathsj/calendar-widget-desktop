const { app, BrowserWindow, ipcMain, screen, shell } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const { getCalendarEvents, checkOutlookAvailable, checkAndGetEvents } = require('./services/outlookService');

let mainWindow;

function createWindow() {
  const { width, height } = screen.getPrimaryDisplay().workAreaSize;

  const windowWidth = 400;
  const windowHeight = 600;

  const iconPath = path.join(__dirname, '..', 'assets', 'icon.ico');
  const iconExists = require('fs').existsSync(iconPath);

  mainWindow = new BrowserWindow({
    width: windowWidth,
    height: windowHeight,
    x: width - windowWidth - 20,
    y: 20,
    frame: false,
    transparent: true,
    backgroundColor: '#00000000',
    alwaysOnTop: false,
    skipTaskbar: false,
    resizable: false,
    icon: iconExists ? iconPath : undefined,
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // Show window immediately, even before content loads
  mainWindow.once('ready-to-show', () => {
    mainWindow.show();
  });

  mainWindow.loadFile('src/renderer/index.html');

  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// APP LIFECYCLE
// ══════════════════════════════════════════════════════════════════════════════

app.whenReady().then(() => {
  createWindow();

  app.on('activate', function () {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit();
});

// No autostart - create desktop shortcut manually
// To create: npm run build, then right-click the exe → Send to → Desktop

// IPC Handlers
ipcMain.handle('check-outlook', async () => {
  try {
    const available = await checkOutlookAvailable();
    return { success: true, available };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('get-events', async (event, startDate, endDate) => {
  try {
    const events = await getCalendarEvents(startDate, endDate);
    return { success: true, events };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('save-note', async (event, meetingId, noteContent, meetingData) => {
  try {
    const notesDir = path.join(app.getPath('userData'), 'meeting-notes');
    await fs.mkdir(notesDir, { recursive: true });

    const fileName = `${meetingId.replace(/[^a-z0-9]/gi, '_')}.md`;
    const filePath = path.join(notesDir, fileName);

    let content = `# ${meetingData.subject}\n\n`;
    content += `**Date:** ${new Date(meetingData.start).toLocaleString()}\n`;
    content += `**Location:** ${meetingData.location || 'N/A'}\n`;
    if (meetingData.teamsLink) {
      content += `**Teams Link:** ${meetingData.teamsLink}\n`;
    }
    content += `**Organizer:** ${meetingData.organizer}\n\n`;
    content += `---\n\n`;
    content += noteContent;

    await fs.writeFile(filePath, content, 'utf-8');

    return { success: true, path: filePath };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('get-note', async (event, meetingId) => {
  try {
    const notesDir = path.join(app.getPath('userData'), 'meeting-notes');
    const fileName = `${meetingId.replace(/[^a-z0-9]/gi, '_')}.md`;
    const filePath = path.join(notesDir, fileName);

    try {
      const content = await fs.readFile(filePath, 'utf-8');
      const noteContent = content.split('---\n\n')[1] || '';
      return { success: true, content: noteContent, path: filePath };
    } catch (err) {
      return { success: true, content: '', path: filePath };
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('minimize-window', () => {
  if (mainWindow) {
    mainWindow.minimize();
  }
});

ipcMain.handle('close-window', () => {
  if (mainWindow) {
    mainWindow.close();
  }
});

ipcMain.handle('open-external', async (event, urlOrPath) => {
  try {
    // Detect if it's a file path (contains backslash or starts with a drive letter)
    const isFilePath = urlOrPath.includes('\\') || urlOrPath.includes('/') && !urlOrPath.startsWith('http');

    if (isFilePath) {
      // Use shell.openPath for file paths (opens in default app)
      await shell.openPath(urlOrPath);
    } else {
      // Use shell.openExternal for URLs
      await shell.openExternal(urlOrPath);
    }
  } catch (error) {
    console.error('Error opening:', error);
  }
});

ipcMain.handle('get-notes-directory', () => {
  return path.join(app.getPath('userData'), 'meeting-notes');
});

// ══════════════════════════════════════════════════════════════════════════════
// ACTIONS API - Shared Business Logic
// ══════════════════════════════════════════════════════════════════════════════
// These functions contain the actual logic for managing actions.json
// They are called by both IPC handlers and the HTTP REST API.

async function apiGetActions() {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    try {
      const data = await fs.readFile(actionsFile, 'utf-8');
      let actions = JSON.parse(data);

      // MIGRATION: Add type field to existing actions that don't have it
      let needsSave = false;
      actions = actions.map(item => {
        if (!item.type) {
          needsSave = true;
          return {
            ...item,
            type: 'action',
            parentId: item.parentId || null,
            path: item.path || '',
            createdAt: item.createdAt || new Date().toISOString()
          };
        }
        return item;
      });

      // Save migrated data back to disk
      if (needsSave) {
        await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');
        console.log('Migrated actions to new format with type field');
      }

      return { success: true, actions };
    } catch (err) {
      if (err.code === 'ENOENT') {
        // File doesn't exist yet — first run, return empty array
        return { success: true, actions: [] };
      }
      // Any other error (permissions, corrupt JSON) should surface
      throw err;
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function apiSaveActions(actions) {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function apiAddAction(action) {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    let actions = [];

    try {
      const data = await fs.readFile(actionsFile, 'utf-8');
      actions = JSON.parse(data);
    } catch (err) {
      // File doesn't exist, start with empty array
    }

    // Add new action with timestamp ID
    const newAction = {
      id: Date.now().toString(),
      type: 'action',
      title: action.title || '',
      url: action.url || '',
      path: action.path || '',
      note: action.note || '',
      pinned: false,
      parentId: action.parentId || null,
      createdAt: new Date().toISOString()
    };

    actions.push(newAction);
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');

    return { success: true, action: newAction };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function apiAddFolder(folder) {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    let actions = [];

    try {
      const data = await fs.readFile(actionsFile, 'utf-8');
      actions = JSON.parse(data);
    } catch (err) {
      // File doesn't exist, start with empty array
    }

    // Add new folder with timestamp ID
    const newFolder = {
      id: Date.now().toString(),
      type: 'folder',
      title: folder.title || '',
      pinned: false,
      parentId: folder.parentId || null,
      createdAt: new Date().toISOString()
    };

    actions.push(newFolder);
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');

    return { success: true, folder: newFolder };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function apiDeleteAction(actionId) {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    const data = await fs.readFile(actionsFile, 'utf-8');
    let actions = JSON.parse(data);

    // Find the item to delete
    const itemToDelete = actions.find(a => a.id === actionId);

    // If it's a folder, check if it has children
    if (itemToDelete && itemToDelete.type === 'folder') {
      const hasChildren = actions.some(a => a.parentId === actionId);
      if (hasChildren) {
        return {
          success: false,
          error: 'Folder is not empty. Please remove or move all items before deleting this folder.'
        };
      }
    }

    actions = actions.filter(a => a.id !== actionId);
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

async function apiValidatePath(filePath) {
  try {
    const fsSync = require('fs');
    const exists = fsSync.existsSync(filePath);
    return { success: true, exists };
  } catch (error) {
    return { success: false, error: error.message };
  }
}

// Actions CRUD handlers (IPC) - delegate to shared API functions
ipcMain.handle('get-actions', async () => apiGetActions());
ipcMain.handle('save-actions', async (event, actions) => apiSaveActions(actions));
ipcMain.handle('add-action', async (event, action) => apiAddAction(action));
ipcMain.handle('add-folder', async (event, folder) => apiAddFolder(folder));
ipcMain.handle('delete-action', async (event, actionId) => apiDeleteAction(actionId));
ipcMain.handle('validate-path', async (event, filePath) => apiValidatePath(filePath));

// ── Fast-startup handlers ─────────────────────────────────────────────────────

// Combined Outlook check + event fetch in one PS process (saves ~2 s vs two calls)
ipcMain.handle('check-and-get-events', async (event, startDate, endDate) => {
  try {
    const result = await checkAndGetEvents(new Date(startDate), new Date(endDate));
    return { success: true, available: result.available, events: result.events };
  } catch (error) {
    return { success: false, available: false, events: [], error: error.message };
  }
});

// Read cached events for a specific date (instant — no PS process needed)
ipcMain.handle('get-cached-events', async (event, dateKey) => {
  try {
    const cacheFile = path.join(app.getPath('userData'), 'events-cache.json');
    try {
      const data = await fs.readFile(cacheFile, 'utf-8');
      const cache = JSON.parse(data);

      // If dateKey provided, return that specific day; otherwise return entire cache
      if (dateKey && cache.days && cache.days[dateKey]) {
        return { success: true, events: cache.days[dateKey], dateKey, fromCache: true };
      } else if (dateKey) {
        return { success: true, events: null, dateKey, fromCache: false };
      } else {
        // Return entire cache (for backward compatibility)
        return { success: true, cache: cache.days || {}, cachedAt: cache.cachedAt };
      }
    } catch (err) {
      if (err.code === 'ENOENT') return { success: true, events: null, dateKey, fromCache: false };
      throw err;
    }
  } catch (error) {
    return { success: false, events: null, dateKey, fromCache: false };
  }
});

// Persist events for a single day
ipcMain.handle('save-events-cache', async (event, events, dateKey) => {
  try {
    const cacheFile = path.join(app.getPath('userData'), 'events-cache.json');

    // Read existing cache
    let cache = { days: {}, cachedAt: new Date().toISOString() };
    try {
      const data = await fs.readFile(cacheFile, 'utf-8');
      cache = JSON.parse(data);
      if (!cache.days) cache.days = {};
    } catch (err) {
      // File doesn't exist, use empty cache
    }

    // Update the specific day
    cache.days[dateKey] = events;
    cache.cachedAt = new Date().toISOString();

    await fs.writeFile(cacheFile, JSON.stringify(cache, null, 2), 'utf-8');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

// Persist events for multiple days (bulk cache for ±5 days)
ipcMain.handle('save-events-range-cache', async (event, eventsMap) => {
  try {
    const cacheFile = path.join(app.getPath('userData'), 'events-cache.json');

    // Read existing cache
    let cache = { days: {}, cachedAt: new Date().toISOString() };
    try {
      const data = await fs.readFile(cacheFile, 'utf-8');
      cache = JSON.parse(data);
      if (!cache.days) cache.days = {};
    } catch (err) {
      // File doesn't exist, use empty cache
    }

    // Merge new events
    Object.assign(cache.days, eventsMap);
    cache.cachedAt = new Date().toISOString();

    await fs.writeFile(cacheFile, JSON.stringify(cache, null, 2), 'utf-8');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});
