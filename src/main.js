const { app, BrowserWindow, ipcMain, screen, shell } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const { getCalendarEvents, checkOutlookAvailable } = require('./services/outlookService');

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

  mainWindow.loadFile('src/renderer/index.html');

  if (process.argv.includes('--dev')) {
    mainWindow.webContents.openDevTools();
  }
}

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

ipcMain.handle('open-external', (event, url) => {
  shell.openExternal(url);
});

ipcMain.handle('get-notes-directory', () => {
  return path.join(app.getPath('userData'), 'meeting-notes');
});

// Actions CRUD handlers
ipcMain.handle('get-actions', async () => {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    try {
      const data = await fs.readFile(actionsFile, 'utf-8');
      return { success: true, actions: JSON.parse(data) };
    } catch (err) {
      // File doesn't exist yet, return empty array
      return { success: true, actions: [] };
    }
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('save-actions', async (event, actions) => {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');
    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('add-action', async (event, action) => {
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
      ...action,
      createdAt: new Date().toISOString()
    };

    actions.push(newAction);
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');

    return { success: true, action: newAction };
  } catch (error) {
    return { success: false, error: error.message };
  }
});

ipcMain.handle('delete-action', async (event, actionId) => {
  try {
    const actionsFile = path.join(app.getPath('userData'), 'actions.json');
    const data = await fs.readFile(actionsFile, 'utf-8');
    let actions = JSON.parse(data);

    actions = actions.filter(a => a.id !== actionId);
    await fs.writeFile(actionsFile, JSON.stringify(actions, null, 2), 'utf-8');

    return { success: true };
  } catch (error) {
    return { success: false, error: error.message };
  }
});
