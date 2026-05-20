const { app, BrowserWindow, ipcMain, screen } = require('electron');
const path = require('path');
const fs = require('fs').promises;
const Store = require('electron-store');
const { getCalendarEvents, checkOutlookAvailable } = require('./services/outlookService');

const store = new Store();

let mainWindow;

function createWindow() {
  const { width, height } = screen.getPrimaryDisplay().workAreaSize;

  const windowWidth = 400;
  const windowHeight = 600;

  mainWindow = new BrowserWindow({
    width: windowWidth,
    height: windowHeight,
    x: width - windowWidth - 20,
    y: 20,
    frame: false,
    transparent: false,
    alwaysOnTop: true,
    skipTaskbar: false,
    resizable: true,
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

// Auto-start configuration
app.setLoginItemSettings({
  openAtLogin: true,
  path: app.getPath('exe')
});

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

ipcMain.handle('get-notes-directory', () => {
  return path.join(app.getPath('userData'), 'meeting-notes');
});
