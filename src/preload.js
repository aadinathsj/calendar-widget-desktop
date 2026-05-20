const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getAuthStatus: () => ipcRenderer.invoke('get-auth-status'),
  login: () => ipcRenderer.invoke('login'),
  getEvents: (startDate, endDate) => ipcRenderer.invoke('get-events', startDate, endDate),
  saveNote: (meetingId, noteContent, meetingData) =>
    ipcRenderer.invoke('save-note', meetingId, noteContent, meetingData),
  getNote: (meetingId) => ipcRenderer.invoke('get-note', meetingId),
  minimizeWindow: () => ipcRenderer.invoke('minimize-window'),
  closeWindow: () => ipcRenderer.invoke('close-window'),
  getNotesDirectory: () => ipcRenderer.invoke('get-notes-directory')
});
