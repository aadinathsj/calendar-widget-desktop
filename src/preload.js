const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  checkOutlook: () => ipcRenderer.invoke('check-outlook'),
  getEvents: (startDate, endDate) => ipcRenderer.invoke('get-events', startDate, endDate),
  saveNote: (meetingId, noteContent, meetingData) =>
    ipcRenderer.invoke('save-note', meetingId, noteContent, meetingData),
  getNote: (meetingId) => ipcRenderer.invoke('get-note', meetingId),
  minimizeWindow: () => ipcRenderer.invoke('minimize-window'),
  closeWindow: () => ipcRenderer.invoke('close-window'),
  getNotesDirectory: () => ipcRenderer.invoke('get-notes-directory'),

  // Actions API
  getActions: () => ipcRenderer.invoke('get-actions'),
  saveActions: (actions) => ipcRenderer.invoke('save-actions', actions),
  addAction: (action) => ipcRenderer.invoke('add-action', action),
  deleteAction: (actionId) => ipcRenderer.invoke('delete-action', actionId),

  // Open URL in system default browser (Chrome)
  openExternal: (url) => ipcRenderer.invoke('open-external', url)
});
