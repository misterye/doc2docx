const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFiles: () => ipcRenderer.invoke('select-files'),
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  convertDoc: (data) => ipcRenderer.invoke('convert-doc', data),
  cancelConversion: (taskId) => ipcRenderer.invoke('cancel-conversion', taskId),
});
