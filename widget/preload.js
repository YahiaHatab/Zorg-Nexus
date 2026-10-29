const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
    close: () => ipcRenderer.invoke('widget-close'),
    minimize: () => ipcRenderer.invoke('widget-minimize'),
    openPortal: () => ipcRenderer.invoke('widget-open-portal'),
    toggleAlwaysOnTop: () => ipcRenderer.invoke('widget-toggle-ontop')
});
