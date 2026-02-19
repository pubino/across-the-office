const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  getReportData: () => ipcRenderer.invoke('get-report-data')
});
