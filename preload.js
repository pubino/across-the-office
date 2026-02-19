const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  selectFolder: () => ipcRenderer.invoke('select-folder'),
  searchFiles: (folderPath, searchText, matchCase) =>
    ipcRenderer.invoke('search-files', { folderPath, searchText, matchCase }),
  replaceInFiles: (files, searchText, replaceText, matchCase, dryRun) =>
    ipcRenderer.invoke('replace-in-files', { files, searchText, replaceText, matchCase, dryRun }),
  revealFile: (filePath) => ipcRenderer.invoke('reveal-file', filePath),
  openReportWindow: (data) => ipcRenderer.invoke('open-report-window', data),
  onSearchProgress: (callback) => {
    ipcRenderer.on('search-progress', (event, data) => callback(data));
  },
  removeSearchProgressListener: () => {
    ipcRenderer.removeAllListeners('search-progress');
  }
});
