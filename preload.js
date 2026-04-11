const { contextBridge, ipcRenderer } = require('electron')

contextBridge.exposeInMainWorld('api', {
  openAndParseFile: () => ipcRenderer.invoke('open-and-parse-file'),
  parseFile: (filePath) => ipcRenderer.invoke('parse-file', filePath),
  onOpenFile: (callback) => ipcRenderer.on('open-file', (_event, filePath) => callback(filePath)),
})
