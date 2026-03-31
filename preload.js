const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('pdfTransformer', {
  libreOfficeStatus: () => ipcRenderer.invoke('libreoffice-status'),
  selectFile: () => ipcRenderer.invoke('select-file'),
  sourcePreviewUrl: (filePath) =>
    ipcRenderer.invoke('source-preview-url', filePath),
  convertToPdf: (filePath) => ipcRenderer.invoke('convert-to-pdf', filePath),
  exportPdf: (sourcePdfPath, defaultFileName) =>
    ipcRenderer.invoke('export-pdf', sourcePdfPath, defaultFileName),
});
