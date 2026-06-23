/**
 * electron/preload.js — Bridge an toàn giữa renderer và main process.
 */
'use strict';
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('edt', {
  pickFile:       ()        => ipcRenderer.invoke('dialog:openFile'),
  saveReport:     (defName) => ipcRenderer.invoke('dialog:saveReport', defName),
  compare:        (args)    => ipcRenderer.invoke('compare:run', args),
  exportReport:   (args)    => ipcRenderer.invoke('report:export', args),
  openPath:       (p)       => ipcRenderer.invoke('shell:openPath', p),
  showInFolder:   (p)       => ipcRenderer.invoke('shell:showInFolder', p),
  onProgress:     (cb)      => {
    const handler = (_e, payload) => cb(payload);
    ipcRenderer.on('progress', handler);
    return () => ipcRenderer.removeListener('progress', handler);
  },
});
