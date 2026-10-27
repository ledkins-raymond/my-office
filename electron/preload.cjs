// Narrow bridge between the editor page and the main process.
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('office', {
  setDirty: (dirty) => ipcRenderer.send('doc:dirty', dirty),
  action: (name) => ipcRenderer.send('menu:action', name),
  onLoad: (cb) => ipcRenderer.on('doc:load', (_e, payload) => cb(payload)),
  onSaved: (cb) => ipcRenderer.on('doc:saved', (_e, payload) => cb(payload)),
  onCommand: (cb) => ipcRenderer.on('editor:command', (_e, name) => cb(name)),
  // Main asks for the document as .docx bytes; the page answers with makeDocx().
  onCollectDocx: (makeDocx) =>
    ipcRenderer.on('docx:collect', async (_e, requestId) => {
      try {
        ipcRenderer.send('docx:bytes', requestId, await makeDocx(), null);
      } catch (err) {
        ipcRenderer.send('docx:bytes', requestId, null, String(err?.message || err));
      }
    }),
});
