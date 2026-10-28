// Narrow bridge between the app pages (My Word, My Excel) and the main process.
const { contextBridge, ipcRenderer } = require('electron');

// Main asks the page for something by request id; the page answers on `replyChannel`.
function answer(channel, replyChannel, produce) {
  ipcRenderer.on(channel, async (_e, requestId) => {
    try {
      ipcRenderer.send(replyChannel, requestId, await produce(), null);
    } catch (err) {
      ipcRenderer.send(replyChannel, requestId, null, String(err?.message || err));
    }
  });
}

contextBridge.exposeInMainWorld('office', {
  setDirty: (dirty) => ipcRenderer.send('doc:dirty', dirty),
  action: (name) => ipcRenderer.send('menu:action', name),
  onLoad: (cb) => ipcRenderer.on('doc:load', (_e, payload) => cb(payload)),
  onSaved: (cb) => ipcRenderer.on('doc:saved', (_e, payload) => cb(payload)),
  onCommand: (cb) => ipcRenderer.on('editor:command', (_e, name) => cb(name)),
  // My Word: the document as .docx bytes, made by the page.
  onCollectDocx: (makeDocx) => answer('docx:collect', 'docx:bytes', makeDocx),
  // My Excel: the workbook as a plain model; the main process writes the .xlsx.
  onCollectModel: (makeModel) => answer('sheet:collect', 'sheet:model', makeModel),
  // My Excel: lay out the printable copy of the sheets before printing or PDF export.
  onPreparePrint: (prepare) => answer('print:prepare', 'print:ready', prepare),
});
