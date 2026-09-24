// Preload roda ANTES da página carregar, com um pé em cada mundo (Node + navegador),
// mas de forma isolada. Como contextIsolation está ativado, a página (index.html)
// não tem acesso direto ao Node — só ao que a gente expor explicitamente aqui.
// Isso evita que uma planilha maliciosa ou um bug no front consiga rodar código Node.
const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('electronAPI', {
  minimizar: () => ipcRenderer.send('janela-minimizar'),
  maximizar: () => ipcRenderer.send('janela-maximizar'),
  fechar: () => ipcRenderer.send('janela-fechar')
});