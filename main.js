// Processo PRINCIPAL do Electron — é Node.js puro (sem DOM, sem HTML).
// É aqui que criamos a janela nativa e subimos o backend Express.
const { app, BrowserWindow, shell, ipcMain, dialog } = require('electron');
const path = require('path');
const { iniciarServidor } = require('./server.js');

const PORTA_DESEJADA = Number(process.env.PORT) || 3000;
const TEMPO_MINIMO_SPLASH_MS = 2200;

let janelaPrincipal;
let janelaSplash;

function criarSplash() {
  janelaSplash = new BrowserWindow({
    width: 380,
    height: 280,
    frame: false,
    resizable: false,
    movable: true,
    alwaysOnTop: true,
    backgroundColor: '#008080',
    icon: path.join(__dirname, 'public', 'icon.ico'),
    webPreferences: {
      nodeIntegration: false,
      contextIsolation: true
    }
  });

  janelaSplash.loadFile(path.join(__dirname, 'public', 'splash.html'));
  janelaSplash.center();
  janelaSplash.setMenuBarVisibility(false);
}

async function criarJanela() {
  criarSplash();
  const inicioSplash = Date.now();

  // Tenta subir o servidor na porta desejada; se estiver ocupada (ex: uma instância
  // antiga que ficou presa), tenta automaticamente as próximas (3001, 3002...) —
  // sem precisar caçar processo no Gerenciador de Tarefas toda vez.
  let porta;
  try {
    const resultado = await iniciarServidor(PORTA_DESEJADA);
    porta = resultado.porta;
  } catch (err) {
    if (janelaSplash) janelaSplash.close();
    dialog.showErrorBox(
      'Erro ao iniciar o servidor',
      `Não foi possível subir o servidor interno do app.\n\n${err.message}`
    );
    app.quit();
    return;
  }

  janelaPrincipal = new BrowserWindow({
    width: 900,
    height: 820,
    minWidth: 700,
    minHeight: 600,
    title: 'Migrador de Planilhas para SQL',
    backgroundColor: '#c0c0c0', // mesma cor de fundo do container do index.html
    icon: path.join(__dirname, 'public', 'icon.ico'), // ícone do app (barra de tarefas, Alt+Tab, atalho)
    frame: false, // remove a barra de título nativa do Windows — a barra Win98 do index.html vira a única
    show: false, // só aparece depois que a splash sumir, pra não "piscar" as duas juntas
    webPreferences: {
      nodeIntegration: false,  // a página carregada NÃO tem acesso direto ao Node (segurança)
      contextIsolation: true,
      preload: path.join(__dirname, 'preload.js')
    }
  });

  // Sem barra nativa, é a própria página que precisa poder minimizar/maximizar/fechar
  // a janela. O preload expõe window.electronAPI pra isso, e aqui a gente escuta os pedidos.
  ipcMain.on('janela-minimizar', () => janelaPrincipal.minimize());
  ipcMain.on('janela-maximizar', () => {
    if (janelaPrincipal.isMaximized()) janelaPrincipal.unmaximize();
    else janelaPrincipal.maximize();
  });
  ipcMain.on('janela-fechar', () => janelaPrincipal.close());

  janelaPrincipal.setMenuBarVisibility(false); // remove o menu padrão (Arquivo/Editar/Ver...)

  // Sem isso, clicar em um link com target="_blank" (como o do portfólio no rodapé)
  // abriria uma segunda janela do próprio Electron. Aqui a gente intercepta,
  // cancela essa janela e manda o link pro navegador padrão do usuário.
  janelaPrincipal.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url);
    return { action: 'deny' };
  });

  // Assim que o conteúdo estiver pronto pra mostrar, espera o tempo mínimo da
  // splash (pra não "piscar" rápido demais se o servidor subir muito rápido),
  // troca a splash pela janela principal.
  janelaPrincipal.once('ready-to-show', () => {
    const tempoDecorrido = Date.now() - inicioSplash;
    const espera = Math.max(0, TEMPO_MINIMO_SPLASH_MS - tempoDecorrido);

    setTimeout(() => {
      if (janelaSplash) {
        janelaSplash.close();
        janelaSplash = null;
      }
      janelaPrincipal.show();
    }, espera);
  });

  // Agora sim: o servidor já está confirmadamente escutando nesta porta.
  janelaPrincipal.loadURL(`http://localhost:${porta}`);

  // Descomente a linha abaixo durante o desenvolvimento pra abrir o DevTools automaticamente
  // janelaPrincipal.webContents.openDevTools();
}

app.whenReady().then(criarJanela);

// Comportamento padrão de app desktop: fecha o processo ao fechar a última janela
// (no Mac, por convenção, os apps costumam continuar "vivos" na dock — daí o check da plataforma)
app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (BrowserWindow.getAllWindows().length === 0) criarJanela();
});