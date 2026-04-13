const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const { spawn } = require('node:child_process');

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    title: 'Doc2Docx Converter',
    frame: true,
    webPreferences: {
      preload: path.join(__dirname, '../preload/index.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
    // Premium Look
    backgroundColor: '#0f172a',
  });

  // In development, load from Vite dev server. In production, load from the built dist.
  const isDev = !app.isPackaged;
  if (isDev) {
    mainWindow.loadURL('http://localhost:5173');
  } else {
    mainWindow.loadFile(path.join(__dirname, '../../dist/index.html'));
  }

  // mainWindow.webContents.openDevTools();
}

app.whenReady().then(() => {
  createWindow();

  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow();
  });
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

// IPC Handlers
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [{ name: 'Word 97-2003 Documents', extensions: ['doc'] }],
  });
  return result.filePaths;
});

ipcMain.handle('select-folder', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openDirectory'],
  });
  if (result.canceled) return [];
  
  const folderPath = result.filePaths[0];
  const files = fs.readdirSync(folderPath);
  return files
    .filter(file => file.toLowerCase().endsWith('.doc'))
    .map(file => path.join(folderPath, file));
});

ipcMain.handle('convert-doc', async (event, filePath) => {
  return new Promise((resolve) => {
    const dir = path.dirname(filePath);
    const ext = path.extname(filePath);
    const base = path.basename(filePath, ext);
    
    let destPath = path.join(dir, `${base}.docx`);
    
    // Auto rename if exists
    let counter = 1;
    while (fs.existsSync(destPath)) {
      destPath = path.join(dir, `${base} (${counter}).docx`);
      counter++;
    }

    // PowerShell script to convert using Word COM
    const psScript = `
      try {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $doc = $word.Documents.Open("${filePath}")
        $doc.SaveAs([ref]"${destPath}", 16) # 16 = wdFormatXMLDocument
        $doc.Close()
        $word.Quit()
        exit 0
      } catch {
        exit 1
      }
    `;

    const ps = spawn('powershell.exe', ['-Command', psScript]);

    ps.on('close', (code) => {
      if (code === 0) {
        resolve({ success: true, path: destPath });
      } else {
        resolve({ success: false, error: 'Conversion failed. Make sure MS Word is installed.' });
      }
    });
  });
});
