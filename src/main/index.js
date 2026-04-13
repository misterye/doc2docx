const { app, BrowserWindow, ipcMain, dialog } = require('electron');
const path = require('node:path');
const fs = require('node:fs');
const { spawn } = require('node:child_process');

function createWindow() {
  const mainWindow = new BrowserWindow({
    width: 900,
    height: 700,
    title: 'Office Legacy Converter',
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

// Map to track active powershell processes by taskId
const activeProcesses = new Map();

// IPC Handlers
ipcMain.handle('select-files', async () => {
  const result = await dialog.showOpenDialog({
    properties: ['openFile', 'multiSelections'],
    filters: [
      { name: 'Office Documents', extensions: ['doc', 'xls', 'ppt'] },
      { name: 'Word 97-2003 Documents', extensions: ['doc'] },
      { name: 'Excel 97-2003 Workbooks', extensions: ['xls'] },
      { name: 'PowerPoint 97-2003 Presentations', extensions: ['ppt'] }
    ],
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
  const extensions = ['.doc', '.xls', '.ppt'];
  return files
    .filter(file => extensions.includes(path.extname(file).toLowerCase()))
    .map(file => path.join(folderPath, file));
});

ipcMain.handle('convert-doc', async (event, { filePath, taskId }) => {
  return new Promise((resolve) => {
    const dir = path.dirname(filePath);
    const ext = path.extname(filePath).toLowerCase();
    const base = path.basename(filePath, ext);
    
    let destExt = '';
    let comObject = '';
    let formatCode = 0;

    if (ext === '.doc') {
      destExt = '.docx';
      comObject = 'Word.Application';
      formatCode = 16; // wdFormatXMLDocument
    } else if (ext === '.xls') {
      destExt = '.xlsx';
      comObject = 'Excel.Application';
      formatCode = 51; // xlOpenXMLWorkbook
    } else if (ext === '.ppt') {
      destExt = '.pptx';
      comObject = 'PowerPoint.Application';
      formatCode = 24; // ppSaveAsOpenXMLPresentation
    } else {
      return resolve({ success: false, error: 'Unsupported file format' });
    }

    let destPath = path.join(dir, `${base}${destExt}`);
    
    // Auto rename if exists
    let counter = 1;
    while (fs.existsSync(destPath)) {
      destPath = path.join(dir, `${base} (${counter})${destExt}`);
      counter++;
    }

    // Dynamic timeout calculation
    const fileSize = fs.statSync(filePath).size;
    const timeoutMs = Math.max(30000, 30000 + (fileSize / (1024 * 1024)) * 10000);

    // PowerShell script logic branched by COM object
    let psScript = '';
    if (ext === '.doc') {
      psScript = `
        $ErrorActionPreference = "Stop"
        try {
          $app = New-Object -ComObject Word.Application
          $app.Visible = $false
          $doc = $app.Documents.Open("${filePath}")
          $doc.SaveAs([ref]"${destPath}", ${formatCode})
          $doc.Close()
          $app.Quit()
          exit 0
        } catch {
          if ($app) { $app.Quit() }
          exit 1
        }
      `;
    } else if (ext === '.xls') {
      psScript = `
        $ErrorActionPreference = "Stop"
        try {
          $app = New-Object -ComObject Excel.Application
          $app.Visible = $false
          $app.DisplayAlerts = $false
          $wb = $app.Workbooks.Open("${filePath}")
          $wb.SaveAs("${destPath}", ${formatCode})
          $wb.Close()
          $app.Quit()
          exit 0
        } catch {
          if ($app) { $app.Quit() }
          exit 1
        }
      `;
    } else if (ext === '.ppt') {
      psScript = `
        $ErrorActionPreference = "Stop"
        try {
          $app = New-Object -ComObject PowerPoint.Application
          # PowerPoint needs a window but we can try to hide it
          $pres = $app.Presentations.Open("${filePath}", [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse, [Microsoft.Office.Core.MsoTriState]::msoFalse)
          $pres.SaveAs("${destPath}", ${formatCode})
          $pres.Close()
          $app.Quit()
          exit 0
        } catch {
          if ($app) { $app.Quit() }
          exit 1
        }
      `;
    }

    const ps = spawn('powershell.exe', ['-Command', psScript]);
    activeProcesses.set(taskId, ps);

    const timeout = setTimeout(() => {
      if (activeProcesses.has(taskId)) {
        ps.kill();
        activeProcesses.delete(taskId);
        resolve({ success: false, error: 'Timeout' });
      }
    }, timeoutMs);

    ps.on('close', (code) => {
      clearTimeout(timeout);
      activeProcesses.delete(taskId);
      if (code === 0) {
        resolve({ success: true, path: destPath });
      } else {
        resolve({ success: false, error: 'Conversion failed or was cancelled.' });
      }
    });

    ps.on('error', (err) => {
      clearTimeout(timeout);
      activeProcesses.delete(taskId);
      resolve({ success: false, error: err.message });
    });
  });
});

ipcMain.handle('cancel-conversion', (event, taskId) => {
  const ps = activeProcesses.get(taskId);
  if (ps) {
    ps.kill();
    activeProcesses.delete(taskId);
    return true;
  }
  return false;
});
