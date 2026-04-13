// Internationalization
const translations = {
  zh: {
    title: 'Office Legacy Converter',
    dropText: '将 .doc, .xls 或 .ppt 文件拖到此处',
    orText: '或',
    selectFiles: '选择文件',
    selectFolder: '选择文件夹',
    queueTitle: '准备转换的文件',
    clearBtn: '清空列表',
    statusReady: '就绪',
    statusConverting: '正在转换...',
    statusDone: '转换完成',
    statusStopped: '已停止',
    startBtn: '开始转换',
    stopBtn: '停止转换',
    langBtn: 'English',
    itemWaiting: '等待中',
    itemConverting: '转换中',
    itemSuccess: '成功',
    itemError: '失败',
    itemTimeout: '超时',
    itemRetry: '重试',
    retryAll: '重试所有失败项',
    itemStopped: '已停止',
    noWordError: '请确保系统中已安装相应的 Office 软件（Word/Excel/PPT）。',
  },
  en: {
    title: 'Office Legacy Converter',
    dropText: 'Drop .doc, .xls or .ppt files here',
    orText: 'or',
    selectFiles: 'Select Files',
    selectFolder: 'Select Folder',
    queueTitle: 'Files in Queue',
    clearBtn: 'Clear List',
    statusReady: 'Ready',
    statusConverting: 'Converting...',
    statusDone: 'All Done',
    statusStopped: 'Stopped',
    startBtn: 'Start Conversion',
    stopBtn: 'Stop All',
    langBtn: '中文',
    itemWaiting: 'Waiting',
    itemConverting: 'Converting',
    itemSuccess: 'Success',
    itemError: 'Error',
    itemTimeout: 'Timeout',
    itemRetry: 'Retry',
    retryAll: 'Retry All Failed',
    itemStopped: 'Stopped',
    noWordError: 'Please ensure Microsoft Word/Excel/PowerPoint is installed.',
  }
};

let currentLang = 'zh';
let fileQueue = [];
let isConverting = false;
let stopRequested = false;
let elements = {};

const MAX_CONCURRENT = 3;

class QueueManager {
  constructor() {
    this.activeTasks = 0;
    this.currentIndex = 0;
  }

  async start() {
    stopRequested = false;
    isConverting = true;
    this.activeTasks = 0;
    this.currentIndex = 0;
    this.process();
  }

  async process() {
    if (stopRequested) {
      this.checkFinished();
      return;
    }

    // Fill up to MAX_CONCURRENT
    while (this.activeTasks < MAX_CONCURRENT && this.currentIndex < fileQueue.length) {
      if (stopRequested) break;
      
      const file = fileQueue[this.currentIndex];
      const index = this.currentIndex;
      this.currentIndex++;

      if (file.status === 'success') continue;

      this.activeTasks++;
      this.runTask(file, index);
    }

    this.checkFinished();
  }

  async runTask(file, index) {
    file.status = 'converting';
    file.taskId = `task-${Date.now()}-${index}`;
    renderList();

    try {
      const result = await window.electronAPI.convertDoc({ 
        filePath: file.path, 
        taskId: file.taskId 
      });

      if (result.success) {
        // Remove from queue on success
        fileQueue = fileQueue.filter(f => f.path !== file.path);
      } else {
        if (result.error === 'Timeout') {
          file.status = 'timeout';
        } else {
          file.status = 'error';
        }
      }
    } catch (err) {
      console.error('Task execution error:', err);
      file.status = 'error';
    }

    this.activeTasks--;
    renderList();
    this.process(); // Try to start next one
  }

  checkFinished() {
    if (this.activeTasks === 0 && (this.currentIndex >= fileQueue.length || stopRequested)) {
      isConverting = false;
      onConversionFinished();
    }
  }
}

const queueManager = new QueueManager();

// Wait for DOM to be ready
document.addEventListener('DOMContentLoaded', () => {
  initElements();
  initListeners();
  updateUI();
});

function initElements() {
  const ids = [
    'title', 'drop-text', 'or-text', 'select-files-btn', 'select-folder-btn',
    'queue-title', 'clear-btn', 'status-text', 'start-btn', 'stop-btn', 'lang-btn',
    'file-list', 'drop-zone', 'retry-all-btn'
  ];
  
  ids.forEach(id => {
    elements[id.replace(/-([a-z])/g, (g) => g[1].toUpperCase()).replace('Btn', 'Btn')] = document.getElementById(id);
  });
}

function updateUI() {
  const t = translations[currentLang];
  Object.keys(t).forEach(key => {
    if (elements[key]) elements[key].textContent = t[key];
  });
  
  const hasFailed = fileQueue.some(f => f.status === 'error' || f.status === 'timeout');
  
  if (!isConverting) {
    elements.statusText.textContent = stopRequested ? t.statusStopped : t.statusReady;
    elements.startBtn.disabled = fileQueue.length === 0;
    elements.stopBtn.disabled = true;
    elements.clearBtn.disabled = false;
    elements.retryAllBtn.style.display = hasFailed ? 'block' : 'none';
  } else {
    elements.statusText.textContent = t.statusConverting;
    elements.startBtn.disabled = true;
    elements.stopBtn.disabled = false;
    elements.clearBtn.disabled = true;
    elements.retryAllBtn.style.display = 'none';
  }
  
  renderList();
}

function renderList() {
  if (!elements.fileList) return;
  
  const t = translations[currentLang];
  elements.fileList.innerHTML = '';
  fileQueue.forEach((file, index) => {
    const li = document.createElement('li');
    
    let statusText = t.itemWaiting;
    let statusClass = 'status-waiting';
    let typeClass = 'type-word';
    let icon = '📘';

    const ext = file.path.toLowerCase().split('.').pop();
    if (ext === 'doc') {
      typeClass = 'type-word';
      icon = '📘';
    } else if (ext === 'xls') {
      typeClass = 'type-excel';
      icon = '📗';
    } else if (ext === 'ppt') {
      typeClass = 'type-ppt';
      icon = '📙';
    }

    li.className = `file-item ${typeClass}`;
    
    switch(file.status) {
      case 'converting': statusText = t.itemConverting; statusClass = 'status-converting'; break;
      case 'success': statusText = t.itemSuccess; statusClass = 'status-success'; break;
      case 'error': statusText = t.itemError; statusClass = 'status-error'; break;
      case 'timeout': statusText = t.itemTimeout; statusClass = 'status-timeout'; break;
      case 'stopped': statusText = t.itemStopped; statusClass = 'status-stopped'; break;
    }

    li.innerHTML = `
      <div class="file-info">
        <span class="file-icon">${icon}</span>
        <span class="file-name" title="${file.path}">${file.name}</span>
      </div>
      <div class="file-actions">
        <span class="file-status ${statusClass}">${statusText}</span>
        ${(file.status === 'error' || file.status === 'timeout') ? 
          `<button class="retry-btn" onclick="retryTask(${index})" title="${t.itemRetry}">🔄</button>` : ''}
        ${(file.status === 'waiting' || file.status === 'converting') ? 
          `<button class="cancel-btn" onclick="cancelTask(${index})" title="Cancel">✕</button>` : ''}
      </div>
    `;
    elements.fileList.appendChild(li);
  });
}

window.retryTask = (index) => {
  const file = fileQueue[index];
  if (file) {
    file.status = 'waiting';
    if (!isConverting) {
      queueManager.start();
    }
    updateUI();
  }
};

window.retryAllFailed = () => {
  fileQueue.forEach(f => {
    if (f.status === 'error' || f.status === 'timeout') {
      f.status = 'waiting';
    }
  });
  if (!isConverting) {
    queueManager.start();
  }
  updateUI();
};

window.cancelTask = async (index) => {
  const file = fileQueue[index];
  if (file.status === 'converting' && file.taskId) {
    await window.electronAPI.cancelConversion(file.taskId);
    file.status = 'stopped';
  } else if (file.status === 'waiting') {
    file.status = 'stopped';
  }
  renderList();
};

function onConversionFinished() {
  updateUI();
}

function initListeners() {
  elements.langBtn.addEventListener('click', () => {
    currentLang = currentLang === 'zh' ? 'en' : 'zh';
    updateUI();
  });

  elements.selectFilesBtn.addEventListener('click', async () => {
    const paths = await window.electronAPI.selectFiles();
    if (paths) addFiles(paths);
  });

  elements.selectFolderBtn.addEventListener('click', async () => {
    const paths = await window.electronAPI.selectFolder();
    if (paths) addFiles(paths);
  });

  elements.clearBtn.addEventListener('click', () => {
    fileQueue = [];
    updateUI();
  });

  elements.startBtn.addEventListener('click', () => {
    queueManager.start();
    updateUI();
  });

  elements.retryAllBtn.addEventListener('click', () => {
    retryAllFailed();
  });

  elements.stopBtn.addEventListener('click', () => {
    stopRequested = true;
    // Mark all waiting as stopped
    fileQueue.forEach(f => {
      if (f.status === 'waiting') f.status = 'stopped';
    });
    // For converting ones, we just let them finish or user can cancel individually,
    // but the queue won't start new ones.
    updateUI();
  });

  elements.dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    elements.dropZone.classList.add('dragover');
  });

  elements.dropZone.addEventListener('dragleave', () => {
    elements.dropZone.classList.remove('dragover');
  });

  elements.dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    elements.dropZone.classList.remove('dragover');
    if (e.dataTransfer && e.dataTransfer.files) {
      const allowedExts = ['.doc', '.xls', '.ppt'];
      const files = Array.from(e.dataTransfer.files)
        .filter(f => allowedExts.includes(f.name.toLowerCase().slice(f.name.lastIndexOf('.'))))
        .map(f => f.path);
      if (files.length > 0) addFiles(files);
    }
  });
}

function addFiles(paths) {
  paths.forEach(path => {
    if (fileQueue.some(f => f.path === path)) return;
    fileQueue.push({
      path,
      name: path.split(/[\\/]/).pop(),
      status: 'waiting'
    });
  });
  updateUI();
}
