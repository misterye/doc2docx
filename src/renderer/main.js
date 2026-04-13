// Internationalization
const translations = {
  zh: {
    title: 'Doc to Docx 转换器',
    dropText: '将 .doc 文件拖到此处',
    orText: '或',
    selectFiles: '选择文件',
    selectFolder: '选择文件夹',
    queueTitle: '准备转换的文件',
    clearBtn: '清空列表',
    statusReady: '就绪',
    statusConverting: '正在转换...',
    statusDone: '转换完成',
    startBtn: '开始转换',
    langBtn: 'English',
    itemWaiting: '等待中',
    itemConverting: '转换中',
    itemSuccess: '成功',
    itemError: '失败',
    noWordError: '请确保系统中已安装 Microsoft Word。',
  },
  en: {
    title: 'Doc to Docx Converter',
    dropText: 'Drop .doc files here',
    orText: 'or',
    selectFiles: 'Select Files',
    selectFolder: 'Select Folder',
    queueTitle: 'Files in Queue',
    clearBtn: 'Clear List',
    statusReady: 'Ready',
    statusConverting: 'Converting...',
    statusDone: 'All Done',
    startBtn: 'Start Conversion',
    langBtn: '中文',
    itemWaiting: 'Waiting',
    itemConverting: 'Converting',
    itemSuccess: 'Success',
    itemError: 'Error',
    noWordError: 'Please ensure Microsoft Word is installed.',
  }
};

let currentLang = 'zh';
let fileQueue = [];
let isConverting = false;
let elements = {};

// Wait for DOM to be ready
document.addEventListener('DOMContentLoaded', () => {
  console.log('DOM Content Loaded. Initializing app...');
  initElements();
  initListeners();
  updateUI();
});

function initElements() {
  const ids = [
    'title', 'drop-text', 'or-text', 'select-files-btn', 'select-folder-btn',
    'queue-title', 'clear-btn', 'status-text', 'start-btn', 'lang-btn',
    'file-list', 'drop-zone'
  ];
  
  ids.forEach(id => {
    const el = document.getElementById(id);
    if (!el) {
      console.warn(`Element with ID "${id}" not found!`);
    }
    // Mapping IDs to camelCase names
    const camelId = id.replace(/-([a-z])/g, (g) => g[1].toUpperCase()).replace('Btn', 'Btn');
    elements[camelId] = el;
  });
  
  // Explicitly ensure critical ones are mapped
  elements.selectFilesBtn = document.getElementById('select-files-btn');
  elements.selectFolderBtn = document.getElementById('select-folder-btn');
  elements.clearBtn = document.getElementById('clear-btn');
  elements.startBtn = document.getElementById('start-btn');
  elements.langBtn = document.getElementById('lang-btn');
  elements.statusText = document.getElementById('status-text');
  elements.dropZone = document.getElementById('drop-zone');
  elements.fileList = document.getElementById('file-list');
}

function updateUI() {
  const t = translations[currentLang];
  if (elements.title) elements.title.textContent = t.title;
  if (elements.dropText) elements.dropText.textContent = t.dropText;
  if (elements.orText) elements.orText.textContent = t.orText;
  if (elements.selectFilesBtn) elements.selectFilesBtn.textContent = t.selectFiles;
  if (elements.selectFolderBtn) elements.selectFolderBtn.textContent = t.selectFolder;
  if (elements.queueTitle) elements.queueTitle.textContent = t.queueTitle;
  if (elements.clearBtn) elements.clearBtn.textContent = t.clearBtn;
  if (elements.langBtn) elements.langBtn.textContent = t.langBtn;
  if (elements.startBtn) elements.startBtn.textContent = t.startBtn;
  
  if (!isConverting && elements.statusText) {
    elements.statusText.textContent = t.statusReady;
  }
  
  renderList();
}

function addFiles(paths) {
  if (!paths || !Array.isArray(paths)) return;
  
  paths.forEach(path => {
    if (fileQueue.some(f => f.path === path)) return;
    fileQueue.push({
      path,
      name: path.split(/[\\/]/).pop(),
      status: 'waiting'
    });
  });
  renderList();
  if (elements.startBtn) {
    elements.startBtn.disabled = fileQueue.length === 0 || isConverting;
  }
}

function renderList() {
  if (!elements.fileList) return;
  
  const t = translations[currentLang];
  elements.fileList.innerHTML = '';
  fileQueue.forEach((file, index) => {
    const li = document.createElement('li');
    li.className = 'file-item';
    
    let statusText = t.itemWaiting;
    let statusClass = 'status-waiting';
    if (file.status === 'converting') { statusText = t.itemConverting; statusClass = 'status-converting'; }
    else if (file.status === 'success') { statusText = t.itemSuccess; statusClass = 'status-success'; }
    else if (file.status === 'error') { statusText = t.itemError; statusClass = 'status-error'; }

    li.innerHTML = `
      <div class="file-info">
        <span class="file-icon">📄</span>
        <span class="file-name" title="${file.path}">${file.name}</span>
      </div>
      <span class="file-status ${statusClass}">${statusText}</span>
    `;
    elements.fileList.appendChild(li);
  });
}

function initListeners() {
  console.log('Initializing listeners...');
  
  if (elements.langBtn) {
    elements.langBtn.addEventListener('click', () => {
      console.log('Language button clicked');
      currentLang = currentLang === 'zh' ? 'en' : 'zh';
      updateUI();
    });
  }

  if (elements.selectFilesBtn) {
    elements.selectFilesBtn.addEventListener('click', async () => {
      console.log('Select files clicked');
      try {
        if (window.electronAPI && window.electronAPI.selectFiles) {
          const paths = await window.electronAPI.selectFiles();
          if (paths) addFiles(paths);
        } else {
          console.error('electronAPI.selectFiles is missing');
          alert('System error: IPC bridge missing.');
        }
      } catch (err) {
        console.error('IPC Error (selectFiles):', err);
      }
    });
  }

  if (elements.selectFolderBtn) {
    elements.selectFolderBtn.addEventListener('click', async () => {
      console.log('Select folder clicked');
      try {
        if (window.electronAPI && window.electronAPI.selectFolder) {
          const paths = await window.electronAPI.selectFolder();
          if (paths) addFiles(paths);
        } else {
          console.error('electronAPI.selectFolder is missing');
        }
      } catch (err) {
        console.error('IPC Error (selectFolder):', err);
      }
    });
  }

  if (elements.clearBtn) {
    elements.clearBtn.addEventListener('click', () => {
      if (isConverting) return;
      fileQueue = [];
      renderList();
      if (elements.startBtn) elements.startBtn.disabled = true;
    });
  }

  if (elements.startBtn) {
    elements.startBtn.addEventListener('click', async () => {
      if (isConverting || fileQueue.length === 0) return;
      
      isConverting = true;
      elements.startBtn.disabled = true;
      if (elements.clearBtn) elements.clearBtn.disabled = true;
      if (elements.statusText) elements.statusText.textContent = translations[currentLang].statusConverting;

      for (let i = 0; i < fileQueue.length; i++) {
        if (fileQueue[i].status === 'success') continue;
        
        fileQueue[i].status = 'converting';
        renderList();
        
        try {
          const result = await window.electronAPI.convertDoc(fileQueue[i].path);
          if (result.success) {
            fileQueue[i].status = 'success';
          } else {
            fileQueue[i].status = 'error';
            if (elements.statusText) elements.statusText.textContent = translations[currentLang].noWordError;
          }
        } catch (err) {
          console.error('Conversion IPC Error:', err);
          fileQueue[i].status = 'error';
        }
        renderList();
      }

      isConverting = false;
      if (elements.clearBtn) elements.clearBtn.disabled = false;
      elements.startBtn.disabled = false;
      if (elements.statusText) elements.statusText.textContent = translations[currentLang].statusDone;
    });
  }

  // Drag & Drop
  if (elements.dropZone) {
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
        const files = Array.from(e.dataTransfer.files)
          .filter(f => f.name.toLowerCase().endsWith('.doc'))
          .map(f => f.path);
        if (files.length > 0) addFiles(files);
      }
    });
  }
}
