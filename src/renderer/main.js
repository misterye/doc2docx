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

// DOM Elements
const elements = {
  title: document.getElementById('title'),
  dropText: document.getElementById('drop-text'),
  orText: document.getElementById('or-text'),
  selectFilesBtn: document.getElementById('select-files-btn'),
  selectFolderBtn: document.getElementById('select-folder-btn'),
  queueTitle: document.getElementById('queue-title'),
  clearBtn: document.getElementById('clear-btn'),
  statusText: document.getElementById('status-text'),
  startBtn: document.getElementById('start-btn'),
  langBtn: document.getElementById('lang-btn'),
  fileList: document.getElementById('file-list'),
  dropZone: document.getElementById('drop-zone'),
};

function updateUI() {
  const t = translations[currentLang];
  elements.title.textContent = t.title;
  elements.dropText.textContent = t.dropText;
  elements.orText.textContent = t.orText;
  elements.selectFilesBtn.textContent = t.selectFiles;
  elements.selectFolderBtn.textContent = t.selectFolder;
  elements.queueTitle.textContent = t.queueTitle;
  elements.clearBtn.textContent = t.clearBtn;
  elements.langBtn.textContent = t.langBtn;
  elements.startBtn.textContent = t.startBtn;
  
  if (!isConverting) {
    elements.statusText.textContent = t.statusReady;
  }
  
  // Update existing list items
  document.querySelectorAll('.file-item').forEach((item, index) => {
    const statusSpan = item.querySelector('.file-status');
    const status = fileQueue[index].status;
    if (status === 'waiting') statusSpan.textContent = t.itemWaiting;
    else if (status === 'converting') statusSpan.textContent = t.itemConverting;
    else if (status === 'success') statusSpan.textContent = t.itemSuccess;
    else if (status === 'error') statusSpan.textContent = t.itemError;
  });
}

function addFiles(paths) {
  paths.forEach(path => {
    if (fileQueue.some(f => f.path === path)) return; // Prevent duplicates
    fileQueue.push({
      path,
      name: path.split(/[\\/]/).pop(),
      status: 'waiting'
    });
  });
  renderList();
  elements.startBtn.disabled = fileQueue.length === 0 || isConverting;
}

function renderList() {
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

// Event Listeners
elements.langBtn.onclick = () => {
  currentLang = currentLang === 'zh' ? 'en' : 'zh';
  updateUI();
};

elements.selectFilesBtn.onclick = async () => {
  const paths = await window.electronAPI.selectFiles();
  if (paths) addFiles(paths);
};

elements.selectFolderBtn.onclick = async () => {
  const paths = await window.electronAPI.selectFolder();
  if (paths) addFiles(paths);
};

elements.clearBtn.onclick = () => {
  if (isConverting) return;
  fileQueue = [];
  renderList();
  elements.startBtn.disabled = true;
};

elements.startBtn.onclick = async () => {
  if (isConverting || fileQueue.length === 0) return;
  
  isConverting = true;
  elements.startBtn.disabled = true;
  elements.clearBtn.disabled = true;
  elements.statusText.textContent = translations[currentLang].statusConverting;

  for (let i = 0; i < fileQueue.length; i++) {
    if (fileQueue[i].status === 'success') continue;
    
    fileQueue[i].status = 'converting';
    renderList();
    
    const result = await window.electronAPI.convertDoc(fileQueue[i].path);
    
    if (result.success) {
      fileQueue[i].status = 'success';
    } else {
      fileQueue[i].status = 'error';
      // If conversion fails, it might be Word missing
      elements.statusText.textContent = translations[currentLang].noWordError;
    }
    renderList();
  }

  isConverting = false;
  elements.clearBtn.disabled = false;
  elements.startBtn.disabled = false;
  elements.statusText.textContent = translations[currentLang].statusDone;
};

// Drag & Drop
elements.dropZone.ondragover = (e) => {
  e.preventDefault();
  elements.dropZone.classList.add('dragover');
};

elements.dropZone.ondragleave = () => {
  elements.dropZone.classList.remove('dragover');
};

elements.dropZone.ondrop = (e) => {
  e.preventDefault();
  elements.dropZone.classList.remove('dragover');
  const files = Array.from(e.dataTransfer.files)
    .filter(f => f.name.toLowerCase().endsWith('.doc'))
    .map(f => f.path);
  if (files.length > 0) addFiles(files);
};

// Init
updateUI();
