# Office Legacy Converter (Office 旧版文档转换器)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/antigravity/office-legacy-converter)

**Office Legacy Converter** 是一款优质、免安装的 Windows 便携应用，支持批量将旧版 Office 文件（.doc, .xls, .ppt）高度还原地转换为现代 OpenXML 格式（.docx, .xlsx, .pptx）。它拥有美观的磨砂玻璃风格界面，并支持并发处理以提升转换效率。

[English Documentation](README.md)

## ✨ 功能特性

- 🚀 **便携免安装**：无需安装，直接运行 `.exe` 即可使用。
- 📂 **多格式支持**：批量处理 Word (.doc)、Excel (.xls) 和 PowerPoint (.ppt)。
- 🖱️ **拖拽支持**：只需将文档拖入界面即可开始。
- 🔄 **并发加速**：支持多线程并行转换（默认同时处理 3 个文件）。
- 🛠️ **稳健容错**：自动超时跳过损坏或响应过慢的文件。
- 🎨 **分类主题**：为不同 Office 应用提供专属图标和配色标识。
- ⚡ **高还原度**：通过调用本地 Office 引擎确保转换效果完美。

## 📋 系统要求

- **操作系统**：Windows 10 或 11。
- **环境依赖**：系统中需安装有对应的 **Microsoft Office** 软件。

## 🚀 快速上手

1.  从 [Releases](https://github.com/antigravity/office-legacy-converter/releases) 页面下载最新的 `Office Legacy Converter 1.0.0.exe`。
2.  **直接运行** 应用。
3.  **添加文件**：
    - 点击 **选择文件** 选取特定文档。
    - 点击 **选择文件夹** 添加目录下的所有旧版文档。
    - 也可以将文档直接**拖拽**到应用界面。
4.  点击 **开始转换** 即可。

## 🛠️ For Developers

### Prerequisites

- [Node.js](https://nodejs.org/) (v16+)
- npm

### Installation

```bash
# 克隆仓库
git clone https://github.com/antigravity/office-legacy-converter.git

# 安装依赖
npm install
```

### Development

```bash
# 启动开发服务器
npm run dev
```

### Build

```bash
# 构建便携式可执行文件
npm run build
```

打包后的文件将生成在 `dist-electron` 目录下。

## 📄 开源协议

本项目采用 MIT 协议开源 - 详情请参阅 [LICENSE](LICENSE) 文件。

---

由 **Antigravity** 精心制作 ❤️
