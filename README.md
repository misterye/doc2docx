# Doc2Docx Converter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/antigravity/doc2docx)

**Doc2Docx** is a premium, portable Windows application designed to convert legacy `.doc` files to the modern `.docx` format with high fidelity. Built with a stunning Glassmorphism UI, it provides a seamless user experience for individual or batch conversions.

[中文文档 (Chinese)](README_zh.md)

## ✨ Features

- 🚀 **Portable & Zero-Install**: Run it directly from the `.exe` without any installation.
- 📂 **Batch Conversion**: Support for selecting multiple files or entire folders.
- 🖱️ **Drag & Drop**: Simply drop your documents into the interface to start.
- 🔄 **Auto-Renaming**: Intelligent handling of file name conflicts to prevent data loss.
- 🌐 **Clean UI**: A modern, bilingual (English/Chinese) interface with dark mode aesthetics.
- ⚡ **High Fidelity**: Uses Microsoft Word's own engine via PowerShell for pixel-perfect conversion.

## 📋 System Requirements

- **Operating System**: Windows 10 or 11.
- **Dependency**: **Microsoft Word** must be installed on your system, as the converter utilizes its COM engine for conversion.

## 🚀 Getting Started

1.  **Download** the latest `Doc2Docx.exe` from the [Releases](https://github.com/antigravity/doc2docx/releases) page.
2.  **Run** the application.
3.  **Add Files**:
    - Click **Select Files** to pick specific documents.
    - Click **Select Folder** to add all `.doc` files from a directory.
    - Drag and drop documents directly into the app.
4.  Click **Start Conversion** to begin the process.

## 🛠️ For Developers

### Prerequisites

- [Node.js](https://nodejs.org/) (v16+)
- npm

### Installation

```bash
# Clone the repository
git clone https://github.com/antigravity/doc2docx.git

# Install dependencies
npm install
```

### Development

```bash
# Run in development mode
npm run dev
```

### Build

```bash
# Build the portable executable
npm run build
```

The output will be located in the `dist-electron` folder.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

Created with ❤️ by **Antigravity**.
