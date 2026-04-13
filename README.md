# Office Legacy Converter

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/antigravity/office-legacy-converter)

**Office Legacy Converter** is a premium, portable Windows application designed to batch convert legacy Office files (.doc, .xls, .ppt) to their modern OpenXML equivalents (.docx, .xlsx, .pptx) with high fidelity. Built with a stunning Glassmorphism UI, it provides a seamless user experience for individual or batch conversions.

[中文文档 (Chinese)](README_zh.md)

## ✨ Features

- 🚀 **Portable & Zero-Install**: Run it directly from the `.exe` without any installation.
- 📂 **Multi-Format Support**: Batch convert Word (.doc), Excel (.xls), and PowerPoint (.ppt).
- 🖱️ **Drag & Drop**: Simply drop your documents into the interface to start.
- 🔄 **Concurrency & Speed**: Parallel processing (3 files at a time) for rapid results.
- 🛠️ **Robust Error Handling**: Automatic timeouts and skipping of problematic files.
- 🎨 **Type-Specific Themes**: Visual distinction with icons and colors for different Office apps.
- ⚡ **High Fidelity**: Uses local Microsoft Office engines for pixel-perfect conversion.

## 📋 System Requirements

- **Operating System**: Windows 10 or 11.
- **Dependencies**: **Microsoft Office** (Word, Excel, PowerPoint) must be installed for respective formats.

## 🚀 Getting Started

1.  **Download** the latest `Office Legacy Converter 1.0.0.exe` from the [Releases](https://github.com/antigravity/office-legacy-converter/releases) page.
2.  **Run** the application.
3.  **Add Files**:
    - Click **Select Files** to pick specific documents.
    - Click **Select Folder** to add all legacy Office files from a directory.
    - Drag and drop documents directly into the app.
4.  Click **Start Conversion** to begin the process.

## 🛠️ For Developers

### Prerequisites

- [Node.js](https://nodejs.org/) (v16+)
- npm

### Installation

```bash
# Clone the repository
git clone https://github.com/antigravity/office-legacy-converter.git

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
