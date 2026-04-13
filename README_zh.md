# Doc2Docx 转换器

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-1.0.0-blue.svg)](https://github.com/antigravity/doc2docx)

**Doc2Docx** 是一款优质的、免安装的 Windows 便携应用，旨在将旧版的 `.doc` 文件高度还原地转换为现代的 `.docx` 格式。它拥有美观的磨砂玻璃风格（Glassmorphism）界面，无论是单文件还是批量转换，都能提供流畅的用户体验。

[English Documentation](README.md)

## ✨ 功能特性

- 🚀 **便携免安装**：无需安装，双击 `.exe` 即可运行。
- 📂 **批量转换**：支持选择多个文件或整个文件夹进行转换。
- 🖱️ **拖拽支持**：只需将文档拖入界面即可开始。
- 🔄 **自动重命名**：智能处理文件名冲突，防止数据覆盖。
- 🌐 **双语界面**：流畅的中英文切换，适配深色模式美学。
- ⚡ **高还原度**：通过 PowerShell 调用系统安装的 Microsoft Word 引擎，确保转换效果完美。

## 📋 系统要求

- **操作系统**：Windows 10 或 11。
- **环境依赖**：系统中必须安装有 **Microsoft Word**，因为转换器利用其 COM 引擎完成格式处理。

## 🚀 快速上手

1.  从 [Releases](https://github.com/antigravity/doc2docx/releases) 页面下载最新的 `Doc2Docx.exe`。
2.  **直接运行** 应用。
3.  **添加文件**：
    - 点击 **选择文件** 选取特定文档。
    - 点击 **选择文件夹** 添加目录下的所有 `.doc` 文件。
    - 也可以将文档直接**拖拽**到应用界面。
4.  点击 **开始转换** 即可。

## 🛠️ 开发者指南

### 前项准备

- [Node.js](https://nodejs.org/) (v16+)
- npm

### 安装

```bash
# 克隆仓库
git clone https://github.com/antigravity/doc2docx.git

# 安装依赖
npm install
```

### 开发环境

```bash
# 启动开发服务器
npm run dev
```

### 编译打包

```bash
# 构建便携式可执行文件
npm run build
```

打包后的文件将生成在 `dist-electron` 目录下。

## 📄 开源协议

本项目采用 MIT 协议开源 - 详情请参阅 [LICENSE](LICENSE) 文件。

---

由 **Antigravity** 精心制作 ❤️
