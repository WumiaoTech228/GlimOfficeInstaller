# Glim Office Installer (GOI)

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Dotnet Framework](https://img.shields.io/badge/.NET%20Framework-4.8-blue.svg)](https://dotnet.microsoft.com/)

一个优雅、现代且符合 Fluent 2.0 设计规范的 Windows 平台多功能 Office 套件极简安装工具。

An elegant, modern, and Fluent 2.0 styled minimalist Office suite deployment utility for Windows.

---

## 💡 项目简介 / Introduction

**Glim Office Installer (GOI)** 旨在为普通用户和 IT 管理员提供一站式的 Office 办公软件部署解决方案。通过极其精致的现代化 UI（基于 Fluent 2.0 设计标准，全面契合 Windows 11 视觉风格），你可以轻松下载、安装、关联并激活各类主流的 Office 办公套件，告别繁琐的命令行和古板的传统安装器。

**GOI** aims to provide a one-stop deployment solution for Office suites. Built with high-fidelity Fluent 2.0 components, it allows you to download, install, associate, and activate various Office packages with a single click, replacing complex commands and outdated legacy installers.

---

## ✨ 核心特性 / Features

* **📦 支持 5 大 Office 办公套件 / 5 Supported Office Suites**:
  - **Microsoft Office** (包括 2024 / Microsoft 365 / 2021 / 2019 / 2016 部署与更新通道自定义)
  - **WPS Office** (最新官方版与历史长期稳定版)
  - **永中 Office 2024 (Yozo Office)**
  - **OnlyOffice Desktop Editors**
  - **LibreOffice** (开源稳定版)
* **🎨 极致的 Fluent 2.0 视觉设计 / Fluent 2.0 Aesthetics**:
  - 完全原生契合 Windows 11 深色模式与浅色模式。
  - 精心定制的版本选择卡片（`VersionCard`），包含微动悬停动画与主题高亮边框。
  - 支持多语种运行时无缝秒切（简体中文、繁体中文、英文），UI 组件就地动态刷新，绝无卡顿与闪烁。
* **🛠️ 强类型状态机驱动 / Strongly-Typed Progress State Machine**:
  - 界面安装状态（准备就绪、下载中、安装中、激活中）全部由底层的强类型枚举回调驱动，杜绝落后的“文本嗅探”机制，过程稳定可靠。
* **🏃 单文件独立运行 / Self-Contained Executable**:
  - 依靠 Costura.Fody 将所有依赖项、字体、甚至激活脚本（Ohook）无损打包进单个 `GOI.exe` 文件中，无任何零碎的外部依赖，即开即用。

---

## 🖥️ 系统要求 / System Requirements

* **操作系统**: Windows 7 SP1, Windows 8, Windows 10, Windows 11 (支持 x86/x64 架构)
* **运行环境**: .NET Framework 4.8 或更高版本
* **权限要求**: 需要管理员权限（部分操作涉及修改注册表与系统服务）

---

## 🛠️ 构建与编译 / How to Build

项目支持通过 .NET SDK 或 Visual Studio 编译。

1. 克隆本仓库到本地：
   ```bash
   git clone https://github.com/WumiaoTech228/GlimOfficeInstaller.git
   cd GlimOfficeInstaller
   ```
2. 使用命令行进行 Release 构建：
   ```bash
   dotnet build -c Release
   ```
3. 编译产物位于：
   `bin/Release/net48/GOI.exe`

---

## 📚 开源依赖项 / Third-Party Libraries

本项目的顺利完成离不开以下优秀开源项目的支持：

* **[iNKORE.UI.WPF.Modern](https://github.com/iNKORE-NET/iNKORE.UI.WPF.Modern)** - 提供极致的 Fluent 2.0 WPF 控件库与原生主题适应。
* **[Microsoft-Activation-Scripts (MAS)](https://github.com/massgravel/Microsoft-Activation-Scripts)** - 提供行业基础的 Windows/Office 开源激活解决方案（本项目集成并修改了其 Ohook 核心脚本）。
* **[Costura.Fody](https://github.com/Fody/Costura)** - 用于将 DLL 依赖项合并到主执行文件中以实现单文件分发。
* **[SharpCompress](https://github.com/adamhathcock/sharpcompress)** - 跨平台的高性能解压缩支持，用于解压与提取本地安装包。

---

## ⚖️ 声明与免责条款 / Disclaimer

1. 本项目所集成的 Microsoft Office 激活脚本（Ohook）来源于公开开源项目，仅供个人技术研究和教育目的使用。
2. 请勿将本项目用于任何商业或非法用途。请支持并购买正版微软 Office 软件。
3. 开发者对使用本工具所造成的任何系统故障、数据丢失或法律纠纷不承担任何责任。

---

## 📄 开源协议 / License

本项目采用 [MIT License](LICENSE) 协议开源。
