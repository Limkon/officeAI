# OfficeAI — Windows 全平台大模型接入 Office / WPS 助手

**OfficeAI** 是一款面向 Windows 系统的开源办公 AI 集成工具，可将国内外主流云端大模型 API 以及本机离线大模型（Ollama、LM Studio、vLLM 等）无缝接入 **Microsoft Office（Word / Excel / PowerPoint）** 与 **WPS Office（文字 / 表格 / 演示）**。

---

## 核心特性

### 1. 全主流大模型与本地模型统一接入
- **云端大模型 API**：
  - OpenAI（GPT-4o / GPT-4.1 / o3-mini）
  - Anthropic Claude（原生 Messages 协议 + 流式输出 + Thinking）
  - Google Gemini（原生 `streamGenerateContent` 协议）
  - DeepSeek 深度求索（支持 V3 与 R1 深度思考折叠展示）
  - 阿里云通义千问 Qwen / 百炼（DashScope）
  - 智谱 GLM（GLM-4 / GLM-Z1）
  - Moonshot Kimi
  - 火山方舟 Doubao（豆包）
  - 硅基流动 SiliconFlow / OpenRouter / Groq / 自定义 OpenAI 兼容网关（OneAPI / NewAPI）
- **本机离线模型（数据不出本地）**：
  - **Ollama**（自动探测本机 `11434` 端口并拉取已安装模型列表）
  - **LM Studio**（自动探测 `1234` 端口）
  - **vLLM / llama.cpp / LocalAI**（自定义本地推理服务）

### 2. 双办公套件深度原生集成（Windows 无需管理员权限）
- **WPS Office（个人版 & 专业版全支持）**：
  - 基于 WPS 原生 JS 加载项（`ribbon.xml` + `WPS JSAPI`），安装后顶部自动出现 **「OfficeAI 助手」** 选项卡与右侧边栏。
  - 支持 WPS 文字（读写选区、润色替换、插入）、WPS 表格（读取单元格区域、自然语言生成公式并写入单元格）、WPS 演示（文本框读写、演讲备注）。
- **Microsoft Office（2016 / 2019 / 2021 / Microsoft 365）**：
  - 基于 Office Web Add-in（`Office.js`）+ 自动生成并信任的本地 HTTPS 证书（`Cert:\CurrentUser\Root`），一键旁加载到 Word / Excel / PowerPoint。

### 3. 生命周期绑定办公软件（不做开机自启，也不后台常驻）
**设计目标：打开 Office / WPS 才启动，全部关闭即退出，不浪费系统资源。**

- **不写任何自启项**：不使用注册表 `Run`、启动文件夹或计划任务，重启电脑后不会自动运行。
- **WPS 全自动**：加载项以**本地离线加载项**安装到 `%APPDATA%\kingsoft\wps\jsaddons`，WPS 启动时直接读取，不依赖网关；点击「OfficeAI 助手」若服务未启动，会自动拉起 OfficeAI（首次可能弹一次 WPS 安全确认）。
- **Microsoft Office 全自动**：二选一（配置中心可切换，只装选中的那个）

| 触发方式 | 实现 | 常驻内存 | Office 启动开销 | 需要登录自启 |
| --- | --- | --- | --- | --- |
| **后台监听器**（默认） | `OfficeAI.Watcher.exe`，每 2s 查一次进程列表 | 约 10MB | **0**（完全不进 Office 进程） | 是（计划任务） |
| Office 唤醒器 | `OfficeAI.Waker.dll`（COM，`LoadBehavior=3`） | 0 | 探活 ≤300ms | 否 |

  两者都不提供界面，侧边栏仍由 Web 加载项提供；也可用 `officeai://` + 桌面快捷方式手动拉起。
- **兜底**：未安装唤醒器时，可用 `officeai://` 协议 + 桌面「OfficeAI 启动」快捷方式一键拉起。
- **用完即走**：退出判定是「**引用计数**」式的——只要还有任一 Office/WPS 进程在跑，或仍有侧边栏在发心跳，网关就继续运行；两者都归零后 15 秒内自动退出。

| 退出条件 | 判定方式 |
| --- | --- |
| Office / WPS 全部关闭 | `tasklist` 检测 `winword / excel / powerpnt / wps / et / wpp` 归零 |
| 所有侧边栏已关闭 | 加载项每 30s `POST /api/office/session` 心跳，超时（默认 2 分钟）即释放 |
| 长时间无人使用 | 空闲 `idleExitMinutes`（默认 10 分钟）无任何请求 |

仅有约 15–30ms 的空转检测开销：轮询间隔 5 秒，仅在 `on-demand` 模式下运行。

### 4. Windows 桌面托盘控制中心
- 提供 Electron Windows 桌面应用（`OfficeAI-Setup.exe` 安装版 & Portable 便携版）与命令行/便携脚本版。
- 内置模型配置中心、连通性测速与模型列表拉取、Office/WPS 一键安装卸载向导、自定义快捷指令库与调用日志统计。

---

## 项目结构

```text
officeAI/
├── packages/
│   ├── shared/          # 共享类型、主流厂商预设、Prompt 模板与流式解析器
│   ├── gateway/         # 本地多协议 AI 网关 (HTTP/HTTPS) 与 Windows 注册表/WPS 集成器
│   ├── addin-ui/        # React + Vite + TailwindCSS 统一侧边栏与配置中心界面
│   ├── office-addin/    # Microsoft Office Web Add-in Manifest 与命令宿主
│   ├── wps-addin/       # WPS Office 原生 JS 加载项 (ribbon.xml + main.js)
│   └── desktop-app/     # Windows Electron 桌面客户端与系统托盘包装
├── scripts/
│   ├── build-portable.mjs
│   └── windows/         # Windows 一键安装/启动/卸载批处理与 PowerShell 脚本
└── .github/workflows/   # GitHub Actions Windows 自动编译与打包工作流
```

---

## 快速开始（Windows）

### 方式一：使用 GitHub Actions 编译好的安装包（推荐）
1. 前往仓库 **Actions** 页面下载最新的 `OfficeAI-Windows-Release` 构建产物（包含 `OfficeAI Setup x.x.x.exe`、`OfficeAI-Portable-win-x64.exe` 以及 `OfficeAI-Windows-Portable.zip`）。
2. 运行 `OfficeAI Setup.exe`，在控制中心的 **「Office / WPS 安装集成」** 标签页点击一键安装即可。

### 方式二：从源码编译与运行
前置要求：Node.js >= 20

```bash
# 1. 安装依赖
npm install

# 2. 编译全部模块
npm run build

# 3. 运行单元测试与网关集成测试
npm test

# 4. 生成 Windows 便携包 (release/OfficeAI-Windows-Portable.zip)
npm run package:portable
```

在 Windows 上安装加载项（默认按需启动，不需要开机自启）：
```powershell
# 一键安装：WPS 本地离线加载项 + officeai:// 一键启动协议 + Microsoft Office 加载项
.\scripts\windows\Install-OfficeAI.bat
```

日常使用：
- **WPS**：直接打开 WPS → 点击顶部「OfficeAI 助手」，服务会自动拉起。
- **Microsoft Office**：双击桌面「OfficeAI 启动」→ 再打开 Word / Excel / PPT。
- 手动启动（调试用）：`.\scripts\windows\Start-OfficeAI.bat`

启动行为可在配置中心 →「Office / WPS 安装集成」→「启动方式」中调整：
| 模式 | 说明 |
| --- | --- |
| 按需启动（默认） | 点击加载项时自动启动；Office / WPS 全部关闭或空闲 N 分钟后自动退出 |
| 常驻托盘 | 启动后一直在后台运行，直到手动退出 |
| 仅手动 | 加载项只做提示，必须手动运行 OfficeAI |

也可填写自定义启动命令（留空时自动探测 `OfficeAI.exe` 或 `Start-OfficeAI.bat`）。

如需构建 Electron Windows `.exe` 安装包（在 Windows 环境下执行）：
```bash
npm run dist:win --workspace=@officeai/desktop-app
```
