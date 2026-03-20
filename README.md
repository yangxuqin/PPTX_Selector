# PPTX 智能打开器

根据 PPTX 文件的创建软件，自动选择 PowerPoint 或 WPS 打开，避免排版错乱。

## 功能特点

- 自动检测 PPTX 文件是由 Microsoft PowerPoint 还是 WPS 创建
- 智能选择对应的软件打开，确保排版正确
- 零控制台窗口，静默运行
- 完整的日志记录，便于问题排查
- 优化的性能，快速启动

## 使用场景

适用于教室电脑等多人共享环境，不同老师使用不同软件制作 PPT，需要用对应软件打开以避免排版错乱。

## 编译

### 编译命令

```bash
g++ main.cpp miniz.c resource.o -o selector.exe -lshell32 -O2 -mwindows -static -static-libgcc -static-libstdc++ -s
```

### 编译步骤

1. 编译资源文件（如果有修改）：
   ```bash
   windres resource.rc -o resource.o
   ```

2. 编译主程序：
   ```bash
   g++ main.cpp miniz.c resource.o -o selector.exe -lshell32 -O2 -mwindows -static -static-libgcc -static-libstdc++ -s
   ```

## 使用方法

### 命令行调用

```bash
selector.exe "path\to\your\presentation.pptx"
```

### 文件关联

可以将 `.pptx` 文件关联到此程序，实现双击自动打开。

## 配置

### 软件路径配置

在 `main.cpp` 中修改以下路径以匹配您的系统：

```cpp
static const char* PATH_POWERPOINT =
    "C:\\Program Files\\Microsoft Office\\root\\Office16\\POWERPNT.EXE";

static const char* PATH_WPS =
    "C:\\Program Files (x86)\\Kingsoft\\WPS Office\\12.1.0.23542\\office6\\wpp.exe";
```

### 日志配置

日志文件位置：`d:\logs\pptx_selecter.log`

可以通过修改以下配置启用或禁用日志：

```cpp
static const bool ENABLE_LOG = true;  // 设置为 false 禁用日志
```

## 工作原理

1. 读取 PPTX 文件（本质是 ZIP 格式）
2. 解压并解析 `docProps/app.xml` 文件
3. 提取 `<Application>` 标签内容
4. 根据应用名称判断创建软件：
   - 包含 "WPS" 或 "Kingsoft" → 使用 WPS 打开
   - 包含 "Microsoft" 或 "PowerPoint" → 使用 PowerPoint 打开
   - 其他情况 → 默认使用 WPS 打开

## 技术栈

- **C++**: 核心逻辑实现
- **miniz**: 轻量级 ZIP 解压缩库
- **Win32 API**: Windows 系统调用
- **ShellExecuteA**: 程序启动

## 文件说明

- `main.cpp`: 主程序源代码
- `miniz.c`: miniz 库实现
- `miniz.h`: miniz 库头文件
- `resource.rc`: 资源文件（图标）
- `icon.ico`: 程序图标
- `resource.o`: 编译后的资源对象文件

## 注意事项

1. 确保 PowerPoint 和 WPS 的安装路径与配置一致
2. 需要管理员权限才能在 `d:\logs` 创建日志目录（或修改日志路径）
3. 程序需要 Windows 系统环境运行

## 故障排查

### 程序无法打开 PPTX

1. 检查日志文件 `d:\logs\pptx_selecter.log`
2. 确认 PowerPoint 和 WPS 的路径是否正确
3. 检查 PPTX 文件是否损坏

### 日志无法创建

1. 检查是否有 `d:\logs` 目录的写入权限
2. 修改日志路径到有权限的位置

### 获取更多信息
[![zread](https://img.shields.io/badge/Ask_Zread-_.svg?style=flat&color=00b0aa&labelColor=000000&logo=data%3Aimage%2Fsvg%2Bxml%3Bbase64%2CPHN2ZyB3aWR0aD0iMTYiIGhlaWdodD0iMTYiIHZpZXdCb3g9IjAgMCAxNiAxNiIgZmlsbD0ibm9uZSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIj4KPHBhdGggZD0iTTQuOTYxNTYgMS42MDAxSDIuMjQxNTZDMS44ODgxIDEuNjAwMSAxLjYwMTU2IDEuODg2NjQgMS42MDE1NiAyLjI0MDFWNC45NjAxQzEuNjAxNTYgNS4zMTM1NiAxLjg4ODEgNS42MDAxIDIuMjQxNTYgNS42MDAxSDQuOTYxNTZDNS4zMTUwMiA1LjYwMDEgNS42MDE1NiA1LjMxMzU2IDUuNjAxNTYgNC45NjAxVjIuMjQwMUM1LjYwMTU2IDEuODg2NjQgNS4zMTUwMiAxLjYwMDEgNC45NjE1NiAxLjYwMDFaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik00Ljk2MTU2IDEwLjM5OTlIMi4yNDE1NkMxLjg4ODEgMTAuMzk5OSAxLjYwMTU2IDEwLjY4NjQgMS42MDE1NiAxMS4wMzk5VjEzLjc1OTlDMS42MDE1NiAxNC4xMTM0IDEuODg4MSAxNC4zOTk5IDIuMjQxNTYgMTQuMzk5OUg0Ljk2MTU2QzUuMzE1MDIgMTQuMzk5OSA1LjYwMTU2IDE0LjExMzQgNS42MDE1NiAxMy43NTk5VjExLjAzOTlDNS42MDE1NiAxMC42ODY0IDUuMzE1MDIgMTAuMzk5OSA0Ljk2MTU2IDEwLjM5OTlaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik0xMy43NTg0IDEuNjAwMUgxMS4wMzg0QzEwLjY4NSAxLjYwMDEgMTAuMzk4NCAxLjg4NjY0IDEwLjM5ODQgMi4yNDAxVjQuOTYwMUMxMC4zOTg0IDUuMzEzNTYgMTAuNjg1IDUuNjAwMSAxMS4wMzg0IDUuNjAwMUgxMy43NTg0QzE0LjExMTkgNS42MDAxIDE0LjM5ODQgNS4zMTM1NiAxNC4zOTg0IDQuOTYwMVYyLjI0MDFDMTQuMzk4NCAxLjg4NjY0IDE0LjExMTkgMS42MDAxIDEzLjc1ODQgMS42MDAxWiIgZmlsbD0iI2ZmZiIvPgo8cGF0aCBkPSJNNCAxMkwxMiA0TDQgMTJaIiBmaWxsPSIjZmZmIi8%2BCjxwYXRoIGQ9Ik00IDEyTDEyIDQiIHN0cm9rZT0iI2ZmZiIgc3Ryb2tlLXdpZHRoPSIxLjUiIHN0cm9rZS1saW5lY2FwPSJyb3VuZCIvPgo8L3N2Zz4K&logoColor=ffffff)](https://zread.ai/yangxuqin/PPTX_Selector)

