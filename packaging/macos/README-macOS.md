# 教学管理系统 · macOS（DMG）打包与使用说明

> 版本 v1.2.0 · 跨平台版（Windows / macOS / Linux）
> 本目录提供 **在 macOS 上生成 `.app` 与 `.dmg` 的完整流水线**。

---

## 0. 重要前提（请先读）

`.dmg` 与 macOS 的 `.app` **只能在 macOS 上生成**：

| 工具 | 作用 | 是否可跨平台 |
|---|---|---|
| PyInstaller | 把 Python 程序打包成可执行文件 | ❌ 官方明确不支持交叉编译，必须在目标系统上构建 |
| `hdiutil` | 生成 DMG 磁盘映像 | ❌ 仅 macOS 自带 |
| `iconutil` / `sips` | PNG → `.icns` 图标 | ❌ 仅 macOS 自带 |
| `codesign` / `notarytool` | 签名与公证 | ❌ 仅 macOS 自带 |

因此正确的做法是二选一：

- **方案 A**：在一台 Mac 上执行本目录的 `build_macos.sh`（约 3~5 分钟，一条命令搞定）。
- **方案 B**：把项目推到 GitHub，用本目录的 CI 工作流（`build-macos.yml`）在云端 macOS 机器上自动产出 DMG，直接下载。

在 Windows 上可以先执行 `python build_macos.py --dry-run` 检查整个流程要跑哪些命令（不需要 Mac）。

---

## 1. 方案 A：在 Mac 上一条命令打包

```bash
# 1) 准备：安装 Xcode 命令行工具（提供 iconutil / sips / codesign）
xcode-select --install

# 2) 进入打包目录（仓库里就是本目录）
cd packaging/macos

# 3) 一键打包（自动建虚拟环境、装依赖、生成图标、打包 .app、制作 .dmg）
chmod +x build_macos.sh
./build_macos.sh
```

产物：

```
dist/教学管理系统.app                               ← 可直接双击运行
教学管理系统-1.2.0-macOS.dmg                        ← 分发用安装包（内含 .app + Applications 快捷方式）
```

安装方式：双击 DMG → 把「教学管理系统」拖进 `Applications`。

### 常用参数

| 参数 | 说明 |
|---|---|
| `--arch arm64` | 仅 Apple Silicon（M 系列） |
| `--arch x86_64` | 仅 Intel Mac |
| `--arch universal2` | 通用二进制（**需要 universal2 版 Python**，见下） |
| `--no-dmg` | 只生成 `.app`，不制作 DMG |
| `--skip-deps` | 复用已有虚拟环境，跳过装依赖 |
| `--sign "Developer ID Application: XXX (TEAMID)"` | 代码签名（含 DMG 签名与公证命令提示） |
| `--venv PATH` | 指定虚拟环境目录 |
| `--dry-run` | 只打印命令不执行（任何系统可用） |

示例：

```bash
./build_macos.sh --arch universal2                      # 通用二进制
./build_macos.sh --no-dmg --skip-deps                   # 快速重打包
./build_macos.sh --sign "Developer ID Application: 张三 (AB12CD34EF)"
```

> **universal2 提示**：`--arch universal2` 需要“通用版”Python。请使用
> [python.org 的 macOS universal2 安装包](https://www.python.org/downloads/macos/)，
> 或 Homebrew 的 `python@3.12`（universal 构建）。用 `python3 -c "import platform;print(platform.machine())"` 无法判断通用性，
> 若打包时报架构相关错误，就用默认（本机架构）分两次分别在 M 系列 Mac 和 Intel Mac 上构建。

---

## 2. 方案 B：用 GitHub Actions 云端产出 DMG（无需 Mac）

1. 把项目推送到 GitHub 仓库（保证目录结构包含 `teaching_manager.py` 与 `packaging/macos/`）。
2. 把本目录下的 `build-macos.yml` 复制到仓库的 `.github/workflows/build-macos.yml`。
3. 在 GitHub 网页上 `Actions → Build macOS DMG → Run workflow`，或在打标签 `v1.2.0` 时自动触发。
4. 构建完成后在该次运行的 **Artifacts** 里下载 `教学管理系统-macOS-arm64.dmg`（Apple 芯片）和 `教学管理系统-macOS-x86_64.dmg`（Intel）。

---

## 3. 方案 C：不打包，直接在 Mac 上运行源码

```bash
python3 -m venv .venv && source .venv/bin/activate
pip install PySide6 matplotlib openpyxl python-docx xlrd
python teaching_manager.py
```

适合先试用；正式分发给老师/同学时建议用方案 A/B 产出 DMG。

---

## 4. macOS 适配说明（v1.3.0）

| 项目 | 说明 |
|---|---|
| **数据目录** | 双击 `.app` 时工作目录是 `/`、且 `.app` 包内只读。程序会按 **环境变量 `TMS_DATA_DIR` → 程序目录（便携模式）→ 用户数据目录** 的顺序自动选择可写目录；macOS 上回退到 `~/Library/Application Support/TeachingManager/`，不会再因无法建库而启动失败。菜单 **文件 → 打开数据目录** 可在访达中直接定位。便签图片存放在该目录的 `feedback_images/` 子目录，备份时请一并复制。 |
| **中文字体** | v1.3.0 起统一优先 **MiSans**：若系统已安装 MiSans 则直接使用；macOS 未安装时自动使用 **苹方 PingFang SC**（→ 华文黑体 → 宋体兜底）。也可把 `MiSans.ttf` 放进程序目录的 `fonts/` 文件夹，程序启动时自动注册加载。界面样式表与 matplotlib 图表（标题/坐标轴/图例）共用同一字体栈，中文不会显示成方块。 |
| **菜单栏** | 菜单角色已标注（`QuitRole`/`AboutRole`/`NoRole`），macOS 下「退出」「关于」会正确归入屏幕顶部的应用菜单，而不是留在「文件」里。 |
| **Retina / 高DPI** | 设置 `HighDpiScaleFactorRoundingPolicy.PassThrough`，并在 `Info.plist` 写入 `NSHighResolutionCapable=true`，高分辨率屏不模糊。 |
| **窗口定位** | 用屏幕「可用区域」居中并自适应尺寸，避免窗口被菜单栏/程序坞遮挡或超出小屏。 |
| **应用标识** | `CFBundleIdentifier=com.teaching.manager`、`CFBundleDisplayName`、`LSMinimumSystemVersion=11.0`、分类为教育类，程序坞与「关于」显示中文名。 |
| **旧版 `.doc`** | Word 的 `.doc` 依赖 Windows COM，macOS 上会给出明确提示（请先另存为 `.docx`），不会崩溃。`.docx/.xlsx/.xls/.csv` 在 macOS 上完全支持。 |
| **关闭行为** | 关闭主窗口即退出应用（单窗口应用符合 macOS 习惯）。 |
| **便签粘贴** | 便签墙的「智能粘贴」读取系统剪贴板（⌘C 复制的图片/表格/文字），macOS 上无需额外权限；首次使用粘贴功能时若系统询问，请允许该应用访问剪贴板。 |

---

## 5. 数据迁移（Windows ⇄ macOS）

数据库就是一个 SQLite 文件 `teaching_management.db`：

- **Mac → Windows**：把 `~/Library/Application Support/TeachingManager/teaching_management.db` 复制到 Windows 端程序目录。
- **Windows → Mac**：反向复制到 `~/Library/Application Support/TeachingManager/`（先用一次程序让该目录生成，或在访达里按 `⌘⇧G` 输入路径创建）。

也可用环境变量临时指定数据目录（例如放在 iCloud/U 盘）：

```bash
TMS_DATA_DIR="$HOME/Documents/教学管理数据" open -a "教学管理系统"
```

---

## 6. 签名与公证（分发给他人时建议做）

```bash
# 1) 查看可用签名身份
security find-identity -v -p codesigning

# 2) 带签名打包（脚本已内置 codesign --deep --options runtime）
./build_macos.sh --sign "Developer ID Application: 你的名字 (TEAMID)"

# 3) 公证（首次需先存凭据：xcrun notarytool store-credentials）
xcrun notarytool submit "教学管理系统-1.2.0-macOS.dmg" \
      --apple-id "你的AppleID" --team-id "TEAMID" --password "App专用密码" --wait
xcrun stapler staple "教学管理系统-1.2.0-macOS.dmg"
```

未签名/未公证时的正常提示与解法：

- 提示 **“无法打开，因为 Apple 无法检查其是否包含恶意软件”** → 在「访达」里 **右键应用 → 打开 → 打开**（只需一次）。
- 或命令行解除隔离：`xattr -dr com.apple.quarantine "/Applications/教学管理系统.app"`。

---

## 7. 常见问题

| 现象 | 处理 |
|---|---|
| 双击 `.app` 没反应 | 终端执行 `"/Applications/教学管理系统.app/Contents/MacOS/教学管理系统"` 查看报错信息 |
| 提示“应用已损坏，无法打开” | 多为下载隔离属性：`xattr -dr com.apple.quarantine <app路径>` |
| 图表中文显示为方块 | 清理 matplotlib 字体缓存：`rm -rf ~/.matplotlib`；确认系统有苹方字体（系统自带） |
| `pip install PySide6` 失败 | 确认 Python ≥ 3.10（推荐 3.12），并先 `pip install -U pip`；Apple 芯片须用 arm64 版 Python |
| 想减小体积 | 默认已排除 QtWebEngine/QtQuick 等无用模块；`.app` 约 250~400 MB 属正常（PySide6 + matplotlib + numpy） |
| 首次启动较慢 | 应用包内含 Qt 与 matplotlib，首次启动需加载；属正常现象 |
| Intel Mac 打不开 arm64 包 | 用 `--arch x86_64` 重新构建，或使用本目录的 CI（会同时产出两种架构） |

---

## 8. 目录结构

```
packaging/macos/
├── build_macos.py         # 一键打包主脚本（依赖→图标→.app→签名→.dmg）
├── build_macos.sh         # 同上，Shell 包装（./build_macos.sh 即可）
├── TeachingManager.spec   # PyInstaller 规格：生成 .app（Info.plist / 排除项 / 架构）
├── build-macos.yml        # GitHub Actions 工作流（复制到 .github/workflows/ 使用）
├── README-macOS.md        # 本文件
└── assets/
    ├── make_icon.py       # 用 PySide6 绘制 1024×1024 应用图标（无需素材）
    └── icon.png           # 已生成好的图标；构建时自动转 .icns
```
