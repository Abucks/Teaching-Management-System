# Windows 安装包构建说明

本目录提供两种 Windows 安装包生成方式，产物都会带上「版本 + 构建时间戳」：

| 产物 | 生成方式 | 特点 |
|---|---|---|
| `TeachingManager-<版本>-windows-x64-setup-<时间戳>.exe` | `installer.iss` + **Inno Setup 6** | 图形化安装向导；默认**按用户安装**（无需管理员/UAC）；创建开始菜单与可选桌面快捷方式；自带卸载器 |
| `TeachingManager-<版本>-windows-x64-<时间戳>.msi` | `build_msi.py` + **Python 标准库 msilib** | 标准 MSI，适合企业分发/组策略；**按机器安装**到 `C:\Program Files\TeachingManager`（安装时需要管理员授权）；支持主版本升级与阻止降级 |
| `TeachingManager-<版本>-windows-x64-portable-<时间戳>.exe` | PyInstaller `--onefile` | 免安装便携版，双击即用 |

---

## 1. MSI（推荐，无需 WiX）

```powershell
# 需要 Windows + Python ≤ 3.12（msilib 在 Python 3.13 被移除）
python packaging\windows\build_msi.py `
  dist\TeachingManager.exe `
  dist\installers\TeachingManager-1.4.0-windows-x64-20260913-1015.msi `
  1.4.0 `
  packaging\macos\assets\icon.ico
```

脚本行为：

- 安装到 `C:\Program Files\TeachingManager`，注册卸载信息（控制面板可见）
- 创建开始菜单与桌面快捷方式（图标自动继承 exe 内嵌图标）
- `Upgrade` 表实现**主版本升级**：装新版本会自动移除旧版本；检测到更高版本时阻止降级
- 组件 GUID / 产品 GUID 均由名称稳定派生，便于升级与修复

> 为什么不用 WiX：`dotnet tool install --global wix` 现在会安装 v7，而 v7 要求接受
> **Open Source Maintenance Fee (OSMF) EULA**，CI 上直接报 `WIX7015` 失败。本目录仍保留
> `installer.wxs`（WiX v5 语法）作为可选方案，使用时请 `--version 5.*` 锁定版本。

验证 MSI 内容（无需管理员，展开到临时目录）：

```powershell
msiexec /a TeachingManager-1.4.0-windows-x64-xxxx.msi /qn TARGETDIR=D:\tmp\msi-out
```

正常安装（会弹 UAC 授权，标准 MSI 行为）：

```powershell
msiexec /i TeachingManager-1.4.0-windows-x64-xxxx.msi        # 图形界面
msiexec /i TeachingManager-1.4.0-windows-x64-xxxx.msi /qn    # 静默（需管理员权限）
msiexec /x TeachingManager-1.4.0-windows-x64-xxxx.msi /qn    # 卸载
```

## 2. Setup.exe（Inno Setup）

```powershell
# 需要安装 Inno Setup 6（ISCC.exe），CI 上通过 choco install innosetup 安装
iscc /DAppVersion=1.4.0 packaging\windows\installer.iss
```

> 注意：Inno 中的相对路径是**相对 `.iss` 文件所在目录**解析的。
> `installer.iss` 内已用 `SourcePath` 拼出绝对路径，直接用上面的命令即可。

安装向导语言使用 Inno 默认（英文）。如需中文向导，请把 `ChineseSimplified.isl`
放入 Inno 的 `Languages` 目录，并在 `installer.iss` 中添加：

```ini
[Languages]
Name: "chinese"; MessagesFile: "compiler:Languages\ChineseSimplified.isl"
```

## 3. 便携版

由 PyInstaller 直接产出（见 `.github/workflows/build-installers.yml`）：

```powershell
pyinstaller --noconfirm --clean --onefile --windowed --name TeachingManager `
  --icon packaging\macos\assets\icon.ico `
  --add-data "packaging\macos\assets\icon.png;." teaching_manager.py
```

---

## 数据存放位置

程序优先把数据库放在**程序所在目录**（便携模式）；该目录不可写时自动回退到
`%APPDATA%\TeachingManager\`。因此装在 Program Files 时数据会落到 `%APPDATA%`，
按用户安装（Inno 默认）时数据就在安装目录，两种情况都无需管理员权限即可日常使用。
