# 发布说明 · v1.4.0

> 已由 `.github/workflows/build-installers.yml` 自动构建并发布到 GitHub Release。
> 命名规则：`TeachingManager-<版本>-<系统>-<架构>-<类型>-<UTC时间戳>`

---

## Release 标题

```
TeachingManager v1.4.0 —— 学生气泡资料页 + 学情管理模块（多系统安装包）
```

## Release 正文（复制以下内容）

### 📦 安装包一览（7 个，覆盖 Windows / macOS / Linux）

| 平台 | 安装包 | 类型 | 说明 |
|---|---|---|---|
| Windows | `*-windows-x64-setup-*.exe` | 安装向导（Inno Setup） | **推荐**：按用户安装，无需管理员/UAC；建开始菜单+可选桌面快捷方式，自带卸载器 |
| Windows | `*-windows-x64-*.msi` | MSI | 企业分发/组策略用；**装到 `C:\Program Files\TeachingManager`，安装时需管理员授权**；支持主版本升级与阻止降级 |
| Windows | `*-windows-x64-portable-*.exe` | 免安装便携版 | 双击即用，U 盘可用 |
| macOS | `*-macos-arm64-*.dmg` | DMG | Apple 芯片（M 系列） |
| macOS | `*-macos-x64-*.dmg` | DMG | Intel 芯片 |
| Linux | `*-linux-x64-*.deb` | DEB | `sudo apt install ./TeachingManager-*.deb`（Ubuntu/Debian） |
| Linux | `*-linux-x64-*.AppImage` | AppImage | `chmod +x` 后直接运行（通用发行版） |

**命名示例**：`TeachingManager-1.4.0-windows-x64-setup-20260913-0428.exe`
（版本 `1.4.0` + 平台架构 + 包类型 + UTC 构建时间 `20260913-0428`）

### 🟠 新功能 1：学生磁贴点开即看「气泡资料页」

单击任意学生磁贴滑入该生资料页，信息以**或大或小的气泡 / 标签云**呈现：

- 👤 姓名（最大）、🏫 班级、🆔 学号、🗓 考试次数
- 🏷️ 每个标签一个气泡：**本班拥有该标签的人越多，气泡越大**（悬停显示占比）
- ⭐ 头衔气泡（金色）
- 📊 平均分（按分数档放大/缩小）、⬆️ 最高、⬇️ 最低、各科目平均分（分数越高越大）、
  🕐 最近考试、💬 评语

资料页提供 **✏️ 编辑资料 / 📈 成绩详情 / 🗑 删除**，并接入导航栏
（面包屑 `🏠 班级总览 › 📚 高一(1)班 › 👤 张三`，可点任意一段跳回）。

### 🟠 新功能 2：📋 学情管理模块（第 5 个模块）

- **自动读取已有学员名单**（每生一行，姓名/学号/班级为只读基础列）
- **自由补充学生任何信息**：`➕ 添加字段` 自建任意列（家庭情况、薄弱科目、学习习惯、
  家长电话…列数不限），支持**重命名 / 删除字段**
- 绿色单元格**双击即可编辑、改完自动保存**（底部显示保存次数）
- 支持关键词**筛选**、**🔄 刷新名单**、**📤 导出 Excel**
- 🔴 **模块隔离**：学情信息**不会**出现在学生磁贴、班级磁贴或气泡资料页；
  删除学生时其学情填写值自动级联清理

### 🟢 修复

- **xlsx 空单元格不再被读成字符串「None」**（此前空标签显示为「🏷️None」、
  空学号变成「None」；空学号现在自动回退为学生姓名）
- Excel 数值型学号去掉多余小数（`20260001.0` → `20260001`）

### ⚙️ 安装包流水线（本次新增）

- 新增 `Build installers` 工作流：**打 `v*` 标签即自动**构建三平台安装包并发布 Release
  - Windows：PyInstaller → **Inno Setup(Setup.exe)** + **msilib(MSI)** + 便携 exe
  - macOS：`macos-14`(arm64) 与 `macos-15-intel`(x64) 双架构 DMG
  - Linux：PyInstaller onedir → **dpkg-deb(.deb)** + **appimagetool(AppImage)**
- MSI 使用 **Python 标准库 msilib** 生成（WiX v7 要求接受 OSMF EULA，CI 会报 WIX7015）

### ⚠️ 首次运行提示

- **Windows**：Setup.exe 按用户安装（无需管理员）；MSI 装 Program Files 需在 UAC 弹窗中授权
- **macOS**：未签名版本首次打开若提示「无法验证开发者」→ 右键 → 打开，或执行
  `xattr -dr com.apple.quarantine "/Applications/教学管理系统.app"`
- **Linux**：`.deb` 用 `sudo apt install ./xxx.deb`；AppImage `chmod +x` 后直接运行
- **数据位置**：优先程序目录（便携模式）；不可写时回退 `~/Library/Application Support/TeachingManager/`
  （macOS）或 `%APPDATA%\TeachingManager\`（Windows）
- **字体**：界面优先 MiSans，未安装时 macOS 用苹方、Windows 用微软雅黑

**完整变更记录见 [CHANGELOG.md](../CHANGELOG.md)。**
