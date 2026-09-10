# 发布说明 · v1.3.1

> 这份文件是给 GitHub Release 页面的**可直接粘贴文案**（标题 + 正文）。
> 发布步骤见文末。

---

## Release 标题

```
v1.3.1 —— 修复「打开课程管理闪出多个窗口」BUG + 学生管理导航栏
```

## Release 正文（复制以下内容）

### 🔴 重要修复

- **修复「打开课程管理 / 切换页面时瞬间闪出多个程序窗口，随后全部关闭」的严重 BUG**
  - 根因：日历格、课程条目、各类磁贴等控件构造时未指定父控件，Qt 会为其创建**真实顶层窗口**；同时清理逻辑使用 `setParent(None)`，使控件**退化为顶层窗口**并延迟销毁、越积越多（诊断时单次操作出现约 180 个 `DayCell` + 30 个 `QWidget` + 6 个 `CourseRow`）。
  - 修复：构造即传父控件 + 清理统一 `hide() + deleteLater()`，**彻底移除 `setParent(None)`**（覆盖全部四个模块）。
  - 实测（真机 Windows，打开课程管理瞬间的顶层窗口数）：**修复前 6 → 30 → 70 → 归零（闪烁）；修复后恒定 6**。
- **修复窗口标题重复**（Windows 上不再出现「教学管理系统 v1.3.1 - 教学管理系统」）。

### 🟠 新增

- **学生管理页面导航栏**：`← 上一页` + **可点击面包屑**（`🏠 班级总览 › 高一(1)班 › 🔍 搜索：张`），点击任意一段即可跳回该页；「🏠 班级总览」一键回首页并清空前进历史。

### ✅ 质量保障

- 新增**防回归测试**：切换页面、重建磁贴不得产生游离顶层窗口（AST 级约束：代码禁止出现 `setParent(None)`）。
- 测试规模：**89 项逻辑 + 23 项界面/新功能 + 跨平台与打包约束**，三平台 CI 全绿。

### 📦 下载与安装

| 平台 | 文件 | 安装方式 |
|---|---|---|
| Windows | `教学管理系统.exe`（单文件，约 84 MB） | 双击运行（首次启动解压需 10~30 秒） |
| macOS (Apple 芯片) | `教学管理系统-1.3.1-macOS-arm64.dmg` | 打开 DMG，把应用拖入「应用程序」 |
| macOS (Intel) | `教学管理系统-1.3.1-macOS-x86_64.dmg` | 同上 |
| 源码 | `Source code (zip/tar.gz)` | `pip install PySide6 matplotlib openpyxl python-docx xlrd && python teaching_manager.py` |

> macOS 的 DMG 由 GitHub Actions 在 macOS runner 上构建（PyInstaller 不支持交叉编译，DMG 也需要 macOS 的 `hdiutil`）。若某个架构的 DMG 未出现在本 Release 中，可在 **Actions → Build macOS DMG → Run workflow** 手动构建后下载工件，或在本机 Mac 上执行 `packaging/macos/build_macos.sh`。

### ⚠️ 首次运行提示

- macOS 未签名版本首次打开若提示「无法验证开发者」：**右键 → 打开**，或执行
  `xattr -dr com.apple.quarantine "/Applications/教学管理系统.app"`
- 数据默认保存在程序目录（便携模式）；不可写时回退 `~/Library/Application Support/TeachingManager/`（macOS）或 `%APPDATA%\TeachingManager\`（Windows）。菜单 **文件 → 打开数据目录** 可直接定位。
- 界面字体优先 **MiSans**；未安装时 macOS 使用苹方、Windows 使用微软雅黑。也可把 `MiSans.ttf` 放进程序目录 `fonts/`。

**完整变更记录见 [CHANGELOG.md](../CHANGELOG.md)。**

---

## 如何在 GitHub 发布这个 Release

1. 打开仓库 → 右侧 **Releases → Draft a new release**。
2. **Choose a tag**：输入 `v1.3.1` → Create new tag on publish。
3. **Release title**：粘贴上面的「Release 标题」。
4. **Describe this release**：粘贴上面「Release 正文」。
5. 拖入安装包（Windows 的 `教学管理系统.exe`；macOS 的 `.dmg` 从 Actions 工件下载后拖入）。
6. 点 **Publish release**。

> macOS 的 `.dmg` 需要在 macOS 上构建（或由本仓库的 Actions 工作流产出）。工作流运行后，在 **Actions → 对应运行 → Artifacts** 中下载 `教学管理系统-macOS-arm64` / `教学管理系统-macOS-x86_64`，再上传到本 Release。
