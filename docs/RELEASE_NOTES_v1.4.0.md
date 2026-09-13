# 发布说明 · v1.4.0

> 可直接粘贴到 GitHub Release 页面的文案（标题 + 正文）。发布步骤见文末。

---

## Release 标题

```
v1.4.0 —— 学生气泡资料页 + 学情管理模块（含 xlsx 空单元格读取修复）
```

## Release 正文（复制以下内容）

### 🟠 新功能 1：学生磁贴点开即看「气泡资料页」

单击任意学生磁贴即可滑入该生的资料页，把已经编辑好的信息以**或大或小的气泡 / 标签云**铺开：

- 👤 姓名（最大）、🏫 班级、🆔 学号（身份气泡）
- 🏷️ 每个标签一个气泡：**本班拥有该标签的人越多，气泡越大**（悬停显示占比）
- ⭐ 头衔气泡（金色）
- 📊 平均分（按分数档自动放大/缩小）、⬆️ 最高、⬇️ 最低、🗓 考试次数
- 各科目平均分（分数越高气泡越大）、🕐 最近考试、💬 评语

资料页右上角提供 **✏️ 编辑资料 / 📈 成绩详情 / 🗑 删除**；页面已接入导航栏，
面包屑显示 `🏠 班级总览 › 📚 高一(1)班 › 👤 张三`，点任意一段即可跳回。

### 🟠 新功能 2：📋 学情管理模块（第 5 个模块）

- **自动读取已有学员名单**（每生一行，姓名/学号/班级为只读基础列）
- **自由补充学生任何信息**：点「➕ 添加字段」自建任意列——家庭情况、薄弱科目、
  学习习惯、家长电话、心理关注点……列数不限；支持**重命名 / 删除字段**
- 绿色单元格**双击即可编辑，改完自动保存**（底部显示保存次数）
- 支持按关键词**筛选行**、**🔄 刷新名单**、**📤 导出 Excel** 归档
- 🔴 **模块隔离（重要）**：学情信息**不会**出现在「学生管理」的学生磁贴、
  班级磁贴或气泡资料页中；删除学生时其学情填写值自动级联清理

### 🟢 修复

- **修复导入 xlsx 时空单元格被读成字符串「None」**：此前空标签会显示为「🏷️None」、
  空学号会变成「None」（空学号现在自动回退为学生姓名）
- 修复 Excel 数值型学号带多余小数的问题（`20260001.0` → `20260001`）

### ✅ 质量保障

- 测试规模：**108 项逻辑 + 29 项界面/新功能 + 跨平台与打包约束**（三平台 CI 全绿）
- 新增关键回归：**「学情信息不得出现在学生磁贴/气泡中」**（模块隔离）、
  **「空单元格不得产生 None」**（导入容错）、防「游离顶层窗口」闪烁

### 📦 下载与安装

| 平台 | 文件 | 安装方式 |
|---|---|---|
| Windows | `教学管理系统.exe`（单文件） | 双击运行（首次启动解压需 10~30 秒） |
| macOS (Apple 芯片) | `教学管理系统-1.4.0-macOS-arm64.dmg` | 打开 DMG，拖入「应用程序」 |
| macOS (Intel) | `教学管理系统-1.4.0-macOS-x86_64.dmg` | 同上 |
| 源码 | `Source code (zip/tar.gz)` | `pip install PySide6 matplotlib openpyxl python-docx xlrd && python teaching_manager.py` |

> macOS 的 DMG 由 GitHub Actions 在 macOS runner 上构建（PyInstaller 不支持交叉编译）。
> 若某架构的 DMG 未出现在本 Release，可在 **Actions → Build macOS DMG → Run workflow**
> 手动构建后，从该次运行的 Artifacts 下载。

### ⚠️ 首次运行提示

- macOS 未签名版本首次打开若提示「无法验证开发者」：**右键 → 打开**，或执行
  `xattr -dr com.apple.quarantine "/Applications/教学管理系统.app"`
- 数据默认保存在程序目录（便携模式）；不可写时回退 `~/Library/Application Support/TeachingManager/`
  （macOS）或 `%APPDATA%\TeachingManager\`（Windows）。菜单 **文件 → 打开数据目录** 可直达。
- 界面字体优先 **MiSans**；未安装时 macOS 用苹方、Windows 用微软雅黑。

**完整变更记录见 [CHANGELOG.md](../CHANGELOG.md)。**

---

## 如何在 GitHub 发布这个 Release

1. 仓库 → **Releases → Draft a new release**
2. **Choose a tag**：输入 `v1.4.0` → Create new tag on publish
3. **Release title / Describe this release**：粘贴上文两段
4. 拖入安装包（Windows exe；macOS 的 `.dmg` 从 Actions 工件下载后拖入）
5. **Publish release**
