# -*- coding: utf-8 -*-
"""
跨平台适配测试（Windows 主机上验证 macOS 分支逻辑）：
  1) 平台中文字体栈（macOS 必须优先苹方，Windows 微软雅黑，Linux 思源/文泉驿）
  2) 数据目录解析优先级：TMS_DATA_DIR > 便携模式 > 用户数据目录
     （模拟 macOS 从 .app 启动、程序目录不可写 → 回退 ~/Library/Application Support）
  3) macOS 不支持的 .doc 旧格式给出明确提示而不是崩溃
  4) 关键常量与版本、界面无残留 Windows 专用字体字面量
运行：python platform_test.py
"""
import importlib.util
import os
import re
import sys
import tempfile
from pathlib import Path

def _find_source() -> Path:
    """定位主程序源文件（teaching_manager.py 或旧的 deepseek_python_*.py）"""
    here = Path(__file__).resolve().parent
    for base in [here, here.parent, *here.parents]:
        for name in ("teaching_manager.py", "app.py"):
            cand = base / name
            if cand.exists():
                return cand
        hits = sorted(base.glob("deepseek_python_*.py"))
        if hits:
            return hits[0]
    raise SystemExit("未找到主程序源文件（teaching_manager.py）")


TARGET = _find_source()
SOURCE = TARGET.read_text(encoding="utf-8")

FAILED = []

def check(name, cond, detail=""):
    if cond:
        print(f"  [PASS] {name}")
    else:
        FAILED.append(name)
        print(f"  [FAIL] {name}  {detail}")


def load_module(data_dir=None):
    if data_dir:
        os.environ["TMS_DATA_DIR"] = data_dir
    else:
        os.environ.pop("TMS_DATA_DIR", None)
    spec = importlib.util.spec_from_file_location("tms_platform", TARGET)
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


def main():
    tmp = tempfile.mkdtemp(prefix="tms_platform_")
    mod = load_module(tmp)

    print("=" * 70)
    print("阶段1: 平台识别与中文字体栈")
    host = "Windows" if mod.IS_WINDOWS else ("macOS" if mod.IS_MACOS else "Linux")
    check("主机平台可识别（win/mac/linux 三选一）",
          sum([mod.IS_WINDOWS, mod.IS_MACOS, mod.IS_LINUX]) == 1, host)
    check("字体栈首选 MiSans（用户要求）", mod.FONT_STACK[0].startswith("MiSans"),
          mod.FONT_STACK[0])
    # 各平台的兜底字体不同：Windows 微软雅黑 / macOS 苹方 / Linux 思源
    host_fallback = {"Windows": "Microsoft YaHei", "macOS": "PingFang SC", "Linux": "Noto Sans CJK SC"}[host]
    check(f"主机字体栈含 {host} 兜底字体（{host_fallback}）",
          host_fallback in mod.FONT_STACK, str(mod.FONT_STACK[:10]))
    check("CSS 字体栈含 sans-serif 兜底", mod.FONT_STACK_CSS.endswith("sans-serif"),
          mod.FONT_STACK_CSS)
    check("CSS 字体栈包含 MiSans", "MiSans" in mod.FONT_STACK_CSS, mod.FONT_STACK_CSS)

    # 模拟 macOS
    old_mac, old_win = mod.IS_MACOS, mod.IS_WINDOWS
    mod.IS_MACOS, mod.IS_WINDOWS = True, False
    mac_stack = mod.get_platform_font_stack()
    check("macOS 字体栈首选 MiSans、其次苹方 PingFang SC",
          mac_stack[0].startswith("MiSans") and mac_stack[5] == "PingFang SC",
          str(mac_stack[:7]))
    check("macOS 字体栈不含微软雅黑在前列", "Microsoft YaHei" not in mac_stack[:5], str(mac_stack[:5]))
    check("macOS 字体栈含华文黑体/冬青兜底",
          any(f in mac_stack for f in ("Heiti SC", "STHeiti", "Hiragino Sans GB")), str(mac_stack))
    mac_home = mod.get_user_data_dir()
    check("macOS 用户数据目录位于 Library/Application Support",
          "Library" in str(mac_home) and "Application Support" in str(mac_home), str(mac_home))

    # 模拟 Linux
    mod.IS_MACOS, mod.IS_WINDOWS = False, False
    lin_stack = mod.get_platform_font_stack()
    check("Linux 字体栈首选 MiSans、含 Noto/思源兜底",
          lin_stack[0].startswith("MiSans")
          and any(f.startswith(("Noto", "Source", "WenQuanYi")) for f in lin_stack),
          str(lin_stack[:8]))
    mod.IS_MACOS, mod.IS_WINDOWS = old_mac, old_win

    print("=" * 70)
    print("阶段1.5: 字体加载与运行时选字（自带 fonts/ 目录 + 已安装字体探测）")
    dirs = [str(d) for d in mod.get_font_search_dirs()]
    check("字体搜索目录含 程序目录/fonts",
          any(d.replace("\\", "/").endswith("/fonts") for d in dirs), str(dirs))
    check("字体搜索目录含 用户数据目录/fonts",
          any("TeachingManager" in d and d.endswith("fonts") for d in dirs), str(dirs))
    check("load_bundled_fonts 可调用且返回列表",
          isinstance(mod.load_bundled_fonts(), list))

    # 真机（或 offscreen）上验证：apply_ui_font 能解析出可用字体并更新全局常量
    old_env = os.environ.get("QT_QPA_PLATFORM")
    os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")
    try:
        from PySide6.QtWidgets import QApplication
        app = QApplication.instance() or QApplication([])
        chosen = mod.apply_ui_font(app)
        check("apply_ui_font 返回字体名", bool(chosen), str(chosen))
        check("apply_ui_font 把选中字体提到栈首", mod.FONT_STACK[0] == chosen,
              f"{mod.FONT_STACK[:3]} chosen={chosen}")
        check("apply_ui_font 同步 matplotlib 字体栈",
              "matplotlib" not in sys.modules
              or mod.plt.rcParams['font.sans-serif'][:1] == [chosen],
              str(getattr(mod, 'plt', None) and mod.plt.rcParams['font.sans-serif'][:2]))
        # 应用默认字体已设置
        check("QApplication 默认字体已设置", app.font().family() != "" , app.font().family())
    except Exception as e:
        check("apply_ui_font 运行正常", False, f"{type(e).__name__}: {e}")
    finally:
        if old_env is None:
            os.environ.pop("QT_QPA_PLATFORM", None)
        else:
            os.environ["QT_QPA_PLATFORM"] = old_env

    print("=" * 70)
    print("阶段2: 数据目录解析（防止 macOS .app 内只读导致无法建库）")
    check("环境变量 TMS_DATA_DIR 生效", str(mod.DATA_DIR) == str(Path(tmp)),
          f"DATA_DIR={mod.DATA_DIR}")
    check("模式说明正确", mod.DATA_DIR_MODE.startswith("环境变量"), mod.DATA_DIR_MODE)
    check("DB_PATH 落在数据目录内", mod.DB_PATH.parent == Path(tmp), str(mod.DB_PATH))
    check("CONFIG_PATH 落在数据目录内", mod.CONFIG_PATH.parent == Path(tmp), str(mod.CONFIG_PATH))

    # 无环境变量 + 程序目录不可写（模拟 .app 包内只读 / macOS CWD=/）
    os.environ.pop("TMS_DATA_DIR", None)
    mod2 = load_module(None)
    check("无覆盖时默认便携模式（程序目录可写）",
          mod2.DATA_DIR == mod2.get_app_dir(), f"{mod2.DATA_DIR} mode={mod2.DATA_DIR_MODE}")

    # 跨平台地模拟“程序目录不可写”：只对伪造的程序目录返回不可写，
    # （不能用 Windows 盘符路径，那些路径在 macOS/Linux 上其实是可写的相对路径）
    real_get_app_dir = mod2.get_app_dir
    real_is_writable = mod2._is_writable_dir
    fake_unwritable = Path(tempfile.mkdtemp(prefix="tms_readonly_app_"))
    mod2.get_app_dir = lambda: fake_unwritable
    mod2._is_writable_dir = lambda p: False if Path(p) == fake_unwritable else real_is_writable(p)
    mod2.IS_MACOS, mod2.IS_WINDOWS = True, False        # 模拟 macOS
    resolved, mode = mod2.resolve_data_dir()
    expected = Path.home() / "Library" / "Application Support" / "TeachingManager"
    check("程序目录不可写时回退 macOS 用户数据目录",
          resolved == expected, f"resolved={resolved} expected={expected}")
    check("回退模式标注为用户数据目录", mode == "系统用户数据目录", mode)
    check("回退目录确实被创建/可用", resolved.exists(), str(resolved))

    # 同样的逻辑在 Linux 上也应回退到 XDG 目录
    mod2.IS_MACOS, mod2.IS_WINDOWS = False, False
    resolved_lin, mode_lin = mod2.resolve_data_dir()
    check("Linux 下程序目录不可写时回退用户数据目录",
          resolved_lin != fake_unwritable and mode_lin == "系统用户数据目录",
          f"resolved={resolved_lin} mode={mode_lin}")

    mod2.get_app_dir = real_get_app_dir
    mod2._is_writable_dir = real_is_writable

    print("=" * 70)
    print("阶段3: 非 Windows 平台的 .doc 兼容提示")
    mod3 = load_module(tmp)
    docfile = Path(tmp) / "legacy.doc"
    docfile.write_bytes(b"\xd0\xcf\x11\xe0")   # 伪 .doc 头
    mod3.IS_WINDOWS = False
    mod3.IS_MACOS = True
    try:
        mod3.FileImporter.import_students_from_file(str(docfile))
        check("macOS 读 .doc 应报错而非崩溃", False, "未抛出异常")
    except Exception as e:
        msg = str(e)
        check("macOS 读 .doc 抛出可读错误", "docx" in msg or ".doc" in msg, msg[:120])
    mod3.IS_WINDOWS = True
    mod3.IS_MACOS = False

    print("=" * 70)
    print("阶段4: 源码层面的跨平台与界面改造检查")
    check("无硬编码 'Microsoft YaHei' 界面字面量残留",
          not re.search(r"font-family: 'Microsoft YaHei'", SOURCE))
    check("字体统一走跨平台常量（无硬编码 QFont 字体名）",
          "UI_FONT_FAMILY" in SOURCE and "QFont('" not in SOURCE)
    check("matplotlib 使用平台字体栈", "plt.rcParams['font.sans-serif'] = FONT_STACK" in SOURCE)
    check("matplotlib family 使用字体栈", "family=FONT_STACK," in SOURCE)
    check("CSS 字体栈常量已注入样式表", SOURCE.count("font-family: {FONT_STACK_CSS};") >= 20,
          f"count={SOURCE.count('font-family: {FONT_STACK_CSS};')}")
    check("菜单角色已适配 macOS", "QuitRole" in SOURCE and "AboutRole" in SOURCE)
    check("Retina/高DPI 舍入策略已设置", "HighDpiScaleFactorRoundingPolicy" in SOURCE)
    check("应用显示名已设置（macOS 程序坞/菜单栏）",
          "setApplicationDisplayName" in SOURCE and "setOrganizationDomain" in SOURCE)
    check("窗口定位使用可用区域（避开菜单栏/程序坞）", "availableGeometry()" in SOURCE)
    check("数据目录可在应用内打开", "open_data_dir" in SOURCE and "QDesktopServices" in SOURCE)
    check("版本号已升到 1.3.1", mod.APP_VERSION == "1.3.1", mod.APP_VERSION)
    check("Bundle ID 常量存在（macOS 打包用）",
          re.fullmatch(r"[a-zA-Z0-9.\-]+", mod.APP_BUNDLE_ID) is not None, mod.APP_BUNDLE_ID)

    print("=" * 70)
    print("阶段5: 新界面功能在源码中的落点")
    check("学生磁贴：姓名在头像下方且不省略",
          "self.name_label" in SOURCE and "setWordWrap(True)" in SOURCE
          and "self.setMinimumHeight(206)" in SOURCE)
    check("学生磁贴不再有最大尺寸限制（避免文字重叠）",
          "setMaximumSize(250, 200)" not in SOURCE)
    check("学生管理：班级总览 → 班级详情的两级导航",
          "class ClassTile" in SOURCE and "def open_class" in SOURCE
          and "def show_class_overview" in SOURCE)
    check("点入动画（AnimatedStack 滑动 + 视差）",
          "class AnimatedStack" in SOURCE and "QEasingCurve.OutCubic" in SOURCE
          and "视差" in SOURCE)
    check("排课：月历控件与有课日期标记",
          "class MonthCalendar" in SOURCE and "class DayCell" in SOURCE
          and "def get_course_day_counts" in SOURCE)
    check("排课：点日期展开当天课程安排",
          "class CourseRow" in SOURCE and "def on_date_selected" in SOURCE
          and "_rebuild_day_panel" in SOURCE)
    check("便签：文字/表格/图片三种内容类型",
          "content_type" in SOURCE and "table_json" in SOURCE and "image_path" in SOURCE
          and "class StickerTile" in SOURCE)
    check("便签：剪贴板智能粘贴",
          "def paste_from_clipboard" in SOURCE and "def _parse_table_text" in SOURCE
          and "QKeySequence.Paste" in SOURCE)
    check("便签：图片落盘存储与清理",
          "def save_qimage_file" in SOURCE and "def delete_image_file" in SOURCE
          and "feedback_images" in SOURCE)
    check("便签墙取消日期/范围筛选（统一展示）",
          "date_filter" not in SOURCE and "class_filter" not in SOURCE
          and "category_filter" not in SOURCE)
    check("每个磁贴都有编辑入口（基类提供）",
          "edit_requested = Signal(object)" in SOURCE
          and "def extend_context_menu" in SOURCE and "self.edit_btn" in SOURCE)
    check("成绩磁贴可编辑（学生成绩编辑对话框）",
          "class StudentGradeEditDialog" in SOURCE and "def edit_student_grades" in SOURCE)

    print("=" * 70)
    print("阶段6: 防“多开多个程序窗口”闪烁的源码约束")
    # 用 AST 精确查找 setParent(None) 调用（忽略注释/文档字符串里的描述文字）
    import ast as _ast
    offenders = []
    for node in _ast.walk(_ast.parse(SOURCE)):
        if (isinstance(node, _ast.Call) and isinstance(node.func, _ast.Attribute)
                and node.func.attr == 'setParent'
                and any(isinstance(a, _ast.Constant) and a.value is None for a in node.args)):
            offenders.append(f"line {node.lineno}")
    check("代码中绝不调用 setParent(None)（会让控件退化为顶层窗口）",
          not offenders, str(offenders[:3]))
    check("磁贴/课程行/日历格构造时均传入父控件",
          "StudentTile(student, parent=self.tiles_container)" in SOURCE
          and "ClassTile(class_name, students, parent=self.class_container)" in SOURCE
          and "StickerTile(fb, parent=self.tiles_container)" in SOURCE
          and "CourseRow(course, parent=self.day_container)" in SOURCE
          and "parent=self)                 # 必须给父控件" in SOURCE)
    check("布局清理统一使用 hide + deleteLater",
          SOURCE.count("w.hide()") >= 5 and SOURCE.count("w.deleteLater()") >= 5,
          f"hide={SOURCE.count('w.hide()')} delete={SOURCE.count('w.deleteLater()')}")
    check("学生管理导航栏：上一页 + 面包屑跳转",
          "def go_back" in SOURCE and "def jump_to" in SOURCE
          and "crumb_layout" in SOURCE and "def _update_nav_bar" in SOURCE
          and "def navigate" in SOURCE)
    check("窗口标题不再重复（仅 macOS 设置显示名）",
          "if IS_MACOS:\n        # 仅 macOS 需要显示名" in SOURCE)

    print("=" * 70)
    print(f"汇总: FAIL={len(FAILED)}")
    if FAILED:
        for f in FAILED:
            print("   -", f)
        sys.exit(1)
    print("跨平台适配测试全部通过 ✅")


if __name__ == "__main__":
    main()
