#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
教学管理系统 - 综合管理平台（跨平台版：Windows / macOS / Linux）
=============================
功能模块：
1. 学生管理系统 - 磁贴式学生信息展示
2. 课程管理系统 - 仿Notion Calendar的拖拽课程管理
3. 反馈系统 - 按时间/班级/类别分类的磁贴反馈
4. 成绩管理系统 - 统计图表与个人成绩报告

技术栈：PySide6 + SQLite + matplotlib + openpyxl + python-docx

跨平台说明：
- 数据目录：优先程序目录（便携模式），不可写时回退用户数据目录
  （macOS: ~/Library/Application Support/TeachingManager；Windows: %APPDATA%\\TeachingManager）
  也可用环境变量 TMS_DATA_DIR 指定。
- 中文字体：按平台自动选择（macOS 苹方 / Windows 微软雅黑 / Linux 思源黑体）。
- 界面字体、菜单角色、Retina 高DPI、窗口定位均已按 macOS 习惯适配。
"""

import sys
import os
import json
import sqlite3
import traceback
from datetime import datetime, timedelta, date
from typing import Optional, List, Dict, Any, Tuple, Union
from dataclasses import dataclass, field, asdict
from pathlib import Path


# ==================== 跨平台适配（Windows / macOS / Linux） ====================
IS_WINDOWS = sys.platform.startswith('win')
IS_MACOS = sys.platform == 'darwin'
IS_LINUX = sys.platform.startswith('linux')


def get_platform_font_stack() -> List[str]:
    """按平台返回界面/图表的中文字体优先级。

    统一规则：**MiSans 优先**（若系统已安装或放入 fonts/ 目录），
    其次 macOS 用苹方（PingFang SC），Windows 用微软雅黑，Linux 用思源/文泉驿。
    这样既满足“MiSans 或 PingFang”的要求，又保证各平台都有可用中文字体。
    """
    misans = ['MiSans', 'MiSans VF', 'MiSans Latin', 'MiSans Normal', 'MiSans Demibold']
    if IS_MACOS:
        # macOS：苹方优先；若装过 Office 也可能有微软雅黑，作为最后兜底
        rest = ['PingFang SC', 'Hiragino Sans GB', 'Heiti SC', 'STHeiti',
                'Songti SC', 'Microsoft YaHei', 'Arial Unicode MS',
                'Helvetica Neue', 'DejaVu Sans']
    elif IS_WINDOWS:
        rest = ['PingFang SC', 'Microsoft YaHei', 'SimHei', 'Segoe UI',
                'Arial Unicode MS', 'DejaVu Sans']
    else:
        rest = ['PingFang SC', 'Noto Sans CJK SC', 'Source Han Sans SC',
                'WenQuanYi Micro Hei', 'DejaVu Sans']
    return misans + rest


FONT_STACK = get_platform_font_stack()
UI_FONT_FAMILY = FONT_STACK[0]          # QFont 用的首选家族名
# Qt 样式表 font-family 值（跨平台回退链）
FONT_STACK_CSS = ", ".join([f"'{f}'" for f in FONT_STACK[:5]] + ["sans-serif"])


def get_font_search_dirs() -> List['Path']:
    """自带字体目录：把 MiSans.ttf 等字体放进这些目录即会自动加载

    查找顺序：程序目录/fonts、数据目录/fonts、用户字体目录
    """
    dirs = [get_app_dir() / "fonts", get_user_data_dir() / "fonts"]
    if IS_MACOS:
        dirs.append(Path.home() / "Library" / "Fonts")
    elif IS_WINDOWS:
        dirs.append(Path(os.environ.get('LOCALAPPDATA', '')) / "Microsoft" / "Windows" / "Fonts")
    else:
        dirs.append(Path.home() / ".local" / "share" / "fonts")
    return dirs


def load_bundled_fonts() -> List[str]:
    """加载 fonts/ 目录内自带的字体文件（ttf/otf/ttc），返回注册成功的中文字体家族名"""
    loaded: List[str] = []
    try:
        from PySide6.QtGui import QFontDatabase
    except Exception:
        return loaded
    for d in get_font_search_dirs()[:2]:
        try:
            if not d.is_dir():
                continue
            for f in sorted(d.iterdir()):
                if f.suffix.lower() in ('.ttf', '.otf', '.ttc'):
                    fid = QFontDatabase.addApplicationFont(str(f))
                    if fid != -1:
                        loaded.extend(QFontDatabase.applicationFontFamilies(fid))
        except Exception:
            continue
    return loaded


def pick_installed_font(preference: Optional[List[str]] = None) -> Optional[str]:
    """在已安装字体中挑出第一个可用的首选字体（需 QApplication 已创建）

    支持模糊匹配：如 MiSans 的家族名可能是 "MiSans VF"、"MiSans Normal" 等。
    """
    try:
        from PySide6.QtGui import QFontDatabase
    except Exception:
        return None
    families = list(QFontDatabase.families())
    lowered = {f.lower(): f for f in families}
    for name in (preference or FONT_STACK):
        if name in families:
            return name
        if name.lower() in lowered:
            return lowered[name.lower()]
        # 前缀模糊匹配（例：'MiSans' 命中 'MiSans VF'）
        for low, real in lowered.items():
            if low.startswith(name.lower()):
                return real
    return None


def apply_ui_font(app=None) -> str:
    """应用界面字体：加载自带字体 → 选出实际可用家族 → 更新全局字体常量与 matplotlib

    必须在 QApplication 创建之后、创建主窗口之前调用。
    返回最终使用的字体家族名。
    """
    global FONT_STACK, UI_FONT_FAMILY, FONT_STACK_CSS

    bundled = load_bundled_fonts()
    chosen = pick_installed_font([*bundled, *FONT_STACK]) if bundled else pick_installed_font()
    if not chosen:
        chosen = UI_FONT_FAMILY          # 都不存在时保留首选名，交给 Qt 回退

    # 选中的字体提到最前，其余保持回退顺序
    rest = [f for f in FONT_STACK if f != chosen]
    FONT_STACK = [chosen, *rest]
    UI_FONT_FAMILY = chosen
    FONT_STACK_CSS = ", ".join([f"'{f}'" for f in FONT_STACK[:5]] + ["sans-serif"])

    # Qt 全局默认字体
    try:
        from PySide6.QtGui import QFont
        if app is not None:
            base = QFont(chosen, 10)
            base.setStyleStrategy(QFont.PreferAntialias)
            app.setFont(base)
    except Exception:
        pass

    # matplotlib 图表字体同步
    try:
        if HAS_MATPLOTLIB:
            plt.rcParams['font.sans-serif'] = FONT_STACK
            plt.rcParams['axes.unicode_minus'] = False
    except Exception:
        pass
    return chosen


APP_ICON_EMOJI = "🎓"


def _is_writable_dir(p: Path) -> bool:
    """探测目录是否可写（不存在则尝试创建）"""
    try:
        p.mkdir(parents=True, exist_ok=True)
        probe = p / f".tms_write_test_{os.getpid()}"
        probe.write_text('ok', encoding='utf-8')
        probe.unlink()
        return True
    except Exception:
        return False


def get_user_data_dir(app_name: str = "TeachingManager") -> Path:
    """平台标准用户数据目录（.app 内部只读时的落点）"""
    if IS_MACOS:
        return Path.home() / "Library" / "Application Support" / app_name
    if IS_WINDOWS:
        base = os.environ.get('APPDATA') or str(Path.home() / 'AppData' / 'Roaming')
        return Path(base) / app_name
    base = os.environ.get('XDG_DATA_HOME') or str(Path.home() / '.local' / 'share')
    return Path(base) / app_name


def get_app_dir() -> Path:
    """程序所在目录：打包后为 exe 同级 / .app 的可执行文件目录；源码运行时为脚本目录"""
    if getattr(sys, 'frozen', False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolve_data_dir() -> Tuple[Path, str]:
    """决定数据库/配置存放目录，返回 (目录, 模式说明)。

    优先级：环境变量 TMS_DATA_DIR > 便携模式（程序目录可写）> 系统用户数据目录。
    macOS 上从访达双击 .app 运行时，工作目录是 '/'、且 .app 包内只读，
    必须回退到 ~/Library/Application Support，否则无法创建数据库。
    """
    env_dir = os.environ.get('TMS_DATA_DIR')
    if env_dir:
        p = Path(env_dir).expanduser()
        if _is_writable_dir(p):
            return p, "环境变量 TMS_DATA_DIR 指定"

    app_dir = get_app_dir()
    if _is_writable_dir(app_dir):
        return app_dir, "便携模式（程序目录）"

    user_dir = get_user_data_dir()
    _is_writable_dir(user_dir)
    return user_dir, "系统用户数据目录"


DATA_DIR, DATA_DIR_MODE = resolve_data_dir()


# ==================== 第三方库导入 ====================
try:
    from PySide6.QtWidgets import (
        QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
        QGridLayout, QScrollArea, QFrame, QLabel, QPushButton, QDialog,
        QLineEdit, QTextEdit, QComboBox, QDateEdit, QTimeEdit, QSpinBox,
        QDoubleSpinBox, QMessageBox, QFileDialog, QStackedWidget, QSizePolicy,
        QGraphicsView, QGraphicsScene, QGraphicsRectItem, QGraphicsTextItem,
        QGraphicsItem, QToolBar, QMenu, QTabWidget, QSplitter, QGroupBox,
        QFormLayout, QTableWidget, QTableWidgetItem, QHeaderView, QAbstractItemView,
        QLayout, QWidgetItem,
        QCheckBox, QRadioButton, QButtonGroup, QListWidget, QListWidgetItem,
        QColorDialog, QInputDialog, QProgressBar, QStatusBar, QSlider, QToolButton,
        QStyle, QStyleFactory, QGraphicsDropShadowEffect, QGraphicsOpacityEffect,
        QButtonGroup, QWidgetAction, QSizeGrip, QRubberBand, QDialogButtonBox,
    )
    from PySide6.QtCore import (
        Qt, QTimer, QPropertyAnimation, QEasingCurve, QRectF, QPointF,
        QSize, Signal, Slot, QDateTime, QDate, QTime, QObject, QEvent,
        QParallelAnimationGroup, QSequentialAnimationGroup, QPauseAnimation,
        QAbstractAnimation, QRect, QMargins, QByteArray, QDataStream, QIODevice,
        QMimeData, QPoint, QModelIndex, QPersistentModelIndex, QAbstractItemModel,
        QSortFilterProxyModel, QRegularExpression, QThread, QMetaObject, Q_ARG,
        QUrl, QStandardPaths,
    )
    from PySide6.QtGui import (
        QColor, QPainter, QPen, QBrush, QLinearGradient, QRadialGradient,
        QFont, QFontMetrics, QIcon, QPixmap, QImage, QCursor, QKeySequence,
        QShortcut, QAction, QPalette, QTextCursor, QTextFormat, QTextCharFormat,
        QSyntaxHighlighter, QGuiApplication, QScreen, QTransform, QPolygonF,
        QPainterPath, QRegion, QCloseEvent, QResizeEvent, QDragEnterEvent,
        QDropEvent, QMouseEvent, QWheelEvent, QKeyEvent, QShowEvent, QHideEvent,
        QFocusEvent, QContextMenuEvent, QInputMethodEvent, QTabletEvent,
        QDesktopServices,
    )
    from PySide6.QtCharts import (
        QChart, QChartView, QPieSeries, QPieSlice, QBarSeries, QBarSet,
        QBarCategoryAxis, QValueAxis, QLineSeries, QAreaSeries, QScatterSeries,
        QSplineSeries, QCategoryAxis, QLegend, QAbstractSeries, QAbstractAxis,
        QDateTimeAxis, QLogValueAxis, QHorizontalBarSeries, QPercentBarSeries,
        QStackedBarSeries, QBoxPlotSeries, QCandlestickSeries,
    )
    HAS_PYSIDE6 = True
except ImportError:
    HAS_PYSIDE6 = False
    print("错误：需要安装PySide6。请运行: pip install PySide6")

try:
    import matplotlib
    matplotlib.use('QtAgg')
    from matplotlib.backends.backend_qtagg import FigureCanvasQTAgg as FigureCanvas
    from matplotlib.figure import Figure
    import matplotlib.pyplot as plt
    # 中文字体支持：默认 DejaVu Sans 不含中文字形，会导致图表标题/坐标轴中文显示为方块
    plt.rcParams['font.sans-serif'] = FONT_STACK
    plt.rcParams['axes.unicode_minus'] = False
    HAS_MATPLOTLIB = True
except ImportError:
    HAS_MATPLOTLIB = False
    print("警告：matplotlib未安装，图表功能将受限。请运行: pip install matplotlib")

try:
    import openpyxl
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False
    print("警告：openpyxl未安装，无法读取xlsx文件。请运行: pip install openpyxl")

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False
    print("警告：python-docx未安装，无法读取docx文件。请运行: pip install python-docx")

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False
    print("警告：xlrd未安装，无法读取xls文件。请运行: pip install xlrd")

import threading
import subprocess
import tempfile
import shutil
import re
import random
import string
import dataclasses
import uuid
import calendar


# ==================== 常量与配置 ====================
APP_NAME = "教学管理系统"
APP_VERSION = "1.4.0"
APP_BUNDLE_ID = "com.teaching.manager"
# 跨平台数据目录：便携模式落在程序目录；.app 只读时自动落到用户数据目录
DB_PATH = DATA_DIR / "teaching_management.db"
CONFIG_PATH = DATA_DIR / "config.json"
IMAGES_DIR = DATA_DIR / "feedback_images"     # 便签图片存放目录

# 便签可选底色（贴纸风格柔和色）
STICKER_COLORS = [
    "#FFF6C9", "#FFE0E0", "#DFF5E1", "#DCEBFF", "#F0E4FF",
    "#FFE9D6", "#D9F7F5", "#F5F0DC", "#E8F0DE", "#FCE1F0",
]

# 颜色方案 - 现代化设计
COLORS = {
    "primary": "#4F6EF7",       # 主色调 - 蓝色
    "primary_dark": "#3A54D4",
    "primary_light": "#E8ECFF",
    "secondary": "#6C5CE7",     # 紫色
    "success": "#00B894",       # 绿色
    "warning": "#FDCB6E",       # 黄色
    "danger": "#FF6B6B",        # 红色
    "info": "#74B9FF",          # 浅蓝
    "background": "#F5F6FA",    # 背景色
    "card_bg": "#FFFFFF",       # 卡片背景
    "text_primary": "#2D3436",  # 主文本
    "text_secondary": "#636E72", # 次要文本
    "border": "#E0E0E0",        # 边框
    "shadow": "#B0BEC5",        # 阴影
    "hover": "#F0F2FF",         # 悬停色
    "selected": "#DFE6FF",      # 选中色
}

# 课程磁贴颜色方案
COURSE_COLORS = [
    "#4F6EF7", "#00B894", "#FF6B6B", "#FDCB6E", "#6C5CE7",
    "#74B9FF", "#E17055", "#00CEC9", "#A29BFE", "#FD79A8",
    "#FAB1A0", "#81ECEC", "#55EFC4", "#FFEAA7", "#DFE6E9",
    "#B2BEC3", "#636E72", "#2D3436", "#0984E3", "#6C5CE7",
]

# 文件类型支持
SUPPORTED_FILE_TYPES = "所有支持的文件 (*.xlsx *.xls *.xlsm *.xlm *.docx *.doc *.csv);;Excel文件 (*.xlsx *.xls *.xlsm *.xlm);;Word文件 (*.docx *.doc);;CSV文件 (*.csv)"
IMAGE_FILE_TYPES = "图片 (*.png *.jpg *.jpeg *.bmp *.gif *.webp);;所有文件 (*)"


# ==================== 便签图片存储 ====================
def ensure_images_dir() -> Path:
    """确保便签图片目录存在"""
    try:
        IMAGES_DIR.mkdir(parents=True, exist_ok=True)
    except Exception:
        pass
    return IMAGES_DIR


def image_abs_path(name: str) -> Path:
    """由数据库中的文件名得到绝对路径"""
    if not name:
        return Path()
    p = Path(name)
    return p if p.is_absolute() else (IMAGES_DIR / p)


def save_image_file(src_path: str, keep_ext: bool = True) -> str:
    """把外部图片文件复制进图片目录，返回保存后的文件名"""
    ensure_images_dir()
    src = Path(src_path)
    ext = src.suffix.lower() if keep_ext and src.suffix else '.png'
    if ext not in ('.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'):
        ext = '.png'
    name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}{ext}"
    shutil.copyfile(str(src), str(IMAGES_DIR / name))
    return name


def save_qimage_file(qimage) -> str:
    """把 QImage（如剪贴板粘贴的图片）保存为 PNG，返回文件名"""
    ensure_images_dir()
    name = f"{datetime.now().strftime('%Y%m%d_%H%M%S')}_{uuid.uuid4().hex[:8]}.png"
    target = IMAGES_DIR / name
    if not qimage.save(str(target), "PNG"):
        raise ValueError("图片保存失败")
    return name


def delete_image_file(name: str) -> bool:
    """删除便签图片文件（忽略缺失）"""
    try:
        p = image_abs_path(name)
        if p and p.exists() and p.is_file():
            p.unlink()
            return True
    except Exception:
        pass
    return False


# ==================== 数据模型 ====================
@dataclass
class Student:
    """学生数据模型"""
    id: Optional[int] = None
    student_no: str = ""
    name: str = ""
    class_name: str = ""
    tags: List[str] = field(default_factory=list)       # 标签列表
    titles: List[str] = field(default_factory=list)     # 头衔列表
    avatar_color: str = "#4F6EF7"
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        d['tags'] = json.dumps(self.tags, ensure_ascii=False)
        d['titles'] = json.dumps(self.titles, ensure_ascii=False)
        return d

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'Student':
        tags = json.loads(d.get('tags', '[]')) if isinstance(d.get('tags'), str) else d.get('tags', [])
        titles = json.loads(d.get('titles', '[]')) if isinstance(d.get('titles'), str) else d.get('titles', [])
        return cls(
            id=d.get('id'),
            student_no=d.get('student_no', ''),
            name=d.get('name', ''),
            class_name=d.get('class_name', ''),
            tags=tags,
            titles=titles,
            avatar_color=d.get('avatar_color', '#4F6EF7'),
            created_at=d.get('created_at', ''),
            updated_at=d.get('updated_at', ''),
        )


@dataclass
class Course:
    """课程数据模型"""
    id: Optional[int] = None
    name: str = ""
    course_date: str = ""          # 日期 YYYY-MM-DD
    start_time: str = "08:00"      # 开始时间 HH:MM
    end_time: str = "09:00"        # 结束时间 HH:MM
    description: str = ""
    color: str = "#4F6EF7"
    location: str = ""
    teacher: str = ""
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())

    @property
    def duration_minutes(self) -> int:
        """课程持续时间（分钟）"""
        try:
            sh, sm = map(int, self.start_time.split(':'))
            eh, em = map(int, self.end_time.split(':'))
            return (eh * 60 + em) - (sh * 60 + sm)
        except (ValueError, AttributeError):
            return 60

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'Course':
        return cls(**d)


@dataclass
class Feedback:
    """便签/反馈数据模型（剪贴板式便签：文字 / 表格 / 图片）"""
    id: Optional[int] = None
    date: str = field(default_factory=lambda: date.today().isoformat())
    time: str = field(default_factory=lambda: datetime.now().strftime("%H:%M"))
    class_name: str = ""
    category: str = "其他"          # 保留字段（新界面不再强制分类）
    content: str = ""              # 文字内容
    student_name: str = ""
    student_id: Optional[int] = None
    status: str = "待处理"          # 保留字段
    priority: int = 1              # 保留字段
    # ---- 便签扩展字段 ----
    content_type: str = "text"     # text | table | image
    table_json: str = ""           # content_type=table 时的表格数据 {"headers":[], "rows":[[]]}
    image_path: str = ""           # content_type=image 时的图片文件名（位于数据目录 feedback_images/）
    image_caption: str = ""        # 图片说明
    color: str = "#FFF6C9"         # 便签底色
    pinned: int = 0                # 置顶 0/1
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())
    updated_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'Feedback':
        # 兼容旧数据：缺失字段使用默认值
        fields = {f.name for f in dataclasses.fields(cls)}
        clean = {k: v for k, v in d.items() if k in fields}
        for key, default in (('content_type', 'text'), ('table_json', ''),
                             ('image_path', ''), ('image_caption', ''),
                             ('color', '#FFF6C9'), ('pinned', 0)):
            if clean.get(key) is None:
                clean[key] = default
        return cls(**clean)

    def table_data(self) -> Dict[str, Any]:
        """解析表格 JSON，失败返回空表"""
        if not self.table_json:
            return {'headers': [], 'rows': []}
        try:
            data = json.loads(self.table_json)
            return {'headers': list(data.get('headers', [])),
                    'rows': [list(r) for r in data.get('rows', [])]}
        except Exception:
            return {'headers': [], 'rows': []}

    def plain_text(self) -> str:
        """便于搜索/复制的纯文本表示"""
        if self.content_type == 'table':
            t = self.table_data()
            lines = ['\t'.join(str(h) for h in t['headers'])]
            lines += ['\t'.join(str(c) for c in row) for row in t['rows']]
            return '\n'.join(lines)
        if self.content_type == 'image':
            return self.image_caption or '[图片]'
        return self.content


@dataclass
class Grade:
    """成绩数据模型"""
    id: Optional[int] = None
    student_id: int = 0
    student_name: str = ""
    course_name: str = ""
    score: float = 0.0
    exam_date: str = field(default_factory=lambda: date.today().isoformat())
    exam_type: str = "期中"          # 期中、期末、月考、周测、其他
    full_score: float = 100.0
    class_rank: Optional[int] = None
    grade_rank: Optional[int] = None
    comment: str = ""
    created_at: str = field(default_factory=lambda: datetime.now().isoformat())

    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'Grade':
        return cls(**d)


# ==================== 数据库管理器 ====================
class DatabaseManager:
    """SQLite数据库管理器 - 负责所有数据持久化"""

    _instance = None

    def __new__(cls, *args, **kwargs):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
            cls._instance._initialized = False
        return cls._instance

    def __init__(self, db_path: str = str(DB_PATH)):
        if self._initialized:
            return
        self.db_path = db_path
        self.conn: Optional[sqlite3.Connection] = None
        self._lock = threading.Lock()
        self._initialized = True
        self._init_db()

    def _get_connection(self) -> sqlite3.Connection:
        """获取数据库连接（线程安全）"""
        if self.conn is None:
            self.conn = sqlite3.connect(self.db_path, check_same_thread=False)
            self.conn.row_factory = sqlite3.Row
            self.conn.execute("PRAGMA foreign_keys = ON")
            self.conn.execute("PRAGMA journal_mode = WAL")
        return self.conn

    def _init_db(self):
        """初始化数据库表"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()

            # 学生表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS students (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_no TEXT UNIQUE NOT NULL,
                    name TEXT NOT NULL,
                    class_name TEXT DEFAULT '',
                    tags TEXT DEFAULT '[]',
                    titles TEXT DEFAULT '[]',
                    avatar_color TEXT DEFAULT '#4F6EF7',
                    created_at TEXT DEFAULT '',
                    updated_at TEXT DEFAULT ''
                )
            ''')

            # 课程表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS courses (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL,
                    course_date TEXT NOT NULL,
                    start_time TEXT DEFAULT '08:00',
                    end_time TEXT DEFAULT '09:00',
                    description TEXT DEFAULT '',
                    color TEXT DEFAULT '#4F6EF7',
                    location TEXT DEFAULT '',
                    teacher TEXT DEFAULT '',
                    created_at TEXT DEFAULT '',
                    updated_at TEXT DEFAULT ''
                )
            ''')

            # 反馈/便签表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS feedbacks (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    date TEXT NOT NULL,
                    time TEXT DEFAULT '',
                    class_name TEXT DEFAULT '',
                    category TEXT DEFAULT '其他',
                    content TEXT DEFAULT '',
                    student_name TEXT DEFAULT '',
                    student_id INTEGER,
                    status TEXT DEFAULT '待处理',
                    priority INTEGER DEFAULT 1,
                    content_type TEXT DEFAULT 'text',
                    table_json TEXT DEFAULT '',
                    image_path TEXT DEFAULT '',
                    image_caption TEXT DEFAULT '',
                    color TEXT DEFAULT '#FFF6C9',
                    pinned INTEGER DEFAULT 0,
                    created_at TEXT DEFAULT '',
                    updated_at TEXT DEFAULT ''
                )
            ''')

            # 成绩表
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS grades (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    student_id INTEGER NOT NULL,
                    student_name TEXT DEFAULT '',
                    course_name TEXT NOT NULL,
                    score REAL DEFAULT 0,
                    exam_date TEXT NOT NULL,
                    exam_type TEXT DEFAULT '期中',
                    full_score REAL DEFAULT 100,
                    class_rank INTEGER,
                    grade_rank INTEGER,
                    comment TEXT DEFAULT '',
                    created_at TEXT DEFAULT ''
                )
            ''')

            # ---- 学情管理（v1.4.0）：自定义字段 + 每个学生的填写值 ----
            # 说明：学情数据只在本模块使用，**不会**出现在学生管理的磁贴/资料气泡里
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS profile_fields (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE,
                    position INTEGER DEFAULT 0,
                    created_at TEXT DEFAULT ''
                )
            ''')
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS profile_values (
                    student_id INTEGER NOT NULL,
                    field_id INTEGER NOT NULL,
                    value TEXT DEFAULT '',
                    updated_at TEXT DEFAULT '',
                    PRIMARY KEY (student_id, field_id)
                )
            ''')

            # 创建索引
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_students_name ON students(name)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_courses_date ON courses(course_date)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_feedbacks_date ON feedbacks(date)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_grades_student ON grades(student_id)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_grades_course ON grades(course_name)')
            cursor.execute('CREATE INDEX IF NOT EXISTS idx_profile_values_student '
                           'ON profile_values(student_id)')

            # ---- 老版本数据库自动迁移（v1.3.0 便签新增字段）----
            self._migrate_schema(cursor)

            conn.commit()

    def _migrate_schema(self, cursor):
        """为已存在的旧数据库补齐新增列（幂等）"""
        migrations = {
            'feedbacks': [
                ('content_type', "TEXT DEFAULT 'text'"),
                ('table_json', "TEXT DEFAULT ''"),
                ('image_path', "TEXT DEFAULT ''"),
                ('image_caption', "TEXT DEFAULT ''"),
                ('color', "TEXT DEFAULT '#FFF6C9'"),
                ('pinned', "INTEGER DEFAULT 0"),
            ],
        }
        for table, columns in migrations.items():
            try:
                cursor.execute(f'PRAGMA table_info({table})')
                existing = {row[1] for row in cursor.fetchall()}
                if not existing:
                    continue
                for col, decl in columns:
                    if col not in existing:
                        cursor.execute(f'ALTER TABLE {table} ADD COLUMN {col} {decl}')
                        print(f"[迁移] {table} 表新增字段: {col}")
            except sqlite3.Error as e:
                print(f"[迁移] {table} 表迁移失败: {e}")

    # ---- 学生操作 ----
    def add_student(self, student: Student) -> int:
        """添加学生，返回ID"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = student.to_dict()
            d.pop('id', None)
            cursor.execute('''
                INSERT INTO students (student_no, name, class_name, tags, titles, avatar_color, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            ''', (d['student_no'], d['name'], d['class_name'], d['tags'], d['titles'],
                  d['avatar_color'], d['created_at'], d['updated_at']))
            conn.commit()
            return cursor.lastrowid

    def update_student(self, student: Student) -> bool:
        """更新学生信息"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = student.to_dict()
            d['updated_at'] = datetime.now().isoformat()
            cursor.execute('''
                UPDATE students SET name=?, class_name=?, tags=?, titles=?,
                    avatar_color=?, updated_at=?
                WHERE id=?
            ''', (d['name'], d['class_name'], d['tags'], d['titles'],
                  d['avatar_color'], d['updated_at'], d['id']))
            conn.commit()
            return cursor.rowcount > 0

    def delete_student(self, student_id: int) -> bool:
        """删除学生（同时清理其学情填写值，避免残留脏数据）"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM profile_values WHERE student_id=?', (student_id,))
            cursor.execute('DELETE FROM students WHERE id=?', (student_id,))
            conn.commit()
            return cursor.rowcount > 0

    def get_all_students(self) -> List[Student]:
        """获取所有学生"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM students ORDER BY class_name, name')
            rows = cursor.fetchall()
            return [Student.from_dict(dict(row)) for row in rows]

    def get_student_by_id(self, student_id: int) -> Optional[Student]:
        """根据ID获取学生"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM students WHERE id=?', (student_id,))
            row = cursor.fetchone()
            if row:
                return Student.from_dict(dict(row))
            return None

    def search_students(self, keyword: str) -> List[Student]:
        """搜索学生"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            like = f'%{keyword}%'
            cursor.execute('''
                SELECT * FROM students WHERE name LIKE ? OR student_no LIKE ? OR class_name LIKE ?
                ORDER BY name
            ''', (like, like, like))
            rows = cursor.fetchall()
            return [Student.from_dict(dict(row)) for row in rows]

    # ---- 学情管理（自定义字段 + 每个学生的填写值）----
    def get_profile_fields(self) -> List[Dict[str, Any]]:
        """学情表格的自定义列（按 position 排序）"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT id, name, position FROM profile_fields '
                           'ORDER BY position, id')
            return [{'id': r[0], 'name': r[1], 'position': r[2]} for r in cursor.fetchall()]

    def add_profile_field(self, name: str) -> Optional[int]:
        """新增学情字段（列）。已存在同名则返回其 id"""
        name = (name or '').strip()
        if not name:
            return None
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM profile_fields WHERE name=?', (name,))
            row = cursor.fetchone()
            if row:
                return row[0]
            cursor.execute('SELECT COALESCE(MAX(position), 0) + 1 FROM profile_fields')
            pos = cursor.fetchone()[0]
            cursor.execute('INSERT INTO profile_fields (name, position, created_at) '
                           'VALUES (?, ?, ?)', (name, pos, datetime.now().isoformat()))
            conn.commit()
            return cursor.lastrowid

    def rename_profile_field(self, field_id: int, new_name: str) -> bool:
        """重命名字段；若新名称已被占用则返回 False"""
        new_name = (new_name or '').strip()
        if not new_name:
            return False
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT id FROM profile_fields WHERE name=? AND id<>?',
                           (new_name, field_id))
            if cursor.fetchone():
                return False
            cursor.execute('UPDATE profile_fields SET name=? WHERE id=?', (new_name, field_id))
            conn.commit()
            return cursor.rowcount > 0

    def delete_profile_field(self, field_id: int) -> bool:
        """删除字段（同时删除该列所有学生的填写值）"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM profile_values WHERE field_id=?', (field_id,))
            cursor.execute('DELETE FROM profile_fields WHERE id=?', (field_id,))
            conn.commit()
            return cursor.rowcount > 0

    def get_profile_values(self) -> Dict[Tuple[int, int], str]:
        """全部学情填写值：{(student_id, field_id): value}"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT student_id, field_id, value FROM profile_values')
            return {(r[0], r[1]): (r[2] or '') for r in cursor.fetchall()}

    def set_profile_value(self, student_id: int, field_id: int, value: str) -> bool:
        """写入/更新某个学生在某字段上的填写值"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO profile_values (student_id, field_id, value, updated_at)
                VALUES (?, ?, ?, ?)
                ON CONFLICT(student_id, field_id) DO UPDATE SET
                    value=excluded.value, updated_at=excluded.updated_at
            ''', (student_id, field_id, value or '', datetime.now().isoformat()))
            conn.commit()
            return True

    def get_student_profile(self, student_id: int) -> Dict[str, str]:
        """某学生的学情档案：{字段名: 值}（对学生/班级磁贴无副作用）"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''
                SELECT f.name, COALESCE(v.value, '')
                FROM profile_fields f
                LEFT JOIN profile_values v
                       ON v.field_id = f.id AND v.student_id = ?
                ORDER BY f.position, f.id
            ''', (student_id,))
            return {r[0]: (r[1] or '') for r in cursor.fetchall()}

    def get_profile_matrix(self) -> Dict[str, Any]:
        """整张学情表：字段列表 + 每名学生每列的值（供界面/导出使用）"""
        fields = self.get_profile_fields()
        values = self.get_profile_values()
        students = self.get_all_students()
        rows = []
        for s in students:
            rows.append({
                'student_id': s.id,
                'name': s.name,
                'student_no': s.student_no,
                'class_name': s.class_name,
                'values': {f['id']: values.get((s.id, f['id']), '') for f in fields},
            })
        return {'fields': fields, 'students': rows}

    # ---- 课程操作 ----
    def add_course(self, course: Course) -> int:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = course.to_dict()
            d.pop('id', None)
            cursor.execute('''
                INSERT INTO courses (name, course_date, start_time, end_time, description,
                    color, location, teacher, created_at, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (d['name'], d['course_date'], d['start_time'], d['end_time'],
                  d['description'], d['color'], d['location'], d['teacher'],
                  d['created_at'], d['updated_at']))
            conn.commit()
            return cursor.lastrowid

    def update_course(self, course: Course) -> bool:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = course.to_dict()
            d['updated_at'] = datetime.now().isoformat()
            cursor.execute('''
                UPDATE courses SET name=?, course_date=?, start_time=?, end_time=?,
                    description=?, color=?, location=?, teacher=?, updated_at=?
                WHERE id=?
            ''', (d['name'], d['course_date'], d['start_time'], d['end_time'],
                  d['description'], d['color'], d['location'], d['teacher'],
                  d['updated_at'], d['id']))
            conn.commit()
            return cursor.rowcount > 0

    def delete_course(self, course_id: int) -> bool:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM courses WHERE id=?', (course_id,))
            conn.commit()
            return cursor.rowcount > 0

    def get_courses_by_date(self, date_str: str) -> List[Course]:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM courses WHERE course_date=? ORDER BY start_time', (date_str,))
            rows = cursor.fetchall()
            return [Course.from_dict(dict(row)) for row in rows]

    def get_all_courses(self) -> List[Course]:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM courses ORDER BY course_date, start_time')
            rows = cursor.fetchall()
            return [Course.from_dict(dict(row)) for row in rows]

    # ---- 日历用：按月份统计每天课程数 ----
    def get_course_day_counts(self, year: int, month: int) -> Dict[str, int]:
        """返回 {'YYYY-MM-DD': 课程数}，用于月历标记有课的日期"""
        prefix = f"{year:04d}-{month:02d}-%"
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute(
                'SELECT course_date, COUNT(*) FROM courses WHERE course_date LIKE ? GROUP BY course_date',
                (prefix,))
            return {row[0]: row[1] for row in cursor.fetchall()}

    def get_all_course_dates(self) -> List[str]:
        """所有有课的日期（升序）"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT DISTINCT course_date FROM courses ORDER BY course_date')
            return [row[0] for row in cursor.fetchall()]

    # ---- 反馈/便签操作 ----
    _FEEDBACK_COLUMNS = ('date', 'time', 'class_name', 'category', 'content', 'student_name',
                         'student_id', 'status', 'priority', 'content_type', 'table_json',
                         'image_path', 'image_caption', 'color', 'pinned')

    def add_feedback(self, feedback: Feedback) -> int:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = feedback.to_dict()
            d.pop('id', None)
            cols = list(self._FEEDBACK_COLUMNS) + ['created_at', 'updated_at']
            placeholders = ', '.join('?' for _ in cols)
            cursor.execute(
                f"INSERT INTO feedbacks ({', '.join(cols)}) VALUES ({placeholders})",
                tuple(d.get(c, '') for c in cols))
            conn.commit()
            return cursor.lastrowid

    def update_feedback(self, feedback: Feedback) -> bool:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = feedback.to_dict()
            d['updated_at'] = datetime.now().isoformat()
            cols = list(self._FEEDBACK_COLUMNS) + ['updated_at']
            assignments = ', '.join(f'{c}=?' for c in cols)
            cursor.execute(
                f"UPDATE feedbacks SET {assignments} WHERE id=?",
                tuple(d.get(c, '') for c in cols) + (d['id'],))
            conn.commit()
            return cursor.rowcount > 0

    def set_feedback_pinned(self, feedback_id: int, pinned: bool) -> bool:
        """置顶/取消置顶便签"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('UPDATE feedbacks SET pinned=? WHERE id=?',
                           (1 if pinned else 0, feedback_id))
            conn.commit()
            return cursor.rowcount > 0

    def delete_feedback(self, feedback_id: int) -> bool:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM feedbacks WHERE id=?', (feedback_id,))
            conn.commit()
            return cursor.rowcount > 0

    def get_all_feedbacks(self) -> List[Feedback]:
        """全部便签：置顶优先，其余按创建时间倒序（统一展示，不按日期筛选）"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''SELECT * FROM feedbacks
                              ORDER BY pinned DESC, date DESC, time DESC, id DESC''')
            rows = cursor.fetchall()
            return [Feedback.from_dict(dict(row)) for row in rows]

    def search_feedbacks(self, keyword: str) -> List[Feedback]:
        """按关键词搜索便签（文字内容/表格/图片说明）"""
        kw = f'%{keyword}%'
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('''SELECT * FROM feedbacks
                              WHERE content LIKE ? OR table_json LIKE ?
                                 OR image_caption LIKE ? OR image_path LIKE ?
                              ORDER BY pinned DESC, date DESC, time DESC, id DESC''',
                           (kw, kw, kw, kw))
            rows = cursor.fetchall()
            return [Feedback.from_dict(dict(row)) for row in rows]

    def get_feedbacks_by_filter(self, date_str: Optional[str] = None,
                                 class_name: Optional[str] = None,
                                 category: Optional[str] = None) -> List[Feedback]:
        """（保留的旧接口）按日期/班级/类别过滤"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            query = 'SELECT * FROM feedbacks WHERE 1=1'
            params = []
            if date_str:
                query += ' AND date=?'
                params.append(date_str)
            if class_name:
                query += ' AND class_name=?'
                params.append(class_name)
            if category:
                query += ' AND category=?'
                params.append(category)
            query += ' ORDER BY pinned DESC, date DESC, time DESC'
            cursor.execute(query, params)
            rows = cursor.fetchall()
            return [Feedback.from_dict(dict(row)) for row in rows]

    # ---- 成绩操作 ----
    def add_grade(self, grade: Grade) -> int:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = grade.to_dict()
            d.pop('id', None)
            cursor.execute('''
                INSERT INTO grades (student_id, student_name, course_name, score,
                    exam_date, exam_type, full_score, class_rank, grade_rank, comment, created_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (d['student_id'], d['student_name'], d['course_name'], d['score'],
                  d['exam_date'], d['exam_type'], d['full_score'], d['class_rank'],
                  d['grade_rank'], d['comment'], d['created_at']))
            conn.commit()
            return cursor.lastrowid

    def update_grade(self, grade: Grade) -> bool:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            d = grade.to_dict()
            cursor.execute('''
                UPDATE grades SET student_id=?, student_name=?, course_name=?, score=?,
                    exam_date=?, exam_type=?, full_score=?, class_rank=?, grade_rank=?,
                    comment=?
                WHERE id=?
            ''', (d['student_id'], d['student_name'], d['course_name'], d['score'],
                  d['exam_date'], d['exam_type'], d['full_score'], d['class_rank'],
                  d['grade_rank'], d['comment'], d['id']))
            conn.commit()
            return cursor.rowcount > 0

    def delete_grade(self, grade_id: int) -> bool:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM grades WHERE id=?', (grade_id,))
            conn.commit()
            return cursor.rowcount > 0

    def delete_grades_by_student(self, student_id: int) -> int:
        """级联删除某学生名下的全部成绩，返回删除条数"""
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('DELETE FROM grades WHERE student_id=?', (student_id,))
            conn.commit()
            return cursor.rowcount

    def get_grades_by_student(self, student_id: int) -> List[Grade]:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM grades WHERE student_id=? ORDER BY exam_date DESC', (student_id,))
            rows = cursor.fetchall()
            return [Grade.from_dict(dict(row)) for row in rows]

    def get_all_grades(self) -> List[Grade]:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM grades ORDER BY exam_date DESC')
            rows = cursor.fetchall()
            return [Grade.from_dict(dict(row)) for row in rows]

    def get_grades_by_course(self, course_name: str) -> List[Grade]:
        with self._lock:
            conn = self._get_connection()
            cursor = conn.cursor()
            cursor.execute('SELECT * FROM grades WHERE course_name=? ORDER BY exam_date DESC', (course_name,))
            rows = cursor.fetchall()
            return [Grade.from_dict(dict(row)) for row in rows]

    def close(self):
        if self.conn:
            self.conn.close()
            self.conn = None


# ==================== 文件导入器 ====================
class FileImporter:
    """文件导入器 - 从各种格式的文件中提取数据"""

    @staticmethod
    def import_students_from_file(file_path: str) -> List[Dict[str, Any]]:
        """从文件导入学生信息
        支持格式：xlsx, xls, docx, csv
        期望列：学号, 姓名, 班级, 标签, 头衔
        """
        ext = Path(file_path).suffix.lower()
        students = []

        try:
            if ext in ['.xlsx', '.xlsm', '.xlm']:
                students = FileImporter._read_students_from_excel(file_path)
            elif ext == '.xls':
                students = FileImporter._read_students_from_xls(file_path)
            elif ext == '.docx':
                students = FileImporter._read_students_from_docx(file_path)
            elif ext == '.csv':
                students = FileImporter._read_students_from_csv(file_path)
            elif ext == '.doc':
                students = FileImporter._read_students_from_doc(file_path)
            else:
                raise ValueError(f"不支持的文件格式: {ext}")
        except Exception as e:
            raise ValueError(f"读取文件失败: {str(e)}")

        return students

    @staticmethod
    def import_grades_from_file(file_path: str) -> List[Dict[str, Any]]:
        """从文件导入成绩信息"""
        ext = Path(file_path).suffix.lower()
        grades = []

        try:
            if ext in ['.xlsx', '.xlsm', '.xlm']:
                grades = FileImporter._read_grades_from_excel(file_path)
            elif ext == '.xls':
                grades = FileImporter._read_grades_from_xls(file_path)
            elif ext == '.csv':
                grades = FileImporter._read_grades_from_csv(file_path)
            else:
                raise ValueError(f"不支持的文件格式: {ext}")
        except Exception as e:
            raise ValueError(f"读取成绩文件失败: {str(e)}")

        return grades

    @staticmethod
    def _read_students_from_excel(file_path: str) -> List[Dict[str, Any]]:
        """从Excel读取学生信息"""
        if not HAS_OPENPYXL:
            raise ImportError("需要安装openpyxl")

        wb = load_workbook(file_path, data_only=True)
        ws = wb.active

        # 识别表头
        headers = []
        for cell in ws[1]:
            headers.append(str(cell.value or '').strip())

        # 映射列
        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if '学号' in h_lower or '编号' in h_lower or 'id' == h_lower or 'no' in h_lower:
                col_map['student_no'] = i
            elif '姓名' in h_lower or '名字' in h_lower or 'name' in h_lower:
                col_map['name'] = i
            elif '班级' in h_lower or 'class' in h_lower:
                col_map['class_name'] = i
            elif '标签' in h_lower or 'tag' in h_lower:
                col_map['tags'] = i
            elif '头衔' in h_lower or 'title' in h_lower or '称号' in h_lower:
                col_map['titles'] = i

        if 'name' not in col_map:
            # 尝试猜测：第一列为学号，第二列为姓名
            if len(headers) >= 2:
                col_map['student_no'] = 0
                col_map['name'] = 1
            else:
                raise ValueError("无法识别学生姓名列")

        students = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or all(v is None or str(v).strip() == '' for v in row):
                continue

            name = FileImporter._cell_text(
                row[col_map['name']] if 'name' in col_map and col_map['name'] < len(row) else None)
            if not name:
                continue

            student_no = FileImporter._cell_text(
                row[col_map['student_no']] if 'student_no' in col_map and col_map['student_no'] < len(row) else None,
                default=name)
            class_name = FileImporter._cell_text(
                row[col_map['class_name']] if 'class_name' in col_map and col_map['class_name'] < len(row) else None)
            tags_str = FileImporter._cell_text(
                row[col_map['tags']] if 'tags' in col_map and col_map['tags'] < len(row) else None)
            titles_str = FileImporter._cell_text(
                row[col_map['titles']] if 'titles' in col_map and col_map['titles'] < len(row) else None)

            tags = [t.strip() for t in tags_str.replace('，', ',').split(',') if t.strip()] if tags_str else []
            titles = [t.strip() for t in titles_str.replace('，', ',').split(',') if t.strip()] if titles_str else []

            students.append({
                'student_no': student_no,
                'name': name,
                'class_name': class_name,
                'tags': tags,
                'titles': titles,
                'avatar_color': random.choice(COURSE_COLORS),
            })

        return students

    @staticmethod
    def _read_students_from_xls(file_path: str) -> List[Dict[str, Any]]:
        """从旧版Excel读取学生信息"""
        if not HAS_XLRD:
            raise ImportError("需要安装xlrd")

        wb = xlrd.open_workbook(file_path)
        ws = wb.sheet_by_index(0)

        headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]
        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if '学号' in h_lower or '编号' in h_lower or 'id' == h_lower:
                col_map['student_no'] = i
            elif '姓名' in h_lower or '名字' in h_lower or 'name' in h_lower:
                col_map['name'] = i
            elif '班级' in h_lower or 'class' in h_lower:
                col_map['class_name'] = i
            elif '标签' in h_lower or 'tag' in h_lower:
                col_map['tags'] = i
            elif '头衔' in h_lower or 'title' in h_lower:
                col_map['titles'] = i

        if 'name' not in col_map:
            if ws.ncols >= 2:
                col_map['student_no'] = 0
                col_map['name'] = 1
            else:
                raise ValueError("无法识别学生姓名列")

        students = []
        for r in range(1, ws.nrows):
            name = str(ws.cell_value(r, col_map.get('name', 1))).strip()
            if not name:
                continue
            student_no = str(ws.cell_value(r, col_map.get('student_no', 0))).strip()
            class_name = str(ws.cell_value(r, col_map.get('class_name', 2))).strip() if 'class_name' in col_map and col_map['class_name'] < ws.ncols else ''
            tags_str = str(ws.cell_value(r, col_map['tags'])).strip() if 'tags' in col_map and col_map['tags'] < ws.ncols else ''
            titles_str = str(ws.cell_value(r, col_map['titles'])).strip() if 'titles' in col_map and col_map['titles'] < ws.ncols else ''

            tags = [t.strip() for t in tags_str.replace('，', ',').split(',') if t.strip()] if tags_str else []
            titles = [t.strip() for t in titles_str.replace('，', ',').split(',') if t.strip()] if titles_str else []

            students.append({
                'student_no': student_no,
                'name': name,
                'class_name': class_name,
                'tags': tags,
                'titles': titles,
                'avatar_color': random.choice(COURSE_COLORS),
            })

        return students

    @staticmethod
    def _read_students_from_docx(file_path: str) -> List[Dict[str, Any]]:
        """从Word文档读取学生信息（表格形式）"""
        if not HAS_DOCX:
            raise ImportError("需要安装python-docx")

        doc = Document(file_path)
        students = []

        for table in doc.tables:
            if not table.rows:
                continue
            headers = [cell.text.strip() for cell in table.rows[0].cells]

            col_map = {}
            for i, h in enumerate(headers):
                h_lower = h.lower()
                if '学号' in h_lower or '编号' in h_lower:
                    col_map['student_no'] = i
                elif '姓名' in h_lower or '名字' in h_lower:
                    col_map['name'] = i
                elif '班级' in h_lower or 'class' in h_lower:
                    col_map['class_name'] = i
                elif '标签' in h_lower or 'tag' in h_lower:
                    col_map['tags'] = i
                elif '头衔' in h_lower or 'title' in h_lower:
                    col_map['titles'] = i

            if 'name' not in col_map and len(headers) >= 2:
                col_map['student_no'] = 0
                col_map['name'] = 1

            if 'name' not in col_map:
                continue

            for row in table.rows[1:]:
                cells = [cell.text.strip() for cell in row.cells]
                if not cells:
                    continue
                name = cells[col_map['name']] if col_map['name'] < len(cells) else ''
                if not name:
                    continue
                student_no = cells[col_map['student_no']] if 'student_no' in col_map and col_map['student_no'] < len(cells) else name
                class_name = cells[col_map['class_name']] if 'class_name' in col_map and col_map['class_name'] < len(cells) else ''
                tags_str = cells[col_map['tags']] if 'tags' in col_map and col_map['tags'] < len(cells) else ''
                titles_str = cells[col_map['titles']] if 'titles' in col_map and col_map['titles'] < len(cells) else ''

                tags = [t.strip() for t in tags_str.replace('，', ',').split(',') if t.strip()] if tags_str else []
                titles = [t.strip() for t in titles_str.replace('，', ',').split(',') if t.strip()] if titles_str else []

                students.append({
                    'student_no': student_no,
                    'name': name,
                    'class_name': class_name,
                    'tags': tags,
                    'titles': titles,
                    'avatar_color': random.choice(COURSE_COLORS),
                })

        return students

    @staticmethod
    def _read_students_from_doc(file_path: str) -> List[Dict[str, Any]]:
        """从旧版Word读取（仅 Windows + Microsoft Word 可用；macOS 请先另存为 .docx）"""
        if not IS_WINDOWS:
            raise ImportError(
                "读取 .doc 旧格式需要 Windows + Microsoft Word。"
                "macOS / Linux 请先用 Word 或 Pages 将文件另存为 .docx 后再导入。")
        # 尝试使用pywin32
        try:
            import win32com.client
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            doc = word.Documents.Open(file_path)
            students = []

            for table in doc.Tables:
                if table.Rows.Count < 2:
                    continue
                headers = [table.Cell(1, c).Range.Text.strip() for c in range(1, table.Columns.Count + 1)]

                col_map = {}
                for i, h in enumerate(headers):
                    h_lower = h.lower()
                    if '学号' in h_lower or '编号' in h_lower:
                        col_map['student_no'] = i
                    elif '姓名' in h_lower or '名字' in h_lower:
                        col_map['name'] = i
                    elif '班级' in h_lower or 'class' in h_lower:
                        col_map['class_name'] = i
                    elif '标签' in h_lower:
                        col_map['tags'] = i
                    elif '头衔' in h_lower:
                        col_map['titles'] = i

                if 'name' not in col_map and len(headers) >= 2:
                    col_map['student_no'] = 0
                    col_map['name'] = 1

                if 'name' not in col_map:
                    continue

                for r in range(2, table.Rows.Count + 1):
                    name = table.Cell(r, col_map['name'] + 1).Range.Text.strip()
                    if not name:
                        continue
                    student_no = table.Cell(r, col_map.get('student_no', 0) + 1).Range.Text.strip()
                    class_name = table.Cell(r, col_map['class_name'] + 1).Range.Text.strip() if 'class_name' in col_map else ''
                    tags_str = table.Cell(r, col_map['tags'] + 1).Range.Text.strip() if 'tags' in col_map else ''
                    titles_str = table.Cell(r, col_map['titles'] + 1).Range.Text.strip() if 'titles' in col_map else ''

                    tags = [t.strip() for t in tags_str.replace('，', ',').split(',') if t.strip()] if tags_str else []
                    titles = [t.strip() for t in titles_str.replace('，', ',').split(',') if t.strip()] if titles_str else []

                    students.append({
                        'student_no': student_no,
                        'name': name,
                        'class_name': class_name,
                        'tags': tags,
                        'titles': titles,
                        'avatar_color': random.choice(COURSE_COLORS),
                    })

            doc.Close()
            word.Quit()
            return students
        except ImportError:
            raise ImportError("读取.doc文件需要安装pywin32。请运行: pip install pywin32")
        except Exception as e:
            raise ValueError(f"读取.doc文件失败: {str(e)}")

    @staticmethod
    def _read_students_from_csv(file_path: str) -> List[Dict[str, Any]]:
        """从CSV读取学生信息"""
        import csv
        students = []
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            if not headers:
                return students

            col_map = {}
            for i, h in enumerate(headers):
                h_lower = h.strip().lower()
                if '学号' in h_lower or '编号' in h_lower:
                    col_map['student_no'] = i
                elif '姓名' in h_lower or '名字' in h_lower:
                    col_map['name'] = i
                elif '班级' in h_lower or 'class' in h_lower:
                    col_map['class_name'] = i
                elif '标签' in h_lower or 'tag' in h_lower:
                    col_map['tags'] = i
                elif '头衔' in h_lower or 'title' in h_lower:
                    col_map['titles'] = i

            if 'name' not in col_map and len(headers) >= 2:
                col_map['student_no'] = 0
                col_map['name'] = 1

            if 'name' not in col_map:
                return students

            for row in reader:
                if not row:
                    continue
                name = row[col_map['name']].strip() if col_map['name'] < len(row) else ''
                if not name:
                    continue
                student_no = row[col_map['student_no']].strip() if 'student_no' in col_map and col_map['student_no'] < len(row) else name
                class_name = row[col_map['class_name']].strip() if 'class_name' in col_map and col_map['class_name'] < len(row) else ''
                tags_str = row[col_map['tags']].strip() if 'tags' in col_map and col_map['tags'] < len(row) else ''
                titles_str = row[col_map['titles']].strip() if 'titles' in col_map and col_map['titles'] < len(row) else ''

                tags = [t.strip() for t in tags_str.replace('，', ',').split(',') if t.strip()] if tags_str else []
                titles = [t.strip() for t in titles_str.replace('，', ',').split(',') if t.strip()] if titles_str else []

                students.append({
                    'student_no': student_no,
                    'name': name,
                    'class_name': class_name,
                    'tags': tags,
                    'titles': titles,
                    'avatar_color': random.choice(COURSE_COLORS),
                })

        return students

    @staticmethod
    def _cell_text(val, default: str = "") -> str:
        """把 Excel 单元格值转为文本。

        openpyxl 对空单元格返回 None，直接 str() 会得到字符串 'None'
        （曾导致导入后标签显示为「None」、学号变成「None」）。
        这里统一把 None/空值转成 default，数字/日期等照常转字符串。
        """
        if val is None:
            return default
        if isinstance(val, (datetime, date)):
            return val.strftime("%Y-%m-%d")
        text = str(val).strip()
        if not text or text == 'None':
            return default
        # Excel 里数值型整数会读成 20260101.0，去掉多余的小数点
        if re.fullmatch(r'-?\d+\.0', text):
            text = text[:-2]
        return text

    @staticmethod
    def _cell_date_text(val, default: str = "") -> str:
        """将 Excel 单元格值规范化为 YYYY-MM-DD 文本。
        Excel 真实日期单元格读出为 datetime/date/时间浮点，直接 str() 会得到
        '2026-06-01 00:00:00' 或 'None'，统一转为 ISO 日期，空值返回 default。
        """
        if val is None:
            return default
        if isinstance(val, datetime):
            return val.strftime("%Y-%m-%d")
        if isinstance(val, date):
            return val.strftime("%Y-%m-%d")
        text = str(val).strip()
        if not text:
            return default
        # 兼容 '2026-06-01 00:00:00' / '2026/6/1' 之类的文本日期
        m = re.match(r'^(\d{4})[-/年](\d{1,2})[-/月](\d{1,2})', text)
        if m:
            try:
                return f"{int(m.group(1)):04d}-{int(m.group(2)):02d}-{int(m.group(3)):02d}"
            except ValueError:
                return text
        return text

    @staticmethod
    def _read_grades_from_excel(file_path: str) -> List[Dict[str, Any]]:
        """从Excel读取成绩"""
        if not HAS_OPENPYXL:
            raise ImportError("需要安装openpyxl")

        wb = load_workbook(file_path, data_only=True)
        ws = wb.active
        headers = [str(cell.value or '').strip() for cell in ws[1]]

        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if '学号' in h_lower or '编号' in h_lower:
                col_map['student_no'] = i
            elif '姓名' in h_lower or '名字' in h_lower or 'name' in h_lower:
                col_map['student_name'] = i
            elif '课程' in h_lower or '科目' in h_lower or 'course' in h_lower:
                col_map['course_name'] = i
            elif '成绩' in h_lower or '分数' in h_lower or 'score' in h_lower:
                col_map['score'] = i
            elif '日期' in h_lower or 'date' in h_lower:
                col_map['exam_date'] = i
            elif '类型' in h_lower or 'type' in h_lower:
                col_map['exam_type'] = i
            elif '满分' in h_lower or 'full' in h_lower:
                col_map['full_score'] = i

        if 'score' not in col_map or 'student_name' not in col_map:
            # 尝试猜测
            if len(headers) >= 3:
                col_map['student_name'] = 0
                col_map['course_name'] = 1
                col_map['score'] = 2
            else:
                raise ValueError("无法识别成绩列")

        grades = []
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or all(v is None or str(v).strip() == '' for v in row):
                continue
            try:
                student_name = FileImporter._cell_text(
                    row[col_map['student_name']] if col_map['student_name'] < len(row) else None)
                score = float(row[col_map['score']]) if col_map['score'] < len(row) and row[col_map['score']] is not None else 0
                if not student_name:
                    continue
                course_name = FileImporter._cell_text(
                    row[col_map['course_name']] if 'course_name' in col_map and col_map['course_name'] < len(row) else None,
                    default='未知课程')
                raw_date = row[col_map['exam_date']] if 'exam_date' in col_map and col_map['exam_date'] < len(row) else None
                exam_date = FileImporter._cell_date_text(raw_date, date.today().isoformat())
                exam_type = FileImporter._cell_text(
                    row[col_map['exam_type']] if 'exam_type' in col_map and col_map['exam_type'] < len(row) else None,
                    default='其他')
                full_score = float(row[col_map['full_score']]) if 'full_score' in col_map and col_map['full_score'] < len(row) and row[col_map['full_score']] else 100.0

                grades.append({
                    'student_name': student_name,
                    'course_name': course_name,
                    'score': score,
                    'exam_date': exam_date,
                    'exam_type': exam_type,
                    'full_score': full_score,
                })
            except (ValueError, IndexError):
                continue

        return grades

    @staticmethod
    def _read_grades_from_xls(file_path: str) -> List[Dict[str, Any]]:
        if not HAS_XLRD:
            raise ImportError("需要安装xlrd")
        wb = xlrd.open_workbook(file_path)
        ws = wb.sheet_by_index(0)
        headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]

        col_map = {}
        for i, h in enumerate(headers):
            h_lower = h.lower()
            if '学号' in h_lower or '编号' in h_lower:
                col_map['student_no'] = i
            elif '姓名' in h_lower or '名字' in h_lower:
                col_map['student_name'] = i
            elif '课程' in h_lower or '科目' in h_lower:
                col_map['course_name'] = i
            elif '成绩' in h_lower or '分数' in h_lower:
                col_map['score'] = i
            elif '日期' in h_lower:
                col_map['exam_date'] = i
            elif '类型' in h_lower:
                col_map['exam_type'] = i
            elif '满分' in h_lower:
                col_map['full_score'] = i

        if 'score' not in col_map or 'student_name' not in col_map:
            if ws.ncols >= 3:
                col_map['student_name'] = 0
                col_map['course_name'] = 1
                col_map['score'] = 2
            else:
                raise ValueError("无法识别成绩列")

        grades = []
        for r in range(1, ws.nrows):
            try:
                student_name = str(ws.cell_value(r, col_map.get('student_name', 0))).strip()
                score = float(ws.cell_value(r, col_map.get('score', 2)))
                if not student_name:
                    continue
                course_name = str(ws.cell_value(r, col_map.get('course_name', 1))).strip()
                exam_date = str(ws.cell_value(r, col_map.get('exam_date', 4))).strip() if 'exam_date' in col_map else date.today().isoformat()
                exam_type = str(ws.cell_value(r, col_map.get('exam_type', 5))).strip() if 'exam_type' in col_map else '其他'
                full_score = float(ws.cell_value(r, col_map.get('full_score', 6))) if 'full_score' in col_map else 100.0
                grades.append({
                    'student_name': student_name,
                    'course_name': course_name,
                    'score': score,
                    'exam_date': exam_date,
                    'exam_type': exam_type,
                    'full_score': full_score,
                })
            except (ValueError, IndexError):
                continue

        return grades

    @staticmethod
    def _read_grades_from_csv(file_path: str) -> List[Dict[str, Any]]:
        import csv
        grades = []
        with open(file_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.reader(f)
            headers = next(reader, None)
            if not headers:
                return grades
            col_map = {}
            for i, h in enumerate(headers):
                h_lower = h.strip().lower()
                if '姓名' in h_lower or 'name' in h_lower:
                    col_map['student_name'] = i
                elif '课程' in h_lower or '科目' in h_lower:
                    col_map['course_name'] = i
                elif '成绩' in h_lower or '分数' in h_lower:
                    col_map['score'] = i
                elif '日期' in h_lower:
                    col_map['exam_date'] = i
                elif '类型' in h_lower:
                    col_map['exam_type'] = i

            for row in reader:
                if not row:
                    continue
                try:
                    student_name = row[col_map.get('student_name', 0)].strip()
                    score = float(row[col_map.get('score', 2)])
                    course_name = row[col_map.get('course_name', 1)].strip()
                    exam_date = row[col_map.get('exam_date', 4)].strip() if 'exam_date' in col_map else date.today().isoformat()
                    exam_type = row[col_map.get('exam_type', 5)].strip() if 'exam_type' in col_map else '其他'
                    grades.append({
                        'student_name': student_name,
                        'course_name': course_name,
                        'score': score,
                        'exam_date': exam_date,
                        'exam_type': exam_type,
                        'full_score': 100.0,
                    })
                except (ValueError, IndexError):
                    continue
        return grades


# ==================== 数据导入服务 ====================
class DataImportService:
    """数据导入服务 - 将导入的数据存入数据库"""

    def __init__(self, db: DatabaseManager):
        self.db = db

    def import_students(self, file_path: str) -> Tuple[int, List[str]]:
        """导入学生数据，返回(成功数量, 错误列表)"""
        try:
            students_data = FileImporter.import_students_from_file(file_path)
        except Exception as e:
            return 0, [str(e)]

        success = 0
        errors = []
        # 一次性取出已有学号，避免每行都全表扫描（原实现 O(n²)，导入大批量时很慢）
        existing_nos = {s.student_no for s in self.db.get_all_students()}

        for sd in students_data:
            try:
                # 检查是否已存在
                if sd['student_no'] in existing_nos:
                    errors.append(f"学号{sd['student_no']}已存在，跳过")
                    continue

                student = Student(
                    student_no=sd['student_no'],
                    name=sd['name'],
                    class_name=sd.get('class_name', ''),
                    tags=sd.get('tags', []),
                    titles=sd.get('titles', []),
                    avatar_color=sd.get('avatar_color', random.choice(COURSE_COLORS)),
                )
                self.db.add_student(student)
                existing_nos.add(sd['student_no'])
                success += 1
            except Exception as e:
                errors.append(f"导入{sd.get('name', '未知')}失败: {str(e)}")

        return success, errors

    def import_grades(self, file_path: str) -> Tuple[int, List[str]]:
        """导入成绩数据"""
        try:
            grades_data = FileImporter.import_grades_from_file(file_path)
        except Exception as e:
            return 0, [str(e)]

        success = 0
        errors = []
        students = self.db.get_all_students()
        student_map = {s.name: s.id for s in students}

        for gd in grades_data:
            try:
                student_name = gd.get('student_name', '')
                student_id = student_map.get(student_name, 0)

                grade = Grade(
                    student_id=student_id,
                    student_name=student_name,
                    course_name=gd.get('course_name', '未知'),
                    score=gd.get('score', 0),
                    exam_date=gd.get('exam_date', date.today().isoformat()),
                    exam_type=gd.get('exam_type', '其他'),
                    full_score=gd.get('full_score', 100.0),
                )
                self.db.add_grade(grade)
                success += 1
            except Exception as e:
                errors.append(f"导入成绩失败({gd.get('student_name', '未知')}): {str(e)}")

        return success, errors


# ==================== 磁贴组件 ====================
class TileWidget(QFrame):
    """通用磁贴组件基类

    - 统一提供「编辑 / 删除」入口：悬停时右上角出现小按钮，右键也有菜单
    - 不再限制最大尺寸，避免内容被压缩导致文字与图标重叠
    """
    clicked = Signal(object)           # 单击（携带本磁贴数据）
    double_clicked = Signal(object)    # 双击
    edit_requested = Signal(object)    # 请求编辑
    delete_requested = Signal(object)  # 请求删除

    def __init__(self, data: Any = None, parent=None):
        super().__init__(parent)
        self.data = data
        self._is_selected = False
        self._hover = False
        self.setCursor(Qt.PointingHandCursor)
        self.setFrameShape(QFrame.Shape.StyledPanel)
        self.setMinimumSize(150, 120)
        self._apply_card_style()
        # 阴影效果
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(15)
        shadow.setOffset(0, 3)
        shadow.setColor(QColor(0, 0, 0, 40))
        self.setGraphicsEffect(shadow)
        # 悬停操作按钮（右上角浮层）
        self._build_hover_actions()

    # ---- 样式 ----
    def _apply_card_style(self, selected=None):
        if selected is None:
            selected = self._is_selected
        if selected:
            self.setStyleSheet(f"""
                QFrame {{
                    background-color: {COLORS['selected']};
                    border: 3px solid {COLORS['primary']};
                    border-radius: 12px;
                }}
            """)
        else:
            self.setStyleSheet(f"""
                QFrame {{
                    background-color: {COLORS['card_bg']};
                    border: 2px solid {COLORS['border']};
                    border-radius: 12px;
                }}
                QFrame:hover {{
                    border-color: {COLORS['primary']};
                    background-color: {COLORS['hover']};
                }}
            """)

    def set_selected(self, selected: bool):
        self._is_selected = selected
        self._apply_card_style(selected)

    # ---- 悬停编辑/删除按钮 ----
    def _build_hover_actions(self):
        self._actions = QWidget(self)
        row = QHBoxLayout(self._actions)
        row.setContentsMargins(0, 0, 0, 0)
        row.setSpacing(4)
        self.edit_btn = QToolButton(self._actions)
        self.edit_btn.setText("✏️")
        self.edit_btn.setToolTip("编辑")
        self.edit_btn.setCursor(Qt.PointingHandCursor)
        self.edit_btn.setStyleSheet(self._mini_action_style(COLORS['primary']))
        self.edit_btn.clicked.connect(lambda: self.edit_requested.emit(self.data))
        row.addWidget(self.edit_btn)

        self.delete_btn = QToolButton(self._actions)
        self.delete_btn.setText("🗑")
        self.delete_btn.setToolTip("删除")
        self.delete_btn.setCursor(Qt.PointingHandCursor)
        self.delete_btn.setStyleSheet(self._mini_action_style(COLORS['danger']))
        self.delete_btn.clicked.connect(lambda: self.delete_requested.emit(self.data))
        row.addWidget(self.delete_btn)

        self._actions.setFixedHeight(24)
        self._actions.setVisible(False)

    @staticmethod
    def _mini_action_style(color: str) -> str:
        return f"""
            QToolButton {{
                background-color: {color}; color: white; border: none;
                border-radius: 11px; font-size: 11px; padding: 0px;
            }}
            QToolButton:hover {{ background-color: {color}; }}
        """

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, '_actions'):
            w = self._actions.sizeHint().width() + 4
            self._actions.setGeometry(max(4, self.width() - w - 6), 6, w, 24)
            self._actions.raise_()

    def enterEvent(self, event):
        self._hover = True
        if hasattr(self, '_actions'):
            self._actions.setVisible(True)
            self._actions.raise_()
        self._apply_card_style()
        super().enterEvent(event)

    def leaveEvent(self, event):
        self._hover = False
        if hasattr(self, '_actions'):
            self._actions.setVisible(False)
        self._apply_card_style()
        super().leaveEvent(event)

    # ---- 事件 ----
    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit(self.data)
        super().mousePressEvent(event)

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.double_clicked.emit(self.data)
        super().mouseDoubleClickEvent(event)

    def contextMenuEvent(self, event):
        """默认右键菜单：编辑 + 删除（子类可扩展）"""
        menu = QMenu(self)
        self.extend_context_menu(menu)
        if not menu.isEmpty():
            menu.exec(event.globalPos())
        event.accept()

    def extend_context_menu(self, menu: QMenu):
        edit_act = menu.addAction("✏️ 编辑")
        edit_act.triggered.connect(lambda: self.edit_requested.emit(self.data))
        menu.addSeparator()
        del_act = menu.addAction("🗑 删除")
        del_act.triggered.connect(lambda: self.delete_requested.emit(self.data))


class StudentTile(TileWidget):
    """学生磁贴 - 头像在上，**完整姓名在头像下方**，学号/班级/标签依次排列"""

    def __init__(self, student: Student, parent=None):
        super().__init__(student, parent)
        self.student = student
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        # 头像（圆形首字）——固定在顶部，不与姓名争抢空间
        self.avatar_label = QLabel(self.student.name[0] if self.student.name else '?')
        self.avatar_label.setAlignment(Qt.AlignCenter)
        self.avatar_label.setFixedSize(48, 48)
        self.avatar_label.setStyleSheet(f"""
            QLabel {{
                background-color: {self.student.avatar_color or COLORS['primary']};
                color: white;
                border-radius: 24px;
                font-size: 20px;
                font-weight: bold;
                font-family: {FONT_STACK_CSS};
            }}
        """)
        avatar_row = QHBoxLayout()
        avatar_row.addStretch()
        avatar_row.addWidget(self.avatar_label, alignment=Qt.AlignTop)
        avatar_row.addStretch()
        layout.addLayout(avatar_row)

        # 姓名：完整显示在头像正下方（自动换行、不省略）
        self.name_label = QLabel(self.student.name or "（未命名）")
        self.name_label.setAlignment(Qt.AlignHCenter | Qt.AlignTop)
        self.name_label.setWordWrap(True)
        self.name_label.setStyleSheet(f"""
            QLabel {{
                font-size: 15px; font-weight: bold; color: {COLORS['text_primary']};
                font-family: {FONT_STACK_CSS};
                border: none; background: transparent; padding: 0px;
            }}
        """)
        self.name_label.setMinimumHeight(24)
        layout.addWidget(self.name_label)

        # 学号
        no_label = QLabel(f"学号 {self.student.student_no}")
        no_label.setAlignment(Qt.AlignCenter)
        no_label.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"border: none; background: transparent;")
        layout.addWidget(no_label)

        # 班级
        if self.student.class_name:
            class_label = QLabel(f"📚 {self.student.class_name}")
            class_label.setAlignment(Qt.AlignCenter)
            class_label.setWordWrap(True)
            class_label.setStyleSheet(
                f"font-size: 11px; color: {COLORS['text_secondary']}; "
                f"border: none; background: transparent;")
            layout.addWidget(class_label)

        # 标签
        if self.student.tags:
            tags_text = ' '.join([f'🏷️{t}' for t in self.student.tags[:3]])
            tags_label = QLabel(tags_text)
            tags_label.setAlignment(Qt.AlignCenter)
            tags_label.setWordWrap(True)
            tags_label.setStyleSheet(
                f"font-size: 10px; color: {COLORS['secondary']}; "
                f"border: none; background: transparent;")
            layout.addWidget(tags_label)

        # 头衔
        if self.student.titles:
            titles_text = ' '.join([f'⭐{t}' for t in self.student.titles[:2]])
            titles_label = QLabel(titles_text)
            titles_label.setAlignment(Qt.AlignCenter)
            titles_label.setWordWrap(True)
            titles_label.setStyleSheet(
                f"font-size: 10px; color: {COLORS['warning']}; "
                f"border: none; background: transparent;")
            layout.addWidget(titles_label)

        layout.addStretch()
        # 保证头像(48)+姓名+学号+班级等都有位置，避免被压缩重叠
        self.setMinimumHeight(206)

    def extend_context_menu(self, menu: QMenu):
        view_act = menu.addAction("📄 查看详情 / 成绩")
        view_act.triggered.connect(lambda: self.double_clicked.emit(self.student))
        edit_act = menu.addAction("✏️ 编辑该学生")
        edit_act.triggered.connect(lambda: self.edit_requested.emit(self.student))
        menu.addSeparator()
        del_act = menu.addAction("🗑 删除该学生")
        del_act.triggered.connect(lambda: self.delete_requested.emit(self.student))


class ClassTile(TileWidget):
    """班级磁贴 - 学生管理首页按班级聚合，点入查看该班学生"""

    def __init__(self, class_name: str, students: List[Student], parent=None):
        super().__init__({'class_name': class_name, 'students': students}, parent)
        self.class_name = class_name
        self.students = students
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(16, 14, 16, 14)
        layout.setSpacing(6)

        icon = QLabel("🏫")
        icon.setAlignment(Qt.AlignCenter)
        icon.setStyleSheet("font-size: 30px; border: none; background: transparent;")
        layout.addWidget(icon)

        title = QLabel(self.class_name or "未分班")
        title.setAlignment(Qt.AlignCenter)
        title.setWordWrap(True)
        title.setStyleSheet(f"""
            QLabel {{
                font-size: 16px; font-weight: bold; color: {COLORS['text_primary']};
                font-family: {FONT_STACK_CSS};
                border: none; background: transparent;
            }}
        """)
        layout.addWidget(title)

        count = QLabel(f"{len(self.students)} 名学生")
        count.setAlignment(Qt.AlignCenter)
        count.setStyleSheet(
            f"font-size: 12px; color: {COLORS['primary']}; font-weight: bold; "
            f"border: none; background: transparent;")
        layout.addWidget(count)

        # 标签/头衔统计（展示班级特征）
        tag_counter: Dict[str, int] = {}
        for s in self.students:
            for t in s.tags[:2]:
                tag_counter[t] = tag_counter.get(t, 0) + 1
        if tag_counter:
            top = sorted(tag_counter.items(), key=lambda kv: -kv[1])[:3]
            tag_txt = '  '.join(f"#{k}×{v}" for k, v in top)
            tags_label = QLabel(tag_txt)
            tags_label.setAlignment(Qt.AlignCenter)
            tags_label.setWordWrap(True)
            tags_label.setStyleSheet(
                f"font-size: 10px; color: {COLORS['secondary']}; "
                f"border: none; background: transparent;")
            layout.addWidget(tags_label)

        hint = QLabel("点击进入班级 →")
        hint.setAlignment(Qt.AlignCenter)
        hint.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"border: none; background: transparent;")
        layout.addWidget(hint)
        layout.addStretch()
        self.setMinimumHeight(190)

    def extend_context_menu(self, menu: QMenu):
        open_act = menu.addAction("📂 查看该班学生")
        open_act.triggered.connect(lambda: self.clicked.emit(self.data))
        rename_act = menu.addAction("✏️ 编辑班级名称")
        rename_act.triggered.connect(lambda: self.edit_requested.emit(self.data))


class StickerTile(TileWidget):
    """便签磁贴（贴纸风格）：支持文字 / 表格 / 图片，可置顶

    视觉上像贴纸：柔和底色 + 顶部胶带 + 圆角卡片 + 阴影。
    """
    pin_toggled = Signal(object)
    open_requested = Signal(object)

    def __init__(self, feedback: Feedback, parent=None):
        super().__init__(feedback, parent)
        self.feedback = feedback
        self._build_card()
        self._build_content()
        self._apply_card_style()      # feedback 就绪后重刷贴纸底色

    # ---- 卡片外观（贴纸质感）----
    def _apply_card_style(self, selected=None):
        # 注意：基类 __init__ 会提前调用本方法（此时 feedback 尚未赋值）
        fb = getattr(self, 'feedback', None)
        bg = (fb.color if fb is not None else None) or "#FFF6C9"
        border = QColor(bg).darker(112).name()
        self.setStyleSheet(f"""
            QFrame#sticker {{
                background-color: {bg};
                border: 2px solid {border};
                border-radius: 14px;
            }}
            QFrame#sticker:hover {{ border-color: {COLORS['primary']}; }}
        """)

    def _build_card(self):
        self.setObjectName("sticker")
        self.setMinimumSize(230, 150)
        self.setMaximumWidth(420)
        self._apply_card_style()
        # 顶部胶带装饰
        self._tape = QLabel("", self)
        self._tape.setFixedSize(64, 16)
        self._tape.setStyleSheet(
            "background-color: rgba(255,255,255,190); border: 1px solid rgba(0,0,0,28);"
            "border-radius: 3px;")
        self._tape.setAttribute(Qt.WA_TransparentForMouseEvents, True)

    def resizeEvent(self, event):
        super().resizeEvent(event)
        if hasattr(self, '_tape'):
            self._tape.setGeometry((self.width() - 64) // 2, -6, 64, 16)
            self._tape.raise_()

    def _build_content(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(14, 18, 14, 12)
        layout.setSpacing(8)

        # 头部：类型标识 + 置顶标记
        head = QHBoxLayout()
        type_map = {'text': "📝 便签", 'table': "📊 表格", 'image': "🖼 图片"}
        self.type_label = QLabel(type_map.get(self.feedback.content_type, "📝 便签"))
        self.type_label.setStyleSheet(
            f"font-size: 10px; color: {COLORS['text_secondary']}; "
            f"border: none; background: transparent;")
        head.addWidget(self.type_label)
        head.addStretch()
        if self.feedback.pinned:
            pin = QLabel("📌")
            pin.setStyleSheet("font-size: 12px; border: none; background: transparent;")
            head.addWidget(pin)
        layout.addLayout(head)

        # 主体内容
        if self.feedback.content_type == 'table':
            layout.addWidget(self._build_table_view())
        elif self.feedback.content_type == 'image':
            layout.addWidget(self._build_image_view())
        else:
            layout.addWidget(self._build_text_view())

        layout.addStretch()
        # 底部：时间
        foot = QLabel(f"🕐 {self.feedback.created_at[:16].replace('T', ' ')}")
        foot.setStyleSheet(
            f"font-size: 10px; color: {COLORS['text_secondary']}; "
            f"border: none; background: transparent;")
        layout.addWidget(foot)

    def _build_text_view(self) -> QWidget:
        text = self.feedback.content.strip() or "（空便签）"
        label = QLabel(text)
        label.setWordWrap(True)
        label.setTextInteractionFlags(Qt.TextSelectableByMouse)
        label.setStyleSheet(f"""
            QLabel {{
                font-size: 13px; color: {COLORS['text_primary']};
                font-family: {FONT_STACK_CSS};
                border: none; background: transparent;
            }}
        """)
        return label

    def _build_table_view(self) -> QWidget:
        data = self.feedback.table_data()
        headers, rows = data['headers'], data['rows']
        if not headers and not rows:
            return self._build_text_view()
        html = ["<table cellspacing='0' cellpadding='4' style='font-size:11px;'>"]
        if headers:
            html.append("<tr>" + "".join(
                f"<th style='background:rgba(255,255,255,150); border:1px solid "
                f"rgba(0,0,0,30);'>{self._esc(h)}</th>" for h in headers) + "</tr>")
        for row in rows[:12]:
            html.append("<tr>" + "".join(
                f"<td style='border:1px solid rgba(0,0,0,25);'>{self._esc(c)}</td>"
                for c in row) + "</tr>")
        html.append("</table>")
        if len(rows) > 12:
            html.append(f"<div style='font-size:10px;color:#636E72;'>…共 {len(rows)} 行</div>")
        label = QLabel("".join(html))
        label.setTextFormat(Qt.RichText)
        label.setWordWrap(True)
        label.setStyleSheet("QLabel { border: none; background: transparent; }")
        return label

    @staticmethod
    def _esc(text: Any) -> str:
        s = "" if text is None else str(text)
        return (s.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
                 .replace('\n', '<br/>'))

    def _build_image_view(self) -> QWidget:
        container = QWidget()
        box = QVBoxLayout(container)
        box.setContentsMargins(0, 0, 0, 0)
        box.setSpacing(4)

        path = image_abs_path(self.feedback.image_path)
        img_label = QLabel()
        img_label.setAlignment(Qt.AlignCenter)
        img_label.setCursor(Qt.PointingHandCursor)
        img_label.setStyleSheet("QLabel { border: none; background: transparent; }")
        if path and path.exists():
            pix = QPixmap(str(path))
            if not pix.isNull():
                img_label.setPixmap(pix.scaled(300, 220, Qt.KeepAspectRatio,
                                               Qt.SmoothTransformation))
            else:
                img_label.setText("（图片无法读取）")
        else:
            img_label.setText("（图片文件缺失）")
        img_label.mousePressEvent = lambda e: self.open_requested.emit(self.feedback)  # type: ignore
        box.addWidget(img_label)

        if self.feedback.image_caption:
            cap = QLabel(self.feedback.image_caption)
            cap.setWordWrap(True)
            cap.setAlignment(Qt.AlignCenter)
            cap.setStyleSheet(
                f"font-size: 11px; color: {COLORS['text_primary']}; "
                f"border: none; background: transparent;")
            box.addWidget(cap)
        return container

    def extend_context_menu(self, menu: QMenu):
        edit_act = menu.addAction("✏️ 编辑便签")
        edit_act.triggered.connect(lambda: self.edit_requested.emit(self.feedback))
        pin_txt = "📍 取消置顶" if self.feedback.pinned else "📌 置顶"
        pin_act = menu.addAction(pin_txt)
        pin_act.triggered.connect(lambda: self.pin_toggled.emit(self.feedback))
        if self.feedback.content_type == 'image':
            open_act = menu.addAction("🔍 查看大图")
            open_act.triggered.connect(lambda: self.open_requested.emit(self.feedback))
        copy_act = menu.addAction("📋 复制内容")
        copy_act.triggered.connect(self._copy_to_clipboard)
        menu.addSeparator()
        del_act = menu.addAction("🗑 删除便签")
        del_act.triggered.connect(lambda: self.delete_requested.emit(self.feedback))

    def _copy_to_clipboard(self):
        try:
            clipboard = QApplication.clipboard()
            if self.feedback.content_type == 'image':
                path = image_abs_path(self.feedback.image_path)
                if path and path.exists():
                    clipboard.setPixmap(QPixmap(str(path)))
                    return
            clipboard.setText(self.feedback.plain_text())
        except Exception:
            pass


class GradeTile(TileWidget):
    """成绩磁贴 - 每个学生一个，展示成绩概览（可右键/悬停编辑）"""

    def __init__(self, student: Student, grades: List[Grade], parent=None):
        super().__init__({'student': student, 'grades': grades}, parent)
        self.student = student
        self.grades = grades
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(4)

        name_label = QLabel(f"👤 {self.student.name}")
        name_label.setWordWrap(True)
        name_label.setStyleSheet(
            f"font-size: 14px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS}; border: none; background: transparent;")
        layout.addWidget(name_label)

        if self.grades:
            avg_score = sum(g.score for g in self.grades) / len(self.grades)
            avg_label = QLabel(f"📊 平均分: {avg_score:.1f}")
            color = '#00B894' if avg_score >= 85 else ('#FDCB6E' if avg_score >= 60 else '#FF6B6B')
            avg_label.setStyleSheet(
                f"font-size: 16px; font-weight: bold; color: {color}; "
                f"border: none; background: transparent;")
            layout.addWidget(avg_label)

            recent = self.grades[0]
            recent_label = QLabel(f"最近: {recent.course_name} {recent.score}")
            recent_label.setWordWrap(True)
            recent_label.setStyleSheet(
                f"font-size: 10px; color: {COLORS['text_secondary']}; "
                f"border: none; background: transparent;")
            layout.addWidget(recent_label)

            cnt = QLabel(f"共 {len(self.grades)} 条成绩记录")
            cnt.setStyleSheet(
                f"font-size: 10px; color: {COLORS['text_secondary']}; "
                f"border: none; background: transparent;")
            layout.addWidget(cnt)
        else:
            no_data = QLabel("暂无成绩数据")
            no_data.setStyleSheet(
                f"font-size: 12px; color: {COLORS['text_secondary']}; "
                f"border: none; background: transparent;")
            layout.addWidget(no_data)

        layout.addStretch()
        self.setMinimumHeight(150)

    def extend_context_menu(self, menu: QMenu):
        edit_act = menu.addAction("✏️ 编辑成绩")
        edit_act.triggered.connect(lambda: self.edit_requested.emit(self.data))
        report_act = menu.addAction("📈 查看成绩报告")
        report_act.triggered.connect(lambda: self.double_clicked.emit(self.data))
        menu.addSeparator()
        del_act = menu.addAction("🗑 删除该生全部成绩")
        del_act.triggered.connect(lambda: self.delete_requested.emit(self.data))


# ==================== 气泡 / 标签云组件 ====================
class FlowLayout(QLayout):
    """流式布局：控件按行排列，放不下自动换行（气泡云的基础）"""

    def __init__(self, parent=None, margin: int = 0, spacing: int = 8):
        super().__init__(parent)
        self._items: List[QWidgetItem] = []
        self.setContentsMargins(margin, margin, margin, margin)
        self.setSpacing(spacing)

    def addItem(self, item):
        self._items.append(item)

    def count(self) -> int:
        return len(self._items)

    def itemAt(self, index: int):
        return self._items[index] if 0 <= index < len(self._items) else None

    def takeAt(self, index: int):
        return self._items.pop(index) if 0 <= index < len(self._items) else None

    def expandingDirections(self):
        return Qt.Orientations(Qt.Orientation(0))

    def hasHeightForWidth(self) -> bool:
        return True

    def heightForWidth(self, width: int) -> int:
        return self._do_layout(QRect(0, 0, width, 0), test_only=True)

    def setGeometry(self, rect: QRect):
        super().setGeometry(rect)
        self._do_layout(rect, test_only=False)

    def sizeHint(self) -> QSize:
        return self.minimumSize()

    def minimumSize(self) -> QSize:
        size = QSize()
        for item in self._items:
            size = size.expandedTo(item.minimumSize())
        margins = self.contentsMargins()
        size += QSize(margins.left() + margins.right(), margins.top() + margins.bottom())
        return size

    def _do_layout(self, rect: QRect, test_only: bool) -> int:
        margins = self.contentsMargins()
        effective = rect.adjusted(margins.left(), margins.top(),
                                  -margins.right(), -margins.bottom())
        x, y, line_height = effective.x(), effective.y(), 0
        spacing = self.spacing()
        for item in self._items:
            hint = item.sizeHint()
            next_x = x + hint.width() + spacing
            if next_x - spacing > effective.right() and line_height > 0:
                x = effective.x()
                y = y + line_height + spacing
                next_x = x + hint.width() + spacing
                line_height = 0
            if not test_only:
                item.setGeometry(QRect(QPoint(x, y), hint))
            x = next_x
            line_height = max(line_height, hint.height())
        return y + line_height - rect.y() + margins.bottom()


class BubbleItem(QFrame):
    """气泡/标签：按权重显示不同大小与配色（类似 tag cloud）"""

    # 尺寸档位：(字号, 内边距, 圆角)
    SIZE_CLASSES = {
        'xl': (22, (18, 10), 20),
        'l': (17, (15, 8), 17),
        'm': (14, (12, 7), 14),
        's': (12, (10, 5), 12),
        'xs': (11, (8, 4), 10),
    }
    # 分类配色（浅底深字，气泡质感）
    CATEGORY_COLORS = {
        'identity': ("#E8ECFF", "#3A54D4"),
        'tag': ("#DFF5E1", "#0F7B5A"),
        'title': ("#FFF3D6", "#B26A00"),
        'score': ("#E4F0FF", "#1B67C6"),
        'subject': ("#F0E4FF", "#6C4BC7"),
        'warn': ("#FFE6E6", "#C0392B"),
        'muted': ("#F1F3F5", "#636E72"),
    }

    def __init__(self, text: str, size_class: str = 'm', category: str = 'muted',
                 tooltip: str = "", parent=None):
        super().__init__(parent)
        self.text = text
        self.size_class = size_class if size_class in self.SIZE_CLASSES else 'm'
        self.category = category if category in self.CATEGORY_COLORS else 'muted'
        font_size, (pad_x, pad_y), radius = self.SIZE_CLASSES[self.size_class]
        bg, fg = self.CATEGORY_COLORS[self.category]

        self.setStyleSheet(f"""
            BubbleItem {{
                background-color: {bg};
                border: 1px solid {QColor(fg).lighter(160).name()};
                border-radius: {radius}px;
            }}
            BubbleItem:hover {{
                border: 2px solid {fg};
                background-color: {QColor(bg).darker(103).name()};
            }}
        """)
        box = QHBoxLayout(self)
        box.setContentsMargins(pad_x, pad_y, pad_x, pad_y)
        box.setSpacing(0)
        label = QLabel(text)
        label.setStyleSheet(
            f"QLabel {{ color: {fg}; font-size: {font_size}px; "
            f"font-weight: {'bold' if self.size_class in ('xl', 'l') else 'normal'}; "
            f"border: none; background: transparent; font-family: {FONT_STACK_CSS}; }}")
        box.addWidget(label)
        if tooltip:
            self.setToolTip(tooltip)
            label.setToolTip(tooltip)
        # 加一点阴影，更像“气泡”
        shadow = QGraphicsDropShadowEffect(self)
        shadow.setBlurRadius(10)
        shadow.setOffset(0, 2)
        shadow.setColor(QColor(0, 0, 0, 26))
        self.setGraphicsEffect(shadow)
        self.setCursor(Qt.WhatsThisCursor if tooltip else Qt.ArrowCursor)


class BubbleCloud(QWidget):
    """气泡云：把若干 BubbleItem 流式排布，支持清空重建"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setStyleSheet("background: transparent;")
        self._layout = FlowLayout(self, margin=4, spacing=9)

    def clear(self):
        while self._layout.count():
            item = self._layout.takeAt(0)
            w = item.widget() if item else None
            if w:
                w.hide()
                w.deleteLater()

    def add_bubble(self, text: str, size_class: str = 'm', category: str = 'muted',
                   tooltip: str = "") -> BubbleItem:
        bubble = BubbleItem(text, size_class, category, tooltip, parent=self)
        self._layout.addWidget(bubble)
        return bubble

    def add_from_specs(self, specs: List[Dict[str, str]]):
        """specs: [{'text','size','category','tooltip'}]"""
        self.clear()
        for spec in specs:
            self.add_bubble(spec.get('text', ''), spec.get('size', 'm'),
                            spec.get('category', 'muted'), spec.get('tooltip', ''))
        self.updateGeometry()


class StudentProfileView(QWidget):
    """学生资料页（点开学生磁贴后进入）

    把该生**已编辑好的各种信息**以「或大或小的气泡/标签」呈现：
    身份信息、标签、头衔、成绩统计、各科目表现、最近考试与评语等。
    气泡大小按重要性/数值自动分档，形成 tag cloud 效果。
    """

    edit_requested = Signal(object)      # 携带 Student
    delete_requested = Signal(object)    # 携带 Student
    detail_requested = Signal(object)    # 携带 Student（打开成绩详情）
    back_requested = Signal()

    def __init__(self, db: 'DatabaseManager', parent=None):
        super().__init__(parent)
        self.db = db
        self.student: Optional[Student] = None
        self.setup_ui()

    # ------------------------------------------------------------------ UI
    def setup_ui(self):
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        outer.setSpacing(8)

        # 顶部：头像 + 姓名 + 操作
        header = QFrame()
        header.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['card_bg']};
                border: 1px solid {COLORS['border']};
                border-radius: 14px;
            }}
        """)
        head_row = QHBoxLayout(header)
        head_row.setContentsMargins(16, 12, 16, 12)
        head_row.setSpacing(14)

        self.avatar_label = QLabel("?")
        self.avatar_label.setFixedSize(64, 64)
        self.avatar_label.setAlignment(Qt.AlignCenter)
        self.avatar_label.setStyleSheet(
            f"QLabel {{ background-color: {COLORS['primary']}; color: white; "
            f"border-radius: 32px; font-size: 26px; font-weight: bold; "
            f"font-family: {FONT_STACK_CSS}; }}")
        head_row.addWidget(self.avatar_label)

        name_box = QVBoxLayout()
        name_box.setSpacing(2)
        self.name_label = QLabel("—")
        self.name_label.setStyleSheet(
            f"font-size: 24px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS}; border: none;")
        name_box.addWidget(self.name_label)
        self.sub_label = QLabel("")
        self.sub_label.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS}; border: none;")
        name_box.addWidget(self.sub_label)
        head_row.addLayout(name_box, stretch=1)

        for text, color, slot in (
                ("✏️ 编辑资料", COLORS['primary'], lambda: self.edit_requested.emit(self.student)),
                ("📈 成绩详情", COLORS['secondary'], lambda: self.detail_requested.emit(self.student)),
                ("🗑 删除", COLORS['danger'], lambda: self.delete_requested.emit(self.student)),
        ):
            btn = QPushButton(text)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: {color}; color: white; border: none;
                    border-radius: 14px; padding: 7px 14px; font-size: 12px;
                    font-family: {FONT_STACK_CSS};
                }}
                QPushButton:hover {{ opacity: 0.85; }}
            """)
            btn.clicked.connect(slot)
            head_row.addWidget(btn)

        outer.addWidget(header)

        # 提示
        self.hint_label = QLabel("💡 下方气泡展示该学生的各类信息；气泡大小表示重要程度或数值高低")
        self.hint_label.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        outer.addWidget(self.hint_label)

        # 气泡云
        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self.scroll.setFrameShape(QFrame.NoFrame)
        self.scroll.setStyleSheet("QScrollArea { background: transparent; }")
        self.cloud = BubbleCloud()
        self.cloud.setStyleSheet(
            f"BubbleCloud {{ background-color: {COLORS['background']}; "
            f"border: 1px solid {COLORS['border']}; border-radius: 14px; }}")
        self.scroll.setWidget(self.cloud)
        outer.addWidget(self.scroll, stretch=1)

    # --------------------------------------------------------------- 渲染
    def render(self, student: Optional[Student]):
        self.student = student
        if student is None:
            self.name_label.setText("—")
            self.cloud.clear()
            return

        self.avatar_label.setText(student.name[0] if student.name else '?')
        self.avatar_label.setStyleSheet(
            f"QLabel {{ background-color: {student.avatar_color or COLORS['primary']}; "
            f"color: white; border-radius: 32px; font-size: 26px; font-weight: bold; "
            f"font-family: {FONT_STACK_CSS}; }}")
        self.name_label.setText(student.name or "（未命名）")
        self.sub_label.setText(
            f"学号 {student.student_no}　|　班级 {student.class_name or '未分班'}"
            f"　|　创建于 {(student.created_at or '')[:10]}")

        grades = self.db.get_grades_by_student(student.id or 0)
        self.cloud.add_from_specs(self._build_specs(student, grades))

    def _build_specs(self, student: Student, grades: List['Grade']) -> List[Dict[str, str]]:
        """把学生信息转换为气泡规格（text/size/category/tooltip）"""
        specs: List[Dict[str, str]] = []

        # 1) 身份
        specs.append({'text': f"👤 {student.name}", 'size': 'xl', 'category': 'identity',
                      'tooltip': f"姓名：{student.name}"})
        specs.append({'text': f"🏫 {student.class_name or '未分班'}", 'size': 'l',
                      'category': 'identity',
                      'tooltip': f"班级：{student.class_name or '未分班'}"})
        specs.append({'text': f"🆔 {student.student_no}", 'size': 'm', 'category': 'identity',
                      'tooltip': f"学号：{student.student_no}"})
        specs.append({'text': f"🗓 {len(grades)} 次考试", 'size': 's', 'category': 'muted',
                      'tooltip': "该生已录入的考试记录条数"})

        # 2) 标签（越多人共有 → 气泡越大，体现班级共性）
        if student.tags:
            peers = self.db.get_all_students()
            same_class = [s for s in peers if s.class_name == student.class_name] or peers
            total = max(len(same_class), 1)
            for tag in student.tags:
                share = sum(1 for s in same_class if tag in (s.tags or [])) / total
                size = 'l' if share >= 0.3 else ('m' if share >= 0.1 else 's')
                specs.append({
                    'text': f"🏷️ {tag}", 'size': size, 'category': 'tag',
                    'tooltip': f"标签：{tag}（本班 {int(share * 100)}% 的学生拥有）"})

        # 3) 头衔
        for title in student.titles:
            specs.append({'text': f"⭐ {title}", 'size': 'l', 'category': 'title',
                          'tooltip': f"头衔：{title}"})

        # 4) 成绩统计
        if grades:
            scores = [g.score for g in grades]
            avg = sum(scores) / len(scores)
            best, worst = max(scores), min(scores)
            subjects = sorted({g.course_name for g in grades})
            avg_size = 'xl' if avg >= 85 else ('l' if avg >= 70 else 'm')
            avg_cat = 'score' if avg >= 85 else ('subject' if avg >= 60 else 'warn')
            specs.append({'text': f"📊 平均 {avg:.1f}", 'size': avg_size, 'category': avg_cat,
                          'tooltip': f"全部 {len(grades)} 条成绩的平均分"})
            specs.append({'text': f"⬆️ 最高 {best:.1f}", 'size': 'm', 'category': 'score',
                          'tooltip': "最高分"})
            specs.append({'text': f"⬇️ 最低 {worst:.1f}", 'size': 'm', 'category': 'score',
                          'tooltip': "最低分"})

            # 各科目平均分：按分数分档大小
            subject_avg: Dict[str, List[float]] = {}
            for g in grades:
                subject_avg.setdefault(g.course_name, []).append(g.score)
            for subject, vals in sorted(subject_avg.items(),
                                        key=lambda kv: -(sum(kv[1]) / len(kv[1]))):
                s_avg = sum(vals) / len(vals)
                size = 'l' if s_avg >= 90 else ('m' if s_avg >= 75 else 's')
                cat = 'subject' if s_avg >= 75 else ('warn' if s_avg < 60 else 'subject')
                specs.append({
                    'text': f"{subject} {s_avg:.0f}", 'size': size, 'category': cat,
                    'tooltip': f"{subject}：{len(vals)} 次记录，平均 {s_avg:.1f} 分"})

            # 最近考试
            latest = sorted(grades, key=lambda g: g.exam_date or '')[-1]
            specs.append({'text': f"🕐 最近 {latest.exam_date} {latest.course_name} "
                                  f"{latest.score:.0f}",
                          'size': 's', 'category': 'muted', 'tooltip': "最近一次考试记录"})
            if latest.comment:
                specs.append({'text': f"💬 {latest.comment}", 'size': 'm', 'category': 'muted',
                              'tooltip': "最近一次考试的评语"})
        else:
            specs.append({'text': "📭 暂无成绩数据（可在成绩管理导入/录入）", 'size': 'm',
                          'category': 'muted', 'tooltip': "尚未录入成绩"})

        return specs


# ==================== 对话框组件 ====================
class StudentDetailDialog(QDialog):
    """学生详情对话框"""

    def __init__(self, student: Student, db: DatabaseManager, parent=None):
        super().__init__(parent)
        self.student = student
        self.db = db
        self.setWindowTitle(f"学生详情 - {student.name}")
        self.setMinimumSize(500, 500)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(12)

        # 头部信息
        header = QFrame()
        header.setStyleSheet(f"background-color: {COLORS['primary_light']}; border-radius: 10px;")
        header_layout = QVBoxLayout(header)
        header_layout.setContentsMargins(20, 15, 20, 15)

        name_label = QLabel(f"{self.student.name}")
        name_label.setStyleSheet(f"font-size: 24px; font-weight: bold; color: {COLORS['primary']};")
        header_layout.addWidget(name_label)

        info_text = f"学号: {self.student.student_no} | 班级: {self.student.class_name or '未分配'}"
        info_label = QLabel(info_text)
        info_label.setStyleSheet(f"font-size: 13px; color: {COLORS['text_secondary']};")
        header_layout.addWidget(info_label)

        if self.student.tags:
            tags_text = ' '.join([f'#{t}' for t in self.student.tags])
            tags_label = QLabel(tags_text)
            tags_label.setStyleSheet(f"font-size: 12px; color: {COLORS['secondary']};")
            header_layout.addWidget(tags_label)

        if self.student.titles:
            titles_text = ' '.join([f'⭐{t}' for t in self.student.titles])
            titles_label = QLabel(titles_text)
            titles_label.setStyleSheet(f"font-size: 12px; color: {COLORS['warning']};")
            header_layout.addWidget(titles_label)

        layout.addWidget(header)

        # 近期科目情况
        grades_label = QLabel("📈 近期科目情况")
        grades_label.setStyleSheet(f"font-size: 16px; font-weight: bold; color: {COLORS['text_primary']};")
        layout.addWidget(grades_label)

        # 成绩表格
        self.grades_table = QTableWidget()
        self.grades_table.setColumnCount(5)
        self.grades_table.setHorizontalHeaderLabels(['科目', '成绩', '满分', '考试日期', '类型'])
        self.grades_table.horizontalHeader().setStretchLastSection(True)
        self.grades_table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.grades_table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.grades_table.setAlternatingRowColors(True)
        self.grades_table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                font-size: 12px;
                font-family: {FONT_STACK_CSS};
            }}
            QHeaderView::section {{
                background-color: {COLORS['primary_light']};
                font-weight: bold;
                padding: 6px;
                border: none;
            }}
        """)
        layout.addWidget(self.grades_table)

        # 关闭按钮
        close_btn = QPushButton("关闭")
        close_btn.setStyleSheet(self._get_button_style())
        close_btn.clicked.connect(self.accept)
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()
        btn_layout.addWidget(close_btn)
        layout.addLayout(btn_layout)

    def _get_button_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['primary']};
                color: white;
                border: none;
                border-radius: 6px;
                padding: 8px 20px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{
                background-color: {COLORS['primary_dark']};
            }}
        """

    def load_data(self):
        grades = self.db.get_grades_by_student(self.student.id or 0)
        self.grades_table.setRowCount(len(grades))
        for i, g in enumerate(grades):
            self.grades_table.setItem(i, 0, QTableWidgetItem(g.course_name))
            score_item = QTableWidgetItem(f"{g.score:.1f}")
            if g.score >= 85:
                score_item.setForeground(QColor('#00B894'))
            elif g.score >= 60:
                score_item.setForeground(QColor('#FDCB6E'))
            else:
                score_item.setForeground(QColor('#FF6B6B'))
            self.grades_table.setItem(i, 1, score_item)
            self.grades_table.setItem(i, 2, QTableWidgetItem(f"{g.full_score:.0f}"))
            self.grades_table.setItem(i, 3, QTableWidgetItem(g.exam_date))
            self.grades_table.setItem(i, 4, QTableWidgetItem(g.exam_type))


class CourseEditDialog(QDialog):
    """课程编辑对话框"""

    def __init__(self, course: Optional[Course] = None, parent=None):
        super().__init__(parent)
        self.course = course
        self.setWindowTitle("新建课程" if course is None else "编辑课程")
        self.setMinimumSize(420, 400)
        self.setup_ui()
        if course:
            self.load_course_data()

    def setup_ui(self):
        layout = QFormLayout(self)
        layout.setSpacing(10)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("输入课程名称")
        layout.addRow("课程名称:", self.name_edit)

        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        layout.addRow("日期:", self.date_edit)

        self.start_time = QTimeEdit(QTime(8, 0))
        self.start_time.setDisplayFormat("HH:mm")
        layout.addRow("开始时间:", self.start_time)

        self.end_time = QTimeEdit(QTime(9, 0))
        self.end_time.setDisplayFormat("HH:mm")
        layout.addRow("结束时间:", self.end_time)

        self.location_edit = QLineEdit()
        self.location_edit.setPlaceholderText("教室/地点")
        layout.addRow("地点:", self.location_edit)

        self.teacher_edit = QLineEdit()
        self.teacher_edit.setPlaceholderText("授课教师")
        layout.addRow("教师:", self.teacher_edit)

        # 颜色选择
        self.color_btn = QPushButton("选择颜色")
        self.color_btn.setStyleSheet(f"background-color: {COLORS['primary']}; color: white; padding: 6px;")
        self.color_btn.clicked.connect(self.choose_color)
        self.selected_color = COLORS['primary']
        layout.addRow("颜色:", self.color_btn)

        self.desc_edit = QTextEdit()
        self.desc_edit.setPlaceholderText("课程描述...")
        self.desc_edit.setMaximumHeight(80)
        layout.addRow("描述:", self.desc_edit)

        # 按钮
        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("保存")
        btn_box.button(QDialogButtonBox.Cancel).setText("取消")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addRow(btn_box)

        # 样式
        for edit in [self.name_edit, self.location_edit, self.teacher_edit, self.desc_edit]:
            edit.setStyleSheet(f"""
                padding: 6px 10px;
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
            """)

    def choose_color(self):
        color = QColorDialog.getColor(QColor(self.selected_color), self, "选择颜色")
        if color.isValid():
            self.selected_color = color.name()
            self.color_btn.setStyleSheet(f"background-color: {self.selected_color}; color: white; padding: 6px;")

    def load_course_data(self):
        if not self.course:
            return
        self.name_edit.setText(self.course.name)
        self.date_edit.setDate(QDate.fromString(self.course.course_date, "yyyy-MM-dd"))
        self.start_time.setTime(QTime.fromString(self.course.start_time, "HH:mm"))
        self.end_time.setTime(QTime.fromString(self.course.end_time, "HH:mm"))
        self.location_edit.setText(self.course.location)
        self.teacher_edit.setText(self.course.teacher)
        self.selected_color = self.course.color
        self.color_btn.setStyleSheet(f"background-color: {self.selected_color}; color: white; padding: 6px;")
        self.desc_edit.setText(self.course.description)

    def get_course_data(self) -> Course:
        return Course(
            id=self.course.id if self.course else None,
            name=self.name_edit.text().strip(),
            course_date=self.date_edit.date().toString("yyyy-MM-dd"),
            start_time=self.start_time.time().toString("HH:mm"),
            end_time=self.end_time.time().toString("HH:mm"),
            location=self.location_edit.text().strip(),
            teacher=self.teacher_edit.text().strip(),
            color=self.selected_color,
            description=self.desc_edit.toPlainText().strip(),
        )


class FeedbackDialog(QDialog):
    """反馈编辑对话框"""

    def __init__(self, feedback: Optional[Feedback] = None, parent=None):
        super().__init__(parent)
        self.feedback = feedback
        self.setWindowTitle("新建反馈" if feedback is None else "编辑反馈")
        self.setMinimumSize(450, 450)
        self.setup_ui()
        if feedback:
            self.load_feedback_data()

    def setup_ui(self):
        layout = QFormLayout(self)
        layout.setSpacing(10)

        self.content_edit = QTextEdit()
        self.content_edit.setPlaceholderText("输入反馈内容...")
        self.content_edit.setMinimumHeight(100)
        layout.addRow("内容:", self.content_edit)

        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        layout.addRow("日期:", self.date_edit)

        self.time_edit = QTimeEdit(QTime.currentTime())
        self.time_edit.setDisplayFormat("HH:mm")
        layout.addRow("时间:", self.time_edit)

        self.class_edit = QLineEdit()
        self.class_edit.setPlaceholderText("班级名称")
        layout.addRow("班级:", self.class_edit)

        self.category_combo = QComboBox()
        self.category_combo.addItems(['学习', '生活', '心理', '安全', '健康', '其他'])
        layout.addRow("类别:", self.category_combo)

        self.status_combo = QComboBox()
        self.status_combo.addItems(['待处理', '处理中', '已解决'])
        layout.addRow("状态:", self.status_combo)

        self.priority_spin = QSpinBox()
        self.priority_spin.setRange(1, 5)
        self.priority_spin.setValue(1)
        layout.addRow("优先级:", self.priority_spin)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("保存")
        btn_box.button(QDialogButtonBox.Cancel).setText("取消")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addRow(btn_box)

    def load_feedback_data(self):
        if not self.feedback:
            return
        self.content_edit.setText(self.feedback.content)
        self.date_edit.setDate(QDate.fromString(self.feedback.date, "yyyy-MM-dd"))
        self.time_edit.setTime(QTime.fromString(self.feedback.time, "HH:mm"))
        self.class_edit.setText(self.feedback.class_name)
        self.category_combo.setCurrentText(self.feedback.category)
        self.status_combo.setCurrentText(self.feedback.status)
        self.priority_spin.setValue(self.feedback.priority)

    def get_feedback_data(self) -> Feedback:
        return Feedback(
            id=self.feedback.id if self.feedback else None,
            date=self.date_edit.date().toString("yyyy-MM-dd"),
            time=self.time_edit.time().toString("HH:mm"),
            class_name=self.class_edit.text().strip(),
            category=self.category_combo.currentText(),
            content=self.content_edit.toPlainText().strip(),
            student_name="",
            status=self.status_combo.currentText(),
            priority=self.priority_spin.value(),
        )


class MultiDeleteDialog(QDialog):
    """通用批量删除选择对话框
    支持：搜索过滤、勾选多条（批量）或只勾一条（单个）、已选计数。
    records: List[Dict] 每项含 id / text / detail(可选说明文字)
    通过 selected_ids() 获取勾选记录 id 列表。
    """

    def __init__(self, records: List[Dict[str, Any]], title: str = "删除记录",
                 instruction: str = "勾选要删除的记录（可多选，支持单个删除）：",
                 empty_text: str = "暂无记录", parent=None):
        super().__init__(parent)
        self.records = list(records)
        self.setWindowTitle(title)
        self.setMinimumSize(460, 480)
        self.setup_ui(instruction, empty_text)

    def setup_ui(self, instruction: str, empty_text: str):
        layout = QVBoxLayout(self)
        layout.setSpacing(8)

        tip = QLabel(instruction)
        tip.setWordWrap(True)
        tip.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        layout.addWidget(tip)

        # 搜索框
        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("🔍 输入关键词过滤（姓名/学号/班级/日期/科目...）")
        self.search_edit.textChanged.connect(self._apply_filter)
        layout.addWidget(self.search_edit)

        # 记录列表
        self.list_widget = QListWidget()
        self.list_widget.setSelectionMode(QAbstractItemView.NoSelection)
        self._items = []      # [(item, record_id)]
        for rec in self.records:
            item = QListWidgetItem(str(rec.get('text', '')))
            item.setFlags(item.flags() | Qt.ItemIsUserCheckable)
            item.setCheckState(Qt.Unchecked)
            detail = rec.get('detail') or ''
            if detail:
                item.setToolTip(detail)
                item.setText(f"{rec.get('text', '')}\n{detail}")
            item.setData(Qt.UserRole, rec.get('id'))
            self.list_widget.addItem(item)
            self._items.append((item, rec.get('id')))
        # 用户勾选/取消勾选任一行时实时刷新计数与按钮状态
        self.list_widget.itemChanged.connect(lambda _item: self._refresh_count())
        layout.addWidget(self.list_widget, stretch=1)

        # 操作行
        row = QHBoxLayout()
        self.select_all_btn = QPushButton("☑ 全选")
        self.select_all_btn.setStyleSheet(self._mini_btn_style())
        self.select_all_btn.clicked.connect(lambda: self._set_all_checked(True))
        row.addWidget(self.select_all_btn)

        self.clear_btn = QPushButton("清空选择")
        self.clear_btn.setStyleSheet(self._mini_btn_style())
        self.clear_btn.clicked.connect(lambda: self._set_all_checked(False))
        row.addWidget(self.clear_btn)

        row.addStretch()
        self.count_label = QLabel()
        self.count_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 12px;")
        row.addWidget(self.count_label)
        layout.addLayout(row)

        # 按钮行
        btn_row = QHBoxLayout()
        btn_row.addStretch()
        cancel_btn = QPushButton("取消")
        cancel_btn.setStyleSheet(self._plain_btn_style())
        cancel_btn.clicked.connect(self.reject)
        btn_row.addWidget(cancel_btn)

        self.ok_btn = QPushButton("删除选中")
        self.ok_btn.setStyleSheet(self._danger_btn_style())
        self.ok_btn.setEnabled(False)
        self.ok_btn.clicked.connect(self.accept)
        btn_row.addWidget(self.ok_btn)
        layout.addLayout(btn_row)

        self._refresh_count()

    def _mini_btn_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['primary_light']}; color: {COLORS['primary']};
                border: none; border-radius: 6px; padding: 5px 12px; font-size: 12px;
            }}
            QPushButton:hover {{ background-color: {COLORS['hover']}; }}
        """

    def _plain_btn_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['card_bg']}; color: {COLORS['text_primary']};
                border: 1px solid {COLORS['border']}; border-radius: 6px; padding: 7px 18px;
                font-size: 13px;
            }}
            QPushButton:hover {{ border-color: {COLORS['primary']}; }}
        """

    def _danger_btn_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['danger']}; color: white; border: none;
                border-radius: 6px; padding: 7px 18px; font-size: 13px; font-weight: bold;
            }}
            QPushButton:hover {{ background-color: #E25555; }}
            QPushButton:disabled {{ background-color: #F0C6C6; }}
        """

    def _apply_filter(self, keyword: str):
        kw = keyword.strip().lower()
        for item, _ in self._items:
            item.setHidden(bool(kw) and kw not in item.text().lower())
        self._refresh_count()

    def _set_all_checked(self, checked: bool):
        """全选(checked=True)只作用于当前可见行；清空选择(checked=False)清空全部（含被过滤隐藏的已选行）"""
        state = Qt.Checked if checked else Qt.Unchecked
        for item, _ in self._items:
            if checked and item.isHidden():
                continue
            item.setCheckState(state)

    def _refresh_count(self):
        total = len(self._items)
        sel = sum(1 for item, _ in self._items if item.checkState() == Qt.Checked)
        self.count_label.setText(f"已选 {sel} / 共 {total} 条")
        self.ok_btn.setEnabled(sel > 0)
        if sel > 0:
            self.ok_btn.setText(f"删除选中 ({sel})")
        else:
            self.ok_btn.setText("删除选中")

    def selected_ids(self) -> List[int]:
        """返回勾选记录的 id 列表"""
        return [rid for item, rid in self._items if item.checkState() == Qt.Checked]

    def selected_records(self) -> List[Dict[str, Any]]:
        """返回勾选记录的原始 dict 列表"""
        ids = set(self.selected_ids())
        return [r for r in self.records if r.get('id') in ids]


# ==================== 过渡动画 ====================
class AnimatedStack(QWidget):
    """带「滑动 + 淡入」过渡的页面容器

    用于学生管理的 班级总览 ↔ 班级详情 等“点入”切换，
    比 QStackedWidget 的硬切换更丝滑（新页滑入 + 旧页视差后退）。
    """

    def __init__(self, parent=None, duration: int = 280):
        super().__init__(parent)
        self._pages: List[QWidget] = []
        self._index = -1
        self._duration = duration
        self._group: Optional[QParallelAnimationGroup] = None

    # ---- 页面管理 ----
    def addWidget(self, page: QWidget) -> int:
        page.setParent(self)
        page.setGeometry(0, 0, self.width(), self.height())
        page.hide()
        self._pages.append(page)
        if self._index < 0:
            self._index = 0
            page.move(0, 0)
            page.show()
        return len(self._pages) - 1

    def count(self) -> int:
        return len(self._pages)

    def currentIndex(self) -> int:
        return self._index

    def currentWidget(self) -> Optional[QWidget]:
        return self._pages[self._index] if 0 <= self._index < len(self._pages) else None

    def widget(self, index: int) -> Optional[QWidget]:
        return self._pages[index] if 0 <= index < len(self._pages) else None

    def resizeEvent(self, event):
        super().resizeEvent(event)
        for p in self._pages:
            p.setGeometry(0, 0, self.width(), self.height())

    # ---- 切换 ----
    def _settle(self):
        """结束正在进行的动画并把非当前页归位隐藏"""
        if self._group is not None:
            try:
                self._group.stop()
            except Exception:
                pass
            self._group = None
        for i, p in enumerate(self._pages):
            if i == self._index:
                p.move(0, 0)
                p.show()
            else:
                p.hide()
                p.move(0, 0)

    def setCurrentIndex(self, index: int, animate: bool = True, direction: int = 1):
        if not (0 <= index < len(self._pages)):
            return
        if index == self._index:
            self._settle()
            return
        old = self.currentWidget()
        new = self._pages[index]
        self._settle()
        self._index = index

        w = max(self.width(), 1)
        new.setGeometry(0, 0, w, self.height())
        new.move(direction * w, 0)
        new.show()
        new.raise_()

        if old is None or not animate or self._duration <= 0:
            if old is not None:
                old.hide()
            new.move(0, 0)
            return

        anim_new = QPropertyAnimation(new, b"pos", self)
        anim_new.setDuration(self._duration)
        anim_new.setStartValue(QPoint(direction * w, 0))
        anim_new.setEndValue(QPoint(0, 0))
        anim_new.setEasingCurve(QEasingCurve.OutCubic)

        anim_old = QPropertyAnimation(old, b"pos", self)
        anim_old.setDuration(self._duration)
        anim_old.setStartValue(QPoint(0, 0))
        anim_old.setEndValue(QPoint(-direction * w // 4, 0))   # 视差：旧页只退 1/4
        anim_old.setEasingCurve(QEasingCurve.OutCubic)

        group = QParallelAnimationGroup(self)
        group.addAnimation(anim_new)
        group.addAnimation(anim_old)
        group.finished.connect(self._settle)
        self._group = group
        group.start()


def cross_fade(widget: QWidget, mid_callback, duration: int = 180):
    """局部内容切换过渡

    先**同步**执行 mid_callback 更新内容（保证行为确定、数据立即生效），
    再对新内容做一次快速淡入，视觉上是平滑过渡而不是硬闪。
    复用同一个 QGraphicsOpacityEffect，避免动画目标被中途销毁。
    """
    if widget is None:
        mid_callback()
        return None
    mid_callback()                      # 内容先更新（同步、可立即断言）
    try:
        effect = getattr(widget, '_fade_effect', None)
        if effect is None:
            effect = QGraphicsOpacityEffect(widget)
            widget.setGraphicsEffect(effect)
            widget._fade_effect = effect
        prev = getattr(widget, '_fade_anim', None)
        if prev is not None:
            try:
                prev.stop()
            except Exception:
                pass
        anim = QPropertyAnimation(effect, b"opacity", widget)
        anim.setDuration(duration)
        anim.setStartValue(0.35)
        anim.setEndValue(1.0)
        anim.setEasingCurve(QEasingCurve.OutCubic)
        widget._fade_anim = anim
        anim.start()
        return anim
    except Exception:
        return None


# ==================== 模块页面 ====================
class BaseModulePage(QWidget):
    """模块页面基类"""

    def __init__(self, db: DatabaseManager, parent=None):
        super().__init__(parent)
        self.db = db
        self.setup_ui()

    def setup_ui(self):
        pass

    def refresh(self):
        """刷新数据"""
        pass


class StudentModule(BaseModulePage):
    """学生管理系统模块

    两级结构：
      第 1 级 —— 班级总览（按班级聚合的磁贴，点入查看该班学生）
      第 2 级 —— 班级详情（该班学生磁贴，可返回）
    切换使用 AnimatedStack 滑动 + 视差过渡，点入更丝滑。
    """

    COLUMNS_CLASS = 4          # 班级网格列数
    COLUMNS_STUDENT = 5        # 学生网格列数

    def __init__(self, db: DatabaseManager, parent=None):
        super().__init__(db, parent)
        self.students: List[Student] = []
        self.current_class: Optional[str] = None
        self._search_keyword = ""
        # 导航历史：('classes', None) | ('class', 班级名) | ('search', 关键词)
        self._view: Tuple[str, Optional[str]] = ('classes', None)
        self._history: List[Tuple[str, Optional[str]]] = []

    # ------------------------------------------------------------------ UI
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        # ---- 顶部工具栏 ----
        toolbar = QHBoxLayout()

        self.title_label = QLabel("👥 学生管理 · 班级总览")
        self.title_label.setStyleSheet(
            f"font-size: 22px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(self.title_label)
        toolbar.addStretch()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索学生姓名/学号/班级...")
        self.search_edit.setFixedWidth(230)
        self.search_edit.textChanged.connect(self.on_search)
        self.search_edit.setStyleSheet(f"""
            QLineEdit {{
                padding: 8px 12px;
                border: 1px solid {COLORS['border']};
                border-radius: 18px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
                background-color: {COLORS['card_bg']};
            }}
            QLineEdit:focus {{ border-color: {COLORS['primary']}; }}
        """)
        toolbar.addWidget(self.search_edit)

        import_btn = QPushButton("📥 导入学生")
        import_btn.setStyleSheet(self._get_btn_style(COLORS['success']))
        import_btn.clicked.connect(self.import_students)
        toolbar.addWidget(import_btn)

        add_btn = QPushButton("➕ 新建学生")
        add_btn.setStyleSheet(self._get_btn_style(COLORS['primary']))
        add_btn.clicked.connect(self.add_student)
        toolbar.addWidget(add_btn)

        delete_btn = QPushButton("🗑 删除学生")
        delete_btn.setStyleSheet(self._get_btn_style(COLORS['danger']))
        delete_btn.clicked.connect(self.delete_students)
        toolbar.addWidget(delete_btn)

        layout.addLayout(toolbar)

        # ---- 页面导航栏：返回上一页 + 面包屑（可点击跳回之前任意一页）----
        nav_bar = QFrame()
        nav_bar.setObjectName("navBar")
        nav_bar.setStyleSheet(f"""
            QFrame#navBar {{
                background-color: {COLORS['primary_light']};
                border: 1px solid {COLORS['border']};
                border-radius: 10px;
            }}
        """)
        nav_row = QHBoxLayout(nav_bar)
        nav_row.setContentsMargins(8, 6, 8, 6)
        nav_row.setSpacing(6)

        self.back_btn = QPushButton("← 上一页")
        self.back_btn.setToolTip("返回上一个浏览过的页面")
        self.back_btn.setCursor(Qt.PointingHandCursor)
        self.back_btn.setStyleSheet(self._get_nav_action_style())
        self.back_btn.clicked.connect(self.go_back)
        nav_row.addWidget(self.back_btn)

        home_btn = QPushButton("🏠 班级总览")
        home_btn.setToolTip("回到班级总览（首页）")
        home_btn.setCursor(Qt.PointingHandCursor)
        home_btn.setStyleSheet(self._get_nav_action_style())
        home_btn.clicked.connect(lambda: self.show_class_overview(animate=True))
        nav_row.addWidget(home_btn)

        # 面包屑容器（动态重建：🏠 班级总览 › 高一(1)班 › 搜索结果 …）
        self.crumb_container = QWidget(nav_bar)
        self.crumb_layout = QHBoxLayout(self.crumb_container)
        self.crumb_layout.setContentsMargins(0, 0, 0, 0)
        self.crumb_layout.setSpacing(2)
        nav_row.addWidget(self.crumb_container)
        nav_row.addStretch()

        self.nav_hint = QLabel("点击路径中的任意一页可跳转回去", parent=nav_bar)
        self.nav_hint.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        nav_row.addWidget(self.nav_hint)

        layout.addWidget(nav_bar)

        # ---- 两级页面容器（带过渡动画）----
        self.stack = AnimatedStack(duration=300)

        # 第 1 级：班级总览
        self.class_page = QWidget()
        class_page_layout = QVBoxLayout(self.class_page)
        class_page_layout.setContentsMargins(0, 0, 0, 0)
        self.class_scroll = QScrollArea()
        self.class_scroll.setWidgetResizable(True)
        self.class_scroll.setFrameShape(QFrame.NoFrame)
        self.class_scroll.setStyleSheet("QScrollArea { background: transparent; }")
        self.class_container = QWidget()
        self.class_container.setStyleSheet("background: transparent;")
        self.class_layout = QGridLayout(self.class_container)
        self.class_layout.setSpacing(16)
        self.class_layout.setContentsMargins(5, 5, 5, 5)
        self.class_scroll.setWidget(self.class_container)
        class_page_layout.addWidget(self.class_scroll)
        self.stack.addWidget(self.class_page)

        # 第 2 级：班级详情（学生磁贴）
        self.detail_page = QWidget()
        detail_layout = QVBoxLayout(self.detail_page)
        detail_layout.setContentsMargins(0, 0, 0, 0)
        self.detail_hint = QLabel()
        self.detail_hint.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        detail_layout.addWidget(self.detail_hint)
        self.student_scroll = QScrollArea()
        self.student_scroll.setWidgetResizable(True)
        self.student_scroll.setFrameShape(QFrame.NoFrame)
        self.student_scroll.setStyleSheet("QScrollArea { background: transparent; }")
        self.tiles_container = QWidget()
        self.tiles_container.setStyleSheet("background: transparent;")
        self.tiles_layout = QGridLayout(self.tiles_container)
        self.tiles_layout.setSpacing(15)
        self.tiles_layout.setContentsMargins(5, 5, 5, 5)
        self.student_scroll.setWidget(self.tiles_container)
        detail_layout.addWidget(self.student_scroll)
        self.stack.addWidget(self.detail_page)

        # 第 3 级：学生资料页（点开磁贴 → 气泡/标签视图）
        self.profile_page = StudentProfileView(self.db)
        self.profile_page.edit_requested.connect(self.on_student_edit)
        self.profile_page.delete_requested.connect(
            lambda st: self._confirm_delete_students([st.id]))
        self.profile_page.detail_requested.connect(
            lambda st: self.on_student_double_clicked(st))
        self.stack.addWidget(self.profile_page)

        layout.addWidget(self.stack, stretch=1)

        # ---- 底部统计 ----
        self.stats_label = QLabel()
        self.stats_label.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(self.stats_label)

    def _get_btn_style(self, color):
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 18px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
                font-weight: 500;
            }}
            QPushButton:hover {{
                opacity: 0.85;
            }}
        """

    # ------------------------------------------------------- 导航栏样式
    def _get_nav_action_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['card_bg']}; color: {COLORS['primary']};
                border: 1px solid {COLORS['border']}; border-radius: 14px;
                padding: 5px 12px; font-size: 12px; font-weight: 500;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{ border-color: {COLORS['primary']}; background-color: {COLORS['hover']}; }}
            QPushButton:disabled {{ color: {COLORS['text_secondary']}; background-color: transparent; }}
        """

    def _get_crumb_style(self, active: bool) -> str:
        if active:
            return f"""
                QPushButton {{
                    background-color: transparent; color: {COLORS['text_primary']};
                    border: none; font-size: 12px; font-weight: bold; padding: 4px 6px;
                    font-family: {FONT_STACK_CSS};
                }}
            """
        return f"""
            QPushButton {{
                background-color: transparent; color: {COLORS['primary']};
                border: none; font-size: 12px; padding: 4px 6px; text-decoration: underline;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{ color: {COLORS['primary_dark']}; }}
        """

    def _crumb_label(self, entry) -> str:
        """把历史条目转成面包屑文字"""
        kind, value = entry
        if kind == 'classes':
            return "🏠 班级总览"
        if kind == 'search':
            return f"🔍 搜索：{value}"
        if kind == 'student':
            stu = self.db.get_student_by_id(int(value)) if value not in (None, '') else None
            return f"👤 {stu.name}" if stu else "👤 学生资料"
        return f"📚 {value}"

    def _update_nav_bar(self):
        """重建面包屑：历史 + 当前页，点击任意一段可跳回该页"""
        while self.crumb_layout.count():
            item = self.crumb_layout.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()

        trail = list(self._history) + [self._view]
        for idx, entry in enumerate(trail):
            if idx > 0:
                sep = QLabel("›", parent=self.crumb_container)
                sep.setStyleSheet(
                    f"color: {COLORS['text_secondary']}; font-size: 12px;")
                self.crumb_layout.addWidget(sep)
            is_current = (idx == len(trail) - 1)
            btn = QPushButton(self._crumb_label(entry), parent=self.crumb_container)
            btn.setCursor(Qt.ArrowCursor if is_current else Qt.PointingHandCursor)
            btn.setStyleSheet(self._get_crumb_style(is_current))
            btn.setEnabled(not is_current)
            if not is_current:
                btn.clicked.connect(lambda _=False, i=idx: self.jump_to(i))
            self.crumb_layout.addWidget(btn)

        self.back_btn.setEnabled(bool(self._history))
        # 「← 上一页」在无可返回页面时隐藏（ui 更干净），面包屑始终展示
        self.back_btn.setVisible(bool(self._history))

    # ------------------------------------------------------- 导航（历史）
    def navigate(self, kind: str, value: Optional[str] = None, animate: bool = True,
                 direction: int = 1):
        """跳转到新页面并记录历史（供返回/面包屑跳转）"""
        entry = (kind, value)
        if entry == self._view:
            self._render_view(entry, animate=False, direction=direction)
            self._update_nav_bar()
            return
        self._history.append(self._view)
        self._history = self._history[-12:]     # 限制历史深度
        self._view = entry
        self._render_view(entry, animate=animate, direction=direction)
        self._update_nav_bar()

    def go_back(self):
        """回到上一页"""
        if not self._history:
            return
        self._view = self._history.pop()
        self._render_view(self._view, animate=True, direction=-1)
        self._update_nav_bar()

    def jump_to(self, index: int):
        """点击面包屑：跳回历史中的第 index 页（之后的记录丢弃）"""
        if not (0 <= index < len(self._history)):
            return
        self._view = self._history[index]
        self._history = self._history[:index]
        self._render_view(self._view, animate=True, direction=-1)
        self._update_nav_bar()

    def _render_view(self, entry, animate: bool = False, direction: int = 1):
        """渲染某个历史条目对应的页面"""
        kind, value = entry
        if kind == 'classes':
            self.current_class = None
            self._rebuild_class_overview()
            self.stack.setCurrentIndex(0, animate=animate, direction=direction)
            self.title_label.setText("👥 学生管理 · 班级总览")
            return
        if kind == 'search':
            results = self.db.search_students(value or "")
            self.current_class = f"搜索：{value}"
            self._rebuild_class_detail(f"搜索结果 · {value}", results)
            self.stack.setCurrentIndex(1, animate=animate, direction=direction)
            self.title_label.setText(f"👥 搜索结果（{len(results)}）")
            return
        if kind == 'student':
            student = self.db.get_student_by_id(int(value)) if value not in (None, '') else None
            if student is None:                       # 学生已被删除 → 回到班级总览
                self._view = ('classes', None)
                self._history = []
                self._render_view(self._view, animate=animate, direction=direction)
                return
            self.current_class = student.class_name or "未分班"
            self.profile_page.render(student)
            self.stack.setCurrentIndex(2, animate=animate, direction=direction)
            self.title_label.setText(f"👤 {student.name} · 学生资料")
            self.stats_label.setText(
                f"{student.name} · 学号 {student.student_no} · "
                f"成绩 {len(self.db.get_grades_by_student(student.id or 0))} 条")
            return
        # 班级详情
        self.current_class = value
        self._rebuild_class_detail(value or "")
        self.stack.setCurrentIndex(1, animate=animate, direction=direction)
        self.title_label.setText(f"👥 学生管理 · {value}")

    # ------------------------------------------------------- 数据与刷新
    def refresh(self):
        """刷新数据：按当前视图重建（班级总览 / 当前班级 / 搜索结果）"""
        self.students = self.db.get_all_students()
        if self._view[0] != 'classes':
            # 保持当前所在页（班级/搜索结果），数据变化后重新渲染
            self._render_view(self._view, animate=False)
        else:
            self._rebuild_class_overview()
            self.stats_label.setText(
                f"共 {len(self.students)} 名学生 · "
                f"{len(self._classes())} 个班级 · 点击班级磁贴查看学生")
        self._update_nav_bar()

    def _classes(self) -> Dict[str, List[Student]]:
        """按班级分组（未填写班级的归入「未分班」）"""
        groups: Dict[str, List[Student]] = {}
        for s in self.students:
            key = s.class_name.strip() if s.class_name and s.class_name.strip() else "未分班"
            groups.setdefault(key, []).append(s)
        # 班级名排序（未分班放最后）
        ordered = {}
        for name in sorted(groups.keys(), key=lambda n: (n == "未分班", n)):
            ordered[name] = sorted(groups[name], key=lambda s: s.name)
        return ordered

    @staticmethod
    def _clear_grid(grid: QGridLayout):
        """清空网格：只 deleteLater，绝不 setParent(None)
        （否则控件会退化为顶层窗口，在 Windows 上会闪出多个“程序窗口”）
        """
        while grid.count():
            item = grid.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()

    # --------------------------------------------------- 第 1 级：班级总览
    def _rebuild_class_overview(self):
        self._clear_grid(self.class_layout)
        groups = self._classes()
        for i, (class_name, students) in enumerate(groups.items()):
            tile = ClassTile(class_name, students, parent=self.class_container)  # 必须给父控件
            tile.clicked.connect(lambda data: self.open_class(data['class_name']))
            tile.edit_requested.connect(lambda data: self.rename_class(data['class_name']))
            row, col = divmod(i, self.COLUMNS_CLASS)
            self.class_layout.addWidget(tile, row, col)
        self.class_layout.setRowStretch(len(groups) // self.COLUMNS_CLASS + 1, 1)
        self.title_label.setText("👥 学生管理 · 班级总览")
        if not groups:
            empty = QLabel("暂无学生数据，请先「导入学生」或「新建学生」",
                           parent=self.class_container)
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(
                f"font-size: 14px; color: {COLORS['text_secondary']}; padding: 40px;")
            self.class_layout.addWidget(empty, 0, 0, 1, self.COLUMNS_CLASS)

    # --------------------------------------------------- 第 2 级：班级详情
    def open_class(self, class_name: str):
        """点入班级（带动画，并记入导航历史）"""
        self.navigate('class', class_name, animate=True, direction=1)
        self.stats_label.setText(
            f"当前班级：{class_name} · "
            f"{len([s for s in self.students if (s.class_name.strip() if s.class_name else '未分班') == class_name])} 名学生")

    def show_class_overview(self, animate: bool = True):
        """返回班级总览（首页）：清空前进历史，可直接回到起点"""
        self.current_class = None
        if self._search_keyword:
            self._search_keyword = ""
            self.search_edit.blockSignals(True)
            self.search_edit.clear()
            self.search_edit.blockSignals(False)
        # 历史中若已存在「班级总览」，回到那一条并丢弃其后的记录
        base_index = None
        for i, entry in enumerate(self._history):
            if entry[0] == 'classes':
                base_index = i
                break
        if base_index is not None:
            self._view = self._history[base_index]
            self._history = self._history[:base_index]
        else:
            self._view = ('classes', None)
            self._history = []
        self._render_view(self._view, animate=animate, direction=-1)
        self._update_nav_bar()
        self.stats_label.setText(
            f"共 {len(self.students)} 名学生 · {len(self._classes())} 个班级 · "
            f"点击班级磁贴查看学生")

    def _rebuild_class_detail(self, class_name: str, students: Optional[List[Student]] = None):
        """构建某个班级（或搜索结果）的学生磁贴"""
        self._clear_grid(self.tiles_layout)
        if students is None:
            students = [s for s in self.students
                        if (s.class_name.strip() if s.class_name else "未分班") == class_name]
        students = sorted(students, key=lambda s: s.name)

        for i, student in enumerate(students):
            tile = StudentTile(student, parent=self.tiles_container)   # 必须给父控件
            tile.clicked.connect(self.on_student_clicked)
            tile.double_clicked.connect(self.on_student_double_clicked)
            tile.edit_requested.connect(self.on_student_edit)
            tile.delete_requested.connect(lambda st: self._confirm_delete_students([st.id]))
            row, col = divmod(i, self.COLUMNS_STUDENT)
            self.tiles_layout.addWidget(tile, row, col)

        empty_slots = (self.COLUMNS_STUDENT - (len(students) % self.COLUMNS_STUDENT)) % self.COLUMNS_STUDENT
        for j in range(empty_slots):
            spacer = QWidget(self.tiles_container)                     # 必须给父控件
            spacer.setMinimumSize(120, 120)
            self.tiles_layout.addWidget(
                spacer, len(students) // self.COLUMNS_STUDENT, (len(students) % self.COLUMNS_STUDENT) + j)
        self.tiles_layout.setRowStretch(len(students) // self.COLUMNS_STUDENT + 1, 1)

        if not students:
            empty = QLabel(f"「{class_name}」暂无学生", parent=self.tiles_container)
            empty.setAlignment(Qt.AlignCenter)
            empty.setStyleSheet(
                f"font-size: 14px; color: {COLORS['text_secondary']}; padding: 40px;")
            self.tiles_layout.addWidget(empty, 0, 0, 1, self.COLUMNS_STUDENT)

        self.detail_hint.setText(
            f"📚 {class_name} · {len(students)} 名学生　"
            f"（双击查看详情，悬停磁贴右上角或右键可编辑/删除）")
        self.stats_label.setText(f"当前班级：{class_name} · {len(students)} 名学生")

    # ------------------------------------------------------------- 交互
    def on_search(self, keyword: str):
        """搜索：跨班级显示结果（带过渡，并记入导航历史）"""
        self._search_keyword = keyword.strip()
        if not self._search_keyword:
            self.show_class_overview(animate=True)
            return
        # 连续输入时只更新当前搜索结果页，不堆叠历史
        if self._view[0] == 'search':
            self._view = ('search', self._search_keyword)
            self._render_view(self._view, animate=False)
            self._update_nav_bar()
            return
        self.navigate('search', self._search_keyword, animate=True, direction=1)

    def on_student_clicked(self, data):
        """单击学生磁贴 → 进入该生资料页（气泡/标签视图，带过渡动画）"""
        student = data
        if student is None or getattr(student, 'id', None) is None:
            return
        self.navigate('student', str(student.id), animate=True, direction=1)

    def open_student_profile(self, student: Student):
        """供外部（测试/其它入口）直接打开学生资料页"""
        if student is None or student.id is None:
            return
        self.navigate('student', str(student.id), animate=True, direction=1)

    def on_student_double_clicked(self, data):
        dialog = StudentDetailDialog(data, self.db, self)
        dialog.exec()

    def on_student_edit(self, data):
        """学生磁贴的编辑入口（悬停 ✏️ 或右键）"""
        dialog = StudentEditDialog(self.db, student=data, parent=self)
        if dialog.exec() == QDialog.Accepted:
            self.refresh()

    def rename_class(self, class_name: str):
        """班级磁贴编辑：重命名班级（批量更新该班学生）"""
        new_name, ok = QInputDialog.getText(
            self, "重命名班级", f"将「{class_name}」重命名为：",
            QLineEdit.Normal, "" if class_name == "未分班" else class_name)
        if not ok:
            return
        new_name = new_name.strip()
        if not new_name or new_name == class_name:
            return
        groups = self._classes()
        for s in groups.get(class_name, []):
            s.class_name = new_name
            self.db.update_student(s)
        self.refresh()
        QMessageBox.information(self, "已更新", f"班级已重命名为「{new_name}」")

    def import_students(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择学生信息文件", "", SUPPORTED_FILE_TYPES
        )
        if not file_path:
            return

        service = DataImportService(self.db)
        success, errors = service.import_students(file_path)

        msg = f"导入完成：成功 {success} 条"
        if errors:
            msg += f"\n失败 {len(errors)} 条：\n" + '\n'.join(errors[:10])
            if len(errors) > 10:
                msg += f"\n... 还有 {len(errors) - 10} 条错误"
        QMessageBox.information(self, "导入结果", msg)
        self.refresh()

    def add_student(self):
        dialog = StudentEditDialog(self.db, parent=self)
        if dialog.exec() == QDialog.Accepted:
            self.refresh()

    def delete_students(self):
        """打开批量/单个删除学生对话框"""
        students = self.db.get_all_students()
        if not students:
            QMessageBox.information(self, "提示", "暂无学生数据可删除")
            return
        records = []
        for s in students:
            tags_txt = '、'.join(s.tags[:3]) if s.tags else '无'
            records.append({
                'id': s.id,
                'text': f"{s.name}（学号 {s.student_no}）",
                'detail': f"班级：{s.class_name or '未分配'}　标签：{tags_txt}",
            })
        dialog = MultiDeleteDialog(records, title="删除学生",
                                   instruction="勾选要删除的学生（可多选；勾 1 个即单个删除）。删除学生将同时删除其名下的全部成绩记录，且不可恢复：",
                                   parent=self)
        if dialog.exec() == QDialog.Accepted:
            ids = dialog.selected_ids()
            if ids:
                self._confirm_delete_students(ids)

    def _confirm_delete_students(self, ids: List[int], confirm: bool = True) -> int:
        """执行删除学生（含级联删除其成绩），返回删除的学生数；confirm=False 用于测试/无弹窗调用"""
        if not ids:
            return 0
        ids = [i for i in ids if i]
        grade_count = 0
        for sid in ids:
            grade_count += len(self.db.get_grades_by_student(sid))
        if confirm:
            msg = (f"确定删除选中的 {len(ids)} 名学生吗？\n"
                   f"将同时删除其名下的 {grade_count} 条成绩记录。\n"
                   f"该操作不可恢复！")
            reply = QMessageBox.question(self, "确认删除", msg,
                                         QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return 0
        removed = 0
        removed_grades = 0
        for sid in ids:
            removed_grades += self.db.delete_grades_by_student(sid)
            if self.db.delete_student(sid):
                removed += 1
        self.refresh()
        if confirm:
            QMessageBox.information(
                self, "删除完成",
                f"已删除学生 {removed} 名，并清理其成绩 {removed_grades} 条。")
        return removed


class StudentEditDialog(QDialog):
    """学生编辑对话框"""

    def __init__(self, db: DatabaseManager, student: Optional[Student] = None, parent=None):
        super().__init__(parent)
        self.db = db
        self.student = student
        self.setWindowTitle("新建学生" if student is None else "编辑学生")
        self.setMinimumSize(400, 350)
        self.setup_ui()
        if student:
            self.load_data()

    def setup_ui(self):
        layout = QFormLayout(self)
        layout.setSpacing(10)

        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("姓名")
        layout.addRow("姓名:", self.name_edit)

        self.no_edit = QLineEdit()
        self.no_edit.setPlaceholderText("学号")
        layout.addRow("学号:", self.no_edit)

        self.class_edit = QLineEdit()
        self.class_edit.setPlaceholderText("班级")
        layout.addRow("班级:", self.class_edit)

        self.tags_edit = QLineEdit()
        self.tags_edit.setPlaceholderText("标签（逗号分隔）")
        layout.addRow("标签:", self.tags_edit)

        self.titles_edit = QLineEdit()
        self.titles_edit.setPlaceholderText("头衔（逗号分隔）")
        layout.addRow("头衔:", self.titles_edit)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("保存")
        btn_box.button(QDialogButtonBox.Cancel).setText("取消")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addRow(btn_box)

    def load_data(self):
        self.name_edit.setText(self.student.name)
        self.no_edit.setText(self.student.student_no)
        self.class_edit.setText(self.student.class_name)
        self.tags_edit.setText(','.join(self.student.tags))
        self.titles_edit.setText(','.join(self.student.titles))

    def accept(self):
        name = self.name_edit.text().strip()
        if not name:
            QMessageBox.warning(self, "提示", "姓名不能为空")
            return
        student_no = self.no_edit.text().strip()
        if not student_no:
            student_no = name

        tags = [t.strip() for t in self.tags_edit.text().replace('，', ',').split(',') if t.strip()]
        titles = [t.strip() for t in self.titles_edit.text().replace('，', ',').split(',') if t.strip()]

        if self.student:
            self.student.name = name
            self.student.student_no = student_no
            self.student.class_name = self.class_edit.text().strip()
            self.student.tags = tags
            self.student.titles = titles
            self.db.update_student(self.student)
        else:
            student = Student(
                student_no=student_no,
                name=name,
                class_name=self.class_edit.text().strip(),
                tags=tags,
                titles=titles,
                avatar_color=random.choice(COURSE_COLORS),
            )
            self.db.add_student(student)

        super().accept()


class MonthCalendar(QWidget):
    """月历控件：标记有课的日期，点击选择日期

    - 有课的日期显示课程数量徽标与彩色圆点
    - 今天有特殊描边，选中日高亮
    - 左右箭头切换月份由外层 CourseModule 提供
    """
    date_selected = Signal(object)      # 参数为 datetime.date

    WEEKDAYS = ["一", "二", "三", "四", "五", "六", "日"]

    def __init__(self, parent=None):
        super().__init__(parent)
        self.year = date.today().year
        self.month = date.today().month
        self.selected = date.today()
        self.day_counts: Dict[str, int] = {}
        self._cells: Dict[str, 'DayCell'] = {}
        self.setup_ui()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(12, 12, 12, 12)
        layout.setSpacing(6)

        # 星期表头
        header = QHBoxLayout()
        header.setSpacing(4)
        for name in self.WEEKDAYS:
            lab = QLabel(name)
            lab.setAlignment(Qt.AlignCenter)
            color = COLORS['danger'] if name in ("六", "日") else COLORS['text_secondary']
            lab.setStyleSheet(
                f"font-size: 11px; font-weight: bold; color: {color}; "
                f"font-family: {FONT_STACK_CSS};")
            header.addWidget(lab, 1)
        layout.addLayout(header)

        # 日期网格
        self.grid = QGridLayout()
        self.grid.setSpacing(4)
        layout.addLayout(self.grid)
        layout.addStretch()

        self.setStyleSheet(f"""
            MonthCalendar {{
                background-color: {COLORS['card_bg']};
                border: 1px solid {COLORS['border']};
                border-radius: 14px;
            }}
        """)

    # ---- 数据 ----
    def set_month(self, year: int, month: int):
        self.year, self.month = year, month
        self.rebuild()

    def set_day_counts(self, counts: Dict[str, int]):
        self.day_counts = counts or {}
        self.rebuild()

    def set_selected(self, d: date):
        self.selected = d
        if (d.year, d.month) != (self.year, self.month):
            self.year, self.month = d.year, d.month
        self.rebuild()

    # ---- 构建 ----
    def rebuild(self):
        # 注意：清理时**不能** setParent(None)，否则控件会退化成顶层窗口，
        # 在 Windows 上会闪出一堆“程序窗口”。直接 deleteLater 即可（父控件仍是月历）。
        while self.grid.count():
            item = self.grid.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()
        self._cells.clear()

        first = date(self.year, self.month, 1)
        days_in_month = calendar.monthrange(self.year, self.month)[1]
        # 周一为一周起点
        leading = first.weekday()
        today = date.today()

        for i in range(leading):
            self.grid.addWidget(self._blank_cell(), 0, i)

        row, col = 0, leading
        for day in range(1, days_in_month + 1):
            d = date(self.year, self.month, day)
            key = d.isoformat()
            cell = DayCell(day, self.day_counts.get(key, 0),
                           is_today=(d == today),
                           is_selected=(d == self.selected),
                           parent=self)                 # 必须给父控件
            cell.clicked.connect(lambda dd=d: self._on_cell_clicked(dd))
            self.grid.addWidget(cell, row, col)
            self._cells[key] = cell
            col += 1
            if col > 6:
                col = 0
                row += 1

        # 补足末行空格
        if col != 0:
            for c in range(col, 7):
                self.grid.addWidget(self._blank_cell(), row, c)

    def _blank_cell(self) -> QWidget:
        w = QWidget(self)                              # 必须有父控件（否则成为顶层窗口）
        w.setMinimumSize(58, 52)
        w.setStyleSheet("background: transparent; border: none;")
        return w

    def _on_cell_clicked(self, d: date):
        self.selected = d
        self.rebuild()
        self.date_selected.emit(d)


class DayCell(QFrame):
    """月历中的单个日期格子（含课程数量徽标）"""
    clicked = Signal()

    def __init__(self, day: int, course_count: int = 0, is_today: bool = False,
                 is_selected: bool = False, parent=None):
        super().__init__(parent)
        self.day = day
        self.course_count = course_count
        self.is_today = is_today
        self.is_selected = is_selected
        self.setMinimumSize(58, 52)
        self.setCursor(Qt.PointingHandCursor)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(4, 4, 4, 4)
        layout.setSpacing(0)

        num = QLabel(str(day))
        num.setAlignment(Qt.AlignCenter)
        num.setStyleSheet(
            f"font-size: 13px; font-weight: bold; border: none; background: transparent; "
            f"color: {COLORS['text_primary']}; font-family: {FONT_STACK_CSS};")
        layout.addWidget(num)

        if course_count > 0:
            badge = QLabel(f"{course_count} 节课")
            badge.setAlignment(Qt.AlignCenter)
            badge.setStyleSheet(
                f"font-size: 9px; color: white; background-color: {COLORS['primary']};"
                f"border-radius: 6px; padding: 1px 4px; border: none;")
            layout.addWidget(badge)
        else:
            layout.addStretch()

        self._apply_style()

    def _apply_style(self):
        if self.is_selected:
            bg, border, width = COLORS['primary_light'], COLORS['primary'], 2
        elif self.course_count > 0:
            bg, border, width = "#FFFFFF", COLORS['info'], 1
        else:
            bg, border, width = "transparent", COLORS['border'], 1
        today_extra = f"border: 2px dashed {COLORS['danger']};" if (
            self.is_today and not self.is_selected) else f"border: {width}px solid {border};"
        self.setStyleSheet(f"""
            DayCell {{
                background-color: {bg};
                {today_extra}
                border-radius: 8px;
            }}
            DayCell:hover {{ border: 2px solid {COLORS['primary']}; }}
        """)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mousePressEvent(event)


class CourseRow(QFrame):
    """当天课程条目：时间 + 名称 + 教师/地点，悬停出现编辑/删除"""
    edit_requested = Signal(object)
    delete_requested = Signal(object)
    double_clicked = Signal(object)

    def __init__(self, course: Course, parent=None):
        super().__init__(parent)
        self.course = course
        self.setObjectName("courseRow")
        self.setMinimumHeight(74)
        self.setCursor(Qt.PointingHandCursor)
        self.setup_ui()

    def setup_ui(self):
        outer = QHBoxLayout(self)
        outer.setContentsMargins(0, 0, 10, 0)
        outer.setSpacing(10)

        # 左侧色条
        bar = QFrame()
        bar.setFixedWidth(6)
        bar.setStyleSheet(
            f"background-color: {self.course.color or COLORS['primary']};"
            f"border-top-left-radius: 6px; border-bottom-left-radius: 6px;")
        outer.addWidget(bar)

        mid = QVBoxLayout()
        mid.setContentsMargins(0, 8, 0, 8)
        mid.setSpacing(2)

        self.time_label = QLabel(
            f"🕐 {self.course.start_time} - {self.course.end_time}"
            f"　（{self.course.duration_minutes} 分钟）")
        self.time_label.setStyleSheet(
            f"font-size: 12px; font-weight: bold; color: {COLORS['primary']}; "
            f"border: none; background: transparent; font-family: {FONT_STACK_CSS};")
        mid.addWidget(self.time_label)

        name_label = QLabel(self.course.name or "（未命名课程）")
        name_label.setWordWrap(True)
        name_label.setStyleSheet(
            f"font-size: 15px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"border: none; background: transparent; font-family: {FONT_STACK_CSS};")
        mid.addWidget(name_label)

        meta = []
        if self.course.teacher:
            meta.append(f"👤 {self.course.teacher}")
        if self.course.location:
            meta.append(f"📍 {self.course.location}")
        if self.course.description:
            meta.append(f"📝 {self.course.description}")
        meta_label = QLabel("　".join(meta) if meta else "（无教师/地点信息）")
        meta_label.setWordWrap(True)
        meta_label.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"border: none; background: transparent; font-family: {FONT_STACK_CSS};")
        mid.addWidget(meta_label)

        outer.addLayout(mid, stretch=1)

        # 操作按钮（常显，便于发现“可编辑”）
        btns = QVBoxLayout()
        btns.setContentsMargins(0, 8, 0, 8)
        btns.setSpacing(4)
        edit_btn = QPushButton("✏️ 编辑")
        edit_btn.setCursor(Qt.PointingHandCursor)
        edit_btn.setStyleSheet(self._btn_style(COLORS['primary']))
        edit_btn.clicked.connect(lambda: self.edit_requested.emit(self.course))
        btns.addWidget(edit_btn)

        del_btn = QPushButton("🗑 删除")
        del_btn.setCursor(Qt.PointingHandCursor)
        del_btn.setStyleSheet(self._btn_style(COLORS['danger']))
        del_btn.clicked.connect(lambda: self.delete_requested.emit(self.course))
        btns.addWidget(del_btn)
        outer.addLayout(btns)

        self.setStyleSheet(f"""
            QFrame#courseRow {{
                background-color: {COLORS['card_bg']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
            }}
            QFrame#courseRow:hover {{
                border: 1px solid {COLORS['primary']};
                background-color: {COLORS['hover']};
            }}
        """)

    @staticmethod
    def _btn_style(color: str) -> str:
        return f"""
            QPushButton {{
                background-color: {color}; color: white; border: none;
                border-radius: 10px; padding: 4px 10px; font-size: 11px;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{ opacity: 0.85; }}
        """

    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.LeftButton:
            self.double_clicked.emit(self.course)
        super().mouseDoubleClickEvent(event)

    def contextMenuEvent(self, event):
        menu = QMenu(self)
        act_edit = menu.addAction("✏️ 编辑课程")
        act_edit.triggered.connect(lambda: self.edit_requested.emit(self.course))
        act_del = menu.addAction("🗑 删除课程")
        act_del.triggered.connect(lambda: self.delete_requested.emit(self.course))
        menu.exec(event.globalPos())
        event.accept()


class StudentProfileModule(BaseModulePage):
    """学情管理模块（v1.4.0 新增）

    - **自动读取**已有学员名单，每个学生一行
    - 以表格形式**自由补充**该生的任何相关信息：可自行添加/重命名/删除字段（列），
      单元格随改随存
    - 这里的学情信息**只在本模块使用**，不会出现在「学生管理」的磁贴或资料气泡中
    - 支持导出 Excel 便于归档
    """

    BASE_HEADERS = ['姓名', '学号', '班级']

    def __init__(self, db: DatabaseManager, parent=None):
        super().__init__(db, parent)
        self.fields: List[Dict[str, Any]] = []
        self.students: List[Student] = []
        self._loading = False
        self._save_count = 0

    # ------------------------------------------------------------------ UI
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        # 顶部工具栏
        toolbar = QHBoxLayout()
        title = QLabel("📋 学情管理")
        title.setStyleSheet(
            f"font-size: 22px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(title)

        sub = QLabel("自动读取学员名单；可自由添加字段并填写任意补充信息（不显示在学生磁贴）")
        sub.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; padding-left: 10px; "
            f"font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(sub)
        toolbar.addStretch()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("按姓名/学号/班级筛选…")
        self.search_edit.setFixedWidth(200)
        self.search_edit.textChanged.connect(self.on_search)
        self.search_edit.setStyleSheet(f"""
            QLineEdit {{
                padding: 8px 12px; border: 1px solid {COLORS['border']};
                border-radius: 18px; font-size: 13px;
                font-family: {FONT_STACK_CSS}; background-color: {COLORS['card_bg']};
            }}
            QLineEdit:focus {{ border-color: {COLORS['primary']}; }}
        """)
        toolbar.addWidget(self.search_edit)
        layout.addLayout(toolbar)

        # 字段操作栏
        field_bar = QHBoxLayout()
        field_bar.addWidget(QLabel("字段："))
        self.field_combo = QComboBox()
        self.field_combo.setMinimumWidth(160)
        field_bar.addWidget(self.field_combo)

        for text, color, slot in (
                ("➕ 添加字段", COLORS['success'], self.add_field),
                ("✏️ 重命名字段", COLORS['primary'], self.rename_field),
                ("🗑 删除字段", COLORS['danger'], self.delete_field),
                ("🔄 刷新名单", COLORS['secondary'], self.refresh),
                ("📤 导出 Excel", COLORS['info'], self.export_excel),
        ):
            btn = QPushButton(text)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(self._get_btn_style(color))
            btn.clicked.connect(slot)
            field_bar.addWidget(btn)
        field_bar.addStretch()
        layout.addLayout(field_bar)

        # 表格
        self.table = QTableWidget(0, len(self.BASE_HEADERS))
        self.table.setHorizontalHeaderLabels(self.BASE_HEADERS)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Interactive)
        self.table.horizontalHeader().setStretchLastSection(True)
        self.table.setAlternatingRowColors(True)
        self.table.setEditTriggers(QAbstractItemView.DoubleClicked |
                                   QAbstractItemView.EditKeyPressed |
                                   QAbstractItemView.AnyKeyPressed)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {COLORS['border']}; border-radius: 10px;
                font-size: 12px; font-family: {FONT_STACK_CSS};
                gridline-color: {COLORS['border']};
            }}
            QHeaderView::section {{
                background-color: {COLORS['primary_light']}; font-weight: bold;
                padding: 6px; border: none;
            }}
        """)
        self.table.cellChanged.connect(self.on_cell_changed)
        layout.addWidget(self.table, stretch=1)

        # 底部状态
        self.status_label = QLabel()
        self.status_label.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(self.status_label)

        hint = QLabel("💡 提示：绿色单元格可直接编辑并自动保存；这些学情信息仅保存在本模块，"
                      "不会显示在「学生管理」的学生磁贴中")
        hint.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(hint)

    def _get_btn_style(self, color):
        return f"""
            QPushButton {{
                background-color: {color}; color: white; border: none;
                border-radius: 14px; padding: 7px 13px; font-size: 12px;
                font-family: {FONT_STACK_CSS}; font-weight: 500;
            }}
            QPushButton:hover {{ opacity: 0.85; }}
        """

    # ------------------------------------------------------------- 数据刷新
    def refresh(self):
        """重新读取字段与学员名单（学员名单自动来自学生管理）"""
        self.fields = self.db.get_profile_fields()
        self.students = self.db.get_all_students()
        self._rebuild_table()
        self._rebuild_field_combo()
        self._update_status()

    def _rebuild_field_combo(self):
        current = self.field_combo.currentData()
        self.field_combo.blockSignals(True)
        self.field_combo.clear()
        for f in self.fields:
            self.field_combo.addItem(f["name"], f["id"])
        if current is not None:
            idx = self.field_combo.findData(current)
            if idx >= 0:
                self.field_combo.setCurrentIndex(idx)
        self.field_combo.blockSignals(False)

    def _rebuild_table(self):
        headers = self.BASE_HEADERS + [f["name"] for f in self.fields]
        values = self.db.get_profile_values()          # {(student_id, field_id): value}

        self._loading = True
        self.table.clear()
        self.table.setColumnCount(len(headers))
        self.table.setHorizontalHeaderLabels(headers)
        self.table.setRowCount(len(self.students))
        for row, stu in enumerate(self.students):
            for col, text in enumerate((stu.name, stu.student_no, stu.class_name)):
                item = QTableWidgetItem(text or '')
                item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
                item.setForeground(QColor(COLORS['text_secondary']))
                item.setData(Qt.UserRole, stu.id)
                self.table.setItem(row, col, item)
            for i, field in enumerate(self.fields):
                col = len(self.BASE_HEADERS) + i
                value = values.get((stu.id, field['id']), '')
                cell = QTableWidgetItem(value)
                cell.setData(Qt.UserRole, stu.id)
                cell.setData(Qt.UserRole + 1, field['id'])
                cell.setBackground(QColor('#F6FFF9'))
                self.table.setItem(row, col, cell)
        self.table.resizeColumnsToContents()
        self._loading = False

    def _update_status(self):
        filled = sum(1 for v in self.db.get_profile_values().values() if str(v).strip())
        self.status_label.setText(
            f"共 {len(self.students)} 名学生 · {len(self.fields)} 个学情字段 · "
            f"已填写 {filled} 项　（自动保存 {self._save_count} 次）")

    # ------------------------------------------------------------- 交互
    def on_search(self, keyword: str):
        kw = keyword.strip().lower()
        for row in range(self.table.rowCount()):
            text = " ".join(
                self.table.item(row, c).text() if self.table.item(row, c) else ''
                for c in range(self.table.columnCount()))
            self.table.setRowHidden(row, bool(kw) and kw not in text.lower())

    def on_cell_changed(self, row: int, col: int):
        """单元格编辑 → 立即写入数据库"""
        if self._loading or col < len(self.BASE_HEADERS):
            return
        cell = self.table.item(row, col)
        if cell is None:
            return
        student_id = cell.data(Qt.UserRole)
        field_id = cell.data(Qt.UserRole + 1)
        if student_id is None or field_id is None:
            return
        self.db.set_profile_value(int(student_id), int(field_id), cell.text())
        self._save_count += 1
        self._update_status()

    def add_field(self):
        name, ok = QInputDialog.getText(
            self, "添加学情字段", "新字段名称（例如：家庭情况 / 薄弱科目 / 家长电话）：",
            QLineEdit.Normal, "")
        if not ok or not name.strip():
            return
        name = name.strip()
        existed = any(f['name'] == name for f in self.fields)
        field_id = self.db.add_profile_field(name)
        self.refresh()
        if field_id is None:
            QMessageBox.warning(self, "提示", "字段名无效")
        elif existed:
            QMessageBox.information(self, "提示", f"字段「{name}」已存在，已为你选中该列")
    def rename_field(self):
        field_id = self.field_combo.currentData()
        if field_id is None:
            QMessageBox.information(self, "提示", "请先添加并选择一个字段")
            return
        old_name = self.field_combo.currentText()
        name, ok = QInputDialog.getText(
            self, "重命名字段", f"把「{old_name}」重命名为：", QLineEdit.Normal, old_name)
        if not ok or not name.strip() or name.strip() == old_name:
            return
        if self.db.rename_profile_field(int(field_id), name.strip()):
            self.refresh()
        else:
            QMessageBox.warning(self, "提示", f"字段「{name.strip()}」已存在")

    def delete_field(self):
        field_id = self.field_combo.currentData()
        if field_id is None:
            QMessageBox.information(self, "提示", "请先添加并选择一个字段")
            return
        name = self.field_combo.currentText()
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定删除字段「{name}」吗？\n该列所有学生已填写的内容都会一并删除，且不可恢复！",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.db.delete_profile_field(int(field_id))
            self.refresh()

    def export_excel(self):
        """导出整张学情表为 xlsx"""
        matrix = self.db.get_profile_matrix()
        if not matrix['students']:
            QMessageBox.information(self, "提示", "暂无学生数据可导出")
            return
        path, _ = QFileDialog.getSaveFileName(
            self, "导出学情表", "学情表.xlsx", "Excel 文件 (*.xlsx)")
        if not path:
            return
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "学情表"
            headers = self.BASE_HEADERS + [f['name'] for f in matrix['fields']]
            ws.append(headers)
            for row in matrix['students']:
                ws.append([row['name'], row['student_no'], row['class_name']] +
                          [row['values'].get(f['id'], '') for f in matrix['fields']])
            wb.save(path)
            QMessageBox.information(self, "导出成功", f"已导出到：\n{path}")
        except Exception as e:
            QMessageBox.warning(self, "导出失败", str(e))


class CourseModule(BaseModulePage):
    """课程/排课管理系统 —— 月历视图

    - 左侧月历：**有课的日期会被标记**（课程数量徽标 + 高亮），今天虚线框
    - 右侧：点选日期后**展开当天课程的详细时间安排**（时间轴列表，可编辑/删除）
    - 日期切换带淡入淡出过渡
    """

    def __init__(self, db: DatabaseManager, parent=None):
        super().__init__(db, parent)
        today = date.today()
        self.view_year = today.year
        self.view_month = today.month
        self.selected_date = today
        self.current_date = today          # 兼容旧接口
        self.courses: List[Course] = []
        self.course_items: List[Any] = []  # 兼容旧接口（旧的图形条目列表）

    # ------------------------------------------------------------------ UI
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        toolbar = QHBoxLayout()

        title = QLabel("📅 课程/排课管理 · 月历")
        title.setStyleSheet(
            f"font-size: 22px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(title)
        toolbar.addStretch()

        self.prev_btn = QPushButton("◀")
        self.prev_btn.setFixedSize(36, 36)
        self.prev_btn.setToolTip("上一个月")
        self.prev_btn.clicked.connect(self.prev_month)
        self.prev_btn.setStyleSheet(self._get_nav_btn_style())
        toolbar.addWidget(self.prev_btn)

        self.date_label = QLabel()
        self.date_label.setAlignment(Qt.AlignCenter)
        self.date_label.setMinimumWidth(150)
        self.date_label.setStyleSheet(
            f"font-size: 16px; font-weight: bold; color: {COLORS['primary']}; "
            f"padding: 0 10px; font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(self.date_label)

        self.next_btn = QPushButton("▶")
        self.next_btn.setFixedSize(36, 36)
        self.next_btn.setToolTip("下一个月")
        self.next_btn.clicked.connect(self.next_month)
        self.next_btn.setStyleSheet(self._get_nav_btn_style())
        toolbar.addWidget(self.next_btn)

        today_btn = QPushButton("今天")
        today_btn.setStyleSheet(self._get_nav_btn_style())
        today_btn.clicked.connect(self.goto_today)
        toolbar.addWidget(today_btn)

        toolbar.addStretch()

        add_course_btn = QPushButton("➕ 新建课程")
        add_course_btn.setStyleSheet(self._get_add_btn_style())
        add_course_btn.clicked.connect(self.add_course)
        toolbar.addWidget(add_course_btn)

        delete_btn = QPushButton("🗑 删除课程")
        delete_btn.setStyleSheet(self._get_del_btn_style())
        delete_btn.clicked.connect(self.delete_courses)
        toolbar.addWidget(delete_btn)

        layout.addLayout(toolbar)

        # 左：月历　右：当天课表
        splitter = QSplitter(Qt.Horizontal)

        self.calendar = MonthCalendar()
        self.calendar.date_selected.connect(self.on_date_selected)
        splitter.addWidget(self.calendar)

        right = QWidget()
        right_layout = QVBoxLayout(right)
        right_layout.setContentsMargins(10, 0, 0, 0)
        right_layout.setSpacing(8)

        self.day_title = QLabel()
        self.day_title.setStyleSheet(
            f"font-size: 16px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS};")
        right_layout.addWidget(self.day_title)

        self.day_hint = QLabel("双击课程可编辑；点「新建课程」为当天添加课程")
        self.day_hint.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        right_layout.addWidget(self.day_hint)

        self.day_scroll = QScrollArea()
        self.day_scroll.setWidgetResizable(True)
        self.day_scroll.setFrameShape(QFrame.NoFrame)
        self.day_scroll.setStyleSheet("QScrollArea { background: transparent; }")
        self.day_container = QWidget()
        self.day_container.setStyleSheet("background: transparent;")
        self.day_layout = QVBoxLayout(self.day_container)
        self.day_layout.setContentsMargins(2, 2, 2, 2)
        self.day_layout.setSpacing(8)
        self.day_scroll.setWidget(self.day_container)
        right_layout.addWidget(self.day_scroll, stretch=1)

        splitter.addWidget(right)
        splitter.setSizes([430, 520])
        layout.addWidget(splitter, stretch=1)

        hint_label = QLabel("💡 月历中有课的日期会显示「N 节课」徽标，点选日期即可查看当天课程时间安排")
        hint_label.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(hint_label)

    def _get_nav_btn_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['card_bg']};
                border: 1px solid {COLORS['border']};
                border-radius: 18px;
                font-size: 13px;
                color: {COLORS['text_primary']};
                font-weight: bold;
                padding: 6px 12px;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{
                border-color: {COLORS['primary']};
                background-color: {COLORS['hover']};
            }}
        """

    def _get_add_btn_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['primary']};
                color: white;
                border: none;
                border-radius: 18px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: {COLORS['primary_dark']};
            }}
        """

    def _get_del_btn_style(self):
        return f"""
            QPushButton {{
                background-color: {COLORS['danger']};
                color: white;
                border: none;
                border-radius: 18px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
                font-weight: 500;
            }}
            QPushButton:hover {{
                background-color: #E25555;
            }}
        """

    # ------------------------------------------------------- 月份/日期导航
    def prev_month(self):
        y, m = self.view_year, self.view_month - 1
        if m < 1:
            y, m = y - 1, 12
        self.view_year, self.view_month = y, m
        self.refresh()

    def next_month(self):
        y, m = self.view_year, self.view_month + 1
        if m > 12:
            y, m = y + 1, 1
        self.view_year, self.view_month = y, m
        self.refresh()

    def prev_day(self):
        """兼容旧接口：选中前一天"""
        self.on_date_selected(self.selected_date - timedelta(days=1))

    def next_day(self):
        """兼容旧接口：选中后一天"""
        self.on_date_selected(self.selected_date + timedelta(days=1))

    def goto_today(self):
        today = date.today()
        self.view_year, self.view_month = today.year, today.month
        self.on_date_selected(today)

    def update_date_label(self):
        self.date_label.setText(f"{self.view_year} 年 {self.view_month} 月")
        if (self.view_year, self.view_month) == (date.today().year, date.today().month):
            self.date_label.setText(f"📌 {self.view_year} 年 {self.view_month} 月")

    def on_date_selected(self, d: date):
        """点选日期：切换月份并展开当天课程（带淡入过渡）"""
        self.selected_date = d
        self.current_date = d
        if (d.year, d.month) != (self.view_year, self.view_month):
            self.view_year, self.view_month = d.year, d.month
        self.calendar.set_selected(d)
        self.update_date_label()

        def _update():
            self._rebuild_day_panel()

        cross_fade(self.day_container, _update, duration=170)

    # ------------------------------------------------------------- 刷新
    def refresh(self):
        self.update_date_label()
        counts = self.db.get_course_day_counts(self.view_year, self.view_month)
        self.calendar.set_month(self.view_year, self.view_month)
        self.calendar.set_day_counts(counts)
        self.calendar.set_selected(self.selected_date)
        self._rebuild_day_panel()

    def _rebuild_day_panel(self):
        """重建右侧「当天课程时间安排」列表"""
        # 只 deleteLater，不 setParent(None)：避免控件变成顶层窗口（会闪出“程序窗口”）
        while self.day_layout.count():
            item = self.day_layout.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()

        date_str = self.selected_date.isoformat()
        self.courses = self.db.get_courses_by_date(date_str)
        self.course_items = list(self.courses)

        weekday_cn = "一二三四五六日"[self.selected_date.weekday()]
        prefix = "📌 今天 · " if self.selected_date == date.today() else ""
        self.day_title.setText(
            f"{prefix}{self.selected_date.strftime('%Y年%m月%d日')}（周{weekday_cn}）"
            f" · {len(self.courses)} 门课程")

        if not self.courses:
            empty = QLabel("当天暂无课程安排\n\n点右上角「➕ 新建课程」为这一天添加课程",
                           parent=self.day_container)
            empty.setAlignment(Qt.AlignCenter)
            empty.setWordWrap(True)
            empty.setStyleSheet(
                f"font-size: 13px; color: {COLORS['text_secondary']}; padding: 40px; "
                f"font-family: {FONT_STACK_CSS};")
            self.day_layout.addWidget(empty)
            self.day_layout.addStretch()
            return

        for course in sorted(self.courses, key=lambda c: (c.start_time or "")):
            row = CourseRow(course, parent=self.day_container)   # 必须给父控件
            row.edit_requested.connect(self.edit_course)
            row.delete_requested.connect(self.delete_course)
            row.double_clicked.connect(self.edit_course)
            self.day_layout.addWidget(row)
        self.day_layout.addStretch()

    # ------------------------------------------------------------- 课程操作
    def add_course(self):
        dialog = CourseEditDialog(parent=self)
        if dialog.exec() == QDialog.Accepted:
            course = dialog.get_course_data()
            course.course_date = self.selected_date.isoformat()
            self.db.add_course(course)
            self.refresh()
            self._highlight_day(course.course_date)

    def edit_course(self, course: Course):
        dialog = CourseEditDialog(course, parent=self)
        if dialog.exec() == QDialog.Accepted:
            updated = dialog.get_course_data()
            self.db.update_course(updated)
            self.refresh()

    def delete_course(self, course: Course):
        reply = QMessageBox.question(
            self, "确认删除", f"确定要删除课程「{course.name}」吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes and course.id:
            self.db.delete_course(course.id)
            self.refresh()

    def delete_courses(self):
        """批量/单个删除课程（列出全部课程）"""
        courses = self.db.get_all_courses()
        if not courses:
            QMessageBox.information(self, "提示", "暂无课程数据可删除")
            return
        records = []
        for c in courses:
            records.append({
                'id': c.id,
                'text': f"{c.name}　{c.course_date} {c.start_time}-{c.end_time}",
                'detail': f"教师：{c.teacher or '未指定'}　地点：{c.location or '未指定'}　{c.description or ''}",
            })
        dialog = MultiDeleteDialog(records, title="删除课程",
                                   instruction="勾选要删除的课程（可多选；勾 1 个即单个删除）。该操作不可恢复：",
                                   parent=self)
        if dialog.exec() == QDialog.Accepted:
            ids = dialog.selected_ids()
            if ids:
                self._confirm_delete_courses(ids)

    def _confirm_delete_courses(self, ids: List[int], confirm: bool = True) -> int:
        """执行删除课程，返回删除数量；confirm=False 用于测试/无弹窗调用"""
        if not ids:
            return 0
        ids = [i for i in ids if i]
        if confirm:
            reply = QMessageBox.question(
                self, "确认删除", f"确定删除选中的 {len(ids)} 门课程吗？\n该操作不可恢复！",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return 0
        removed = 0
        for cid in ids:
            if self.db.delete_course(cid):
                removed += 1
        self.refresh()
        if confirm:
            QMessageBox.information(self, "删除完成", f"已删除课程 {removed} 门。")
        return removed

    def _highlight_day(self, date_str: str):
        """新建课程后把视图定位到该日期"""
        try:
            d = date.fromisoformat(date_str)
        except ValueError:
            return
        self.on_date_selected(d)

class StickerEditDialog(QDialog):
    """便签编辑对话框：文字 / 表格 / 图片（支持剪贴板粘贴，像贴便签一样）

    - 「📋 智能粘贴」：自动识别剪贴板是图片 / 表格(TSV) / 文字，切到对应页并填入
    - 表格页：可直接编辑单元格，支持加行加列、从剪贴板粘贴整张表
    - 图片页：粘贴剪贴板图片、选择文件、拖拽图片文件、加说明文字
    """

    def __init__(self, db: DatabaseManager, feedback: Optional[Feedback] = None, parent=None):
        super().__init__(parent)
        self.db = db
        self.feedback = feedback
        self.setWindowTitle("新建便签" if feedback is None else "编辑便签")
        self.setMinimumSize(620, 520)
        self.setAcceptDrops(True)
        self._pending_image_name: str = ""      # 新粘贴/选择的图片（保存时生效）
        self._new_image_temp: Optional[Path] = None
        self._build_ui()
        if feedback is not None:
            self._load(feedback)

    # ------------------------------------------------------------------ UI
    def _build_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        # 顶部：颜色 + 智能粘贴
        top = QHBoxLayout()
        top.addWidget(QLabel("便签颜色："))
        self.color_buttons = []
        for color in STICKER_COLORS[:6]:
            btn = QPushButton()
            btn.setFixedSize(24, 24)
            btn.setCursor(Qt.PointingHandCursor)
            btn.setStyleSheet(
                f"background-color: {color}; border: 2px solid {COLORS['border']};"
                f"border-radius: 12px;")
            btn.clicked.connect(lambda _=False, c=color: self._set_color(c))
            self.color_buttons.append(btn)
            top.addWidget(btn)
        self.color_label = QLabel(STICKER_COLORS[0])
        self.color_label.setStyleSheet(f"color: {COLORS['text_secondary']}; font-size: 11px;")
        top.addWidget(self.color_label)
        top.addStretch()

        paste_btn = QPushButton("📋 智能粘贴（图片/表格/文字）")
        paste_btn.setStyleSheet(self._primary_btn_style())
        paste_btn.clicked.connect(lambda: self.paste_from_clipboard(smart=True))
        top.addWidget(paste_btn)
        layout.addLayout(top)

        self.selected_color = STICKER_COLORS[0]

        # 内容页
        self.tabs = QTabWidget()
        self.tabs.addTab(self._build_text_tab(), "📝 文字")
        self.tabs.addTab(self._build_table_tab(), "📊 表格")
        self.tabs.addTab(self._build_image_tab(), "🖼 图片")
        layout.addWidget(self.tabs, stretch=1)

        self.hint_label = QLabel("提示：可先用微信/QQ/截图工具复制内容，再点「智能粘贴」或按 Ctrl+V")
        self.hint_label.setStyleSheet(f"font-size: 11px; color: {COLORS['text_secondary']};")
        layout.addWidget(self.hint_label)

        # 底部按钮
        row = QHBoxLayout()
        row.addStretch()
        cancel = QPushButton("取消")
        cancel.setStyleSheet(self._plain_btn_style())
        cancel.clicked.connect(self.reject)
        row.addWidget(cancel)
        save = QPushButton("保存便签")
        save.setStyleSheet(self._primary_btn_style())
        save.clicked.connect(self.accept)
        row.addWidget(save)
        layout.addLayout(row)

        # Ctrl+V 智能粘贴
        shortcut = QShortcut(QKeySequence.Paste, self)
        shortcut.activated.connect(lambda: self.paste_from_clipboard(smart=True))

    def _build_text_tab(self) -> QWidget:
        page = QWidget()
        box = QVBoxLayout(page)
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("在这里写点什么…（可直接 Ctrl+V 粘贴文字）")
        self.text_edit.setStyleSheet(self._editor_style())
        box.addWidget(self.text_edit)
        return page

    def _build_table_tab(self) -> QWidget:
        page = QWidget()
        box = QVBoxLayout(page)
        bar = QHBoxLayout()
        for text, slot in (("➕ 加行", self._add_row), ("➕ 加列", self._add_column),
                           ("➖ 删行", self._remove_row), ("➖ 删列", self._remove_column),
                           ("📋 粘贴表格", lambda: self.paste_from_clipboard(smart=False)),
                           ("清空", self._clear_table)):
            btn = QPushButton(text)
            btn.setStyleSheet(self._mini_btn_style())
            btn.clicked.connect(slot)
            bar.addWidget(btn)
        bar.addStretch()
        box.addLayout(bar)

        self.table = QTableWidget(4, 3)
        self.table.setHorizontalHeaderLabels([f"列{i+1}" for i in range(3)])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {COLORS['border']}; border-radius: 8px; font-size: 12px;
                font-family: {FONT_STACK_CSS};
            }}
            QHeaderView::section {{
                background-color: {COLORS['primary_light']}; font-weight: bold;
                padding: 4px; border: none;
            }}
        """)
        box.addWidget(self.table)
        return page

    def _build_image_tab(self) -> QWidget:
        page = QWidget()
        box = QVBoxLayout(page)
        bar = QHBoxLayout()
        for text, slot in (("📋 粘贴剪贴板图片", lambda: self.paste_from_clipboard(smart=False)),
                           ("📂 选择图片文件", self._choose_image_file),
                           ("🗑 清除图片", self._clear_image)):
            btn = QPushButton(text)
            btn.setStyleSheet(self._mini_btn_style())
            btn.clicked.connect(slot)
            bar.addWidget(btn)
        bar.addStretch()
        box.addLayout(bar)

        self.image_preview = QLabel("把图片拖到这里，或点上面的按钮粘贴/选择图片")
        self.image_preview.setAlignment(Qt.AlignCenter)
        self.image_preview.setMinimumHeight(260)
        self.image_preview.setStyleSheet(f"""
            QLabel {{
                border: 2px dashed {COLORS['border']}; border-radius: 10px;
                color: {COLORS['text_secondary']}; font-size: 12px;
                background-color: {COLORS['card_bg']};
            }}
        """)
        box.addWidget(self.image_preview, stretch=1)

        cap_row = QHBoxLayout()
        cap_row.addWidget(QLabel("图片说明："))
        self.caption_edit = QLineEdit()
        self.caption_edit.setPlaceholderText("可选，例如：期中试卷第 3 题截图")
        self.caption_edit.setStyleSheet(self._editor_style())
        cap_row.addWidget(self.caption_edit, stretch=1)
        box.addLayout(cap_row)
        return page

    # --------------------------------------------------------------- 样式
    @staticmethod
    def _primary_btn_style() -> str:
        return f"""
            QPushButton {{
                background-color: {COLORS['primary']}; color: white; border: none;
                border-radius: 14px; padding: 7px 16px; font-size: 13px;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{ background-color: {COLORS['primary_dark']}; }}
        """

    @staticmethod
    def _plain_btn_style() -> str:
        return f"""
            QPushButton {{
                background-color: {COLORS['card_bg']}; color: {COLORS['text_primary']};
                border: 1px solid {COLORS['border']}; border-radius: 14px;
                padding: 7px 16px; font-size: 13px; font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{ border-color: {COLORS['primary']}; }}
        """

    @staticmethod
    def _mini_btn_style() -> str:
        return f"""
            QPushButton {{
                background-color: {COLORS['primary_light']}; color: {COLORS['primary']};
                border: none; border-radius: 10px; padding: 5px 10px; font-size: 11px;
                font-family: {FONT_STACK_CSS};
            }}
            QPushButton:hover {{ background-color: {COLORS['hover']}; }}
        """

    @staticmethod
    def _editor_style() -> str:
        return f"""
            padding: 8px; border: 1px solid {COLORS['border']}; border-radius: 8px;
            font-size: 13px; font-family: {FONT_STACK_CSS};
        """

    def _set_color(self, color: str):
        self.selected_color = color
        self.color_label.setText(color)

    # ------------------------------------------------------- 加载已有便签
    def _load(self, fb: Feedback):
        self.selected_color = fb.color or STICKER_COLORS[0]
        self.color_label.setText(self.selected_color)
        self.text_edit.setPlainText(fb.content or "")
        data = fb.table_data()
        if data['headers'] or data['rows']:
            self._set_table(data['headers'], data['rows'])
        if fb.content_type == 'image':
            self.tabs.setCurrentIndex(2)
            self.caption_edit.setText(fb.image_caption or "")
            if fb.image_path:
                self._pending_image_name = fb.image_path
                path = image_abs_path(fb.image_path)
                if path.exists():
                    self._show_preview(QPixmap(str(path)))
                else:
                    self.image_preview.setText("（原图片文件缺失）")
        elif fb.content_type == 'table':
            self.tabs.setCurrentIndex(1)
        else:
            self.tabs.setCurrentIndex(0)

    # ---------------------------------------------------------- 表格操作
    def _set_table(self, headers: List[Any], rows: List[List[Any]]):
        cols = max(len(headers), max((len(r) for r in rows), default=0), 1)
        self.table.setColumnCount(cols)
        labels = []
        for i in range(cols):
            if i < len(headers) and str(headers[i]).strip():
                labels.append(str(headers[i]))
            else:
                labels.append(f"列{i+1}")
        self.table.setHorizontalHeaderLabels(labels)
        self.table.setRowCount(max(len(rows), 1))
        for r, row in enumerate(rows):
            for c in range(cols):
                val = row[c] if c < len(row) else ''
                self.table.setItem(r, c, QTableWidgetItem('' if val is None else str(val)))

    def _table_data(self) -> Dict[str, Any]:
        headers = []
        for c in range(self.table.columnCount()):
            item = self.table.horizontalHeaderItem(c)
            headers.append(item.text() if item else f"列{c+1}")
        rows = []
        for r in range(self.table.rowCount()):
            row = []
            for c in range(self.table.columnCount()):
                item = self.table.item(r, c)
                row.append(item.text() if item else '')
            if any(str(v).strip() for v in row):
                rows.append(row)
        return {'headers': headers, 'rows': rows}

    def _add_row(self):
        self.table.insertRow(self.table.rowCount())

    def _add_column(self):
        c = self.table.columnCount()
        self.table.insertColumn(c)
        self.table.setHorizontalHeaderItem(c, QTableWidgetItem(f"列{c+1}"))

    def _remove_row(self):
        r = self.table.currentRow()
        if r < 0:
            r = self.table.rowCount() - 1
        if r >= 0 and self.table.rowCount() > 1:
            self.table.removeRow(r)

    def _remove_column(self):
        c = self.table.currentColumn()
        if c < 0:
            c = self.table.columnCount() - 1
        if c >= 0 and self.table.columnCount() > 1:
            self.table.removeColumn(c)

    def _clear_table(self):
        self.table.clearContents()
        self.table.setRowCount(3)
        self.table.setColumnCount(3)
        self.table.setHorizontalHeaderLabels(["列1", "列2", "列3"])

    # ---------------------------------------------------------- 图片操作
    def _show_preview(self, pixmap: QPixmap):
        if pixmap.isNull():
            return
        self.image_preview.setPixmap(
            pixmap.scaled(520, 300, Qt.KeepAspectRatio, Qt.SmoothTransformation))

    def _assign_clipboard_image(self, image) -> bool:
        """把 QImage 保存为临时文件并在预览中显示"""
        if image is None or image.isNull():
            return False
        try:
            name = save_qimage_file(image)          # 直接落到图片目录
            self._pending_image_name = name
            self._new_image_temp = image_abs_path(name)
            self._show_preview(QPixmap(str(self._new_image_temp)))
            self.tabs.setCurrentIndex(2)
            return True
        except Exception as e:
            QMessageBox.warning(self, "提示", f"图片保存失败：{e}")
            return False

    def _choose_image_file(self):
        path, _ = QFileDialog.getOpenFileName(self, "选择图片", "", IMAGE_FILE_TYPES)
        if not path:
            return
        try:
            name = save_image_file(path)
            self._pending_image_name = name
            self._new_image_temp = image_abs_path(name)
            self._show_preview(QPixmap(str(self._new_image_temp)))
            self.tabs.setCurrentIndex(2)
        except Exception as e:
            QMessageBox.warning(self, "提示", f"图片读取失败：{e}")

    def _clear_image(self):
        self._pending_image_name = ""
        self._new_image_temp = None
        self.image_preview.setPixmap(QPixmap())
        self.image_preview.setText("已清除图片（保存后将变为文字便签）")

    # ------------------------------------------------------- 剪贴板粘贴
    def paste_from_clipboard(self, smart: bool = True):
        """智能粘贴：自动识别剪贴板内容类型（图片 > 表格 > 文字）"""
        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData()

        # 1) 图片
        if mime is not None and mime.hasImage():
            image = clipboard.image()
            if self._assign_clipboard_image(image):
                self.hint_label.setText("✅ 已粘贴剪贴板图片")
                return

        # 2) 表格 / 文字
        text = clipboard.text() or ""
        if text.strip():
            tsv = self._parse_table_text(text)
            if smart and tsv and len(tsv) > 1:
                self._set_table(tsv[0], tsv[1:])
                self.tabs.setCurrentIndex(1)
                self.hint_label.setText(f"✅ 已识别为表格（{len(tsv)-1} 行）")
                return
            if tsv and not smart and self.tabs.currentIndex() == 1:
                self._set_table(tsv[0], tsv[1:] if len(tsv) > 1 else [])
                self.hint_label.setText("✅ 已粘贴为表格")
                return
            # 普通文字
            self.tabs.setCurrentIndex(0)
            if self.text_edit.toPlainText().strip():
                self.text_edit.append("")
            self.text_edit.append(text)
            self.hint_label.setText("✅ 已粘贴文字")
            return

        self.hint_label.setText("⚠️ 剪贴板里没有可粘贴的内容")

    @staticmethod
    def _parse_table_text(text: str) -> Optional[List[List[str]]]:
        """把制表符/多空格分隔的文本解析成表格（Excel/表格软件复制的内容）"""
        rows = []
        for raw in text.splitlines():
            if not raw.strip():
                continue
            if '\t' in raw:
                cells = [c.strip() for c in raw.split('\t')]
            elif '|' in raw and raw.count('|') >= 2:
                cells = [c.strip() for c in raw.strip('|').split('|')]
            else:
                cells = [raw.strip()]
            rows.append(cells)
        if len(rows) < 2:
            return None
        width = max(len(r) for r in rows)
        if width < 2:
            return None
        rows = [r + [''] * (width - len(r)) for r in rows]
        return rows

    # -------------------------------------------------------- 拖拽图片文件
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls() or event.mimeData().hasImage():
            event.acceptProposedAction()

    def dropEvent(self, event):
        mime = event.mimeData()
        if mime.hasImage():
            self._assign_clipboard_image(QImage(mime.imageData()))
            return
        for url in mime.urls():
            path = url.toLocalFile()
            if path and Path(path).suffix.lower() in (
                    '.png', '.jpg', '.jpeg', '.bmp', '.gif', '.webp'):
                try:
                    name = save_image_file(path)
                    self._pending_image_name = name
                    self._new_image_temp = image_abs_path(name)
                    self._show_preview(QPixmap(str(self._new_image_temp)))
                    self.tabs.setCurrentIndex(2)
                    self.hint_label.setText("✅ 已导入拖入的图片")
                except Exception as e:
                    QMessageBox.warning(self, "提示", f"图片读取失败：{e}")
                break

    # ------------------------------------------------------------- 输出
    def get_feedback(self) -> Feedback:
        """组装便签数据（新建或编辑）"""
        fb = self.feedback or Feedback()
        tab = self.tabs.currentIndex()
        old_image = (self.feedback.image_path if self.feedback else "")

        if tab == 2 and self._pending_image_name:
            fb.content_type = 'image'
            fb.image_path = self._pending_image_name
            fb.image_caption = self.caption_edit.text().strip()
            fb.content = fb.image_caption
            fb.table_json = ''
            if old_image and old_image != fb.image_path:
                delete_image_file(old_image)
        elif tab == 1:
            data = self._table_data()
            if data['rows'] or any(h.strip() and not h.startswith('列') for h in data['headers']):
                fb.content_type = 'table'
                fb.table_json = json.dumps(data, ensure_ascii=False)
                fb.content = ''
                fb.image_caption = ''
                if old_image:
                    delete_image_file(old_image)
                fb.image_path = ''
            else:
                fb.content_type = 'text'
                fb.content = self.text_edit.toPlainText().strip()
                fb.table_json = ''
        else:
            fb.content_type = 'text'
            fb.content = self.text_edit.toPlainText().strip()
            fb.table_json = ''
            fb.image_caption = ''
            if old_image:
                delete_image_file(old_image)
            fb.image_path = ''

        fb.color = self.selected_color
        fb.date = date.today().isoformat()
        fb.time = datetime.now().strftime("%H:%M")
        fb.updated_at = datetime.now().isoformat()
        return fb

    def accept(self):
        fb = self.get_feedback()
        if fb.content_type == 'text' and not fb.content.strip():
            QMessageBox.information(self, "提示", "便签内容不能为空：请写点文字、贴张表格或图片")
            return
        if fb.content_type == 'image' and not fb.image_path:
            QMessageBox.information(self, "提示", "请先粘贴或选择一张图片")
            return
        super().accept()


class ImageViewerDialog(QDialog):
    """图片便签查看大图（可缩放）"""

    def __init__(self, feedback: Feedback, parent=None):
        super().__init__(parent)
        self.setWindowTitle(feedback.image_caption or "查看图片")
        self.setMinimumSize(700, 560)
        layout = QVBoxLayout(self)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        label = QLabel()
        label.setAlignment(Qt.AlignCenter)
        path = image_abs_path(feedback.image_path)
        if path.exists():
            pix = QPixmap(str(path))
            label.setPixmap(pix)
            self._pixmap = pix
        else:
            label.setText("图片文件缺失")
        self.scroll.setWidget(label)
        layout.addWidget(self.scroll, stretch=1)

        row = QHBoxLayout()
        row.addStretch()
        copy_btn = QPushButton("复制图片到剪贴板")
        copy_btn.setStyleSheet(StickerEditDialog._mini_btn_style())
        copy_btn.clicked.connect(lambda: self._copy())
        row.addWidget(copy_btn)
        close_btn = QPushButton("关闭")
        close_btn.setStyleSheet(StickerEditDialog._primary_btn_style())
        close_btn.clicked.connect(self.accept)
        row.addWidget(close_btn)
        layout.addLayout(row)

    def _copy(self):
        try:
            QApplication.clipboard().setPixmap(self._pixmap)
        except Exception:
            pass


class FeedbackModule(BaseModulePage):
    """反馈 / 便签墙模块

    - 像剪贴板 / sticker 便签：贴**文字、表格、图片**都行
    - **不再按日期/班级/范围筛选**：所有便签统一展示在一面便签墙上
    - 支持智能粘贴（Ctrl+V 或按钮）、置顶、编辑、删除、复制内容
    """

    COLUMNS = 3

    def __init__(self, db: DatabaseManager, parent=None):
        super().__init__(db, parent)
        self.feedbacks: List[Feedback] = []
        self.current_filter_date = None
        self.current_filter_class = None
        self.current_filter_category = None

    # ------------------------------------------------------------------ UI
    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        toolbar = QHBoxLayout()
        title = QLabel("💬 便签墙（反馈）")
        title.setStyleSheet(
            f"font-size: 22px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(title)

        sub = QLabel("文字 / 表格 / 图片都可以贴，像便签一样统一展示")
        sub.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; padding-left: 10px; "
            f"font-family: {FONT_STACK_CSS};")
        toolbar.addWidget(sub)
        toolbar.addStretch()

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("搜索便签内容…")
        self.search_edit.setFixedWidth(220)
        self.search_edit.textChanged.connect(self.on_search)
        self.search_edit.setStyleSheet(f"""
            QLineEdit {{
                padding: 8px 12px; border: 1px solid {COLORS['border']};
                border-radius: 18px; font-size: 13px;
                font-family: {FONT_STACK_CSS}; background-color: {COLORS['card_bg']};
            }}
            QLineEdit:focus {{ border-color: {COLORS['primary']}; }}
        """)
        toolbar.addWidget(self.search_edit)

        paste_new_btn = QPushButton("📋 粘贴新建")
        paste_new_btn.setToolTip("读取剪贴板内容（图片/表格/文字）直接创建便签")
        paste_new_btn.setStyleSheet(self._get_btn_style(COLORS['secondary']))
        paste_new_btn.clicked.connect(self.add_from_clipboard)
        toolbar.addWidget(paste_new_btn)

        add_btn = QPushButton("➕ 新建便签")
        add_btn.setStyleSheet(self._get_btn_style(COLORS['success']))
        add_btn.clicked.connect(self.add_feedback)
        toolbar.addWidget(add_btn)

        delete_btn = QPushButton("🗑 删除便签")
        delete_btn.setStyleSheet(self._get_btn_style(COLORS['danger']))
        delete_btn.clicked.connect(self.delete_feedbacks)
        toolbar.addWidget(delete_btn)

        layout.addLayout(toolbar)

        # 便签墙（统一展示，无日期/范围限制）
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QFrame.NoFrame)
        self.scroll_area.setStyleSheet("QScrollArea { background: transparent; }")
        self.tiles_container = QWidget()
        self.tiles_container.setStyleSheet("background: transparent;")
        self.tiles_layout = QGridLayout(self.tiles_container)
        self.tiles_layout.setSpacing(16)
        self.tiles_layout.setContentsMargins(5, 5, 5, 5)
        self.tiles_layout.setAlignment(Qt.AlignTop | Qt.AlignLeft)
        self.scroll_area.setWidget(self.tiles_container)
        layout.addWidget(self.scroll_area, stretch=1)

        self.stats_label = QLabel()
        self.stats_label.setStyleSheet(
            f"font-size: 12px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(self.stats_label)

    def _get_btn_style(self, color):
        return f"""
            QPushButton {{
                background-color: {color}; color: white; border: none;
                border-radius: 16px; padding: 7px 14px; font-size: 12px;
                font-family: {FONT_STACK_CSS}; font-weight: 500;
            }}
            QPushButton:hover {{ opacity: 0.85; }}
        """

    # ------------------------------------------------------------- 数据
    def refresh(self):
        """统一展示全部便签（置顶在前）"""
        keyword = self.search_edit.text().strip() if hasattr(self, 'search_edit') else ''
        self.feedbacks = (self.db.search_feedbacks(keyword) if keyword
                          else self.db.get_all_feedbacks())
        self.display_feedbacks(self.feedbacks)

    def on_search(self, keyword: str):
        self.feedbacks = (self.db.search_feedbacks(keyword.strip()) if keyword.strip()
                          else self.db.get_all_feedbacks())
        self.display_feedbacks(self.feedbacks)

    def display_feedbacks(self, feedbacks: List[Feedback]):
        # 只 deleteLater，不 setParent(None)：避免便签磁贴变成顶层窗口
        while self.tiles_layout.count():
            item = self.tiles_layout.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()

        for i, fb in enumerate(feedbacks):
            tile = StickerTile(fb, parent=self.tiles_container)   # 必须给父控件
            tile.edit_requested.connect(self.on_edit_sticker)
            tile.delete_requested.connect(lambda f: self._confirm_delete_feedbacks([f.id]))
            tile.pin_toggled.connect(self.toggle_pin)
            tile.open_requested.connect(self.show_image)
            tile.double_clicked.connect(self.on_edit_sticker)
            row, col = divmod(i, self.COLUMNS)
            self.tiles_layout.addWidget(tile, row, col, alignment=Qt.AlignTop)

        counts = {'text': 0, 'table': 0, 'image': 0}
        for fb in feedbacks:
            counts[fb.content_type] = counts.get(fb.content_type, 0) + 1
        self.stats_label.setText(
            f"共 {len(feedbacks)} 条便签 · 文字 {counts.get('text', 0)} · "
            f"表格 {counts.get('table', 0)} · 图片 {counts.get('image', 0)}"
            f"　（Ctrl+V 或「粘贴新建」可快速贴图/贴表）")

    # ------------------------------------------------------------- 交互
    def add_feedback(self):
        dialog = StickerEditDialog(self.db, parent=self)
        if dialog.exec() == QDialog.Accepted:
            self.db.add_feedback(dialog.get_feedback())
            self.refresh()

    def add_from_clipboard(self):
        """把剪贴板内容直接变成一张便签"""
        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData()
        try:
            if mime is not None and mime.hasImage():
                image = clipboard.image()
                if image is not None and not image.isNull():
                    name = save_qimage_file(image)
                    fb = Feedback(content_type='image', image_path=name,
                                  content='', color=random.choice(STICKER_COLORS))
                    self.db.add_feedback(fb)
                    self.refresh()
                    return

            text = clipboard.text() or ""
            if text.strip():
                parsed = StickerEditDialog._parse_table_text(text)
                if parsed:
                    fb = Feedback(content_type='table',
                                  table_json=json.dumps(
                                      {'headers': parsed[0], 'rows': parsed[1:]},
                                      ensure_ascii=False),
                                  color=random.choice(STICKER_COLORS))
                else:
                    fb = Feedback(content_type='text', content=text.strip(),
                                  color=random.choice(STICKER_COLORS))
                self.db.add_feedback(fb)
                self.refresh()
                return
        except Exception as e:
            QMessageBox.warning(self, "粘贴失败", f"无法读取剪贴板内容：{e}")
            return
        QMessageBox.information(self, "提示", "剪贴板里没有可粘贴的图片、表格或文字")

    def on_edit_sticker(self, data):
        fb = data.feedback if isinstance(data, StickerTile) else data
        dialog = StickerEditDialog(self.db, feedback=fb, parent=self)
        if dialog.exec() == QDialog.Accepted:
            updated = dialog.get_feedback()
            updated.id = fb.id
            updated.pinned = fb.pinned
            self.db.update_feedback(updated)
            self.refresh()

    def toggle_pin(self, data):
        fb = data.feedback if isinstance(data, StickerTile) else data
        self.db.set_feedback_pinned(fb.id, not bool(fb.pinned))
        self.refresh()

    def show_image(self, data):
        fb = data.feedback if isinstance(data, StickerTile) else data
        dialog = ImageViewerDialog(fb, parent=self)
        dialog.exec()

    # 兼容旧接口（第一版反馈模块的筛选按钮）——现在统一展示
    def apply_filter(self):
        self.refresh()

    def reset_filter(self):
        if hasattr(self, 'search_edit'):
            self.search_edit.blockSignals(True)
            self.search_edit.clear()
            self.search_edit.blockSignals(False)
        self.refresh()

    def delete_feedbacks(self):
        """批量/单个删除便签"""
        feedbacks = list(self.feedbacks)
        if not feedbacks:
            feedbacks = self.db.get_all_feedbacks()
        if not feedbacks:
            QMessageBox.information(self, "提示", "暂无便签可删除")
            return
        records = []
        for f in feedbacks:
            type_map = {'text': "文字", 'table': "表格", 'image': "图片"}
            preview = f.plain_text().replace('\n', ' / ')[:34]
            records.append({
                'id': f.id,
                'text': f"[{type_map.get(f.content_type, '便签')}] {preview}"
                        f"{'…' if len(f.plain_text()) > 34 else ''}",
                'detail': f"{f.created_at[:16].replace('T', ' ')}"
                          f"{'　📌已置顶' if f.pinned else ''}",
            })
        dialog = MultiDeleteDialog(records, title="删除便签",
                                   instruction="勾选要删除的便签（可多选；勾 1 个即单个删除）。该操作不可恢复：",
                                   parent=self)
        if dialog.exec() == QDialog.Accepted:
            ids = dialog.selected_ids()
            if ids:
                self._confirm_delete_feedbacks(ids)

    def _confirm_delete_feedbacks(self, ids: List[int], confirm: bool = True) -> int:
        """执行删除便签（含清理图片文件），返回删除数量；confirm=False 用于测试"""
        if not ids:
            return 0
        ids = [i for i in ids if i]
        if confirm:
            reply = QMessageBox.question(
                self, "确认删除", f"确定删除选中的 {len(ids)} 条便签吗？\n该操作不可恢复！",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return 0
        removed = 0
        for fid in ids:
            fb = next((f for f in self.db.get_all_feedbacks() if f.id == fid), None)
            if self.db.delete_feedback(fid):
                removed += 1
                if fb is not None and fb.content_type == 'image' and fb.image_path:
                    delete_image_file(fb.image_path)
        self.refresh()
        if confirm:
            QMessageBox.information(self, "删除完成", f"已删除便签 {removed} 条。")
        return removed

class GradeModule(BaseModulePage):
    """成绩管理系统模块"""

    def __init__(self, db: DatabaseManager, parent=None):
        super().__init__(db, parent)
        self.grades: List[Grade] = []
        self.students: List[Student] = []
        self.current_student_id: Optional[int] = None

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setContentsMargins(20, 15, 20, 15)
        layout.setSpacing(10)

        # 顶部工具栏
        toolbar = QHBoxLayout()
        title = QLabel("📊 成绩管理系统")
        title.setStyleSheet(f"font-size: 22px; font-weight: bold; color: {COLORS['text_primary']};")
        toolbar.addWidget(title)
        toolbar.addStretch()

        import_btn = QPushButton("📥 导入成绩")
        import_btn.setStyleSheet(self._get_btn_style(COLORS['success']))
        import_btn.clicked.connect(self.import_grades)
        toolbar.addWidget(import_btn)

        add_btn = QPushButton("➕ 手动录入")
        add_btn.setStyleSheet(self._get_btn_style(COLORS['primary']))
        add_btn.clicked.connect(self.add_grade)
        toolbar.addWidget(add_btn)

        delete_btn = QPushButton("🗑 删除成绩")
        delete_btn.setStyleSheet(self._get_btn_style(COLORS['danger']))
        delete_btn.clicked.connect(self.delete_grades)
        toolbar.addWidget(delete_btn)

        layout.addLayout(toolbar)

        # 分割视图
        splitter = QSplitter(Qt.Horizontal)

        # 左侧：学生成绩磁贴
        left_panel = QWidget()
        left_layout = QVBoxLayout(left_panel)
        left_layout.setContentsMargins(0, 0, 5, 0)

        student_label = QLabel("学生成绩概览")
        student_label.setStyleSheet(f"font-size: 15px; font-weight: bold; color: {COLORS['text_primary']};")
        left_layout.addWidget(student_label)

        self.student_scroll = QScrollArea()
        self.student_scroll.setWidgetResizable(True)
        self.student_scroll.setFrameShape(QFrame.NoFrame)
        self.student_tiles_container = QWidget()
        self.student_tiles_layout = QGridLayout(self.student_tiles_container)
        self.student_tiles_layout.setSpacing(10)
        self.student_tiles_layout.setContentsMargins(0, 5, 0, 5)
        self.student_scroll.setWidget(self.student_tiles_container)
        left_layout.addWidget(self.student_scroll)

        splitter.addWidget(left_panel)

        # 右侧：图表区域
        right_panel = QWidget()
        right_layout = QVBoxLayout(right_panel)
        right_layout.setContentsMargins(5, 0, 0, 0)

        self.chart_tabs = QTabWidget()
        self.chart_tabs.setStyleSheet(f"""
            QTabWidget::pane {{
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                background: {COLORS['card_bg']};
            }}
            QTabBar::tab {{
                padding: 8px 16px;
                font-size: 12px;
                font-family: {FONT_STACK_CSS};
            }}
            QTabBar::tab:selected {{
                color: {COLORS['primary']};
                border-bottom: 3px solid {COLORS['primary']};
            }}
        """)
        right_layout.addWidget(self.chart_tabs)

        splitter.addWidget(right_panel)
        splitter.setSizes([350, 600])
        layout.addWidget(splitter, stretch=1)

    def _get_btn_style(self, color):
        return f"""
            QPushButton {{
                background-color: {color};
                color: white;
                border: none;
                border-radius: 18px;
                padding: 8px 16px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
                font-weight: 500;
            }}
            QPushButton:hover {{ opacity: 0.85; }}
        """

    def refresh(self):
        self.students = self.db.get_all_students()
        self.grades = self.db.get_all_grades()

        # 刷新学生磁贴（只 deleteLater，避免控件退化为顶层窗口）
        while self.student_tiles_layout.count():
            item = self.student_tiles_layout.takeAt(0)
            w = item.widget()
            if w:
                w.hide()
                w.deleteLater()

        for i, student in enumerate(self.students):
            student_grades = [g for g in self.grades if g.student_id == student.id]
            tile = GradeTile(student, student_grades, parent=self.student_tiles_container)
            tile.clicked.connect(lambda sd, sid=student.id: self.on_student_selected(sid))
            tile.double_clicked.connect(lambda sd, sid=student.id: self.show_student_report(sid))
            tile.edit_requested.connect(
                lambda data, sid=student.id: self.edit_student_grades(sid))
            tile.delete_requested.connect(
                lambda data, sid=student.id: self.delete_student_grades(sid))
            row = i // 3
            col = i % 3
            self.student_tiles_layout.addWidget(tile, row, col)

        # 刷新图表
        self.refresh_charts()

    def edit_student_grades(self, student_id: int):
        """成绩磁贴的编辑入口：打开该生成绩编辑器（可逐条改/删/新增）"""
        student = self.db.get_student_by_id(student_id)
        if student is None:
            return
        dialog = StudentGradeEditDialog(self.db, student, parent=self)
        dialog.exec()
        self.refresh()

    def delete_student_grades(self, student_id: int, confirm: bool = True) -> int:
        """删除某学生的全部成绩（成绩磁贴右键）"""
        grades = self.db.get_grades_by_student(student_id)
        if not grades:
            if confirm:
                QMessageBox.information(self, "提示", "该学生暂无成绩记录")
            return 0
        if confirm:
            reply = QMessageBox.question(
                self, "确认删除",
                f"确定删除该学生的全部 {len(grades)} 条成绩记录吗？\n该操作不可恢复！",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return 0
        removed = self.db.delete_grades_by_student(student_id)
        self.refresh()
        if confirm:
            QMessageBox.information(self, "删除完成", f"已删除 {removed} 条成绩。")
        return removed

    def on_student_selected(self, student_id: int):
        self.current_student_id = student_id
        self.refresh_charts()

    def refresh_charts(self):
        """刷新统计图表"""
        # 清除现有图表
        self.chart_tabs.clear()

        if not self.grades:
            empty_label = QLabel("暂无成绩数据，请先导入成绩", parent=self.chart_tabs)
            empty_label.setAlignment(Qt.AlignCenter)
            empty_label.setStyleSheet(f"font-size: 16px; color: {COLORS['text_secondary']};")
            self.chart_tabs.addTab(empty_label, "统计图表")
            return

        # 创建图表
        self.create_pie_chart()
        self.create_bar_chart()
        self.create_trend_chart()

        # 如果有选中的学生，生成个人报告
        if self.current_student_id:
            self.create_student_report(self.current_student_id)

    def create_pie_chart(self):
        """创建扇形统计图 - 各科目平均分分布"""
        if not HAS_MATPLOTLIB:
            return

        fig = Figure(figsize=(6, 5), facecolor=COLORS['card_bg'])
        ax = fig.add_subplot(111)

        course_scores = {}
        for g in self.grades:
            if g.course_name not in course_scores:
                course_scores[g.course_name] = []
            course_scores[g.course_name].append(g.score)

        avg_scores = {k: sum(v) / len(v) for k, v in course_scores.items()}
        labels = list(avg_scores.keys())
        values = list(avg_scores.values())

        colors = COURSE_COLORS[:len(labels)]
        wedges, texts, autotexts = ax.pie(
            values, labels=labels, colors=colors, autopct='%1.1f%%',
            startangle=90, pctdistance=0.75
        )
        for t in autotexts:
            t.set_fontsize(9)
        ax.set_title('各科目平均分分布', fontsize=14, fontweight='bold')
        ax.axis('equal')

        canvas = FigureCanvas(fig)   # 由 addTab 接管父子关系，避免瞬时顶层窗口
        self.chart_tabs.addTab(canvas, "📊 分数分布")

    def create_bar_chart(self):
        """创建条形统计图"""
        if not HAS_MATPLOTLIB:
            return

        fig = Figure(figsize=(6, 5), facecolor=COLORS['card_bg'])
        ax = fig.add_subplot(111)

        # 按考试类型分组
        exam_types = sorted(set(g.exam_type for g in self.grades))
        courses = sorted(set(g.course_name for g in self.grades))

        data = {}
        for et in exam_types:
            data[et] = []
            for c in courses:
                scores = [g.score for g in self.grades if g.exam_type == et and g.course_name == c]
                data[et].append(sum(scores) / len(scores) if scores else 0)

        x = range(len(courses))
        width = 0.8 / len(exam_types) if exam_types else 0.4

        for i, et in enumerate(exam_types):
            offset = (i - len(exam_types) / 2) * width
            bars = ax.bar([xi + offset for xi in x], data[et], width, label=et, alpha=0.8)

        ax.set_xticks(list(x))
        ax.set_xticklabels(courses, fontsize=9)
        ax.set_ylabel('平均分', fontsize=11)
        ax.set_title('各科目平均成绩对比', fontsize=14, fontweight='bold')
        ax.legend(fontsize=9)
        ax.grid(axis='y', alpha=0.3)
        fig.tight_layout()

        canvas = FigureCanvas(fig)   # 由 addTab 接管父子关系，避免瞬时顶层窗口
        self.chart_tabs.addTab(canvas, "📈 成绩对比")

    def create_trend_chart(self):
        """创建趋势图"""
        if not HAS_MATPLOTLIB:
            return

        fig = Figure(figsize=(6, 5), facecolor=COLORS['card_bg'])
        ax = fig.add_subplot(111)

        if self.current_student_id:
            student_grades = [g for g in self.grades if g.student_id == self.current_student_id]
            student = self.db.get_student_by_id(self.current_student_id)
            title = f"{student.name if student else '学生'} 的成绩趋势"
        else:
            student_grades = self.grades
            title = "全体学生成绩趋势"

        if student_grades:
            dates = sorted(set(g.exam_date for g in student_grades))
            courses = sorted(set(g.course_name for g in student_grades))

            for course in courses:
                course_grades = [g for g in student_grades if g.course_name == course]
                course_grades.sort(key=lambda g: g.exam_date)
                x = [g.exam_date for g in course_grades]
                y = [g.score for g in course_grades]
                ax.plot(x, y, marker='o', label=course, linewidth=2)

            ax.set_xlabel('考试日期', fontsize=10)
            ax.set_ylabel('分数', fontsize=10)
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.legend(fontsize=9)
            ax.grid(alpha=0.3)
            plt.setp(ax.get_xticklabels(), rotation=45, ha='right', fontsize=8)
            fig.tight_layout()
        else:
            ax.text(0.5, 0.5, '暂无数据', ha='center', va='center', fontsize=16,
                    color=COLORS['text_secondary'])
            ax.set_title(title, fontsize=14, fontweight='bold')
            ax.axis('off')

        canvas = FigureCanvas(fig)   # 由 addTab 接管父子关系，避免瞬时顶层窗口
        self.chart_tabs.addTab(canvas, "📉 成绩趋势")

    def create_student_report(self, student_id: int):
        """创建学生个人成绩报告"""
        if not HAS_MATPLOTLIB:
            return

        student = self.db.get_student_by_id(student_id)
        if not student:
            return

        student_grades = self.db.get_grades_by_student(student_id)
        if not student_grades:
            return

        fig = Figure(figsize=(7, 6), facecolor=COLORS['card_bg'])
        gs = fig.add_gridspec(2, 2, hspace=0.35, wspace=0.3)

        # 左上：成绩汇总
        ax1 = fig.add_subplot(gs[0, 0])
        course_avg = {}
        for g in student_grades:
            if g.course_name not in course_avg:
                course_avg[g.course_name] = []
            course_avg[g.course_name].append(g.score)
        courses = list(course_avg.keys())
        avgs = [sum(v) / len(v) for v in course_avg.values()]
        colors = [COURSE_COLORS[i % len(COURSE_COLORS)] for i in range(len(courses))]
        bars = ax1.bar(courses, avgs, color=colors, alpha=0.85)
        ax1.set_title('各科目平均分', fontsize=11, fontweight='bold')
        ax1.set_ylabel('平均分', fontsize=9)
        ax1.tick_params(axis='x', rotation=30, labelsize=8)
        ax1.grid(axis='y', alpha=0.3)

        # 右上：成绩变化趋势
        ax2 = fig.add_subplot(gs[0, 1])
        if student_grades:
            student_grades_sorted = sorted(student_grades, key=lambda g: g.exam_date)
            x = list(range(len(student_grades_sorted)))
            y = [g.score for g in student_grades_sorted]
            ax2.plot(x, y, marker='o', color=COLORS['primary'], linewidth=2, markersize=6)
            ax2.fill_between(x, y, alpha=0.2, color=COLORS['primary'])
            ax2.set_xticks(x)
            ax2.set_xticklabels([g.exam_date for g in student_grades_sorted], rotation=30, fontsize=7)
            ax2.set_title('成绩变化趋势', fontsize=11, fontweight='bold')
            ax2.set_ylabel('分数', fontsize=9)
            ax2.grid(alpha=0.3)

        # 下方：分析报告文本
        ax3 = fig.add_subplot(gs[1, :])
        ax3.axis('off')

        # 生成分析报告
        report_lines = []
        report_lines.append(f"📋 成绩分析报告 - {student.name}")
        report_lines.append(f"学号：{student.student_no} | 班级：{student.class_name or '未分配'}")
        report_lines.append("")

        if student_grades:
            total_avg = sum(g.score for g in student_grades) / len(student_grades)
            max_score = max(g.score for g in student_grades)
            min_score = min(g.score for g in student_grades)

            report_lines.append(f"📊 总体表现：")
            report_lines.append(f"  • 平均分：{total_avg:.1f}")
            report_lines.append(f"  • 最高分：{max_score:.1f}")
            report_lines.append(f"  • 最低分：{min_score:.1f}")

            if total_avg >= 85:
                report_lines.append("  • 评价：优秀，继续保持！🌟")
            elif total_avg >= 70:
                report_lines.append("  • 评价：良好，还有提升空间 💪")
            elif total_avg >= 60:
                report_lines.append("  • 评价：及格，需要加强学习 📖")
            else:
                report_lines.append("  • 评价：需要重点关注 ⚠️")

            report_lines.append("")
            report_lines.append("📚 各科目表现：")
            for course in courses:
                scores = course_avg[course]
                avg = sum(scores) / len(scores)
                trend = "📈上升" if len(scores) > 1 and scores[-1] > scores[0] else "📉下降" if len(scores) > 1 and scores[-1] < scores[0] else "➡️平稳"
                report_lines.append(f"  • {course}：平均{avg:.1f}分 {trend}")

        report_text = '\n'.join(report_lines)
        ax3.text(0.05, 0.95, report_text, transform=ax3.transAxes, fontsize=9,
                 verticalalignment='top', family=FONT_STACK,
                 bbox=dict(boxstyle='round', facecolor='#F8F9FA', alpha=0.8))

        fig.suptitle(f'{student.name} - 个人成绩报告', fontsize=14, fontweight='bold', y=0.98)
        canvas = FigureCanvas(fig)   # 由 addTab 接管父子关系，避免瞬时顶层窗口
        self.chart_tabs.addTab(canvas, f"👤 {student.name}")

    def show_student_report(self, student_id: int):
        self.current_student_id = student_id
        self.refresh_charts()
        # 切换到报告标签
        for i in range(self.chart_tabs.count()):
            if self.chart_tabs.tabText(i).startswith("👤"):
                self.chart_tabs.setCurrentIndex(i)
                break

    def import_grades(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "选择成绩文件", "", "Excel文件 (*.xlsx *.xls *.xlsm *.csv);;所有文件 (*)"
        )
        if not file_path:
            return

        service = DataImportService(self.db)
        success, errors = service.import_grades(file_path)

        msg = f"导入完成：成功 {success} 条"
        if errors:
            msg += f"\n失败 {len(errors)} 条：\n" + '\n'.join(errors[:10])
        QMessageBox.information(self, "导入结果", msg)
        self.refresh()

    def add_grade(self):
        dialog = GradeEditDialog(self.db, parent=self)
        if dialog.exec() == QDialog.Accepted:
            self.refresh()

    def delete_grades(self):
        """打开批量/单个删除成绩记录对话框（列出全部成绩）"""
        grades = self.db.get_all_grades()
        if not grades:
            QMessageBox.information(self, "提示", "暂无成绩数据可删除")
            return
        students = self.db.get_all_students()
        student_ids = {s.id: s.name for s in students}
        records = []
        for g in grades:
            name_txt = g.student_name or student_ids.get(g.student_id, '未知学生')
            records.append({
                'id': g.id,
                'text': f"{name_txt}｜{g.course_name}　{g.score:.1f}分",
                'detail': f"{g.exam_date}　{g.exam_type}　满分 {g.full_score:.0f}　{g.comment or ''}",
            })
        dialog = MultiDeleteDialog(records, title="删除成绩",
                                   instruction="勾选要删除的成绩记录（可多选；勾 1 条即单个删除）。按学生删除请用「学生管理→删除学生」。该操作不可恢复：",
                                   parent=self)
        if dialog.exec() == QDialog.Accepted:
            ids = dialog.selected_ids()
            if ids:
                self._confirm_delete_grades(ids)

    def _confirm_delete_grades(self, ids: List[int], confirm: bool = True) -> int:
        """执行删除成绩记录，返回删除数量；confirm=False 用于测试/无弹窗调用"""
        if not ids:
            return 0
        ids = [i for i in ids if i]
        if confirm:
            reply = QMessageBox.question(
                self, "确认删除", f"确定删除选中的 {len(ids)} 条成绩记录吗？\n该操作不可恢复！",
                QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
            if reply != QMessageBox.Yes:
                return 0
        removed = 0
        for gid in ids:
            if self.db.delete_grade(gid):
                removed += 1
        self.current_student_id = None
        self.refresh()
        if confirm:
            QMessageBox.information(self, "删除完成", f"已删除成绩 {removed} 条。")
        return removed


class GradeEditDialog(QDialog):
    """成绩录入 / 编辑对话框

    - 不传 grade 时为「录入成绩」（可选择学生）
    - 传 grade（可带 student）时为「编辑成绩」，保存时更新原记录
    """

    def __init__(self, db: DatabaseManager, parent=None,
                 grade: Optional[Grade] = None, student: Optional[Student] = None):
        super().__init__(parent)
        self.db = db
        self.grade = grade
        self.student = student
        self.setWindowTitle("编辑成绩" if grade is not None else "录入成绩")
        self.setMinimumSize(400, 360)
        self.setup_ui()
        if grade is not None:
            self.load_grade(grade, student)

    def setup_ui(self):
        layout = QFormLayout(self)
        layout.setSpacing(10)

        self.student_combo = QComboBox()
        students = self.db.get_all_students()
        for s in students:
            self.student_combo.addItem(f"{s.name} ({s.student_no})", s.id)
        layout.addRow("学生:", self.student_combo)

        self.course_edit = QLineEdit()
        self.course_edit.setPlaceholderText("科目名称")
        layout.addRow("科目:", self.course_edit)

        self.score_spin = QDoubleSpinBox()
        self.score_spin.setRange(0, 150)
        self.score_spin.setValue(0)
        self.score_spin.setDecimals(1)
        layout.addRow("分数:", self.score_spin)

        self.full_score_spin = QDoubleSpinBox()
        self.full_score_spin.setRange(1, 200)
        self.full_score_spin.setValue(100)
        self.full_score_spin.setDecimals(1)
        layout.addRow("满分:", self.full_score_spin)

        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("yyyy-MM-dd")
        layout.addRow("考试日期:", self.date_edit)

        self.type_combo = QComboBox()
        self.type_combo.addItems(['期中', '期末', '月考', '周测', '其他'])
        layout.addRow("考试类型:", self.type_combo)

        self.comment_edit = QLineEdit()
        self.comment_edit.setPlaceholderText("评语（可选）")
        layout.addRow("评语:", self.comment_edit)

        btn_box = QDialogButtonBox(QDialogButtonBox.Ok | QDialogButtonBox.Cancel)
        btn_box.button(QDialogButtonBox.Ok).setText("保存")
        btn_box.button(QDialogButtonBox.Cancel).setText("取消")
        btn_box.accepted.connect(self.accept)
        btn_box.rejected.connect(self.reject)
        layout.addRow(btn_box)

    def load_grade(self, grade: Grade, student: Optional[Student] = None):
        """把已有成绩填入表单"""
        target_id = grade.student_id
        if student is not None:
            target_id = student.id
        idx = self.student_combo.findData(target_id)
        if idx >= 0:
            self.student_combo.setCurrentIndex(idx)
        self.course_edit.setText(grade.course_name)
        self.score_spin.setValue(float(grade.score or 0))
        self.full_score_spin.setValue(float(grade.full_score or 100))
        d = QDate.fromString(grade.exam_date, "yyyy-MM-dd")
        if d.isValid():
            self.date_edit.setDate(d)
        ti = self.type_combo.findText(grade.exam_type)
        if ti >= 0:
            self.type_combo.setCurrentIndex(ti)
        self.comment_edit.setText(grade.comment or "")

    def accept(self):
        student_id = self.student_combo.currentData()
        course_name = self.course_edit.text().strip()
        if not course_name:
            QMessageBox.warning(self, "提示", "请输入科目名称")
            return

        students = self.db.get_all_students()
        student_name = next((s.name for s in students if s.id == student_id), '')

        if self.grade is not None:
            self.grade.student_id = student_id or self.grade.student_id
            self.grade.student_name = student_name
            self.grade.course_name = course_name
            self.grade.score = self.score_spin.value()
            self.grade.exam_date = self.date_edit.date().toString("yyyy-MM-dd")
            self.grade.exam_type = self.type_combo.currentText()
            self.grade.full_score = self.full_score_spin.value()
            self.grade.comment = self.comment_edit.text().strip()
            self.db.update_grade(self.grade)
        else:
            grade = Grade(
                student_id=student_id,
                student_name=student_name,
                course_name=course_name,
                score=self.score_spin.value(),
                exam_date=self.date_edit.date().toString("yyyy-MM-dd"),
                exam_type=self.type_combo.currentText(),
                full_score=self.full_score_spin.value(),
                comment=self.comment_edit.text().strip(),
            )
            self.db.add_grade(grade)
        super().accept()


class StudentGradeEditDialog(QDialog):
    """学生成绩编辑对话框（成绩磁贴的「✏️ 编辑」入口）

    以表格列出该生全部成绩，可逐条编辑 / 删除 / 新增，保存后写回数据库。
    """

    def __init__(self, db: DatabaseManager, student: Student, parent=None):
        super().__init__(parent)
        self.db = db
        self.student = student
        self.setWindowTitle(f"编辑成绩 - {student.name}")
        self.setMinimumSize(700, 460)
        self.setup_ui()
        self.load_data()

    def setup_ui(self):
        layout = QVBoxLayout(self)
        layout.setSpacing(10)

        head = QLabel(
            f"👤 {self.student.name}　学号 {self.student.student_no}　"
            f"班级 {self.student.class_name or '未分配'}")
        head.setStyleSheet(
            f"font-size: 14px; font-weight: bold; color: {COLORS['text_primary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(head)

        tip = QLabel("双击某行可直接编辑；也可选中后用下方按钮操作。")
        tip.setStyleSheet(
            f"font-size: 11px; color: {COLORS['text_secondary']}; "
            f"font-family: {FONT_STACK_CSS};")
        layout.addWidget(tip)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(['科目', '分数', '满分', '考试日期', '类型'])
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.table.setAlternatingRowColors(True)
        self.table.doubleClicked.connect(lambda _idx: self.edit_selected())
        self.table.setStyleSheet(f"""
            QTableWidget {{
                border: 1px solid {COLORS['border']}; border-radius: 8px;
                font-size: 12px; font-family: {FONT_STACK_CSS};
            }}
            QHeaderView::section {{
                background-color: {COLORS['primary_light']}; font-weight: bold;
                padding: 6px; border: none;
            }}
        """)
        layout.addWidget(self.table, stretch=1)

        row = QHBoxLayout()
        add_btn = QPushButton("➕ 新增成绩")
        add_btn.setStyleSheet(StickerEditDialog._primary_btn_style())
        add_btn.clicked.connect(self.add_grade)
        row.addWidget(add_btn)

        edit_btn = QPushButton("✏️ 编辑选中")
        edit_btn.setStyleSheet(StickerEditDialog._mini_btn_style())
        edit_btn.clicked.connect(self.edit_selected)
        row.addWidget(edit_btn)

        del_btn = QPushButton("🗑 删除选中")
        del_btn.setStyleSheet(StickerEditDialog._mini_btn_style())
        del_btn.clicked.connect(self.delete_selected)
        row.addWidget(del_btn)

        row.addStretch()
        close_btn = QPushButton("完成")
        close_btn.setStyleSheet(StickerEditDialog._primary_btn_style())
        close_btn.clicked.connect(self.accept)
        row.addWidget(close_btn)
        layout.addLayout(row)

    # ---- 数据 ----
    def _grades(self) -> List[Grade]:
        return self.db.get_grades_by_student(self.student.id or 0)

    def load_data(self):
        grades = self._grades()
        self.table.setRowCount(len(grades))
        for i, g in enumerate(grades):
            self.table.setItem(i, 0, QTableWidgetItem(g.course_name))
            score_item = QTableWidgetItem(f"{g.score:.1f}")
            score_item.setData(Qt.UserRole, g.id)
            if g.score >= 85:
                score_item.setForeground(QColor('#00B894'))
            elif g.score >= 60:
                score_item.setForeground(QColor('#FDCB6E'))
            else:
                score_item.setForeground(QColor('#FF6B6B'))
            self.table.setItem(i, 1, score_item)
            self.table.setItem(i, 2, QTableWidgetItem(f"{g.full_score:.0f}"))
            self.table.setItem(i, 3, QTableWidgetItem(g.exam_date))
            self.table.setItem(i, 4, QTableWidgetItem(g.exam_type))

    def _selected_grade(self) -> Optional[Grade]:
        row = self.table.currentRow()
        if row < 0:
            return None
        item = self.table.item(row, 1)
        gid = item.data(Qt.UserRole) if item else None
        return next((g for g in self._grades() if g.id == gid), None)

    def add_grade(self):
        dialog = GradeEditDialog(self.db, parent=self, student=self.student)
        idx = dialog.student_combo.findData(self.student.id)
        if idx >= 0:
            dialog.student_combo.setCurrentIndex(idx)
        if dialog.exec() == QDialog.Accepted:
            self.load_data()

    def edit_selected(self):
        grade = self._selected_grade()
        if grade is None:
            QMessageBox.information(self, "提示", "请先选中一行成绩")
            return
        dialog = GradeEditDialog(self.db, parent=self, grade=grade, student=self.student)
        if dialog.exec() == QDialog.Accepted:
            self.load_data()

    def delete_selected(self):
        grade = self._selected_grade()
        if grade is None:
            QMessageBox.information(self, "提示", "请先选中一行成绩")
            return
        reply = QMessageBox.question(
            self, "确认删除",
            f"确定删除「{grade.course_name} {grade.score:.1f} 分（{grade.exam_date}）」吗？",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No)
        if reply == QMessageBox.Yes:
            self.db.delete_grade(grade.id)
            self.load_data()


# ==================== 主窗口 ====================
class MainWindow(QMainWindow):
    """主窗口"""

    def __init__(self):
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} v{APP_VERSION}")
        self.db = DatabaseManager()

        # 居中显示（用可用区域，避开 macOS 菜单栏/程序坞与 Windows 任务栏）
        win_w, win_h = 1100, 700
        screen_obj = QGuiApplication.primaryScreen()
        if screen_obj is not None:
            avail = screen_obj.availableGeometry()
            win_w = max(900, min(win_w, avail.width() - 40))
            win_h = max(600, min(win_h, avail.height() - 40))
            self.setGeometry(
                avail.x() + (avail.width() - win_w) // 2,
                avail.y() + (avail.height() - win_h) // 2,
                win_w, win_h
            )
        self.setMinimumSize(min(900, win_w), min(600, win_h))

        self.setup_ui()
        self.apply_styles()
        self.refresh_all()

    def setup_ui(self):
        # 中央部件
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QHBoxLayout(central)
        main_layout.setContentsMargins(0, 0, 0, 0)
        main_layout.setSpacing(0)

        # 侧边栏
        self.sidebar = QFrame()
        self.sidebar.setFixedWidth(200)
        self.sidebar.setStyleSheet(f"""
            QFrame {{
                background-color: {COLORS['card_bg']};
                border-right: 1px solid {COLORS['border']};
            }}
        """)
        sidebar_layout = QVBoxLayout(self.sidebar)
        sidebar_layout.setContentsMargins(12, 20, 12, 20)
        sidebar_layout.setSpacing(6)

        # Logo
        logo = QLabel("🎓 教学管理")
        logo.setStyleSheet(f"font-size: 20px; font-weight: bold; color: {COLORS['primary']}; padding: 10px 5px;")
        logo.setAlignment(Qt.AlignCenter)
        sidebar_layout.addWidget(logo)

        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setStyleSheet(f"background-color: {COLORS['border']}; max-height: 1px;")
        sidebar_layout.addWidget(separator)
        sidebar_layout.addSpacing(10)

        # 导航按钮
        self.nav_buttons = []
        nav_items = [
            ("👥", "学生管理", 0),
            ("📅", "课程管理", 1),
            ("💬", "反馈系统", 2),
            ("📊", "成绩管理", 3),
            ("📋", "学情管理", 4),
        ]

        for icon, text, index in nav_items:
            btn = QPushButton(f"{icon}  {text}")
            btn.setCheckable(True)
            btn.setFixedHeight(44)
            btn.clicked.connect(lambda checked, i=index: self.switch_page(i))
            self.nav_buttons.append(btn)
            sidebar_layout.addWidget(btn)

        sidebar_layout.addStretch()

        # 版本信息
        version_label = QLabel(f"v{APP_VERSION}")
        version_label.setAlignment(Qt.AlignCenter)
        version_label.setStyleSheet(f"font-size: 11px; color: {COLORS['text_secondary']};")
        sidebar_layout.addWidget(version_label)

        main_layout.addWidget(self.sidebar)

        # 内容区域
        self.content_stack = QStackedWidget()
        main_layout.addWidget(self.content_stack, stretch=1)

        # 创建模块页面
        self.student_module = StudentModule(self.db)
        self.course_module = CourseModule(self.db)
        self.feedback_module = FeedbackModule(self.db)
        self.grade_module = GradeModule(self.db)
        self.profile_module = StudentProfileModule(self.db)

        self.content_stack.addWidget(self.student_module)
        self.content_stack.addWidget(self.course_module)
        self.content_stack.addWidget(self.feedback_module)
        self.content_stack.addWidget(self.grade_module)
        self.content_stack.addWidget(self.profile_module)

        # 默认选中第一个
        self.nav_buttons[0].setChecked(True)
        self.switch_page(0)

        # 状态栏
        self.statusBar().showMessage("就绪")

        # 菜单栏
        self.setup_menus()

    def setup_menus(self):
        menubar = self.menuBar()
        # macOS 会把菜单栏移到屏幕顶部全局菜单，需正确标注菜单角色
        try:
            NoRole = QAction.MenuRole.NoRole
            QuitRole = QAction.MenuRole.QuitRole
            AboutRole = QAction.MenuRole.AboutRole
        except AttributeError:   # 兼容旧版枚举写法
            NoRole, QuitRole, AboutRole = QAction.NoRole, QAction.QuitRole, QAction.AboutRole

        file_menu = menubar.addMenu("文件")
        import_students_action = QAction("导入学生数据", self)
        import_students_action.setMenuRole(NoRole)
        import_students_action.triggered.connect(self.student_module.import_students)
        file_menu.addAction(import_students_action)

        import_grades_action = QAction("导入成绩数据", self)
        import_grades_action.setMenuRole(NoRole)
        import_grades_action.triggered.connect(self.grade_module.import_grades)
        file_menu.addAction(import_grades_action)

        file_menu.addSeparator()
        open_dir_action = QAction("打开数据目录", self)
        open_dir_action.setMenuRole(NoRole)
        open_dir_action.triggered.connect(self.open_data_dir)
        file_menu.addAction(open_dir_action)

        file_menu.addSeparator()
        exit_action = QAction("退出", self)
        exit_action.setMenuRole(QuitRole)          # macOS: 归入应用菜单「退出」
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        help_menu = menubar.addMenu("帮助")
        about_action = QAction("关于", self)
        about_action.setMenuRole(AboutRole)        # macOS: 归入应用菜单「关于」
        about_action.triggered.connect(self.show_about)
        help_menu.addAction(about_action)

    def open_data_dir(self):
        """在系统文件管理器中打开数据目录（macOS=Finder, Windows=资源管理器）"""
        try:
            path = str(DATA_DIR)
            if not QDesktopServices.openUrl(QUrl.fromLocalFile(path)):
                raise RuntimeError("系统未接受打开请求")
        except Exception as e:
            QMessageBox.information(self, "数据目录", f"数据目录：\n{DATA_DIR}\n\n（打开失败：{e}）")

    def switch_page(self, index: int):
        """切换模块页面"""
        self.content_stack.setCurrentIndex(index)

        # 更新按钮状态
        for i, btn in enumerate(self.nav_buttons):
            btn.setChecked(i == index)

        # 刷新对应模块
        if index == 0:
            self.student_module.refresh()
        elif index == 1:
            self.course_module.refresh()
        elif index == 2:
            self.feedback_module.refresh()
        elif index == 3:
            self.grade_module.refresh()
        elif index == 4:
            self.profile_module.refresh()

        # 更新状态栏
        module_names = ["学生管理系统", "课程管理系统", "反馈系统", "成绩管理系统", "学情管理"]
        self.statusBar().showMessage(f"当前模块：{module_names[index]}")

    def refresh_all(self):
        """刷新所有模块"""
        self.student_module.refresh()
        self.course_module.refresh()
        self.feedback_module.refresh()
        self.grade_module.refresh()
        self.profile_module.refresh()

    def apply_styles(self):
        """应用全局样式"""
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {COLORS['background']};
                font-family: {FONT_STACK_CSS};
            }}
            QWidget {{
                font-family: {FONT_STACK_CSS};
                font-size: 13px;
            }}
            QLabel {{
                color: {COLORS['text_primary']};
            }}
            QPushButton {{
                font-family: {FONT_STACK_CSS};
            }}
            QComboBox, QLineEdit, QTextEdit, QDateEdit, QTimeEdit, QSpinBox, QDoubleSpinBox {{
                padding: 6px 10px;
                border: 1px solid {COLORS['border']};
                border-radius: 6px;
                font-size: 13px;
                font-family: {FONT_STACK_CSS};
                background-color: {COLORS['card_bg']};
            }}
            QComboBox:focus, QLineEdit:focus, QTextEdit:focus {{
                border-color: {COLORS['primary']};
            }}
            QScrollArea {{
                border: none;
            }}
            QMenuBar {{
                background-color: {COLORS['card_bg']};
                border-bottom: 1px solid {COLORS['border']};
                font-family: {FONT_STACK_CSS};
            }}
            QMenuBar::item {{
                padding: 6px 12px;
                background: transparent;
            }}
            QMenuBar::item:selected {{
                background-color: {COLORS['hover']};
                border-radius: 4px;
            }}
            QMenu {{
                background-color: {COLORS['card_bg']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                font-family: {FONT_STACK_CSS};
            }}
            QMenu::item {{
                padding: 6px 20px;
            }}
            QMenu::item:selected {{
                background-color: {COLORS['hover']};
            }}
            QStatusBar {{
                background-color: {COLORS['card_bg']};
                border-top: 1px solid {COLORS['border']};
                font-size: 12px;
                font-family: {FONT_STACK_CSS};
            }}
        """)

        # 导航按钮样式
        for btn in self.nav_buttons:
            btn.setStyleSheet(f"""
                QPushButton {{
                    background-color: transparent;
                    border: none;
                    border-radius: 10px;
                    padding: 10px 15px;
                    text-align: left;
                    font-size: 14px;
                    color: {COLORS['text_secondary']};
                    font-family: {FONT_STACK_CSS};
                    font-weight: 500;
                }}
                QPushButton:hover {{
                    background-color: {COLORS['hover']};
                    color: {COLORS['text_primary']};
                }}
                QPushButton:checked {{
                    background-color: {COLORS['primary_light']};
                    color: {COLORS['primary']};
                    font-weight: bold;
                    border-left: 4px solid {COLORS['primary']};
                }}
            """)

    def show_about(self):
        platform_name = "macOS" if IS_MACOS else ("Windows" if IS_WINDOWS else "Linux")
        QMessageBox.about(
            self,
            "关于",
            f"<h2>{APP_NAME}</h2>"
            f"<p>版本：{APP_VERSION}</p>"
            f"<p>功能：学生管理、课程管理、反馈系统、成绩管理</p>"
            f"<p>技术栈：Python + PySide6 + SQLite</p>"
            f"<p>运行平台：{platform_name}（跨平台版）</p>"
            f"<p>数据目录：{DATA_DIR}<br/>（{DATA_DIR_MODE}）</p>"
        )

    def closeEvent(self, event: QCloseEvent):
        """关闭事件处理"""
        self.db.close()
        event.accept()


# ==================== 主程序入口 ====================
def main():
    """主函数"""
    if not HAS_PYSIDE6:
        print("错误：缺少PySide6依赖。请运行: pip install PySide6")
        sys.exit(1)

    # Retina / 高DPI 适配（必须在 QApplication 创建前设置）
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    except Exception:
        pass
    if IS_MACOS:
        os.environ.setdefault("QT_MAC_WANTS_LAYER", "1")

    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    if IS_MACOS:
        # 仅 macOS 需要显示名（Windows 上设置会让窗口标题重复成“X - X”）
        app.setApplicationDisplayName(APP_NAME)
    app.setApplicationVersion(APP_VERSION)
    app.setOrganizationName("TeachingManager")
    app.setOrganizationDomain(APP_BUNDLE_ID)
    app.setQuitOnLastWindowClosed(True)         # macOS 关窗即退出（单窗口应用）

    # 应用图标：打包时随包提供 icon.png / .icns
    icon_path = Path(getattr(sys, '_MEIPASS', get_app_dir())) / "icon.png"
    if icon_path.exists():
        app.setWindowIcon(QIcon(str(icon_path)))

    # 字体：优先 MiSans（若已安装或放入 fonts/ 目录），否则 PingFang SC 等兜底
    chosen_font = apply_ui_font(app)
    print(f"[字体] 界面字体：{chosen_font}（候选顺序：{' → '.join(FONT_STACK[:4])}）")
    print(f"[数据目录] {DATA_DIR}（{DATA_DIR_MODE}）")

    ensure_images_dir()
    window = MainWindow()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()