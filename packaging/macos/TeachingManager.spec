# -*- mode: python ; coding: utf-8 -*-
"""
PyInstaller 规格文件 —— 构建 macOS 应用包 教学管理系统.app
================================================================
必须在一台 macOS 机器上执行（PyInstaller 不支持交叉编译）：

    cd packaging/macos
    python3 build_macos.py            # 推荐：自动装依赖、生成图标、打包、产出 DMG
    或手动：
    TMS_SOURCE=/path/to/deepseek_python_xxx.py pyinstaller --noconfirm --clean TeachingManager.spec

可用环境变量：
    TMS_SOURCE        源脚本路径（默认自动向上查找 teaching_manager.py）
    TMS_VERSION       版本号（默认 1.2.0）
    TMS_TARGET_ARCH   arm64 / x86_64 / universal2（默认本机架构）
    TMS_APP_NAME      应用名（默认 教学管理系统）
"""
import os
import sys
from pathlib import Path

if sys.platform != "darwin":
    raise SystemExit(
        "\n[TeachingManager.spec] 该规格文件仅用于 macOS 打包。\n"
        "PyInstaller 不支持交叉编译，DMG 也需要 macOS 的 hdiutil。\n"
        "请在 Mac 上运行 python3 build_macos.py；Windows 打包请使用 Windows 专用命令。\n")

SPEC_DIR = Path(SPECPATH).resolve()  # noqa: F821  (PyInstaller 注入)
ASSETS = SPEC_DIR / "assets"

APP_NAME = os.environ.get("TMS_APP_NAME", "教学管理系统")
BUNDLE_ID = os.environ.get("TMS_BUNDLE_ID", "com.teaching.manager")
VERSION = os.environ.get("TMS_VERSION", "1.2.0")
TARGET_ARCH = os.environ.get("TMS_TARGET_ARCH") or None


def find_source() -> Path:
    """定位主程序源文件：teaching_manager.py（兼容旧的 deepseek_python_*.py）"""
    env = os.environ.get("TMS_SOURCE")
    if env:
        p = Path(env).expanduser().resolve()
        if p.exists():
            return p
        raise SystemExit(f"TMS_SOURCE 指向的文件不存在: {p}")
    for base in [SPEC_DIR, *SPEC_DIR.parents]:
        for name in ("teaching_manager.py", "app.py"):
            cand = base / name
            if cand.exists():
                return cand.resolve()
        hits = sorted(base.glob("deepseek_python_*.py"))
        if hits:
            return hits[0].resolve()
    raise SystemExit("未找到源脚本 teaching_manager.py，请用 TMS_SOURCE 环境变量指定其路径")


SRC = find_source()

icon_png = ASSETS / "icon.png"
icon_icns = ASSETS / "icon.icns"
icon_arg = str(icon_icns) if icon_icns.exists() else (str(icon_png) if icon_png.exists() else None)

datas = []
if icon_png.exists():
    datas.append((str(icon_png), "."))       # 运行时窗口/Dock 图标

# macOS 上完全用不到的 Qt 模块，排除以显著减小体积
excludes = [
    "tkinter", "unittest",
    "PySide6.QtWebEngineCore", "PySide6.QtWebEngineWidgets", "PySide6.QtWebEngineQuick",
    "PySide6.QtQml", "PySide6.QtQuick", "PySide6.QtQuick3D", "PySide6.QtQuickWidgets",
    "PySide6.Qt3DCore", "PySide6.Qt3DRender", "PySide6.Qt3DAnimation", "PySide6.Qt3DExtras",
    "PySide6.QtMultimedia", "PySide6.QtMultimediaWidgets",
    "PySide6.QtBluetooth", "PySide6.QtNfc", "PySide6.QtPositioning",
    "PySide6.QtSensors", "PySide6.QtSerialPort", "PySide6.QtWebSockets",
    "PySide6.QtDesigner", "PySide6.QtHelp", "PySide6.QtTest", "PySide6.QtSql",
    "PySide6.QtPdf", "PySide6.QtPdfWidgets", "PySide6.QtSpatialAudio",
]

a = Analysis(  # noqa: F821
    [str(SRC)],
    pathex=[str(SRC.parent)],
    binaries=[],
    datas=datas,
    hiddenimports=[],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=excludes,
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)  # noqa: F821

exe = EXE(  # noqa: F821
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name=APP_NAME,
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,                 # GUI 程序，不弹终端窗口
    disable_windowed_traceback=False,
    argv_emulation=False,          # 我们不处理拖拽到图标打开文件
    target_arch=TARGET_ARCH,
    codesign_identity=os.environ.get("TMS_CODESIGN_IDENTITY") or None,
    entitlements_file=os.environ.get("TMS_ENTITLEMENTS") or None,
    icon=icon_arg,
)

coll = COLLECT(  # noqa: F821
    exe,
    a.binaries,
    a.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name=APP_NAME,
)

app = BUNDLE(  # noqa: F821
    coll,
    name=f"{APP_NAME}.app",
    icon=icon_arg,
    bundle_identifier=BUNDLE_ID,
    version=VERSION,
    info_plist={
        "CFBundleName": APP_NAME,
        "CFBundleDisplayName": APP_NAME,
        "CFBundleExecutable": APP_NAME,
        "CFBundleIdentifier": BUNDLE_ID,
        "CFBundleShortVersionString": VERSION,
        "CFBundleVersion": VERSION,
        "CFBundlePackageType": "APPL",
        "CFBundleSignature": "????",
        "NSHumanReadableCopyright": "教学管理系统",
        "NSPrincipalClass": "NSApplication",
        # Retina：不声明会在高分屏上模糊
        "NSHighResolutionCapable": True,
        # macOS 11+ 起支持（Apple Silicon / Intel 通用）
        "LSMinimumSystemVersion": "11.0",
        "LSApplicationCategoryType": "public.app-category.education",
        # 中文界面
        "CFBundleDevelopmentRegion": "zh_CN",
        "NSRequiresAquaSystemAppearance": False,
    },
)
