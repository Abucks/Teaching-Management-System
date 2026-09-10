#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
macOS 一键打包脚本 —— 生成 教学管理系统.app 与 教学管理系统-<版本>-macOS.dmg
==========================================================================
⚠️ 必须在 macOS 上运行（PyInstaller 无法交叉编译；DMG 需要 macOS 的 hdiutil）。

用法（在 Mac 的终端里）：
    cd <项目目录>/packaging/macos
    python3 build_macos.py                 # 全流程：依赖 → 图标 → .app → .dmg
    python3 build_macos.py --dry-run       # 只打印将要执行的命令（任何系统都可跑）
    python3 build_macos.py --no-dmg        # 只产出 .app
    python3 build_macos.py --arch universal2
    python3 build_macos.py --sign "Developer ID Application: XXX (TEAMID)"
    python3 build_macos.py --skip-deps     # 复用已有 venv

常用参数：
    --source PATH   主程序源文件（默认自动查找 teaching_manager.py）
    --version STR   版本号（默认读源码中的 APP_VERSION）
    --venv PATH     虚拟环境目录（默认 ./.venv-mac）
"""
import argparse
import os
import platform
import re
import shutil
import subprocess
import sys
from pathlib import Path

SPEC_DIR = Path(__file__).resolve().parent
ASSETS = SPEC_DIR / "assets"
DIST = SPEC_DIR / "dist"
BUILD = SPEC_DIR / "build"
APP_NAME = "教学管理系统"

REQUIRED_PKGS = ["PySide6", "matplotlib", "openpyxl", "python-docx", "xlrd", "pyinstaller"]


# --------------------------------------------------------------------------- #
# 工具函数
# --------------------------------------------------------------------------- #
def log(msg: str):
    print(f"\033[1;34m==>\033[0m {msg}" if sys.stdout.isatty() else f"==> {msg}")


def warn(msg: str):
    print(f"\033[1;33m[警告]\033[0m {msg}", file=sys.stderr)


def die(msg: str, code: int = 1):
    print(f"\033[1;31m[错误]\033[0m {msg}", file=sys.stderr)
    sys.exit(code)


def run(cmd, dry_run=False, env=None, check=True, cwd=None):
    printable = " ".join(str(c) for c in cmd)
    print(f"    $ {printable}")
    if dry_run:
        return 0
    proc = subprocess.run([str(c) for c in cmd], env=env, cwd=cwd)
    if check and proc.returncode != 0:
        die(f"命令执行失败（退出码 {proc.returncode}）: {printable}")
    return proc.returncode


def find_source(explicit: str | None) -> Path:
    """定位主程序源文件：优先 teaching_manager.py，兼容 deepseek_python_*.py"""
    if explicit:
        p = Path(explicit).expanduser().resolve()
        if not p.exists():
            die(f"--source 指定的文件不存在: {p}")
        return p
    for base in [SPEC_DIR, *SPEC_DIR.parents]:
        for name in ("teaching_manager.py", "app.py"):
            cand = base / name
            if cand.exists():
                return cand.resolve()
        hits = sorted(base.glob("deepseek_python_*.py"))
        if hits:
            return hits[0].resolve()
    die("未找到主程序源文件，请用 --source 指定 teaching_manager.py 的路径")


def detect_version(src: Path) -> str:
    try:
        m = re.search(r'APP_VERSION\s*=\s*"([^"]+)"', src.read_text(encoding="utf-8"))
        if m:
            return m.group(1)
    except Exception:
        pass
    return "1.2.0"


def host_arch() -> str:
    machine = platform.machine().lower()
    return "arm64" if machine in ("arm64", "aarch64") else "x86_64"


# --------------------------------------------------------------------------- #
# 各步骤
# --------------------------------------------------------------------------- #
def step_env(py: str, venv: Path, skip_deps: bool, dry_run: bool):
    if skip_deps and venv.exists():
        log(f"跳过依赖安装（复用 {venv}）")
        return venv / "bin" / "python"
    log(f"创建虚拟环境: {venv}")
    run([py, "-m", "venv", str(venv)], dry_run=dry_run)
    vpy = venv / "bin" / "python"
    log("安装依赖（PySide6 / matplotlib / openpyxl / python-docx / xlrd / pyinstaller）")
    run([vpy, "-m", "pip", "install", "--upgrade", "pip"], dry_run=dry_run)
    run([vpy, "-m", "pip", "install", *REQUIRED_PKGS], dry_run=dry_run)
    return vpy


def step_icon(vpy: Path, dry_run: bool):
    log("生成应用图标 icon.png（1024×1024）")
    run([vpy, str(ASSETS / "make_icon.py")], dry_run=dry_run)

    png = ASSETS / "icon.png"
    icns = ASSETS / "icon.icns"
    log("将 PNG 转换为 .icns（sips + iconutil）")
    if dry_run:
        print("    $ sips -z <size> ... icon.png --out icon.iconset/... ; iconutil -c icns icon.iconset")
        return icns
    if not png.exists():
        warn("icon.png 缺失，跳过 .icns 生成（应用将使用默认图标）")
        return None
    iconset = ASSETS / "icon.iconset"
    if iconset.exists():
        shutil.rmtree(iconset)
    iconset.mkdir(parents=True)

    # macOS 图标规范要求的尺寸组合
    sizes = [(16, "16x16", 1), (32, "16x16", 2), (32, "32x32", 1), (64, "32x32", 2),
             (128, "128x128", 1), (256, "128x128", 2), (256, "256x256", 1),
             (512, "256x256", 2), (512, "512x512", 1), (1024, "512x512", 2)]
    for px, name, scale in sizes:
        suffix = "" if scale == 1 else "@2x"
        out = iconset / f"icon_{name}{suffix}.png"
        run(["sips", "-z", str(px), str(px), str(png), "--out", str(out)], dry_run=dry_run)
    run(["iconutil", "-c", "icns", str(iconset), "-o", str(icns)], dry_run=dry_run)
    if not dry_run and icns.exists():
        print(f"    已生成 {icns.name} ({icns.stat().st_size // 1024} KB)")
    return icns


def step_build(vpy: Path, src: Path, version: str, arch: str | None, sign: str | None, dry_run: bool):
    log(f"PyInstaller 打包 .app（版本 {version}，架构 {arch or '本机'}）")
    env = os.environ.copy()
    env["TMS_SOURCE"] = str(src)
    env["TMS_VERSION"] = version
    if arch:
        env["TMS_TARGET_ARCH"] = arch
    if sign:
        env["TMS_CODESIGN_IDENTITY"] = sign
    # 注意：使用 .spec 文件时不能再传 --specpath（PyInstaller 会报
    # "makespec options not valid when a .spec file is given"）
    run([vpy, "-m", "PyInstaller", "--noconfirm", "--clean",
         "--distpath", str(DIST), "--workpath", str(BUILD),
         str(SPEC_DIR / "TeachingManager.spec")],
        dry_run=dry_run, env=env)
    return DIST / f"{APP_NAME}.app"


def step_sign(app_path: Path, sign: str | None, dry_run: bool):
    if not sign:
        log("未提供 --sign，跳过代码签名（未签名应用首次打开需右键→打开）")
        return
    log(f"代码签名: {sign}")
    run(["codesign", "--force", "--deep", "--options", "runtime",
         "--timestamp", "--sign", sign, str(app_path)], dry_run=dry_run)
    run(["codesign", "--verify", "--deep", "--strict", "--verbose=2", str(app_path)],
        dry_run=dry_run, check=False)


def step_dmg(app_path: Path, version: str, sign: str | None, dry_run: bool):
    dmg = SPEC_DIR / f"教学管理系统-{version}-macOS.dmg"
    staging = BUILD / "dmg-staging"
    log(f"制作 DMG 安装包: {dmg.name}")
    if dry_run:
        print(f"    $ rm -rf {staging} && mkdir -p {staging}")
        print(f"    $ cp -R {app_path} {staging}/")
        print(f"    $ ln -s /Applications {staging}/Applications")
        print(f"    $ hdiutil create -volname {APP_NAME} -srcfolder {staging} -ov -format UDZO {dmg}")
        if sign:
            print(f"    $ codesign --sign {sign} {dmg}")
            print(f"    $ xcrun notarytool submit {dmg} --keychain-profile <profile> --wait")
            print(f"    $ xcrun stapler staple {dmg}")
        return dmg

    if staging.exists():
        shutil.rmtree(staging)
    staging.mkdir(parents=True)
    run(["cp", "-R", str(app_path), str(staging / app_path.name)])
    # 拖拽安装用的 Applications 快捷方式
    run(["ln", "-s", "/Applications", str(staging / "Applications")])
    if dmg.exists():
        dmg.unlink()
    run(["hdiutil", "create", "-volname", APP_NAME, "-srcfolder", str(staging),
         "-ov", "-format", "UDZO", str(dmg)])
    if sign:
        run(["codesign", "--sign", sign, "--timestamp", str(dmg)], check=False)
        warn("如需通过 Gatekeeper，请再执行公证：")
        warn(f"  xcrun notarytool submit '{dmg}' --apple-id <你的AppleID> --team-id <TEAMID> "
             f"--password <App专用密码> --wait")
        warn(f"  xcrun stapler staple '{dmg}'")
    return dmg


# --------------------------------------------------------------------------- #
# 主流程
# --------------------------------------------------------------------------- #
def main():
    ap = argparse.ArgumentParser(description="macOS 打包脚本（生成 .app 与 .dmg）")
    ap.add_argument("--source", help="主程序源文件路径")
    ap.add_argument("--version", help="版本号（默认读源码 APP_VERSION）")
    ap.add_argument("--arch", choices=["arm64", "x86_64", "universal2"], help="目标架构")
    ap.add_argument("--sign", help='代码签名身份，如 "Developer ID Application: X (TEAM)"')
    ap.add_argument("--venv", default=str(SPEC_DIR / ".venv-mac"), help="虚拟环境目录")
    ap.add_argument("--skip-deps", action="store_true", help="跳过依赖安装")
    ap.add_argument("--no-dmg", action="store_true", help="只生成 .app，不制作 DMG")
    ap.add_argument("--keep-build", action="store_true", help="保留中间构建目录")
    ap.add_argument("--dry-run", action="store_true", help="只打印命令不执行")
    args = ap.parse_args()

    src = find_source(args.source)
    version = args.version or detect_version(src)
    dry = args.dry_run

    print("=" * 68)
    print(f" macOS 打包：{APP_NAME} v{version}")
    print(f" 源文件   : {src}")
    print(f" 平台     : {platform.system()} {platform.release()} ({platform.machine()})")
    print(f" 目标架构 : {args.arch or host_arch() + '（本机）'}")
    print(f" 模式     : {'DRY-RUN（仅打印命令）' if dry else '实际构建'}")
    print("=" * 68)

    if not dry and platform.system() != "Darwin":
        die("本脚本必须在 macOS 上运行：PyInstaller 不支持交叉编译，DMG 也需要 macOS 的 hdiutil。\n"
            "    请在一台 Mac 上执行，或使用仓库内的 GitHub Actions 工作流（macos runner）自动产出 DMG。\n"
            "    Windows 端可先用 python build_macos.py --dry-run 检查流程。")
    if not dry and shutil.which("hdiutil") is None and not args.no_dmg:
        die("未找到 hdiutil（DMG 制作工具），请确认在 macOS 上运行，或加 --no-dmg 只生成 .app")

    # 1) 依赖环境
    py = sys.executable if args.skip_deps else None
    if args.skip_deps:
        log(f"跳过依赖安装，使用当前解释器: {py}")
    else:
        py = step_env(sys.executable, Path(args.venv), args.skip_deps, dry)
    py = str(py)

    # 2) 图标
    step_icon(Path(py), dry)

    # 3) 打包 .app
    app_path = step_build(Path(py), src, version, args.arch, args.sign, dry)

    # 4) 签名
    step_sign(app_path, args.sign, dry)

    # 5) DMG
    dmg = None
    if not args.no_dmg:
        dmg = step_dmg(app_path, version, args.sign, dry)

    # 6) 清理
    if not (args.keep_build or dry):
        for junk in [SPEC_DIR / "build" / "dmg-staging"]:
            if junk.exists():
                shutil.rmtree(junk, ignore_errors=True)

    print("=" * 68)
    print(" 构建完成 🎉")
    print(f"  应用包 : {app_path}")
    if dmg:
        print(f"  安装包 : {dmg}")
    print("  安装方式：打开 DMG，把「教学管理系统」拖入 Applications 文件夹。")
    print("  首次运行若被 Gatekeeper 拦截：右键应用 → 打开；或执行")
    print("      xattr -dr com.apple.quarantine '/Applications/教学管理系统.app'")
    print("=" * 68)
    return 0


if __name__ == "__main__":
    sys.exit(main())
