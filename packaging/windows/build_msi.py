#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""用 Python 标准库 msilib 生成 Windows MSI 安装包

为什么不用 WiX：`dotnet tool install wix` 现在会装到 v7，而 v7 要求接受
Open Source Maintenance Fee (OSMF) EULA，CI 上直接报 WIX7015。msilib 是
Windows Installer API 的标准库封装，零外部依赖、无授权问题。

用法:
    python build_msi.py <TeachingManager.exe> <输出.msi> <版本> [图标.ico]

限制：仅 Windows + Python ≤ 3.12（msilib 在 Python 3.13 被移除）。
特性：安装到 Program Files、创建开始菜单/桌面快捷方式、注册卸载信息、
      支持主版本升级（新版本自动移除旧版本）、阻止降级安装。
"""
import os
import sys
import uuid

import msilib
from msilib import (CAB, Binary, Directory, Feature, add_data, init_database, schema,
                    sequence)

PRODUCT_NAME = "教学管理系统"
MANUFACTURER = "Asell Bucks"
UPGRADE_CODE = "{B2C4A1F6-9D3E-4A58-8F7C-1E6D4B0A92F3}"
APP_EXE = "TeachingManager.exe"
INSTALL_DIR = "TeachingManager"
MENU_DIR = "教学管理系统"
SHORTCUT_NAME = "教学管理系统"


def _stable_guid(key: str) -> str:
    """按名称生成稳定 GUID（同一次构建/升级之间保持一致）"""
    return "{" + str(uuid.uuid5(uuid.NAMESPACE_DNS, "teaching-manager/" + key)).upper() + "}"


class _DirNode:
    """给 msilib.Directory 当父节点用的轻量替身
    （3.12 的 msilib.Directory 没有 add_subdir，必须自己传父对象）"""

    def __init__(self, logical: str, absolute: str = ""):
        self.logical = logical
        self.absolute = absolute


def build(exe_path: str, out_msi: str, version: str, icon_path: str | None = None) -> str:
    exe_path = os.path.abspath(exe_path)
    out_msi = os.path.abspath(out_msi)
    if not os.path.isfile(exe_path):
        raise SystemExit(f"未找到源 exe：{exe_path}")
    os.makedirs(os.path.dirname(out_msi) or ".", exist_ok=True)

    product_code = _stable_guid("product-" + version)

    db = init_database(out_msi, schema, PRODUCT_NAME, product_code, version, MANUFACTURER)
    msilib.add_tables(db, sequence)

    # ---------------- 属性（卸载信息 / 安装上下文）----------------
    props = [
        ("AllUsers", "1"),
        ("ARPNOMODIFY", "1"),
        ("ARPNOREPAIR", "1"),
        ("ARPCOMMENTS", "本地离线的教学管理系统（学生/排课/便签/成绩/学情）"),
        ("ARPCONTACT", "https://github.com/Abucks/Teaching-Management-System"),
    ]
    icon_ok = False
    if icon_path and os.path.isfile(icon_path):
        # Icon 表的 Data 是 OBJECT(二进制流) 列：先 add_stream 写入流，再用 Binary(流名) 引用。
        # 该写入在部分 MSI 实现下会失败（error 1101），失败时跳过即可：
        # 快捷方式会自动继承 exe 内嵌的图标，不影响安装与使用。
        try:
            msilib.add_stream(db, "ProductIcon.ico", os.path.abspath(icon_path))
            add_data(db, "Icon", [("ProductIcon", Binary("ProductIcon.ico"))])
            props.append(("ARPPRODUCTICON", "ProductIcon"))
            icon_ok = True
        except Exception as exc:      # noqa: BLE001
            print(f"[warn] 写入 MSI 图标失败（已跳过）：{exc}", file=sys.stderr)
    add_data(db, "Property", props)

    # ---------------- 目录树 ----------------
    # Directory 表必须是一棵完整可链接的树（否则安装时报 2705 Invalid table: Directory），
    # 标准目录（TARGETDIR / ProgramFiles64Folder / ProgramMenuFolder / DesktopFolder）
    # 都要显式建行；根目录用 physical="" 使 absolute 为字符串，便于子目录拼接。
    cab = CAB("product.cab")
    root = Directory(db, cab, None, "", "TARGETDIR", "SourceDir")
    program_files = Directory(db, cab, root, ".", "ProgramFiles64Folder", "PFiles")
    app_dir = Directory(db, cab, program_files, ".", "INSTALLDIR", INSTALL_DIR)
    menu_dir = Directory(db, cab,
                         Directory(db, cab, root, ".", "ProgramMenuFolder", "."),
                         ".", "ProgramMenuDir", MENU_DIR)
    desktop_dir = Directory(db, cab, root, ".", "DesktopFolder", ".")

    # ---------------- 功能与组件 ----------------
    feature = Feature(db, "DefaultFeature", PRODUCT_NAME, "教学管理系统主程序", 1, 1,
                      None, "INSTALLDIR")

    # 只建一个组件（拥有 exe，KeyPath=exe）；快捷方式挂到同一组件上，
    # 避免"只有快捷方式的组件没有 KeyPath"导致的安装失败
    app_dir.start_component(component="MainExecutable", feature=feature,
                            flags=0, uuid=_stable_guid("component/main-exe"))
    app_dir.add_file(APP_EXE, src=exe_path)
    _ = (menu_dir, desktop_dir)          # 目录行已建立，供下面的快捷方式引用

    # ---------------- 快捷方式（挂在 exe 组件上，保证有 KeyPath）----------------
    icon_fields = ("ProductIcon", 0) if icon_ok else (None, None)
    add_data(db, "Shortcut", [
        ("StartMenuShortcut", "ProgramMenuDir", SHORTCUT_NAME, "MainExecutable",
         f"[INSTALLDIR]{APP_EXE}", None, "教学管理系统", None, icon_fields[0], icon_fields[1], 1, None),
        ("DesktopShortcut", "DesktopFolder", SHORTCUT_NAME, "MainExecutable",
         f"[INSTALLDIR]{APP_EXE}", None, "教学管理系统", None, icon_fields[0], icon_fields[1], 1, None),
    ])

    # ---------------- 主版本升级（新版覆盖旧版、阻止降级）----------------
    # 注意：msilib 的标准 InstallExecuteSequence 已包含
    #   FindRelatedProducts(200) 与 RemoveExistingProducts(6700)，
    # 因此这里只需提供 Upgrade 表，无需再插入动作（重复插入会因主键冲突失败）
    add_data(db, "Upgrade", [
        # 已安装 < 当前版本 → 记录到 OLDVERSIONFOUND，稍后由 RemoveExistingProducts 移除
        (UPGRADE_CODE, None, version, None, 0, None, "OLDVERSIONFOUND"),
        # 已安装 >= 当前版本 → 仅检测（阻止降级）
        (UPGRADE_CODE, version, None, None, 3, None, "NEWERVERSIONFOUND"),
    ])

    # 阻止降级：若检测到更新版本则终止安装
    add_data(db, "LaunchCondition", [
        ("NOT NEWERVERSIONFOUND", "已安装更高版本的 教学管理系统，无需重复安装。"),
    ])

    cab.commit(db)
    db.Commit()
    return out_msi


def main(argv: list[str]) -> int:
    # CI（Windows Runner）控制台默认是 cp1252，直接打印中文会抛 UnicodeEncodeError
    # 并让构建"看起来失败"；这里统一把标准输出改成 UTF-8 容错模式。
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")   # type: ignore[attr-defined]
        except Exception:
            pass

    if len(argv) < 4:
        print(__doc__)
        return 2
    exe, out, version = argv[1], argv[2], argv[3]
    icon = argv[4] if len(argv) > 4 else None
    if sys.platform != "win32":
        print("错误：MSI 只能在 Windows 上构建", file=sys.stderr)
        return 1
    path = build(exe, out, version, icon)
    size = os.path.getsize(path) / 1024 / 1024
    print(f"[ok] 已生成 MSI：{path}  ({size:.1f} MB)")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
