#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成应用图标 icon.png（1024×1024，跨平台，无需外部图片素材）。

在 macOS 打包流程中由 build_macos.py 自动调用，随后用 sips/iconutil 转成 .icns。
Windows 打包时同一张 PNG 作为窗口/任务栏图标一并打入程序。
"""
import os
import sys
from pathlib import Path

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

from PySide6.QtCore import QRectF, Qt
from PySide6.QtGui import (QColor, QFont, QGuiApplication, QImage, QLinearGradient,
                           QPainter, QPainterPath, QPen)

SIZE = 1024
OUT = Path(__file__).resolve().parent / "icon.png"
OUT_ICO = Path(__file__).resolve().parent / "icon.ico"


def pick_cjk_font() -> str:
    """选一个本机存在的中文字体用于绘制图标文字"""
    if sys.platform == "darwin":
        candidates = ["PingFang SC", "Hiragino Sans GB", "Heiti SC", "STHeiti"]
    elif sys.platform.startswith("win"):
        candidates = ["Microsoft YaHei", "SimHei", "Segoe UI"]
    else:
        candidates = ["Noto Sans CJK SC", "Source Han Sans SC", "WenQuanYi Micro Hei"]
    # 由 Qt 自行回退，这里只返回首选名
    return candidates[0]


def draw_icon() -> QImage:
    img = QImage(SIZE, SIZE, QImage.Format_ARGB32)
    img.fill(Qt.transparent)
    p = QPainter(img)
    p.setRenderHint(QPainter.Antialiasing, True)
    p.setRenderHint(QPainter.TextAntialiasing, True)

    # 圆角方形底板 + 渐变（与程序主色调一致）
    radius = SIZE * 0.22
    path = QPainterPath()
    path.addRoundedRect(QRectF(SIZE * 0.06, SIZE * 0.06, SIZE * 0.88, SIZE * 0.88), radius, radius)
    grad = QLinearGradient(0, 0, SIZE, SIZE)
    grad.setColorAt(0.0, QColor("#4F6EF7"))
    grad.setColorAt(0.55, QColor("#6C5CE7"))
    grad.setColorAt(1.0, QColor("#00B894"))
    p.fillPath(path, grad)

    # 内描边高光
    p.setPen(QPen(QColor(255, 255, 255, 60), SIZE * 0.012))
    p.drawPath(path)

    # 学士帽图形（简洁几何绘制，避免依赖 emoji 字体）
    cx, cy = SIZE * 0.5, SIZE * 0.40
    cap_w, cap_h = SIZE * 0.46, SIZE * 0.20
    cap = QPainterPath()
    cap.moveTo(cx - cap_w / 2, cy)
    cap.lineTo(cx, cy - cap_h / 2)
    cap.lineTo(cx + cap_w / 2, cy)
    cap.lineTo(cx, cy + cap_h / 2)
    cap.closeSubpath()
    p.fillPath(cap, QColor(255, 255, 255, 235))

    # 帽穗
    p.setPen(QPen(QColor(255, 255, 255, 220), SIZE * 0.018, Qt.SolidLine, Qt.RoundCap))
    p.drawLine(int(cx + cap_w * 0.42), int(cy + cap_h * 0.05),
               int(cx + cap_w * 0.46), int(cy + cap_h * 0.95))
    p.setBrush(QColor("#FDCB6E"))
    p.setPen(Qt.NoPen)
    p.drawEllipse(QRectF(cx + cap_w * 0.40, cy + cap_h * 0.90, SIZE * 0.055, SIZE * 0.055))

    # 文字「教」
    p.setPen(QColor(255, 255, 255, 250))
    font = QFont(pick_cjk_font())
    font.setPixelSize(int(SIZE * 0.30))
    font.setBold(True)
    p.setFont(font)
    p.drawText(QRectF(0, SIZE * 0.56, SIZE, SIZE * 0.32), Qt.AlignCenter, "教")

    p.end()
    return img


def main():
    app = QGuiApplication(sys.argv)   # noqa: F841  (QPainter 需要 GUI 应用上下文)
    img = draw_icon()
    if not img.save(str(OUT), "PNG"):
        print(f"[error] 无法写入 {OUT}", file=sys.stderr)
        return 1
    print(f"[ok] 图标已生成: {OUT} ({img.width()}x{img.height()})")

    # Windows 可执行文件图标（.ico 使用标准最大尺寸 256，过大 Windows 不显示）
    ico_img = img.scaled(256, 256, Qt.KeepAspectRatio, Qt.SmoothTransformation)
    if ico_img.save(str(OUT_ICO), "ICO"):
        print(f"[ok] Windows 图标已生成: {OUT_ICO} ({ico_img.width()}x{ico_img.height()})")
    else:
        print("[warn] 当前 Qt 不支持写 ICO，跳过 Windows 图标", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main())
