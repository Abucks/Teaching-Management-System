#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# macOS 一键打包包装脚本（等价于 python3 build_macos.py）
#
#   ./build_macos.sh                 # 产出 .app 与 .dmg
#   ./build_macos.sh --no-dmg        # 只产出 .app
#   ./build_macos.sh --arch universal2
#   ./build_macos.sh --sign "Developer ID Application: XX (TEAMID)"
#
# 需要 macOS + Xcode 命令行工具（xcode-select --install）
# ---------------------------------------------------------------------------
set -euo pipefail

cd "$(dirname "$0")"

PY="${PYTHON:-python3}"

if ! command -v "$PY" >/dev/null 2>&1; then
  echo "错误：未找到 $PY。" >&2
  echo "请安装 Python 3.10+（推荐 https://www.python.org/downloads/macos/ 的 universal2 安装包），" >&2
  echo "或先安装 Xcode 命令行工具：xcode-select --install" >&2
  exit 1
fi

if [ "$(uname -s)" != "Darwin" ]; then
  echo "警告：当前系统不是 macOS（$(uname -s)），PyInstaller 无法交叉编译，请改用 --dry-run 检查流程。" >&2
fi

exec "$PY" build_macos.py "$@"
