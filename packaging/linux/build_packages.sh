#!/usr/bin/env bash
# ---------------------------------------------------------------------------
# Linux 打包：生成 .deb 与 AppImage（在 Linux 上运行；GitHub Actions 也用它）
#
#   用法: packaging/linux/build_packages.sh <PyInstaller输出目录> <版本> <输出目录>
#   例:   packaging/linux/build_packages.sh dist/TeachingManager 1.4.0 dist/installers
#
# 依赖: dpkg-deb（deb 构建）；appimagetool 会被自动下载（AppImage，可选）
# ---------------------------------------------------------------------------
set -euo pipefail

APP_DIR="${1:-dist/TeachingManager}"
VERSION="${2:-1.4.0}"
OUT_DIR="${3:-dist/installers}"
APP_NAME="教学管理系统"
PKG_NAME="teachingmanager"
EXEC_NAME="TeachingManager"          # PyInstaller --name 生成的入口可执行文件名
STAMP="$(date -u +%Y%m%d-%H%M)"
DEB_NAME="TeachingManager-${VERSION}-linux-x64-${STAMP}.deb"
APPIMAGE_NAME="TeachingManager-${VERSION}-linux-x64-${STAMP}.AppImage"

REPO_ROOT="$(cd "$(dirname "$0")/../.." && pwd)"
mkdir -p "$OUT_DIR"
OUT_DIR="$(cd "$OUT_DIR" && pwd)"

if [ ! -x "$APP_DIR/$EXEC_NAME" ]; then
  echo "错误：未找到 PyInstaller 产物 $APP_DIR/$EXEC_NAME" >&2
  exit 1
fi

ICON_SRC="$REPO_ROOT/packaging/macos/assets/icon.png"

# ------------------------------------------------------------------ .deb
echo "==> 构建 .deb"
STAGE="$(mktemp -d)"
trap 'rm -rf "$STAGE"' EXIT
mkdir -p "$STAGE/DEBIAN" \
         "$STAGE/opt/$PKG_NAME" \
         "$STAGE/usr/share/applications" \
         "$STAGE/usr/share/icons/hicolor/256x256/apps"

cp -r "$APP_DIR"/. "$STAGE/opt/$PKG_NAME/"
chmod +x "$STAGE/opt/$PKG_NAME/$EXEC_NAME"
[ -f "$ICON_SRC" ] && cp "$ICON_SRC" "$STAGE/usr/share/icons/hicolor/256x256/apps/$PKG_NAME.png"

cat > "$STAGE/usr/share/applications/$PKG_NAME.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=$APP_NAME
Name[en]=Teaching Management System
Comment=本地离线的教学管理系统（学生/排课/便签/成绩/学情）
Exec=/opt/$PKG_NAME/$EXEC_NAME
Icon=$PKG_NAME
Terminal=false
Categories=Education;Office;
StartupWMClass=$EXEC_NAME
EOF

INSTALLED_SIZE="$(du -sk "$STAGE/opt" | cut -f1)"
cat > "$STAGE/DEBIAN/control" <<EOF
Package: $PKG_NAME
Version: $VERSION
Section: education
Priority: optional
Architecture: amd64
Installed-Size: $INSTALLED_SIZE
Maintainer: Asell Bucks <104905686+Abucks@users.noreply.github.com>
Homepage: https://github.com/Abucks/Teaching-Management-System
Depends: libc6, libegl1 | libegl-mesa0, libgl1, libxkbcommon0, libdbus-1-3, libfontconfig1, libglib2.0-0, libxcb-cursor0, libxcb-xinerama0, libxcb-icccm4, libxcb-keysyms1, libxcb-render-util0, libxcb-shape0
Description: 教学管理系统（Teaching Management System）
 本地离线运行的教师工作台：学生管理、排课月历、便签墙、成绩图表与学情管理。
 数据保存在本机 SQLite，不联网、不上传。
EOF

dpkg-deb --build --root-owner-group "$STAGE" "$OUT_DIR/$DEB_NAME" >/dev/null
echo "    ✓ $DEB_NAME"

# ------------------------------------------------------------- AppImage
echo "==> 构建 AppImage"
APPDIR="$(mktemp -d)/AppDir"
mkdir -p "$APPDIR/usr/bin" "$APPDIR/usr/share/icons/hicolor/256x256/apps"
cp -r "$APP_DIR"/. "$APPDIR/usr/bin/"
chmod +x "$APPDIR/usr/bin/$EXEC_NAME"
[ -f "$ICON_SRC" ] && cp "$ICON_SRC" "$APPDIR/$PKG_NAME.png" \
  && cp "$ICON_SRC" "$APPDIR/usr/share/icons/hicolor/256x256/apps/$PKG_NAME.png"

cat > "$APPDIR/$PKG_NAME.desktop" <<EOF
[Desktop Entry]
Type=Application
Name=$APP_NAME
Comment=本地离线的教学管理系统
Exec=$EXEC_NAME
Icon=$PKG_NAME
Terminal=false
Categories=Education;Office;
EOF

cat > "$APPDIR/AppRun" <<EOF
#!/bin/sh
HERE="\$(dirname "\$(readlink -f "\$0")")"
exec "\$HERE/usr/bin/$EXEC_NAME" "\$@"
EOF
chmod +x "$APPDIR/AppRun"

TOOL="$(mktemp -d)/appimagetool"
if curl -fsSL -o "$TOOL" \
    "https://github.com/AppImage/appimagetool/releases/download/continuous/appimagetool-x86_64.AppImage"; then
  chmod +x "$TOOL"
  # --appimage-extract-and-run 让它在无 FUSE 的容器/CI 里也能运行
  if ARCH=x86_64 "$TOOL" --appimage-extract-and-run "$APPDIR" "$OUT_DIR/$APPIMAGE_NAME" >/dev/null 2>&1; then
    echo "    ✓ $APPIMAGE_NAME"
  else
    echo "    ! AppImage 构建失败（跳过，不影响 .deb）" >&2
  fi
else
  echo "    ! 无法下载 appimagetool（跳过 AppImage）" >&2
fi

echo "==> Linux 产物:"
ls -lh "$OUT_DIR" | tail -n +2
