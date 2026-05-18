#!/usr/bin/env bash
# ============================================================
# build.sh — 离线打包脚本
# ============================================================
# 用法:
#   chmod +x build.sh
#   ./build.sh              # 构建所有格式
#   ./build.sh pyinstaller  # 仅构建 PyInstaller 单文件
#   ./build.sh wheels       # 仅构建 pip wheel 离线包
#   ./build.sh clean        # 清理构建产物
# ============================================================
#
# 产出:
#   dist/table_field_compare          # Linux 单文件可执行 (pyinstaller)
#   dist/table_field_compare.exe      # Windows 交叉构建 (wine)
#   dist/wheels/                      # pip 离线安装包
#   dist/table_field_compare_部署包.zip # 完整部署包 (wheels + 脚本 + README)
#
# 前置依赖:
#   pip install pyinstaller openpyxl
#   Windows交叉构建需要: wine + Windows版Python

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
PROJECT_DIR="$SCRIPT_DIR/.."
DIST_DIR="$PROJECT_DIR/dist"
SRC="$PROJECT_DIR/src/table_field_compare.py"

mkdir -p "$DIST_DIR"

case "${1:-all}" in

  pyinstaller)
    echo "=== 构建 PyInstaller 单文件 ==="
    pip install pyinstaller 2>/dev/null

    # Linux 单文件
    pyinstaller \
      --onefile \
      --name table_field_compare \
      --distpath "$DIST_DIR" \
      --workpath /tmp/pyinstaller_build \
      --clean \
      --noconfirm \
      "$SRC"

    echo "✅ Linux:  $DIST_DIR/table_field_compare"

    # Windows 交叉构建（可选，需要 wine + Windows Python）
    # pyinstaller --onefile --name table_field_compare.exe \
    #   --distpath "$DIST_DIR" --workpath /tmp/pyinstaller_winbuild "$SRC"
    # echo "✅ Windows: $DIST_DIR/table_field_compare.exe"
    ;;

  wheels)
    echo "=== 构建 pip wheel 离线包 ==="
    WHEELS_DIR="$DIST_DIR/wheels"
    rm -rf "$WHEELS_DIR"
    mkdir -p "$WHEELS_DIR"

    # 下载所有依赖的 wheel 文件
    pip download \
      --dest "$WHEELS_DIR" \
      openpyxl

    # 同时下载 et_xmlfile（openpyxl 的依赖）
    pip download \
      --dest "$WHEELS_DIR" \
      et-xmlfile 2>/dev/null || true

    echo "✅ 离线包: $WHEELS_DIR ($(ls -1 "$WHEELS_DIR" | wc -l) 个文件)"
    ls -lh "$WHEELS_DIR"
    ;;

  bundle)
    # 构建完整部署包
    ./build.sh wheels
    BUNDLE_DIR="$DIST_DIR/部署包"
    rm -rf "$BUNDLE_DIR"
    mkdir -p "$BUNDLE_DIR"

    # 复制文件
    cp "$SRC" "$BUNDLE_DIR/"
    cp -r "$DIST_DIR/wheels" "$BUNDLE_DIR/wheels"

    # 生成安装脚本
    cat > "$BUNDLE_DIR/install.sh" << 'INSTALL_SCRIPT'
#!/usr/bin/env bash
set -e
DIR="$(cd "$(dirname "$0")" && pwd)"
echo "=== 安装 openpyxl 依赖 ==="
pip install --no-index --find-links "$DIR/wheels" openpyxl
echo "✅ 安装完成"
echo ""
echo "使用方式:"
echo "  python $DIR/table_field_compare.py TEST_DEV.xlsx standard.xlsx -o result.xlsx"
echo "  python $DIR/table_field_compare.py --help"
INSTALL_SCRIPT
    chmod +x "$BUNDLE_DIR/install.sh"

    # Windows 安装脚本
    cat > "$BUNDLE_DIR/install.bat" << 'INSTALL_BAT'
@echo off
echo === 安装 openpyxl 依赖 ===
pip install --no-index --find-links "%~dp0wheels" openpyxl
echo.
echo 安装完成！
echo 使用方式:
echo   python "%~dp0table_field_compare.py" TEST_DEV.xlsx standard.xlsx -o result.xlsx
pause
INSTALL_BAT

    # 打包
    cd "$DIST_DIR"
    zip -r "table_field_compare_部署包.zip" "部署包/"
    rm -rf "$BUNDLE_DIR"
    echo "✅ 部署包: $DIST_DIR/table_field_compare_部署包.zip"
    ;;

  clean)
    echo "=== 清理构建产物 ==="
    rm -rf "$DIST_DIR"
    rm -rf /tmp/pyinstaller_build
    echo "✅ 已清理"
    ;;

  all)
    "$0" wheels
    "$0" bundle
    # pyinstaller 跳过（需要 GUI 环境，适合在有桌面环境的机器上构建）
    echo ""
    echo "=== 可选的 PyInstaller 构建 ==="
    echo "  在有桌面环境的机器上执行: ./build.sh pyinstaller"
    ;;

  *)
    echo "用法: $0 {pyinstaller|wheels|bundle|clean|all}"
    exit 1
    ;;
esac
