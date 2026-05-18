#!/usr/bin/env bash
# ============================================================
# build_portable.sh — 构建自包含便携包
# ============================================================
# 用法:
#   ./build_portable.sh linux    # Linux 便携包
#   ./build_portable.sh windows  # Windows 便携包
#   ./build_portable.sh all      # 两者
#   ./build_portable.sh download # 下载 Python 运行时
#   ./build_portable.sh clean
#
# 产出:
#   dist/table_field_compare_portable_linux.tar.gz    (~50MB)
#   dist/table_field_compare_portable_windows.zip     (~25MB)

set -euo pipefail
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
DIST_DIR="$SCRIPT_DIR/../dist"
WHEELS_DIR="$DIST_DIR/wheels"
RUNTIME_DIR="$DIST_DIR/runtime"
SRC="$SCRIPT_DIR/../src/table_field_compare.py"
PACKAGES_DIR="$SCRIPT_DIR"

GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m'

clean() {
    echo -e "${YELLOW}=== 清理 ===${NC}"
    rm -rf "$DIST_DIR/portable_linux" "$DIST_DIR/portable_windows"
    rm -f "$DIST_DIR/table_field_compare_portable_linux.tar.gz"
    rm -f "$DIST_DIR/table_field_compare_portable_windows.zip"
    echo -e "${GREEN}✅ 已清理${NC}"
}

check_prereqs() {
    if [ ! -d "$WHEELS_DIR" ] || [ -z "$(ls -A "$WHEELS_DIR" 2>/dev/null)" ]; then
        echo -e "${YELLOW}wheels 为空，先下载...${NC}"
        "$PACKAGES_DIR/build.sh" wheels
    fi
    if [ ! -f "$SRC" ]; then
        echo "❌ 找不到: $SRC"
        exit 1
    fi
    mkdir -p "$DIST_DIR"
}

build_linux() {
    echo -e "${GREEN}=== 构建 Linux 便携包 ===${NC}"
    check_prereqs

    local TARBALL
    TARBALL=$(ls "$RUNTIME_DIR"/cpython-*x86_64*linux*install_only.tar.gz 2>/dev/null | head -1 || true)
    if [ -z "$TARBALL" ]; then
        echo "❌ 未找到 Linux Python 运行时，先执行: ./build_portable.sh download"
        exit 1
    fi

    local PORTABLE="$DIST_DIR/portable_linux"
    rm -rf "$PORTABLE"
    mkdir -p "$PORTABLE"

    echo "  解压 Python..."
    tar xzf "$TARBALL" -C "$PORTABLE"

    # 找 python bin
    local PYTHON_BIN
    PYTHON_BIN=$(find "$PORTABLE" -name python3.11 -o -name python3 | head -1)
    if [ -z "$PYTHON_BIN" ] || [ ! -f "$PYTHON_BIN" ]; then
        echo "❌ 找不到 python 可执行文件"
        find "$PORTABLE" -name 'python*' -type f 2>/dev/null | head -5
        exit 1
    fi
    echo "  Python: $PYTHON_BIN"

    # 安装依赖
    echo "  安装 openpyxl..."
    "$PYTHON_BIN" -m pip install \
        --no-index --find-links "$WHEELS_DIR" \
        --target "$PORTABLE/site-packages" \
        openpyxl 2>&1 | tail -1

    # 复制脚本
    cp "$SRC" "$PORTABLE/table_field_compare.py"

    # 启动脚本
    cat > "$PORTABLE/run.sh" << 'LAUNCHER'
#!/usr/bin/env bash
DIR="$(cd "$(dirname "$0")" && pwd)"
PYTHON_BIN=$(find "$DIR" -type f -name python3.11 -o -type f -name python3 | head -1)
export PYTHONPATH="$DIR/site-packages:$PYTHONPATH"
exec "$PYTHON_BIN" "$DIR/table_field_compare.py" "$@"
LAUNCHER
    chmod +x "$PORTABLE/run.sh"

    # 打包
    echo "  打包..."
    local OUTPUT="$DIST_DIR/table_field_compare_portable_linux.tar.gz"
    cd "$DIST_DIR"
    tar czf "$OUTPUT" portable_linux/
    local SIZE; SIZE=$(du -h "$OUTPUT" | cut -f1)
    echo -e "${GREEN}✅ ${OUTPUT} (${SIZE})${NC}"
    rm -rf "$PORTABLE"
}

build_windows() {
    echo -e "${GREEN}=== 构建 Windows 便携包 ===${NC}"
    check_prereqs

    local WIN_ZIP
    WIN_ZIP=$(ls "$RUNTIME_DIR"/python-*-embed-amd64.zip 2>/dev/null | head -1 || true)
    if [ -z "$WIN_ZIP" ]; then
        echo "❌ 未找到 Windows Python embeddable，先执行: ./build_portable.sh download"
        exit 1
    fi

    local PORTABLE="$DIST_DIR/portable_windows"
    rm -rf "$PORTABLE"
    mkdir -p "$PORTABLE"

    echo "  解压 Python embeddable..."
    unzip -qo "$WIN_ZIP" -d "$PORTABLE"

    # 启用 site-packages
    local PTH_FILE
    PTH_FILE=$(ls "$PORTABLE"/python*._pth 2>/dev/null | head -1)
    if [ -n "$PTH_FILE" ]; then
        echo "Lib/site-packages" >> "$PTH_FILE"
        sed -i 's/^#import site/import site/' "$PTH_FILE" 2>/dev/null || true
    fi

    # 安装 wheels 到 site-packages
    local SITE="$PORTABLE/Lib/site-packages"
    mkdir -p "$SITE"
    for whl in "$WHEELS_DIR"/*.whl; do
        echo "  安装: $(basename "$whl")"
        unzip -qo "$whl" -d "$SITE"
    done

    # 复制脚本
    cp "$SRC" "$PORTABLE/table_field_compare.py"

    # 启动批处理（用独立脚本生成，避免 bash 转义问题）
    python3 "$PACKAGES_DIR/generate_bat.py" "$PORTABLE"

    # 打包
    echo "  打包..."
    local OUTPUT="$DIST_DIR/table_field_compare_portable_windows.zip"
    cd "$DIST_DIR"
    rm -f "$OUTPUT"
    zip -qr "$OUTPUT" portable_windows/
    local SIZE; SIZE=$(du -h "$OUTPUT" | cut -f1)
    echo -e "${GREEN}✅ ${OUTPUT} (${SIZE})${NC}"
    rm -rf "$PORTABLE"
}

download() {
    echo -e "${GREEN}=== 下载 Python 运行时 ===${NC}"
    mkdir -p "$RUNTIME_DIR"

    # Windows embeddable
    local WIN_DEST="$RUNTIME_DIR/python-3.11.9-embed-amd64.zip"
    if [ ! -f "$WIN_DEST" ]; then
        echo "  下载 Windows embeddable (10MB)..."
        curl -sL "https://www.python.org/ftp/python/3.11.9/python-3.11.9-embed-amd64.zip" -o "$WIN_DEST"
        echo "  ✅ $(du -h "$WIN_DEST" | cut -f1)"
    else
        echo "  Windows: 已存在"
    fi

    # Linux standalone
    if ! ls "$RUNTIME_DIR"/cpython-*install_only.tar.gz 2>/dev/null | head -1 | grep -q .; then
        echo "  下载 Linux standalone..."
        local LINUX_URL
        LINUX_URL=$(curl -sL \
            -H "User-Agent: Mozilla/5.0" \
            https://api.github.com/repos/indygreg/python-build-standalone/releases/latest \
            | python3 -c "
import json,sys
r=json.load(sys.stdin)
for a in r['assets']:
    n=a['name']
    if 'x86_64-unknown-linux-gnu' in n and 'cpython-3.11' in n and 'install_only' in n and 'debug' not in n:
        print(a['browser_download_url'])
        break
")
        if [ -n "$LINUX_URL" ]; then
            local FNAME; FNAME=$(basename "$LINUX_URL")
            curl -sL "$LINUX_URL" -o "$RUNTIME_DIR/$FNAME"
            echo "  ✅ $(du -h "$RUNTIME_DIR/$FNAME" | cut -f1)"
        else
            echo "  ❌ 获取失败"
        fi
    else
        echo "  Linux: 已存在"
    fi
}

case "${1:-all}" in
    linux)    build_linux ;;
    windows)  build_windows ;;
    all)      build_linux; build_windows ;;
    clean)    clean ;;
    download) download ;;
    *)
        echo "用法: $0 {linux|windows|all|clean|download}"
        exit 1
        ;;
esac
