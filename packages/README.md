# 表结构字段对比工具 — 离线部署指南

## 📦 产物说明

```
dist/
├── table_field_compare_部署包.zip               # Python 依赖包（需要机子有 Python）
├── table_field_compare_portable_linux.tar.gz     # Linux 自包含便携包（无需 Python）
└── table_field_compare_portable_windows.zip     # Windows 自包含便携包（无需 Python）
```

## 🚀 目标机选择

| 目标机情况 | 用什么 |
|-----------|--------|
| 有 Python 3.8+ | `部署包.zip`（~300KB） |
| 无 Python — Windows | `portable_windows.zip`（~12MB） |
| 无 Python — Linux | `portable_linux.tar.gz`（~51MB） |

---

## 方式一：pip wheel 离线包（目标机有 Python）

```bash
unzip table_field_compare_部署包.zip
cd 部署包/
./install.sh              # Linux/macOS
# install.bat             # Windows 双击
```

---

## 方式二：便携包（目标机无 Python，解压即用）

### Linux

```bash
tar xzf table_field_compare_portable_linux.tar.gz
cd portable_linux/
./run.sh TEST_DEV.xlsx standard.xlsx -o result.xlsx
```

### Windows

解压 `table_field_compare_portable_windows.zip`，两种用法：

1. **命令行** — 打开 cmd，cd 到解压目录：
   ```
   运行对比.bat TEST_DEV.xlsx standard.xlsx -o result.xlsx
   ```

2. **拖拽** — 同时选中两个 Excel，拖到 `拖拽文件运行.bat` 上

---

## 📋 使用

```bash
# 查看帮助
python table_field_compare.py --help

# 基本用法
python table_field_compare.py TEST_DEV.xlsx standard.xlsx -o result.xlsx

# 仅控制台输出
python table_field_compare.py dev.xlsx std.xlsx --console-only

# 自定义排除字段
python table_field_compare.py dev.xlsx std.xlsx --exclude DEL_FLG,SRC_SYS_CD
```

## 🔤 字体说明

自动检测系统中文字体：Windows → 微软雅黑 / macOS → PingFang SC / Linux → Noto Sans CJK SC

## 🔄 重新构建

```bash
cd packages/

# 轻量级（需要 Python）
./build.sh all

# 便携包（解压即用）
./build_portable.sh all
```
