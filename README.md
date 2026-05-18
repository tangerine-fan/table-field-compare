# table-field-compare

表结构字段对比工具 —— 对比 TEST_DEV 与标准化 Excel 文件的字段差异，生成带样式的对比报告。

## 快速开始

```bash
# 安装依赖
pip install openpyxl

# 运行
python table_field_compare.py TEST_DEV.xlsx standard.xlsx -o result.xlsx
```

## 功能

- 📊 自动识别 TEST_DEV 和标准化文件的表结构和字段列
- 🔍 智能检测字段名所在列，兼容多种 Excel 格式
- 📝 输出双 Sheet Excel 报告：汇总对比 + 字段级详细对比
- 🎨 条件着色（绿/红/黄）一目了然
- 📈 底部统计汇总行
- 🚫 可配置排除字段（默认排除 ETL 时间戳类字段）
- 🖥️ `--console-only` 模式，不生成 Excel 仅控制台输出
- 🔤 跨平台中文字体自动检测（Windows/macOS/Linux）

## 用法

```
用法: python table_field_compare.py TEST_DEV.xlsx standard.xlsx [选项]

选项:
  -o, --output PATH     输出文件路径（默认: 表结构字段对比结果.xlsx）
  --exclude FIELDS      需要排除的字段，逗号分隔
  --console-only        仅控制台输出，不生成 Excel
  -v, --verbose         详细日志
  -h, --help            帮助信息
```

### 示例

```bash
# 基本用法
python table_field_compare.py TEST_DEV.xlsx standard.xlsx -o result.xlsx

# 仅控制台查看
python table_field_compare.py dev.xlsx std.xlsx --console-only

# 自定义排除字段
python table_field_compare.py dev.xlsx std.xlsx --exclude DEL_FLG,SRC_SYS_CD

# 查看所有选项
python table_field_compare.py --help
```

## 输出

### Sheet 1: 汇总对比

| 标准表名 | DEV原表名 | 标准化字段数 | DEV字段数 | 对比结果 | 差异说明 |
|---------|----------|------------|----------|---------|---------|
| CCS_CARD_ACCT | S_CCS_CCS_CARD_ACCT_MALL | 4 | 4 | 存在差异 | 标准化独有1个：CARD_STS \| DEV独有1个：OPN_DT |
| CRD_CUST_INFO | S_CRD_CRD_CUST_INFO_MTH | 4 | 4 | 完全一致 | 字段数量和字段名完全一致（4个字段） |

底部包含统计汇总：共 N 表 | 完全一致: X | 存在差异: Y | DEV缺失: Z | 标准化缺失: W

### Sheet 2: 字段级详细对比

逐字段列出每个字段的来源：两边共有 / 仅标准化有 / 仅DEV有

## 离线部署

详见 [packages/README.md](packages/README.md)，提供三种交付方式：

| 方式 | 大小 | 目标机要求 |
|------|------|-----------|
| pip wheel 离线包 | ~300KB | 需要 Python 3.8+ |
| Linux 便携包 | ~50MB | 解压即用，无需 Python |
| Windows 便携包 | ~12MB | 解压即用，无需 Python |

## 依赖

- Python ≥ 3.8
- [openpyxl](https://openpyxl.readthedocs.io/) ≥ 3.1

## 许可证

MIT
