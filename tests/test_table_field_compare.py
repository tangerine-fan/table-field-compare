"""
table_field_compare 测试套件
用法:
    pip install pytest
    pytest tests/ -v
    pytest tests/ -v --tb=long
"""
from __future__ import annotations

import os
import subprocess
import sys
import tempfile
from pathlib import Path

import openpyxl
import pytest

PROJECT_ROOT = Path(__file__).resolve().parent.parent
SCRIPT = PROJECT_ROOT / "src" / "table_field_compare.py"


# ── Fixtures ──

@pytest.fixture(scope="module")
def dev_xlsx():
    """创建 TEST_DEV 测试数据"""
    tmpdir = tempfile.mkdtemp(prefix="tfc_")
    path = os.path.join(tmpdir, "dev.xlsx")
    _make_test_dev(path, {
        "S_CCS_CCS_CARD_ACCT_MALL": ["ACCT_NO", "CARD_NO", "OVLMT_DATE", "OPN_DT", "DEL_FLG"],
        "S_CRD_CRD_CUST_INFO_MTH": ["CUST_ID", "CUST_NAME", "PART_DT"],
    })
    return path


@pytest.fixture(scope="module")
def std_xlsx():
    """创建标准化测试数据"""
    tmpdir = tempfile.mkdtemp(prefix="tfc_")
    path = os.path.join(tmpdir, "std.xlsx")
    _make_std_file(path, {
        "CCS_CARD_ACCT": ["ACCT_NO", "CARD_NO", "OVLMT_DATE", "CARD_STS"],
        "CRD_CUST_INFO": ["CUST_ID", "CUST_NAME", "BIRTH_DT"],
    })
    return path


@pytest.fixture(scope="module")
def empty_dev_xlsx():
    tmpdir = tempfile.mkdtemp(prefix="tfc_")
    path = os.path.join(tmpdir, "empty_dev.xlsx")
    _make_test_dev(path, {})
    return path


@pytest.fixture(scope="module")
def empty_std_xlsx():
    tmpdir = tempfile.mkdtemp(prefix="tfc_")
    path = os.path.join(tmpdir, "empty_std.xlsx")
    _make_std_file(path, {})
    return path


@pytest.fixture
def out_xlsx():
    """临时输出文件"""
    return os.path.join(tempfile.mkdtemp(prefix="tfc_out_"), "result.xlsx")


# ── 数据工厂 ──

def _make_test_dev(path, tables):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "字段信息"
    ws.cell(row=1, column=1, value="SCHEMA标识")
    ws.cell(row=1, column=2, value="表名")
    ws.cell(row=1, column=3, value="字段名")
    row_idx = 2
    for table, fields in tables.items():
        for field in fields:
            ws.cell(row=row_idx, column=2, value=table)
            ws.cell(row=row_idx, column=3, value=field)
            row_idx += 1
    wb.save(path)
    return path


def _make_std_file(path, tables):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "标准化字段映射"
    ws.cell(row=1, column=1, value="标准化字段映射表")
    ws.cell(row=2, column=2, value="表名")
    ws.cell(row=2, column=3, value="字段名称")
    row_idx = 3
    for table, fields in tables.items():
        for field in fields:
            ws.cell(row=row_idx, column=1, value=row_idx - 2)
            ws.cell(row=row_idx, column=2, value=table)
            ws.cell(row=row_idx, column=3, value=field)
            row_idx += 1
    wb.save(path)
    return path


# ── 运行辅助 ──

def run(*args, expect_ok=True):
    """运行脚本，失败时附带完整 stderr"""
    result = subprocess.run(
        [sys.executable, str(SCRIPT)] + list(args),
        capture_output=True, text=True, timeout=30,
    )
    if expect_ok and result.returncode != 0:
        # pytest 会展示这段信息，不用藏着了
        raise AssertionError(
            f"脚本返回 {result.returncode}\n"
            f"=== STDERR ===\n{result.stderr}\n"
            f"=== STDOUT ===\n{result.stderr[-500:]}"
        )
    return result


# ── 测试 ──

class TestNormalRun:
    """正常场景"""

    def test_generates_excel(self, dev_xlsx, std_xlsx, out_xlsx):
        result = run(dev_xlsx, std_xlsx, "-o", out_xlsx)
        assert os.path.exists(out_xlsx)
        assert "统计" in result.stderr

    def test_output_structure(self, dev_xlsx, std_xlsx, out_xlsx):
        run(dev_xlsx, std_xlsx, "-o", out_xlsx)
        wb = openpyxl.load_workbook(out_xlsx)
        assert wb.sheetnames == ["汇总对比", "字段级详细对比"]

    def test_summary_content(self, dev_xlsx, std_xlsx, out_xlsx):
        run(dev_xlsx, std_xlsx, "-o", out_xlsx)
        ws = openpyxl.load_workbook(out_xlsx)["汇总对比"]
        assert ws.cell(row=1, column=1).value == "标准表名"
        assert ws.cell(row=2, column=1).value == "CCS_CARD_ACCT"
        assert ws.cell(row=2, column=5).value == "存在差异"

    def test_hermes_blue_style(self, dev_xlsx, std_xlsx, out_xlsx):
        run(dev_xlsx, std_xlsx, "-o", out_xlsx)
        ws = openpyxl.load_workbook(out_xlsx)["汇总对比"]
        fill = ws.cell(row=1, column=1).fill.start_color.rgb
        assert fill == "004472C4", f"期望 004472C4，实际 {fill}"


class TestEdgeCases:
    """边界情况"""

    def test_exclude_fields(self, dev_xlsx, std_xlsx, out_xlsx):
        run(dev_xlsx, std_xlsx, "-o", out_xlsx, "--exclude", "DEL_FLG,PART_DT")
        ws = openpyxl.load_workbook(out_xlsx)["汇总对比"]
        assert ws.cell(row=2, column=4).value == 4  # CCS_CARD_ACCT 排除 DEL_FLG 后剩 4 个

    def test_console_only(self, dev_xlsx, std_xlsx, out_xlsx):
        run(dev_xlsx, std_xlsx, "-o", out_xlsx, "--console-only")
        assert not os.path.exists(out_xlsx), "--console-only 不应生成文件"

    def test_empty_dev(self, empty_dev_xlsx, std_xlsx, out_xlsx):
        result = run(empty_dev_xlsx, std_xlsx, "-o", out_xlsx)
        assert "DEV缺失" in result.stderr

    def test_empty_std(self, dev_xlsx, empty_std_xlsx, out_xlsx):
        result = run(dev_xlsx, empty_std_xlsx, "-o", out_xlsx)
        assert "标准化缺失" in result.stderr

    def test_file_not_found(self, out_xlsx):
        result = run("/tmp/nonexistent_12345.xlsx", "/tmp/also_nope.xlsx",
                     "-o", out_xlsx, expect_ok=False)
        assert result.returncode != 0


class TestCLI:
    """命令行"""

    def test_help(self):
        result = run("--help")
        output = result.stdout + result.stderr  # argparse help 在 stdout
        assert "--exclude" in output
        assert "--console-only" in output
        assert "--theme" not in output  # 不支持运行时切换

    def test_missing_args(self):
        result = run(expect_ok=False)
        assert result.returncode != 0
