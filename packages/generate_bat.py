#!/usr/bin/env python3
"""Generate Windows .bat launcher files for portable package."""
import sys
import os

def main():
    portable_dir = sys.argv[1]

    # 1. 命令行运行
    run_bat = r'''@echo off
chcp 65001 >nul 2>&1
set "DIR=%~dp0"
set "PYTHONPATH=%DIR%Lib\site-packages;%PYTHONPATH%"
"%DIR%python.exe" "%DIR%table_field_compare.py" %*
if %ERRORLEVEL% NEQ 0 pause
'''
    path = os.path.join(portable_dir, '运行对比.bat')
    with open(path, 'w', encoding='gbk') as f:
        f.write(run_bat)
    print(f"  -> {path}")

    # 2. 拖拽运行（两个文件拖到 bat 上）
    drag_bat = r'''@echo off
chcp 65001 >nul 2>&1
set "DIR=%~dp0"
if "%~2"=="" (
    echo 用法：同时选中 TEST_DEV.xlsx 和 standard.xlsx
    echo 拖到这个图标上运行
    pause
    exit /b 1
)
set "PYTHONPATH=%DIR%Lib\site-packages;%PYTHONPATH%"
"%DIR%python.exe" "%DIR%table_field_compare.py" "%~1" "%~2" -o "%DIR%结果.xlsx"
echo ============================================
echo 完成！结果保存在: 结果.xlsx
echo ============================================
pause
'''
    path = os.path.join(portable_dir, '拖拽文件运行.bat')
    with open(path, 'w', encoding='gbk') as f:
        f.write(drag_bat)
    print(f"  -> {path}")

if __name__ == '__main__':
    main()
