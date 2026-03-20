@echo off
chcp 65001 >nul
title BiddingCount 单文件打包工具 (分离架构版)
color 0B

echo ======================================================
echo           正在打包 BiddingCount (分离架构)
echo ======================================================

:: 1. 环境准备
echo [1/4] 清理旧的构建目录...
if exist dist rd /s /q dist
if exist build rd /s /q build

:: 2. 检查必要文件
echo [2/4] 检查资源文件...
set ICON_CMD=
if exist "icon.ico" (
    echo [确认] 发现图标: icon.ico
    set ICON_CMD=--icon="icon.ico" --add-data "icon.ico;."
)

:: 核心：检查 main.py 是否在场
if not exist "main.py" (
    echo [错误] 未发现 main.py，逻辑层缺失，打包终止！
    pause
    exit
)

:: 3. 这里的打包命令是关键
echo [3/4] 正在执行 PyInstaller 封装 (全量打包)...
echo ------------------------------------------------------

:: --onefile: 生成单个 exe
:: --noconsole: 界面运行不带黑框
:: --add-data "main.py;.": 这一行最重要！把逻辑文件打包进 exe 内部
:: --collect-all customtkinter: 确保 UI 组件不丢失
pyinstaller --noconsole --onefile ^
    --collect-all customtkinter ^
    --add-data "main.py;." ^
    %ICON_CMD% ^
    --name "BiddingCount_v1.0" ^
    "GUI.py"

echo.
echo [4/4] 正在检查打包结果...
if exist "dist\BiddingCount_v1.0.exe" (
    echo ======================================================
    echo [成功] 最终文件: dist\BiddingCount_v1.0.exe
    echo 你现在只需要把这一个 EXE 发给同事即可。
    echo ======================================================
) else (
    echo [失败] 打包过程中可能出现错误，请检查控制台输出。
)

pause