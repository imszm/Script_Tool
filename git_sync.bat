@echo off
:: 强制使用 UTF-8 编码以解决乱码
chcp 65001 > nul
setlocal enabledelayedexpansion

title Git 自动化同步工具

echo ===========================================
echo       Git 仓库自动提交与推送工具
echo ===========================================
echo.

:: 1. 检查状态
echo [1/4] 正在检查文件状态...
git status
echo.

:: 2. 添加所有更改
echo [2/4] 正在添加更改到暂存区...
git add .
echo 添加完成。
echo.

:: 3. 获取用户输入的提交说明
set /p msg="请输入本次修改的详细说明 (按回车跳过使用默认说明): "

:: 如果用户直接回车，则生成默认的时间戳说明
if "%msg%"=="" (
    set msg=Routine_update_%date%_%time%
    :: 去除时间戳中可能导致 Git 报错的空格和特殊符号
    set msg=!msg: =_!
    set msg=!msg:/=-!
    set msg=!msg::=-!
)

echo.
echo [3/4] 正在提交本地更改...
:: 使用双引号包裹变量，防止空格导致 pathspec 错误
git commit -m "!msg!"
echo.

:: 4. 推送到远程仓库
echo [4/4] 正在推送到 GitHub (origin main)...
git push origin main

if %ERRORLEVEL% equ 0 (
    echo.
    echo ===========================================
    echo       操作成功！代码已同步至 GitHub。
    echo ===========================================
) else (
    echo.
    echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
    echo       错误：推送失败，请检查网络或 Token 状态。
    echo !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
)

echo.
echo 按任意键退出...
pause > nul