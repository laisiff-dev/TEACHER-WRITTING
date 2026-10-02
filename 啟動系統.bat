@echo off
chcp 65001 >nul
title 輔英科技大學 - 教師論文著作檢索與自動匯入系統

cls
echo ========================================================
echo   輔英科技大學 - 教師論文著作檢索與自動匯入系統 啟動控制台
echo ========================================================
echo.

set PYCMD=none

:: 檢查系統 PATH 中的 python 或 py
where python >nul 2>&1
if %errorlevel% equ 0 set PYCMD=python

if "%PYCMD%"=="none" (
    where py >nul 2>&1
    if %errorlevel% equ 0 set PYCMD=py
)

:: 檢查 Windows 常規預設安裝路徑
if "%PYCMD%"=="none" (
    for /d %%d in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
        if exist "%%d\python.exe" set PYCMD="%%d\python.exe"
    )
)

if "%PYCMD%"=="none" (
    for /d %%d in ("C:\Python3*") do (
        if exist "%%d\python.exe" set PYCMD="%%d\python.exe"
    )
)

:: 若未安裝 Python 提示下載指引
if "%PYCMD%"=="none" (
    echo [ERROR] 尚未偵測到 Python 環境！
    echo.
    echo 請在該台電腦下載並安裝 Python:
    echo   1. 下載官方安裝檔: https://www.python.org/downloads/
    echo   2. 安裝時【務必勾選】"Add Python.exe to PATH" (將 Python 加入系統環境變數)
    echo   3. 安裝完成後，請重新雙擊執行此「啟動系統.bat」。
    echo.
    pause
    exit /b 1
)

echo [1/3] 已成功偵測 Python 環境: %PYCMD%
echo [2/3] 正在檢查並自動安裝必要套件 (Flask, Pandas, Excel 模組)...
%PYCMD% -m pip install -r requirements.txt --prefer-binary
if %errorlevel% neq 0 (
    echo [! 提示] 自動安裝套件遇到問題，嘗試補強安裝...
    %PYCMD% -m pip install flask requests pandas beautifulsoup4 openpyxl xlwt lxml urllib3 --prefer-binary
)

echo.
echo ========================================================
echo [3/3] 正在啟動 Web 網頁系統...
echo.
echo 觀察黑色主控台視窗中的提示網址：
echo 提示網址: http://127.0.0.1:5000   (若被佔用自動改用 8080)
echo.
echo  【說明】系統將會自動開啟預設瀏覽器 (Chrome / Edge)
echo  【提醒】使用期間請保持此黑色主控台視窗開啟！
echo ========================================================
echo.

%PYCMD% main.py --web

if %errorlevel% neq 0 (
    echo.
    echo [!] 系統執行發生異常，請參考上方提示訊息。
    echo.
)

pause
