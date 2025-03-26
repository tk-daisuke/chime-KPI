@echo off
title KPI 自動通知システム
echo KPI 自動通知システムを起動します...
echo.

:: Pythonがインストールされているか確認
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo エラー: Pythonがインストールされていないか、環境変数が設定されていません。
    echo Python 3.6以上をインストールしてから再試行してください。
    pause
    exit /b 1
)

:: startup.pyスクリプトが存在するか確認。なければ作成
if not exist startup.py (
    echo startup.pyが見つかりません。
)

:: startup.pyを実行
python startup.py