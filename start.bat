@echo off
title KPI 自動通知システム
echo KPI 自動通知システムを起動します...
echo.

:: 変数の定義（必要に応じて変更）
set "VENV_NAME=kpi_venv"                                           :: 仮想環境の名前
set "VENV_PATH=%USERPROFILE%\Documents\%VENV_NAME%"                 :: 仮想環境のパス（ユーザーのドキュメントフォルダ内）
set "PYTHON_SCRIPT_PATH=\\network_share\path\to\startup.py"        :: ネットワーク上の startup.py のパス（必ず変更してください）
set "REQUIREMENTS_PATH=%~dp0requirements.txt" :: requirements.txtのパス（バッチファイルと同じフォルダ内）
set "PYTHON_PATH=%VENV_PATH%\Scripts\python.exe"                     :: 仮想環境の Python 実行ファイルパス

:: Pythonがインストールされているか確認
echo Pythonのインストール状態を確認しています...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo エラー: Pythonがインストールされていないか、環境変数が設定されていません。
    echo Python 3.6以上をインストールしてから再試行してください。
    pause
    exit /b 1
)
echo Pythonは正常にインストールされています。

:: ネットワーク上のstartup.pyが存在するか確認
echo ネットワーク上のstartup.pyの存在を確認しています...
if not exist "%PYTHON_SCRIPT_PATH%" (
    echo エラー: ネットワーク上のstartup.pyが見つかりません。パスを確認してください: "%PYTHON_SCRIPT_PATH%"
    pause
    exit /b 1
)
echo ネットワーク上のstartup.pyは正常に存在しています。

:: 仮想環境が存在するか確認
echo 仮想環境の存在を確認しています...
if not exist "%VENV_PATH%" (
    echo 仮想環境 "%VENV_NAME%" が存在しないため、作成します...
    python -m venv "%VENV_PATH%"
    if %errorlevel% neq 0 (
        echo エラー: 仮想環境の作成に失敗しました。
        pause
        exit /b 1
    )
    echo 仮想環境 "%VENV_NAME%" を作成しました。
    
    :: requirements.txtが存在するか確認
    if not exist "%REQUIREMENTS_PATH%" (
        echo エラー: requirements.txtが見つかりません。パスを確認してください: "%REQUIREMENTS_PATH%"
        pause
        exit /b 1
    )
    
    :: 必要なライブラリをインストール
    echo 必要なライブラリをインストールします...
    "%VENV_PATH%\Scripts\pip.exe" install -r "%REQUIREMENTS_PATH%"
    if %errorlevel% neq 0 (
        echo エラー: ライブラリのインストールに失敗しました。
        pause
        exit /b 1
    )
    echo 必要なライブラリは正常にインストールされました。
)
echo 仮想環境は正常に存在しています。

:: 仮想環境をアクティベートし、スクリプトを実行
echo 仮想環境をアクティベートし、スクリプトを実行します...
"%PYTHON_PATH%" "%PYTHON_SCRIPT_PATH%"
if %errorlevel% neq 0 (
    echo エラー: Pythonスクリプトの実行に失敗しました。
    pause
    exit /b 1
)

echo 完了しました。
pause
