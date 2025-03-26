import os
import sys
import subprocess
import time
import platform

def show_loading_animation(text):
    """ローディングアニメーションを表示する関数"""
    for i in range(4):
        sys.stdout.write(f"\r{text}" + "." * i + " " * (3 - i))
        sys.stdout.flush()
        time.sleep(0.5)
    print(f"\r{text}.... 完了")

def check_venv():
    """仮想環境が存在するか確認し、なければ作成する"""
    venv_dir = ".venv"
    
    # 仮想環境のディレクトリが存在するか確認
    if not os.path.exists(venv_dir):
        print("仮想環境が見つかりません。新しく作成します...")
        try:
            subprocess.run([sys.executable, "-m", "venv", venv_dir], check=True)
            print("仮想環境を作成しました。")
        except subprocess.CalledProcessError:
            print("仮想環境の作成に失敗しました。Python 3.3以上がインストールされているか確認してください。")
            return False
    return True

def get_activate_command():
    """OSに応じたactivateコマンドを返す"""
    if platform.system() == "Windows":
        return os.path.join(".venv", "Scripts", "activate")
    else:  # Linux or MacOS
        return f"source {os.path.join('.venv', 'bin', 'activate')}"

def install_requirements():
    """requirements.txtからパッケージをインストールする"""
    # requirements.txtが存在するか確認
    if not os.path.exists("requirements.txt"):
        # requirements.txtがない場合は作成する
        with open("requirements.txt", "w", encoding="utf-8") as f:
            f.write("pandas\nxlwings\nrequests\nmatplotlib\njapanize-matplotlib\npyfiglet\nschedule\nconfigparser\n")
        print("requirements.txtを作成しました。")
    
    print("必要なパッケージをインストールしています...")
    
    # 仮想環境内でpipを使ってパッケージをインストール
    if platform.system() == "Windows":
        activate_cmd = f"call {os.path.join('.venv', 'Scripts', 'activate')}"
        pip_cmd = "pip install -r requirements.txt"
        full_cmd = f"{activate_cmd} && {pip_cmd}"
        subprocess.run(full_cmd, shell=True, check=True)
    else:  # Linux or MacOS
        activate_cmd = f"source {os.path.join('.venv', 'bin', 'activate')}"
        pip_cmd = "pip install -r requirements.txt"
        full_cmd = f"{activate_cmd} && {pip_cmd}"
        subprocess.run(full_cmd, shell=True, executable="/bin/bash", check=True)
    
    show_loading_animation("パッケージのインストール中")

        
def run_main_script():
    """メインスクリプトを実行する"""
    print("メインシステムを起動します...")
    

    # 仮想環境内でメインスクリプトを実行
    if platform.system() == "Windows":
        activate_cmd = f"call {os.path.join('.venv', 'Scripts', 'activate')}"
        run_cmd = "python main.py"
        full_cmd = f"{activate_cmd} && {run_cmd}"
        subprocess.run(full_cmd, shell=True)


def main():
    """メイン関数"""
    print("======================================")
    print("KPI 自動通知システム セットアップツール")
    print("======================================")
    
    # 仮想環境のチェックと作成
    if not check_venv():
        input("エンターキーを押して終了...")
        return
    
    # パッケージのインストール
    try:
        install_requirements()
    except subprocess.CalledProcessError:
        print("パッケージのインストールに失敗しました。")
        input("エンターキーを押して終了...")
        return
    
    # メインスクリプトの実行
    try:
        run_main_script()
    except Exception as e:
        print(f"メインスクリプトの実行中にエラーが発生しました: {e}")
    
    input("エンターキーを押して終了...")

if __name__ == "__main__":
    main()