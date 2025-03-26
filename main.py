# 仮想環境の構築手順
# 仮想環境を作成: python -m venv .venv
# 仮想環境をアクティベート: .venv\Scripts\activate

# 仮想環境のアクティベート確認
# ターミナルに (.venv) と表示されれば成功
#pip freeze > requirements.txt
#pip install -r requirements.txt
#python -m pip install setuptools
import time
import pandas as pd
import xlwings as xw
import requests
import pyfiglet
import schedule
from datetime import datetime
import configparser

# 設定ファイルの読み込み
config = configparser.ConfigParser()
# UTF-8エンコーディングを指定して読み込む
try:
    config.read('config.ini', encoding='utf-8')
    if not config.sections():
        raise FileNotFoundError("config.iniが見つからないか、空です。")
except FileNotFoundError as e:
    print(f"設定ファイルの読み込みエラー: {e}")
    print("スクリプトと同じディレクトリに config.ini を作成し、設定を記述してください。")
    exit()
except Exception as e:
    print(f"設定ファイルの解析中に予期せぬエラーが発生しました: {e}")
    exit()

# 設定値 (configparserから読み込むように変更)
class Settings:
    # Excelファイル設定
    excel_files = {
        'book3': {
            'path': config.get('Excel', 'Book3Path', fallback='data.xlsx'),
            'sheet': config.get('Excel', 'Book3Sheet', fallback='Sheet1')
        }
        # 必要に応じて他のブックの設定も config.ini から読み込むように追加
    }
    
    # Chime設定
    hook_url = config.get('Chime', 'HookURL', fallback='') # デフォルトは空文字列
    max_msg_size = config.getint('Chime', 'MaxMsgSize', fallback=4096)

# ローディングアニメーションを表示する関数
def showLoad():
    """ローディング表示"""
    loading_text = "Loading"
    for i in range(4):
        print(loading_text + '.' * i, end='\r')
        time.sleep(0.5)
    print(loading_text + '....')

# Excelデータ取得クラス
class ExcelSvc:
    def __init__(self):
        self.settings = Settings.excel_files

    def getData(self, book):
        """Excelデータ取得"""
        file_info = self.settings.get(book)
        if not file_info:
            raise ValueError(f"Unknown book key: {book}")
            
        return pd.read_excel(
            file_info['path'],
            sheet_name=file_info['sheet']
        )

    def forceRefreshWorkbook(self, filePath):
        """Excelブックの内部関数更新のため、強制的にExcelを開いてブックを更新します"""
        try:
            app = xw.App(visible=False)
            wb = app.books.open(filePath)
            wb.api.RefreshAll()
            app.api.CalculateUntilAsyncQueriesDone()
            wb.save()
            wb.close()
            app.quit()
        except Exception as e:
            print(f"エラーが発生しました: {e}")
          #  os.system("taskkill /f /im EXCEL.EXE")

# Chimeメッセージ送信クラス
class ChimeSvc:
    def __init__(self):
        self.hook_url = Settings.hook_url
        self.max_size = Settings.max_msg_size

    def sendMsg(self, content):
        """メッセージ送信"""
        content = self.fixMsgSize(content)
        message = {"Content": f"/md\n{content}"}
        
        try:
            response = requests.post(self.hook_url, json=message)
            if response.status_code == 200:
                print("メッセージが正常に投稿されました")
            else:
                print("メッセージの投稿に失敗しました")
        except Exception as e:
            print(f"エラーが発生しました: {e}")

    def fixMsgSize(self, content):
        """メッセージサイズ調整"""
        while len(content.encode('utf-8')) > self.max_size:
            lines = content.split('\n')
            content = '\n'.join(lines[:-1])
        return content

# サービスのインスタンス化
excelSvc = ExcelSvc()
chimeSvc = ChimeSvc()

def runData(book):
    """データ作成と送信"""
    try:
        print(f"データ処理を開始します: {book}")
        excelSvc.forceRefreshWorkbook(Settings.excel_files[book]['path'])
        
        data = excelSvc.getData(book)
        print("データを取得しました")
        
        # データの一部を送信
        if Settings.hook_url:
            # 送信する列が存在するか確認
            columns_to_send = ['出勤時間', 'あ']
            available_columns = [col for col in columns_to_send if col in data.columns]
            if available_columns:
                msg_content = data[available_columns].to_markdown(index=False)
                chimeSvc.sendMsg(msg_content)
            else:
                print(f"指定された列 {columns_to_send} がデータフレームに存在しないため、メッセージは送信されませんでした。")
        else:
            print("Chime Webhook URLが設定されていないため、データメッセージは送信されませんでした。")
        
        print("処理が完了しました")
    except KeyError as e:
        print(f"エラー発生: 指定された列が見つかりません - {e}")
    except Exception as e:
        print(f"エラー発生: {e}")

def remindSupply():
    """備品補充のリマインダー送信"""
    if Settings.hook_url:
        chimeSvc.sendMsg("今日は備品補充の日です。必要な備品を確認してください。")
    else:
        print("Chime Webhook URLが設定されていないため、リマインダーは送信されませんでした。")

def setSched():
    """スケジュール設定"""
    # 平日の11時にデータを送信
    schedule.every().monday.at("11:00").do(lambda: runData('book3'))
    schedule.every().tuesday.at("11:00").do(lambda: runData('book3'))
    schedule.every().wednesday.at("11:00").do(lambda: runData('book3'))
    schedule.every().thursday.at("11:00").do(lambda: runData('book3'))
    schedule.every().friday.at("11:00").do(lambda: runData('book3'))
    
    # 火曜と金曜に備品リマインダー
    schedule.every().tuesday.at("10:00").do(remindSupply)
    schedule.every().friday.at("10:00").do(remindSupply)
    
    print("スケジュールを設定しました")

def main():
    """メイン関数"""
    # タイトル表示
    titleStr = pyfiglet.figlet_format("KPI 自動通知システム", font="slant")
    print(titleStr)
    
    # 設定ファイルの内容を表示（デバッグ用、必要に応じて削除）
    print("--- 設定ファイル (config.ini) の内容 ---")
    for section in config.sections():
        print(f"[{section}]")
        for key, val in config.items(section):
            # HookURLは表示しないようにする
            if section == 'Chime' and key.lower() == 'hookurl':
                print(f"{key} = ***非表示***")
            else:
                print(f"{key} = {val}")
    print("------------------------------------")
    
    # ローディング表示
    showLoad()
    
    # Chime Webhook URLが設定されているか確認
    if not Settings.hook_url:
        print("\n警告: Chime Webhook URLがconfig.iniに設定されていません。")
        print("Chimeへの通知機能は無効になります。\n")
    
    # 起動時に一度実行
    runData('book3')
    
    # スケジュール設定
    setSched()
    
    print(f"システムが起動しました - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print("このウィンドウを閉じないでください。閉じると通知が停止します。")
    print("Ctrl+Cで終了できます。")
    
    # スケジュールを継続的に実行
    try:
        while True:
            schedule.run_pending()
            time.sleep(1)
    except KeyboardInterrupt:
        print("\nシステムを終了します...")
    except Exception as e:
        print(f"エラーが発生しました: {e}")

if __name__ == "__main__":
    main()